/**
 * Contas do Agrupamento — backend (Google Apps Script)
 *
 * Tudo é um movimento de dinheiro. Cada chefe tem um saldo por secção:
 *   saldo > 0  tem consigo dinheiro da secção que ainda não entregou ao fundo
 *   saldo < 0  adiantou dinheiro; o agrupamento deve-lhe esse valor
 * O fundo de cada secção (guardado pelo tesoureiro) só muda com movimentos
 * confirmados pelo tesoureiro.
 *
 * Configuração (Definições do projeto > Propriedades do script):
 *   PIN             PIN partilhado pelos chefes (criado pelo setup())
 *   PIN_TESOURARIA  PIN do tesoureiro (criado pelo setup())
 *   CONTAS_ID       folha de cálculo "Contas do Agrupamento" (criada pelo setup())
 *   PASTA_ID        pasta da Drive onde ficam as faturas (criada pelo setup())
 */

const SECCOES = ['Agrupamento', 'Lobitos', 'Exploradores', 'Pioneiros', 'Caminheiros'];
const TIPOS_DESPESA = [
  'Alimentação', 'Combustível', 'Material de desgaste', 'Papelaria', 'Aluguer',
  'Portagens', 'Transportes', 'Material de limpeza', 'Material de campo', 'Outros',
];
const TIPO = {
  ENTRADA: 'Entrada',
  DESPESA: 'Despesa',
  ENTREGA: 'Entrega ao fundo',
  REEMBOLSO: 'Reembolso',
  PASSAGEM: 'Passagem entre chefes',
};
const FORMAS = ['Dinheiro', 'Eletrónico'];
const ESTADO = {
  CONFIRMADO: 'Confirmado',
  POR_CONFIRMAR: 'Por confirmar',
  RECUSADO: 'Recusado',
  ANULADO: 'Anulado',
};
const TESOURARIA = 'Tesouraria';

const CABECALHO = [
  'ID', 'Registado em', 'Data', 'Ano letivo', 'Tipo', 'Secção', 'Pessoa', 'Valor (€)', 'Forma',
  'De / Para', 'Atividade', 'Tipo de despesa', 'Detalhe tipo (Outros)', 'Descrição', 'NIF do emissor',
  'Fatura', 'Efeito na pessoa', 'Efeito no fundo', 'Estado', 'Registado por', 'Tratado em', 'Ligado a',
];
const C = {};
CABECALHO.forEach((c, i) => { C[c] = i; });
const letra = nome => String.fromCharCode(65 + C[nome]);

const FOLHA_MOV = 'Movimentos';
const FOLHA_PESSOAS = 'Pessoas';
const FOLHA_FUNDOS = 'Fundos';
const FOLHA_SALDOS = 'Saldos dos chefes';

const NOME_PASTA_RAIZ = 'Faturas do Agrupamento';
const NOME_FOLHA = 'Contas do Agrupamento';
const FUSO = 'Europe/Lisbon';

const TIPOS_FICHEIRO = {
  'image/jpeg': 'jpg',
  'image/png': 'png',
  'image/webp': 'webp',
  'image/heic': 'heic',
  'image/heif': 'heif',
  'application/pdf': 'pdf',
};
const TAMANHO_MAX_FICHEIRO = 10 * 1024 * 1024;

const MAX_FALHAS_PIN = 10;          // tentativas erradas permitidas...
const JANELA_FALHAS_PIN = 10 * 60;  // ...em 10 minutos (segundos)
const MAX_MOVIMENTOS_PAINEL = 20;
const MAX_FEITOS_TESOURARIA = 20;

// ---------------------------------------------------------------------------
// Configuração inicial — correr a partir do editor (pode ser corrido várias vezes)
// ---------------------------------------------------------------------------

function setup() {
  const props = PropertiesService.getScriptProperties();

  let pasta;
  if (props.getProperty('PASTA_ID')) {
    pasta = DriveApp.getFolderById(props.getProperty('PASTA_ID'));
  } else {
    pasta = DriveApp.createFolder(NOME_PASTA_RAIZ);
    props.setProperty('PASTA_ID', pasta.getId());
  }

  let ss;
  if (props.getProperty('CONTAS_ID')) {
    ss = SpreadsheetApp.openById(props.getProperty('CONTAS_ID'));
  } else {
    // A folha da versão anterior (só faturas) fica guardada como arquivo
    const antiga = props.getProperty('SPREADSHEET_ID');
    if (antiga) {
      try {
        const f = DriveApp.getFileById(antiga);
        if (!/\(arquivo\)$/.test(f.getName())) f.setName(f.getName() + ' (arquivo)');
      } catch (err) {
        // Já não existe; não faz mal
      }
    }
    ss = SpreadsheetApp.create(NOME_FOLHA);
    ss.setSpreadsheetLocale('pt_PT');
    ss.setSpreadsheetTimeZone(FUSO);
    DriveApp.getFileById(ss.getId()).moveTo(pasta);
    props.setProperty('CONTAS_ID', ss.getId());
  }

  prepararFolhas(ss);

  ['PIN', 'PIN_TESOURARIA'].forEach(chave => {
    if (!props.getProperty(chave)) props.setProperty(chave, String(Math.floor(100000 + Math.random() * 900000)));
  });

  console.log('Pasta na Drive: ' + pasta.getUrl());
  console.log('Folha de cálculo: ' + ss.getUrl());
  console.log('PIN dos chefes: ' + props.getProperty('PIN'));
  console.log('PIN da tesouraria: ' + props.getProperty('PIN_TESOURARIA'));
  console.log('Próximo passo: preencher as abas "Pessoas" e "Fundos" (saldo inicial) na folha de cálculo.');
}

function prepararFolhas(ss) {
  const cabecalhoVerde = r => r.setFontWeight('bold').setBackground('#e8f0e8');

  let mov = ss.getSheetByName(FOLHA_MOV);
  if (!mov) {
    mov = ss.insertSheet(FOLHA_MOV);
    cabecalhoVerde(mov.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]));
    mov.setFrozenRows(1);
    const col = nome => mov.getRange(letra(nome) + '2:' + letra(nome));
    col('Registado em').setNumberFormat('dd/mm/yyyy hh:mm');
    col('Tratado em').setNumberFormat('dd/mm/yyyy hh:mm');
    col('Data').setNumberFormat('dd/mm/yyyy');
    ['Valor (€)', 'Efeito na pessoa', 'Efeito no fundo'].forEach(n => col(n).setNumberFormat('#,##0.00 "€"'));
    col('NIF do emissor').setNumberFormat('@');
    col('Ano letivo').setNumberFormat('@');
    mov.hideColumns(C['ID'] + 1);
    mov.hideColumns(C['Ligado a'] + 1);
  }

  let pessoas = ss.getSheetByName(FOLHA_PESSOAS);
  if (!pessoas) {
    pessoas = ss.insertSheet(FOLHA_PESSOAS);
    cabecalhoVerde(pessoas.getRange('A1:C1').setValues([['Nome', 'Secção', 'Ativo']]));
    pessoas.setFrozenRows(1);
    pessoas.getRange('A1').setNote(
      'Escreve aqui os chefes (um por linha), a secção de cada um e "Sim" em Ativo.\n' +
      'Quem sair do agrupamento: muda Ativo para "Não" (não apagues a linha).'
    );
    const lista = l => SpreadsheetApp.newDataValidation().requireValueInList(l, true).build();
    pessoas.getRange('B2:B').setDataValidation(lista(SECCOES));
    pessoas.getRange('C2:C').setDataValidation(lista(['Sim', 'Não']));
    pessoas.setColumnWidth(1, 200);
  }

  let saldos = ss.getSheetByName(FOLHA_SALDOS);
  if (!saldos) {
    saldos = ss.insertSheet(FOLHA_SALDOS);
    const L = n => letra(n);
    saldos.getRange('A1').setFormula(
      '=QUERY(' + FOLHA_MOV + '!A:' + L('Ligado a') + ',"select ' + L('Pessoa') + ', ' + L('Secção') +
      ', sum(' + L('Efeito na pessoa') + ") where " + L('Estado') + " = 'Confirmado' and " + L('Pessoa') +
      " <> '' group by " + L('Pessoa') + ', ' + L('Secção') +
      " label " + L('Pessoa') + " 'Pessoa', " + L('Secção') + " 'Secção', sum(" + L('Efeito na pessoa') +
      ") 'Saldo (€)'\",1)"
    );
    saldos.getRange('A1:C1').setFontWeight('bold');
    saldos.getRange('C2:C').setNumberFormat('#,##0.00 "€"');
    saldos.getRange('E1').setValue('Positivo: o chefe tem dinheiro da secção consigo. Negativo: o agrupamento deve-lhe.');
  }

  let fundos = ss.getSheetByName(FOLHA_FUNDOS);
  if (!fundos) {
    fundos = ss.insertSheet(FOLHA_FUNDOS, 0);
    const cab = ['Secção', 'Saldo inicial (€)', 'Movimentos (€)', 'No fundo (€)', 'Com os chefes (€)',
      'Em dívida aos chefes (€)', 'Total real (€)'];
    cabecalhoVerde(fundos.getRange(1, 1, 1, cab.length).setValues([cab]));
    const m = FOLHA_MOV + '!';
    const s = "'" + FOLHA_SALDOS + "'!";
    const linhas = SECCOES.map((sec, i) => {
      const r = i + 2;
      return [
        sec,
        '',
        '=SUMIFS(' + m + letra('Efeito no fundo') + ':' + letra('Efeito no fundo') + ',' + m + letra('Secção') + ':' +
          letra('Secção') + ',A' + r + ',' + m + letra('Estado') + ':' + letra('Estado') + ',"Confirmado")',
        '=B' + r + '+C' + r,
        '=SUMIFS(' + s + 'C:C,' + s + 'B:B,A' + r + ',' + s + 'C:C,">0")',
        '=-SUMIFS(' + s + 'C:C,' + s + 'B:B,A' + r + ',' + s + 'C:C,"<0")',
        '=D' + r + '+E' + r + '-F' + r,
      ];
    });
    const ultima = SECCOES.length + 1;
    const total = ['Total'];
    for (let c = 1; c < cab.length; c++) {
      const L = String.fromCharCode(65 + c);
      total.push('=SUM(' + L + '2:' + L + ultima + ')');
    }
    linhas.push(total);
    fundos.getRange(2, 1, linhas.length, cab.length).setValues(linhas);
    fundos.getRange(2, 2, linhas.length, cab.length - 1).setNumberFormat('#,##0.00 "€"');
    fundos.getRange(ultima + 1, 1, 1, cab.length).setFontWeight('bold');
    fundos.getRange(2, 2, SECCOES.length, 1).setBackground('#fff8dc');
    fundos.getRange('B1').setNote('Escreve aqui quanto dinheiro cada fundo tinha quando se começou a usar a app.');
    fundos.setColumnWidths(1, cab.length, 150);
  }

  ss.getSheets()
    .filter(f => /^(Folha|Sheet)\d+$/.test(f.getName()) && f.getLastRow() === 0)
    .forEach(f => ss.deleteSheet(f));
}

// ---------------------------------------------------------------------------
// Endpoints da app web
// ---------------------------------------------------------------------------

function doGet() {
  return ContentService.createTextOutput('A API das contas está a funcionar.');
}

function doPost(e) {
  let pedido;
  try {
    pedido = JSON.parse(e.postData.contents);
  } catch (err) {
    return resposta({ ok: false, erro: 'Pedido inválido.' });
  }

  try {
    if (pinBloqueado()) {
      return resposta({ ok: false, erro: 'Demasiadas tentativas com PIN errado. Tente daqui a 10 minutos.' });
    }
    const acao = String(pedido.acao || '');
    const tesouraria = acao.indexOf('tesouraria.') === 0;
    if (!pinCorreto(pedido.pin, tesouraria ? 'PIN_TESOURARIA' : 'PIN')) {
      registarFalhaPin();
      return resposta({ ok: false, codigo: 'PIN', erro: 'PIN errado.' });
    }

    const acoes = {
      'iniciar': iniciar,
      'painel': painel,
      'registar': registar,
      'anular': anular,
      'atividade': relatorioAtividade,
      'tesouraria.verificarPin': () => ({ ok: true }),
      'tesouraria.listar': tesourariaListar,
      'tesouraria.foto': tesourariaFoto,
      'tesouraria.confirmar': tesourariaConfirmar,
      'tesouraria.acertar': tesourariaAcertar,
      'tesouraria.desfazer': tesourariaDesfazer,
    };
    if (!acoes[acao]) return resposta({ ok: false, erro: 'Ação desconhecida.' });
    return resposta(acoes[acao](pedido));
  } catch (err) {
    if (err instanceof ErroValidacao) return resposta({ ok: false, erro: err.message });
    console.error(err);
    return resposta({ ok: false, erro: 'Erro no servidor. Tente novamente.' });
  }
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ErroValidacao(mensagem) {
  this.message = mensagem;
}

// ---------------------------------------------------------------------------
// PIN
// ---------------------------------------------------------------------------

function pinCorreto(pin, chave) {
  const esperado = PropertiesService.getScriptProperties().getProperty(chave);
  return !!esperado && String(pin || '').trim() === esperado.trim();
}

function pinBloqueado() {
  return Number(CacheService.getScriptCache().get('falhasPin') || 0) >= MAX_FALHAS_PIN;
}

function registarFalhaPin() {
  const cache = CacheService.getScriptCache();
  const falhas = Number(cache.get('falhasPin') || 0) + 1;
  cache.put('falhasPin', String(falhas), JANELA_FALHAS_PIN);
}

// ---------------------------------------------------------------------------
// Leitura de dados
// ---------------------------------------------------------------------------

function abrirContas() {
  const id = PropertiesService.getScriptProperties().getProperty('CONTAS_ID');
  if (!id) throw new ErroValidacao('A aplicação ainda não foi configurada (falta correr o setup).');
  return SpreadsheetApp.openById(id);
}

function comLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function lerPessoas(ss) {
  const folha = ss.getSheetByName(FOLHA_PESSOAS);
  const n = folha.getLastRow() - 1;
  if (n < 1) return [];
  return folha.getRange(2, 1, n, 3).getValues()
    .filter(l => String(l[0]).trim() && SECCOES.indexOf(String(l[1]).trim()) !== -1 && String(l[2]).trim() !== 'Não')
    .map(l => ({ nome: String(l[0]).trim(), seccao: String(l[1]).trim() }));
}

function lerMovimentos(ss) {
  const folha = ss.getSheetByName(FOLHA_MOV);
  const n = folha.getLastRow() - 1;
  if (n < 1) return [];
  const valores = folha.getRange(2, 1, n, CABECALHO.length).getValues();
  const links = folha.getRange(2, C['Fatura'] + 1, n, 1).getRichTextValues();
  const dataTexto = v => (v instanceof Date ? Utilities.formatDate(v, FUSO, 'dd/MM/yyyy') : '');
  const ms = v => (v instanceof Date ? v.getTime() : 0);
  return valores.map((l, i) => {
    const url = links[i][0] ? links[i][0].getLinkUrl() || '' : '';
    const m = /\/d\/([A-Za-z0-9_-]+)/.exec(url);
    return {
      linha: i + 2,
      id: String(l[C['ID']]),
      registadoEm: ms(l[C['Registado em']]),
      data: dataTexto(l[C['Data']]),
      dataMs: ms(l[C['Data']]),
      ano: String(l[C['Ano letivo']]),
      tipo: String(l[C['Tipo']]),
      seccao: String(l[C['Secção']]),
      pessoa: String(l[C['Pessoa']]),
      valor: Number(l[C['Valor (€)']]) || 0,
      forma: String(l[C['Forma']]),
      dePara: String(l[C['De / Para']]),
      atividade: String(l[C['Atividade']]),
      tipoDespesa: String(l[C['Tipo de despesa']]),
      tipoOutro: String(l[C['Detalhe tipo (Outros)']]),
      descricao: String(l[C['Descrição']]),
      url,
      fileId: m ? m[1] : '',
      efeitoPessoa: Number(l[C['Efeito na pessoa']]) || 0,
      efeitoFundo: Number(l[C['Efeito no fundo']]) || 0,
      estado: String(l[C['Estado']]),
      registadoPor: String(l[C['Registado por']]),
      tratadoEm: ms(l[C['Tratado em']]),
      ligadoA: String(l[C['Ligado a']]),
    };
  }).filter(m => m.id);
}

const chave = (pessoa, seccao) => pessoa + '|' + seccao;
const arred = v => Math.round(v * 100) / 100;

/** Saldo confirmado de cada pessoa por secção, e entregas à espera de confirmação. */
function calcularSaldos(movs) {
  const saldos = {};
  const obter = (pessoa, seccao) => {
    const k = chave(pessoa, seccao);
    if (!saldos[k]) saldos[k] = { pessoa, seccao, saldo: 0, entregaPendente: 0 };
    return saldos[k];
  };
  movs.forEach(m => {
    if (!m.pessoa) return;
    if (m.estado === ESTADO.CONFIRMADO) obter(m.pessoa, m.seccao).saldo += m.efeitoPessoa;
    if (m.estado === ESTADO.POR_CONFIRMAR && m.tipo === TIPO.ENTREGA) obter(m.pessoa, m.seccao).entregaPendente += m.valor;
  });
  Object.keys(saldos).forEach(k => {
    saldos[k].saldo = arred(saldos[k].saldo);
    saldos[k].entregaPendente = arred(saldos[k].entregaPendente);
  });
  return saldos;
}

function obterPessoa(ss, nome) {
  const p = lerPessoas(ss).find(x => x.nome === String(nome || '').trim());
  if (!p) throw new ErroValidacao('Pessoa desconhecida. Escolhe outra vez quem és.');
  return p;
}

/** Um chefe trabalha com a sua secção e com o Agrupamento. */
function seccoesDe(pessoa) {
  return pessoa.seccao === 'Agrupamento' ? ['Agrupamento'] : [pessoa.seccao, 'Agrupamento'];
}

// ---------------------------------------------------------------------------
// App dos chefes
// ---------------------------------------------------------------------------

function iniciar() {
  const ss = abrirContas();
  return { ok: true, pessoas: lerPessoas(ss), tiposDespesa: TIPOS_DESPESA };
}

function painel(p) {
  const ss = abrirContas();
  const pessoa = obterPessoa(ss, p.pessoa);
  const movs = lerMovimentos(ss);
  const saldos = calcularSaldos(movs);
  const minhas = seccoesDe(pessoa);

  // A secção do chefe primeiro, depois o Agrupamento e outras onde ainda tenha saldo
  const resumo = minhas.concat(SECCOES.filter(s => minhas.indexOf(s) === -1))
    .filter(s => minhas.indexOf(s) !== -1 || (saldos[chave(pessoa.nome, s)] || {}).saldo)
    .map(s => saldos[chave(pessoa.nome, s)] || { pessoa: pessoa.nome, seccao: s, saldo: 0, entregaPendente: 0 });

  const meus = movs
    .filter(m => (m.pessoa === pessoa.nome || m.registadoPor === pessoa.nome) && m.estado !== ESTADO.ANULADO)
    // Numa passagem, quem a registou só vê a sua metade
    .filter(m => !(m.ligadoA && m.pessoa !== pessoa.nome))
    .sort((a, b) => b.registadoEm - a.registadoEm)
    .slice(0, MAX_MOVIMENTOS_PAINEL)
    .map(m => ({
      id: m.id,
      data: m.data,
      tipo: m.tipo,
      seccao: m.seccao,
      valor: m.valor,
      efeito: m.pessoa === pessoa.nome ? m.efeitoPessoa : 0,
      pessoa: m.pessoa,
      dePara: m.dePara,
      atividade: m.atividade,
      descricao: m.descricao || (m.tipo === TIPO.DESPESA ? textoTipoDespesa(m) : ''),
      estado: m.estado,
      registadoPor: m.registadoPor,
      podeAnular: podeAnular(m, pessoa.nome),
    }));

  // Atividades já usadas (para sugerir o mesmo nome e para os relatórios)
  const atividades = {};
  movs
    .filter(m => m.atividade && minhas.indexOf(m.seccao) !== -1 && m.estado !== ESTADO.ANULADO)
    .sort((a, b) => b.registadoEm - a.registadoEm)
    .forEach(m => {
      atividades[m.seccao] = atividades[m.seccao] || [];
      if (atividades[m.seccao].indexOf(m.atividade) === -1) atividades[m.seccao].push(m.atividade);
    });

  return { ok: true, pessoa, seccoes: minhas, saldos: resumo, movimentos: meus, atividades };
}

function textoTipoDespesa(m) {
  return m.tipoDespesa === 'Outros' && m.tipoOutro ? m.tipoOutro : m.tipoDespesa;
}

function podeAnular(m, nome) {
  if (m.registadoPor !== nome) return false;
  if (m.estado === ESTADO.ANULADO || m.estado === ESTADO.RECUSADO) return false;
  // Depois de o tesoureiro confirmar, só ele pode desfazer
  if (m.efeitoFundo !== 0 && m.estado === ESTADO.CONFIRMADO) return false;
  return true;
}

function registar(p) {
  const ss = abrirContas();
  const pessoas = lerPessoas(ss);
  const eu = pessoas.find(x => x.nome === String(p.pessoa || '').trim());
  if (!eu) throw new ErroValidacao('Pessoa desconhecida. Escolhe outra vez quem és.');

  const str = (v, max) => String(v == null ? '' : v).trim().slice(0, max);
  const id = str(p.id, 64);
  if (!/^[A-Za-z0-9-]{8,64}$/.test(id)) throw new ErroValidacao('Identificador inválido.');

  const tipo = str(p.tipo, 40);
  if ([TIPO.ENTRADA, TIPO.DESPESA, TIPO.ENTREGA, TIPO.PASSAGEM].indexOf(tipo) === -1) {
    throw new ErroValidacao('Tipo de movimento inválido.');
  }
  const seccao = str(p.seccao, 40);
  if (seccoesDe(eu).indexOf(seccao) === -1) throw new ErroValidacao('Secção inválida.');
  const valor = arred(Number(p.valor));
  if (!isFinite(valor) || valor <= 0 || valor > 100000) throw new ErroValidacao('Valor inválido.');
  const forma = str(p.forma, 20);
  if (FORMAS.indexOf(forma) === -1) throw new ErroValidacao('Indica se foi em dinheiro ou eletrónico.');
  const data = lerData(str(p.data, 10));
  const atividade = str(p.atividade, 200);
  const descricao = str(p.descricao, 1000);

  const base = {
    id, data, tipo, seccao, valor, forma, atividade, descricao,
    pessoa: eu.nome, dePara: '', tipoDespesa: '', tipoOutro: '', nif: '', ficheiro: null,
    efeitoPessoa: 0, efeitoFundo: 0, estado: ESTADO.CONFIRMADO, registadoPor: eu.nome, ligadoA: '',
  };
  const linhas = [];

  if (tipo === TIPO.ENTRADA) {
    const deQuem = str(p.deQuem, 200);
    if (!deQuem) throw new ErroValidacao('Indica de quem recebeste o dinheiro.');
    base.dePara = 'De: ' + deQuem;
    const ficouCom = str(p.ficouCom, 100);
    if (ficouCom === '__agrupamento__') {
      base.pessoa = '';
      base.efeitoFundo = valor;
      base.estado = ESTADO.POR_CONFIRMAR;
    } else {
      const dono = pessoas.find(x => x.nome === ficouCom);
      if (!dono) throw new ErroValidacao('Indica com quem ficou o dinheiro.');
      base.pessoa = dono.nome;
      base.efeitoPessoa = valor;
    }
    linhas.push(base);
  }

  if (tipo === TIPO.DESPESA) {
    base.tipoDespesa = str(p.tipoDespesa, 50);
    if (TIPOS_DESPESA.indexOf(base.tipoDespesa) === -1) throw new ErroValidacao('Escolhe o tipo de despesa.');
    base.tipoOutro = base.tipoDespesa === 'Outros' ? str(p.tipoOutro, 200) : '';
    if (base.tipoDespesa === 'Outros' && !base.tipoOutro) throw new ErroValidacao('Especifica o tipo de despesa.');
    base.nif = str(p.nif, 20).replace(/\s/g, '');
    if (base.nif && !nifValido(base.nif)) throw new ErroValidacao('NIF do emissor inválido.');
    base.ficheiro = lerFicheiro(p.ficheiro);
    if (str(p.pagoCom, 20) === 'agrupamento') {
      base.pessoa = '';
      base.dePara = 'Pago com cartão/conta do agrupamento';
      base.efeitoFundo = -valor;
      base.estado = ESTADO.POR_CONFIRMAR;
    } else {
      base.efeitoPessoa = -valor;
    }
    linhas.push(base);
  }

  if (tipo === TIPO.ENTREGA) {
    base.dePara = 'Para: tesoureiro';
    base.efeitoPessoa = -valor;
    base.efeitoFundo = valor;
    base.estado = ESTADO.POR_CONFIRMAR;
    linhas.push(base);
  }

  if (tipo === TIPO.PASSAGEM) {
    const para = pessoas.find(x => x.nome === str(p.para, 100));
    if (!para || para.nome === eu.nome) throw new ErroValidacao('Escolhe a quem passaste o dinheiro.');
    const idB = id.slice(0, 60) + '-b';
    linhas.push(Object.assign({}, base, { dePara: 'Para: ' + para.nome, efeitoPessoa: -valor, ligadoA: idB }));
    linhas.push(Object.assign({}, base, {
      id: idB, pessoa: para.nome, dePara: 'De: ' + eu.nome, efeitoPessoa: valor, ligadoA: id,
    }));
  }

  return comLock(() => {
    const folha = ss.getSheetByName(FOLHA_MOV);
    // A app repete o pedido se a rede falhar; o ID evita linhas duplicadas.
    if (idExiste(folha, id)) return { ok: true, duplicado: true };

    linhas.forEach(l => {
      if (l.ficheiro) l.url = guardarFatura(l);
      escreverMovimento(folha, l);
    });
    SpreadsheetApp.flush();
    return { ok: true, estado: linhas[0].estado };
  });
}

function lerData(texto) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(texto);
  if (!m) throw new ErroValidacao('Data inválida.');
  const data = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const amanha = new Date(Date.now() + 24 * 3600 * 1000);
  if (data.getMonth() !== Number(m[2]) - 1 || data > amanha || data.getFullYear() < 2000) {
    throw new ErroValidacao('Data inválida.');
  }
  return data;
}

function lerFicheiro(f) {
  f = f || {};
  const mimeType = String(f.mimeType || '').toLowerCase();
  if (!TIPOS_FICHEIRO[mimeType]) throw new ErroValidacao('Falta a foto da fatura (ou o tipo de ficheiro não é suportado).');
  let bytes;
  try {
    bytes = Utilities.base64Decode(String(f.base64 || ''));
  } catch (err) {
    throw new ErroValidacao('Ficheiro corrompido.');
  }
  if (!bytes.length) throw new ErroValidacao('Falta a foto da fatura.');
  if (bytes.length > TAMANHO_MAX_FICHEIRO) throw new ErroValidacao('Ficheiro demasiado grande (máx. 10 MB).');
  return { mimeType, bytes };
}

function guardarFatura(l) {
  const raiz = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PASTA_ID'));
  const pasta = obterPastaAno(raiz, anoLetivo(l.data));
  const nome = [
    Utilities.formatDate(l.data, FUSO, 'yyyy-MM-dd'),
    l.seccao,
    l.atividade || l.tipoDespesa,
    l.valor.toFixed(2).replace('.', ',') + 'EUR',
    l.id.slice(0, 8),
  ].join('_').replace(/[\\/:*?"<>|\n\r]+/g, '-').slice(0, 150) + '.' + TIPOS_FICHEIRO[l.ficheiro.mimeType];
  const ficheiro = pasta.createFile(Utilities.newBlob(l.ficheiro.bytes, l.ficheiro.mimeType, nome));
  ficheiro.setDescription(l.registadoPor + ' — ' + (l.atividade || '') + ' — ' + l.tipoDespesa);
  return ficheiro.getUrl();
}

function escreverMovimento(folha, l) {
  const linha = new Array(CABECALHO.length).fill('');
  const agora = new Date();
  linha[C['ID']] = texto(l.id);
  linha[C['Registado em']] = agora;
  linha[C['Data']] = l.data;
  linha[C['Ano letivo']] = anoLetivo(l.data);
  linha[C['Tipo']] = l.tipo;
  linha[C['Secção']] = l.seccao;
  linha[C['Pessoa']] = texto(l.pessoa);
  linha[C['Valor (€)']] = l.valor;
  linha[C['Forma']] = l.forma;
  linha[C['De / Para']] = texto(l.dePara);
  linha[C['Atividade']] = texto(l.atividade);
  linha[C['Tipo de despesa']] = l.tipoDespesa;
  linha[C['Detalhe tipo (Outros)']] = texto(l.tipoOutro);
  linha[C['Descrição']] = texto(l.descricao);
  linha[C['NIF do emissor']] = texto(l.nif);
  linha[C['Efeito na pessoa']] = l.efeitoPessoa;
  linha[C['Efeito no fundo']] = l.efeitoFundo;
  linha[C['Estado']] = l.estado;
  linha[C['Registado por']] = texto(l.registadoPor);
  linha[C['Tratado em']] = l.tratadoEm || '';
  linha[C['Ligado a']] = texto(l.ligadoA);
  folha.appendRow(linha);
  if (l.url) {
    folha.getRange(folha.getLastRow(), C['Fatura'] + 1).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText('Ver fatura').setLinkUrl(l.url).build()
    );
  }
}

function anular(p) {
  const ss = abrirContas();
  const pessoa = obterPessoa(ss, p.pessoa);
  return comLock(() => {
    const movs = lerMovimentos(ss);
    const m = movs.find(x => x.id === String(p.id || ''));
    if (!m) throw new ErroValidacao('Movimento não encontrado.');
    if (!podeAnular(m, pessoa.nome)) {
      throw new ErroValidacao('Este movimento já não pode ser anulado aqui. Fala com o tesoureiro.');
    }
    const folha = ss.getSheetByName(FOLHA_MOV);
    [m].concat(movs.filter(x => m.ligadoA && x.id === m.ligadoA)).forEach(x => {
      folha.getRange(x.linha, C['Estado'] + 1).setValue(ESTADO.ANULADO);
    });
    SpreadsheetApp.flush();
    return { ok: true };
  });
}

/** Contas de uma atividade: o que entrou (e como) e o que saiu. */
function relatorioAtividade(p) {
  const ss = abrirContas();
  const pessoa = obterPessoa(ss, p.pessoa);
  const seccao = String(p.seccao || '');
  if (seccoesDe(pessoa).indexOf(seccao) === -1) throw new ErroValidacao('Secção inválida.');
  const nome = String(p.atividade || '').trim().toLowerCase();
  if (!nome) throw new ErroValidacao('Escolhe a atividade.');

  const movs = lerMovimentos(ss).filter(m =>
    m.seccao === seccao && m.atividade.trim().toLowerCase() === nome &&
    (m.tipo === TIPO.ENTRADA || m.tipo === TIPO.DESPESA) &&
    m.estado !== ESTADO.ANULADO && m.estado !== ESTADO.RECUSADO
  ).sort((a, b) => a.dataMs - b.dataMs);

  const item = m => ({
    data: m.data,
    valor: m.valor,
    forma: m.forma,
    quem: m.tipo === TIPO.ENTRADA ? m.dePara.replace(/^De: /, '') : (m.pessoa || 'Fundo do agrupamento'),
    ficouCom: m.tipo === TIPO.ENTRADA ? (m.pessoa || 'Conta do agrupamento') : '',
    tipoDespesa: m.tipo === TIPO.DESPESA ? textoTipoDespesa(m) : '',
    descricao: m.descricao,
    porConfirmar: m.estado === ESTADO.POR_CONFIRMAR,
  });
  const entradas = movs.filter(m => m.tipo === TIPO.ENTRADA).map(item);
  const despesas = movs.filter(m => m.tipo === TIPO.DESPESA).map(item);
  const soma = (l, f) => arred(l.filter(f || (() => true)).reduce((s, x) => s + x.valor, 0));

  return {
    ok: true,
    seccao,
    atividade: movs.length ? movs[0].atividade : p.atividade,
    entradas,
    despesas,
    totalEntradas: soma(entradas),
    entradasDinheiro: soma(entradas, x => x.forma === 'Dinheiro'),
    entradasEletronico: soma(entradas, x => x.forma === 'Eletrónico'),
    totalDespesas: soma(despesas),
    resultado: arred(soma(entradas) - soma(despesas)),
  };
}

// ---------------------------------------------------------------------------
// App da tesouraria
// ---------------------------------------------------------------------------

function tesourariaListar() {
  const ss = abrirContas();
  const movs = lerMovimentos(ss);
  const saldos = calcularSaldos(movs);

  const paraConfirmar = movs
    .filter(m => m.estado === ESTADO.POR_CONFIRMAR)
    .sort((a, b) => a.registadoEm - b.registadoEm)
    .map(m => resumoMovimento(m));

  // O que falta acertar com cada chefe (descontando entregas que ainda faltam confirmar)
  const acertos = Object.keys(saldos)
    .map(k => saldos[k])
    .map(s => ({ pessoa: s.pessoa, seccao: s.seccao, valor: arred(s.saldo - s.entregaPendente) }))
    .filter(s => Math.abs(s.valor) >= 0.01)
    .sort((a, b) => SECCOES.indexOf(a.seccao) - SECCOES.indexOf(b.seccao) || a.pessoa.localeCompare(b.pessoa));

  const feitos = movs
    .filter(m => m.tratadoEm && m.estado !== ESTADO.ANULADO)
    .sort((a, b) => b.tratadoEm - a.tratadoEm)
    .slice(0, MAX_FEITOS_TESOURARIA)
    .map(m => resumoMovimento(m));

  return { ok: true, paraConfirmar, acertos, resumo: resumoFundos(ss, movs, saldos), feitos };
}

function resumoMovimento(m) {
  return {
    id: m.id, tipo: m.tipo, seccao: m.seccao, pessoa: m.pessoa, registadoPor: m.registadoPor,
    valor: m.valor, forma: m.forma, dePara: m.dePara, atividade: m.atividade,
    tipoDespesa: textoTipoDespesa(m), descricao: m.descricao, data: m.data,
    fileId: m.fileId, url: m.url, estado: m.estado,
  };
}

function resumoFundos(ss, movs, saldos) {
  const iniciais = {};
  ss.getSheetByName(FOLHA_FUNDOS).getRange(2, 1, SECCOES.length, 2).getValues()
    .forEach(l => { iniciais[l[0]] = Number(l[1]) || 0; });

  return SECCOES.map(seccao => {
    const noFundo = arred((iniciais[seccao] || 0) + movs
      .filter(m => m.seccao === seccao && m.estado === ESTADO.CONFIRMADO)
      .reduce((s, m) => s + m.efeitoFundo, 0));
    let comChefes = 0;
    let divida = 0;
    Object.keys(saldos).forEach(k => {
      const s = saldos[k];
      if (s.seccao !== seccao) return;
      if (s.saldo > 0) comChefes += s.saldo;
      else divida -= s.saldo;
    });
    comChefes = arred(comChefes);
    divida = arred(divida);
    return { seccao, noFundo, comChefes, divida, total: arred(noFundo + comChefes - divida) };
  });
}

/** Devolve a imagem da fatura (ou uma miniatura, no caso de PDF/HEIC). */
function tesourariaFoto(p) {
  const fileId = String(p.fileId || '');
  if (!/^[A-Za-z0-9_-]{10,}$/.test(fileId)) throw new ErroValidacao('Fatura não encontrada.');
  let ficheiro;
  try {
    ficheiro = DriveApp.getFileById(fileId);
  } catch (err) {
    throw new ErroValidacao('A fotografia desta fatura já não existe na Drive.');
  }
  if (!estaNaPastaDasFaturas(ficheiro)) throw new ErroValidacao('Fatura não encontrada.');

  const mime = ficheiro.getMimeType();
  const resultado = { ok: true, pdf: mime === 'application/pdf', url: ficheiro.getUrl() };
  if (['image/jpeg', 'image/png', 'image/webp'].indexOf(mime) !== -1) {
    resultado.mimeType = mime;
    resultado.base64 = Utilities.base64Encode(ficheiro.getBlob().getBytes());
  } else {
    try {
      const mini = ficheiro.getThumbnail();
      if (mini) {
        resultado.mimeType = mini.getContentType();
        resultado.base64 = Utilities.base64Encode(mini.getBytes());
      }
    } catch (err) {
      // Sem miniatura: a app mostra só o link para abrir o ficheiro
    }
  }
  return resultado;
}

function estaNaPastaDasFaturas(ficheiro) {
  const raiz = PropertiesService.getScriptProperties().getProperty('PASTA_ID');
  const pais = ficheiro.getParents();
  while (pais.hasNext()) {
    const pai = pais.next();
    if (pai.getId() === raiz) return true;
    const avos = pai.getParents();
    while (avos.hasNext()) if (avos.next().getId() === raiz) return true;
  }
  return false;
}

function tesourariaConfirmar(p) {
  const ss = abrirContas();
  return comLock(() => {
    const m = lerMovimentos(ss).find(x => x.id === String(p.id || ''));
    if (!m) throw new ErroValidacao('Movimento não encontrado.');
    if (m.estado !== ESTADO.POR_CONFIRMAR) throw new ErroValidacao('Este movimento já foi tratado.');
    const folha = ss.getSheetByName(FOLHA_MOV);
    folha.getRange(m.linha, C['Estado'] + 1).setValue(p.aceitar === true ? ESTADO.CONFIRMADO : ESTADO.RECUSADO);
    folha.getRange(m.linha, C['Tratado em'] + 1).setValue(new Date());
    SpreadsheetApp.flush();
    return { ok: true };
  });
}

/** Acerta o saldo de um chefe: reembolsa-o (saldo negativo) ou recebe o dinheiro dele (saldo positivo). */
function tesourariaAcertar(p) {
  const ss = abrirContas();
  const forma = String(p.forma || '');
  if (FORMAS.indexOf(forma) === -1) throw new ErroValidacao('Indique se foi em dinheiro ou por transferência.');
  const pessoa = String(p.pessoa || '');
  const seccao = String(p.seccao || '');
  const id = String(p.id || '');
  if (!/^[A-Za-z0-9-]{8,64}$/.test(id)) throw new ErroValidacao('Identificador inválido.');

  return comLock(() => {
    const folha = ss.getSheetByName(FOLHA_MOV);
    if (idExiste(folha, id)) return { ok: true, duplicado: true };

    const s = calcularSaldos(lerMovimentos(ss))[chave(pessoa, seccao)];
    const valor = s ? arred(s.saldo - s.entregaPendente) : 0;
    if (Math.abs(valor) < 0.01) throw new ErroValidacao('As contas com esta pessoa já estão acertadas.');
    if (Math.abs(valor - Number(p.valor)) >= 0.01) {
      throw new ErroValidacao('Os valores mudaram entretanto. Carregue em "Atualizar" e veja outra vez.');
    }

    const hoje = new Date();
    escreverMovimento(folha, {
      id,
      data: new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()),
      tipo: valor < 0 ? TIPO.REEMBOLSO : TIPO.ENTREGA,
      seccao,
      pessoa,
      valor: Math.abs(valor),
      forma,
      dePara: valor < 0 ? 'Pago pelo tesoureiro' : 'Recebido pelo tesoureiro',
      atividade: '', tipoDespesa: '', tipoOutro: '', descricao: '', nif: '',
      efeitoPessoa: -valor,
      efeitoFundo: valor,
      estado: ESTADO.CONFIRMADO,
      registadoPor: TESOURARIA,
      tratadoEm: hoje,
      ligadoA: '',
    });
    SpreadsheetApp.flush();
    return { ok: true };
  });
}

/** Desfaz uma ação do tesoureiro: volta a pôr por confirmar, ou anula o acerto que ele registou. */
function tesourariaDesfazer(p) {
  const ss = abrirContas();
  return comLock(() => {
    const m = lerMovimentos(ss).find(x => x.id === String(p.id || ''));
    if (!m || !m.tratadoEm) throw new ErroValidacao('Movimento não encontrado.');
    const folha = ss.getSheetByName(FOLHA_MOV);
    if (m.registadoPor === TESOURARIA) {
      folha.getRange(m.linha, C['Estado'] + 1).setValue(ESTADO.ANULADO);
    } else {
      folha.getRange(m.linha, C['Estado'] + 1).setValue(ESTADO.POR_CONFIRMAR);
      folha.getRange(m.linha, C['Tratado em'] + 1).setValue('');
    }
    SpreadsheetApp.flush();
    return { ok: true };
  });
}

// ---------------------------------------------------------------------------
// Utilitários
// ---------------------------------------------------------------------------

/** Impede que texto começado por "=", "+", etc. seja interpretado como fórmula. */
function texto(v) {
  return v ? "'" + v : '';
}

function nifValido(nif) {
  if (!/^\d{9}$/.test(nif)) return false;
  let soma = 0;
  for (let i = 0; i < 8; i++) soma += Number(nif[i]) * (9 - i);
  const resto = soma % 11;
  return (resto < 2 ? 0 : 11 - resto) === Number(nif[8]);
}

/** Ano letivo: de 1 de setembro a 31 de agosto. */
function anoLetivo(data) {
  const inicio = data.getMonth() >= 8 ? data.getFullYear() : data.getFullYear() - 1;
  return inicio + '-' + String((inicio + 1) % 100).padStart(2, '0');
}

function obterPastaAno(raiz, ano) {
  const it = raiz.getFoldersByName(ano);
  return it.hasNext() ? it.next() : raiz.createFolder(ano);
}

function idExiste(folha, id) {
  const n = folha.getLastRow() - 1;
  if (n < 1) return false;
  return folha.getRange(2, 1, n, 1).getValues().some(l => String(l[0]) === id);
}
