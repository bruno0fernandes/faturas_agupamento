/**
 * Faturas do Agrupamento — backend (Google Apps Script)
 *
 * Recebe as submissões da app, guarda a foto/PDF na Google Drive
 * e acrescenta uma linha na Google Sheet do ano letivo respetivo.
 * Serve também a app da tesouraria, que marca as faturas como reembolsadas.
 *
 * Configuração (Definições do projeto > Propriedades do script):
 *   PIN             PIN partilhado pelos dirigentes (criado pelo setup())
 *   PIN_TESOURARIA  PIN do tesoureiro (criado pelo setup())
 *   SPREADSHEET_ID  criado pelo setup()
 *   PASTA_ID        criado pelo setup()
 */

const SECCOES = ['Agrupamento', 'Lobitos', 'Exploradores', 'Pioneiros', 'Caminheiros', 'Outros'];
const TIPOS_DESPESA = [
  'Alimentação', 'Combustível', 'Material de desgaste', 'Papelaria', 'Aluguer',
  'Portagens', 'Transportes', 'Material de limpeza', 'Material de campo', 'Outros',
];
const CABECALHO = [
  'ID', 'Data de carregamento', 'Data da fatura', 'Pago por', 'Secção', 'Detalhe secção (Outros)',
  'Atividade', 'Tipo de despesa', 'Detalhe tipo (Outros)', 'Descrição', 'Valor (€)', 'NIF do emissor', 'Fatura',
  'Reembolso', 'Forma de reembolso', 'Tratada em',
];
const COL_SECCAO = 'E';
const COL_TIPO = 'H';
const COL_VALOR = 'K';
const COL_FATURA = 13;
const COL_REEMBOLSO = 14;

const ESTADOS_REEMBOLSO = ['Por tratar', 'Reembolsado', 'Sem reembolso'];
const FORMAS_REEMBOLSO = ['Dinheiro', 'Transferência'];
const MAX_TRATADAS = 40;

const NOME_PASTA_RAIZ = 'Faturas do Agrupamento';
const NOME_FOLHA = 'Faturas do Agrupamento';
const FUSO = 'Europe/Lisbon';

const TIPOS_PERMITIDOS = {
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

// ---------------------------------------------------------------------------
// Configuração inicial — correr UMA vez a partir do editor
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
  if (props.getProperty('SPREADSHEET_ID')) {
    ss = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
  } else {
    ss = SpreadsheetApp.create(NOME_FOLHA);
    ss.setSpreadsheetLocale('pt_PT');
    ss.setSpreadsheetTimeZone(FUSO);
    DriveApp.getFileById(ss.getId()).moveTo(pasta);
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  obterFolhaAno(ss, anoLetivo(new Date()));
  // Remove a folha vazia que vem por defeito numa folha de cálculo nova
  ss.getSheets()
    .filter(f => /^(Folha|Sheet)\d+$/.test(f.getName()) && f.getLastRow() === 0)
    .forEach(f => ss.deleteSheet(f));

  ['PIN', 'PIN_TESOURARIA'].forEach(chave => {
    if (!props.getProperty(chave)) props.setProperty(chave, String(Math.floor(100000 + Math.random() * 900000)));
  });

  console.log('Pasta na Drive: ' + pasta.getUrl());
  console.log('Folha de cálculo: ' + ss.getUrl());
  console.log('PIN dos dirigentes: ' + props.getProperty('PIN'));
  console.log('PIN da tesouraria: ' + props.getProperty('PIN_TESOURARIA'));
  console.log('(Podes mudar os PINs em Definições do projeto > Propriedades do script)');
}

/** Faz uma submissão de teste para confirmar que está tudo bem configurado. */
function testeSubmissao() {
  // PNG de 1x1 píxel
  const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const resultado = submeter({
    id: 'teste-' + Utilities.getUuid(),
    pagoPor: 'Teste',
    seccao: 'Outros',
    seccaoOutro: 'Teste de configuração (podes apagar esta linha)',
    valor: 0.01,
    atividade: 'Teste',
    tipoDespesa: 'Outros',
    tipoOutro: 'Teste',
    descricao: 'Submissão de teste',
    dataFatura: Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd'),
    nif: '',
    ficheiro: { base64: png, mimeType: 'image/png' },
  });
  console.log(JSON.stringify(resultado));
}

// ---------------------------------------------------------------------------
// Endpoints da app web
// ---------------------------------------------------------------------------

function doGet() {
  return ContentService.createTextOutput('A API das faturas está a funcionar.');
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
      return resposta({ ok: false, erro: 'Demasiadas tentativas com PIN errado. Tenta daqui a 10 minutos.' });
    }
    const acao = String(pedido.acao || '');
    const tesouraria = acao.indexOf('tesouraria.') === 0;
    if (!pinCorreto(pedido.pin, tesouraria ? 'PIN_TESOURARIA' : 'PIN')) {
      registarFalhaPin();
      return resposta({ ok: false, codigo: 'PIN', erro: 'PIN errado.' });
    }

    if (acao === 'verificarPin' || acao === 'tesouraria.verificarPin') return resposta({ ok: true });
    if (acao === 'submeter') return resposta(submeter(pedido));
    if (acao === 'tesouraria.listar') return resposta(tesourariaListar());
    if (acao === 'tesouraria.foto') return resposta(tesourariaFoto(pedido));
    if (acao === 'tesouraria.marcar') return resposta(tesourariaMarcar(pedido));
    return resposta({ ok: false, erro: 'Ação desconhecida.' });
  } catch (err) {
    if (err instanceof ErroValidacao) return resposta({ ok: false, erro: err.message });
    console.error(err);
    return resposta({ ok: false, erro: 'Erro no servidor. Tenta novamente.' });
  }
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
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
// Submissão
// ---------------------------------------------------------------------------

function ErroValidacao(mensagem) {
  this.message = mensagem;
}

function submeter(p) {
  const d = validar(p);
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
  const ano = anoLetivo(d.dataFatura);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const folha = obterFolhaAno(ss, ano);

    // A app repete o pedido se a rede falhar; o ID evita linhas duplicadas.
    if (idJaExiste(folha, d.id)) return { ok: true, anoLetivo: ano, duplicado: true };

    const pastaAno = obterPastaAno(DriveApp.getFolderById(props.getProperty('PASTA_ID')), ano);
    const nome = [
      Utilities.formatDate(d.dataFatura, FUSO, 'yyyy-MM-dd'),
      d.seccao === 'Outros' ? 'Outros-' + d.seccaoOutro : d.seccao,
      d.valor.toFixed(2).replace('.', ',') + 'EUR',
      d.id.slice(0, 8),
    ].join('_').replace(/[\\/:*?"<>|\n\r]+/g, '-').slice(0, 150) + '.' + TIPOS_PERMITIDOS[d.mimeType];
    const ficheiro = pastaAno.createFile(Utilities.newBlob(d.bytes, d.mimeType, nome));
    ficheiro.setDescription(d.pagoPor + ' — ' + d.atividade + ' — ' + d.tipoDespesa);

    folha.appendRow([
      texto(d.id),
      new Date(),
      d.dataFatura,
      texto(d.pagoPor),
      d.seccao,
      texto(d.seccaoOutro),
      texto(d.atividade),
      d.tipoDespesa,
      texto(d.tipoOutro),
      texto(d.descricao),
      d.valor,
      texto(d.nif),
      '',
      'Por tratar',
      '',
      '',
    ]);
    folha.getRange(folha.getLastRow(), COL_FATURA).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText('Ver fatura').setLinkUrl(ficheiro.getUrl()).build()
    );
    SpreadsheetApp.flush();
    return { ok: true, anoLetivo: ano };
  } finally {
    lock.releaseLock();
  }
}

function validar(p) {
  const str = (v, max) => String(v == null ? '' : v).trim().slice(0, max);

  const id = str(p.id, 64);
  if (!/^[A-Za-z0-9-]{8,64}$/.test(id)) throw new ErroValidacao('Identificador inválido.');

  const pagoPor = str(p.pagoPor, 100);
  if (!pagoPor) throw new ErroValidacao('Indica quem pagou.');

  const seccao = str(p.seccao, 50);
  if (SECCOES.indexOf(seccao) === -1) throw new ErroValidacao('Secção inválida.');
  const seccaoOutro = seccao === 'Outros' ? str(p.seccaoOutro, 200) : '';
  if (seccao === 'Outros' && !seccaoOutro) throw new ErroValidacao('Descreve a secção "Outros".');

  const valor = Math.round(Number(p.valor) * 100) / 100;
  if (!isFinite(valor) || valor <= 0 || valor > 100000) throw new ErroValidacao('Valor inválido.');

  const atividade = str(p.atividade, 200);
  if (!atividade) throw new ErroValidacao('Indica o nome da atividade.');

  const tipoDespesa = str(p.tipoDespesa, 50);
  if (TIPOS_DESPESA.indexOf(tipoDespesa) === -1) throw new ErroValidacao('Tipo de despesa inválido.');
  const tipoOutro = tipoDespesa === 'Outros' ? str(p.tipoOutro, 200) : '';
  if (tipoDespesa === 'Outros' && !tipoOutro) throw new ErroValidacao('Especifica o tipo de despesa "Outros".');

  const descricao = str(p.descricao, 1000);

  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str(p.dataFatura, 10));
  if (!m) throw new ErroValidacao('Data da fatura inválida.');
  const dataFatura = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const amanha = new Date(Date.now() + 24 * 3600 * 1000);
  if (dataFatura.getMonth() !== Number(m[2]) - 1 || dataFatura > amanha || dataFatura.getFullYear() < 2000) {
    throw new ErroValidacao('Data da fatura inválida.');
  }

  const nif = str(p.nif, 20).replace(/\s/g, '');
  if (nif && !nifValido(nif)) throw new ErroValidacao('NIF do emissor inválido.');

  const f = p.ficheiro || {};
  const mimeType = str(f.mimeType, 50).toLowerCase();
  if (!TIPOS_PERMITIDOS[mimeType]) throw new ErroValidacao('Tipo de ficheiro não suportado.');
  let bytes;
  try {
    bytes = Utilities.base64Decode(String(f.base64 || ''));
  } catch (err) {
    throw new ErroValidacao('Ficheiro corrompido.');
  }
  if (!bytes.length) throw new ErroValidacao('Falta a foto da fatura.');
  if (bytes.length > TAMANHO_MAX_FICHEIRO) throw new ErroValidacao('Ficheiro demasiado grande (máx. 10 MB).');

  return { id, pagoPor, seccao, seccaoOutro, valor, atividade, tipoDespesa, tipoOutro, descricao, dataFatura, nif, mimeType, bytes };
}

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

// ---------------------------------------------------------------------------
// Ano letivo: de 1 de setembro a 31 de agosto (pela data da fatura)
// ---------------------------------------------------------------------------

function anoLetivo(data) {
  const inicio = data.getMonth() >= 8 ? data.getFullYear() : data.getFullYear() - 1;
  return inicio + '-' + String((inicio + 1) % 100).padStart(2, '0');
}

function obterFolhaAno(ss, ano) {
  const existente = ss.getSheetByName(ano);
  if (existente) {
    const largura = existente.getLastColumn();
    const cabecalho = existente.getRange(1, 1, 1, largura).getValues()[0];
    if (cabecalho.every((c, i) => c === CABECALHO[i])) {
      if (largura < CABECALHO.length) acrescentarColunas(existente, largura);
      return existente;
    }
    // Aba criada com uma versão antiga das colunas: arquiva-a e cria uma nova
    let sufixo = ' (antigo)';
    for (let n = 2; ss.getSheetByName(ano + sufixo); n++) sufixo = ' (antigo ' + n + ')';
    existente.setName(ano + sufixo);
    const resumoAntigo = ss.getSheetByName('Resumo ' + ano);
    if (resumoAntigo) resumoAntigo.setName('Resumo ' + ano + sufixo);
  }

  const folha = ss.insertSheet(ano, 0);
  folha.appendRow(CABECALHO);
  folha.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold').setBackground('#e8f0e8');
  folha.setFrozenRows(1);
  folha.getRange('B2:B').setNumberFormat('dd/mm/yyyy hh:mm');
  folha.getRange('C2:C').setNumberFormat('dd/mm/yyyy');
  folha.getRange(COL_VALOR + '2:' + COL_VALOR).setNumberFormat('#,##0.00 "€"');
  folha.getRange('L2:L').setNumberFormat('@');
  formatarColunasReembolso(folha);
  folha.hideColumns(1);
  folha.setColumnWidth(7, 220);
  folha.setColumnWidth(10, 280);

  const resumo = ss.insertSheet('Resumo ' + ano, 0);
  escreverResumo(resumo, ano, 1, 'Secção', SECCOES, COL_SECCAO);
  escreverResumo(resumo, ano, 5, 'Tipo de despesa', TIPOS_DESPESA, COL_TIPO);
  resumo.setColumnWidth(1, 160);
  resumo.setColumnWidth(5, 180);

  return folha;
}

/** Acrescenta a uma aba existente as colunas que foram adicionadas depois de ela ser criada. */
function acrescentarColunas(folha, largura) {
  const novas = CABECALHO.slice(largura);
  folha.getRange(1, largura + 1, 1, novas.length).setValues([novas])
    .setFontWeight('bold').setBackground('#e8f0e8');
  const n = folha.getLastRow() - 1;
  if (largura < COL_REEMBOLSO && n > 0) folha.getRange(2, COL_REEMBOLSO, n, 1).setValue('Por tratar');
  formatarColunasReembolso(folha);
}

function formatarColunasReembolso(folha) {
  const regra = lista => SpreadsheetApp.newDataValidation().requireValueInList(lista, true).build();
  folha.getRange(2, COL_REEMBOLSO, folha.getMaxRows() - 1, 1).setDataValidation(regra(ESTADOS_REEMBOLSO));
  folha.getRange(2, COL_REEMBOLSO + 1, folha.getMaxRows() - 1, 1).setDataValidation(regra(FORMAS_REEMBOLSO));
  folha.getRange(2, COL_REEMBOLSO + 2, folha.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
}

/** Tabela de totais e nº de faturas por categoria, a começar na coluna `col`. */
function escreverResumo(resumo, ano, col, titulo, categorias, colCategoria) {
  const ref = "'" + ano + "'!";
  const letra = String.fromCharCode(64 + col);
  const intervaloCat = ref + colCategoria + ':' + colCategoria;
  const intervaloValor = ref + COL_VALOR + ':' + COL_VALOR;
  const linhas = [[titulo, 'Total (€)', 'Nº de faturas']];
  categorias.forEach((c, i) => {
    const r = i + 2;
    linhas.push([
      c,
      '=SUMIF(' + intervaloCat + ',' + letra + r + ',' + intervaloValor + ')',
      '=COUNTIF(' + intervaloCat + ',' + letra + r + ')',
    ]);
  });
  const ultima = categorias.length + 1;
  const total = String.fromCharCode(65 + col);
  const conta = String.fromCharCode(66 + col);
  linhas.push([
    'Total',
    '=SUM(' + total + '2:' + total + ultima + ')',
    '=SUM(' + conta + '2:' + conta + ultima + ')',
  ]);
  resumo.getRange(1, col, linhas.length, 3).setValues(linhas);
  resumo.getRange(1, col, 1, 3).setFontWeight('bold').setBackground('#e8f0e8');
  resumo.getRange(linhas.length, col, 1, 3).setFontWeight('bold');
  resumo.getRange(2, col + 1, linhas.length - 1, 1).setNumberFormat('#,##0.00 "€"');
}

function obterPastaAno(raiz, ano) {
  const it = raiz.getFoldersByName(ano);
  return it.hasNext() ? it.next() : raiz.createFolder(ano);
}

function idJaExiste(folha, id) {
  const n = folha.getLastRow() - 1;
  if (n < 1) return false;
  return folha.getRange(2, 1, n, 1).getValues().some(l => l[0] === id);
}

// ---------------------------------------------------------------------------
// Tesouraria
// ---------------------------------------------------------------------------

const COL = {};
CABECALHO.forEach((c, i) => { COL[c] = i; });

/** Faturas por tratar (mais antigas primeiro) e as últimas já tratadas. */
function tesourariaListar() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  const pendentes = [];
  const tratadas = [];

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ss.getSheets().map(f => f.getName()).filter(nome => /^\d{4}-\d{2}$/.test(nome)).forEach(ano => {
      const folha = obterFolhaAno(ss, ano); // acrescenta as colunas de reembolso se faltarem
      const n = folha.getLastRow() - 1;
      if (n < 1) return;
      const valores = folha.getRange(2, 1, n, CABECALHO.length).getValues();
      const links = folha.getRange(2, COL_FATURA, n, 1).getRichTextValues();
      valores.forEach((l, i) => {
        if (!l[COL['ID']]) return;
        const fatura = linhaParaFatura(l, links[i][0], ano);
        (fatura.reembolso === 'Por tratar' ? pendentes : tratadas).push(fatura);
      });
    });
  } finally {
    lock.releaseLock();
  }

  pendentes.sort((a, b) => a.ordem - b.ordem);
  tratadas.sort((a, b) => b.ordemTratada - a.ordemTratada);
  return {
    ok: true,
    pendentes,
    tratadas: tratadas.slice(0, MAX_TRATADAS),
    totalPendente: Math.round(pendentes.reduce((s, f) => s + f.valor, 0) * 100) / 100,
  };
}

function linhaParaFatura(l, link, ano) {
  const data = v => (v instanceof Date ? Utilities.formatDate(v, FUSO, 'dd/MM/yyyy') : '');
  const ms = v => (v instanceof Date ? v.getTime() : 0);
  const url = link ? link.getLinkUrl() || '' : '';
  const m = /\/d\/([A-Za-z0-9_-]+)/.exec(url);
  const estado = String(l[COL['Reembolso']] || '');
  return {
    id: String(l[COL['ID']]),
    ano,
    dataFatura: data(l[COL['Data da fatura']]),
    ordem: ms(l[COL['Data da fatura']]) || ms(l[COL['Data de carregamento']]),
    pagoPor: String(l[COL['Pago por']]),
    seccao: String(l[COL['Secção']]),
    seccaoOutro: String(l[COL['Detalhe secção (Outros)']]),
    atividade: String(l[COL['Atividade']]),
    tipoDespesa: String(l[COL['Tipo de despesa']]),
    tipoOutro: String(l[COL['Detalhe tipo (Outros)']]),
    descricao: String(l[COL['Descrição']]),
    valor: Number(l[COL['Valor (€)']]) || 0,
    url,
    fileId: m ? m[1] : '',
    reembolso: ESTADOS_REEMBOLSO.indexOf(estado) === -1 ? 'Por tratar' : estado,
    forma: String(l[COL['Forma de reembolso']] || ''),
    tratadaEm: data(l[COL['Tratada em']]),
    ordemTratada: ms(l[COL['Tratada em']]),
  };
}

/** Devolve a imagem da fatura (ou uma miniatura, no caso de PDF/HEIC). */
function tesourariaFoto(p) {
  const fileId = String(p.fileId || '');
  if (!/^[A-Za-z0-9_-]{10,}$/.test(fileId)) throw new ErroValidacao('Fatura não encontrada.');
  let ficheiro;
  try {
    ficheiro = DriveApp.getFileById(fileId);
  } catch (err) {
    throw new ErroValidacao('A foto desta fatura já não existe na Drive.');
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

function tesourariaMarcar(p) {
  const estado = String(p.estado || '');
  if (ESTADOS_REEMBOLSO.indexOf(estado) === -1) throw new ErroValidacao('Opção inválida.');
  const forma = estado === 'Reembolsado' ? String(p.forma || '') : '';
  if (estado === 'Reembolsado' && FORMAS_REEMBOLSO.indexOf(forma) === -1) {
    throw new ErroValidacao('Indique como foi feito o reembolso.');
  }
  const ano = String(p.ano || '');
  const id = String(p.id || '');
  if (!/^\d{4}-\d{2}$/.test(ano) || !id) throw new ErroValidacao('Fatura não encontrada.');

  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const folha = ss.getSheetByName(ano);
    const n = folha ? folha.getLastRow() - 1 : 0;
    if (n < 1) throw new ErroValidacao('Fatura não encontrada.');
    const linha = folha.getRange(2, 1, n, 1).getValues().findIndex(l => String(l[0]) === id);
    if (linha === -1) throw new ErroValidacao('Fatura não encontrada. Pode ter sido apagada da folha.');
    folha.getRange(linha + 2, COL_REEMBOLSO, 1, 3).setValues([[
      estado,
      forma,
      estado === 'Por tratar' ? '' : new Date(),
    ]]);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
