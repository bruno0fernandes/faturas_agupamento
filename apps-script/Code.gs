/**
 * Faturas do Agrupamento — backend (Google Apps Script)
 *
 * Recebe as submissões da app, guarda a foto/PDF na Google Drive
 * e acrescenta uma linha na Google Sheet do ano letivo respetivo.
 *
 * Configuração (Definições do projeto > Propriedades do script):
 *   PIN             PIN partilhado pelos dirigentes (criado pelo setup())
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
];
const COL_SECCAO = 'E';
const COL_TIPO = 'H';
const COL_VALOR = 'K';
const COL_FATURA = 13;

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

  if (!props.getProperty('PIN')) {
    props.setProperty('PIN', String(Math.floor(100000 + Math.random() * 900000)));
  }

  console.log('Pasta na Drive: ' + pasta.getUrl());
  console.log('Folha de cálculo: ' + ss.getUrl());
  console.log('PIN: ' + props.getProperty('PIN') + '  (podes mudá-lo em Propriedades do script)');
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
    if (!pinCorreto(pedido.pin)) {
      registarFalhaPin();
      return resposta({ ok: false, codigo: 'PIN', erro: 'PIN errado.' });
    }

    if (pedido.acao === 'verificarPin') return resposta({ ok: true });
    if (pedido.acao === 'submeter') return resposta(submeter(pedido));
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

function pinCorreto(pin) {
  const esperado = PropertiesService.getScriptProperties().getProperty('PIN');
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
    const cabecalho = existente.getRange(1, 1, 1, existente.getLastColumn()).getValues()[0];
    if (cabecalho.join('|') === CABECALHO.join('|')) return existente;
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
