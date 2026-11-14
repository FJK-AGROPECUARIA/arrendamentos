/**
 * Gestão de Arrendamentos FJK — Google Apps Script (servidor)
 * Guarda tudo numa planilha própria e as fotos das mortes numa pasta do Drive.
 * Publicar: Implantar > Nova implantação > App da Web
 *   Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 * Copie o link /exec e cole no app (Gestão > Conexão com a planilha) ou no API_URL do index.html.
 */

const TABELAS = {
  arr:   { aba: 'Arrendamentos', cols: ['id', 'nome', 'cidade', 'ativo'] },
  lotes: { aba: 'Lotes',         cols: ['id', 'arrId', 'nome', 'prop', 'valorCab', 'salario', 'ativo'] },
  movs:  { aba: 'Movimentos',    cols: ['id', 'data', 'tipo', 'loteId', 'loteDest', 'qtd', 'sexo', 'peso', 'valor', 'brinco', 'foto', 'obs', 'criadoEm', 'subtipo', 'precoArroba', 'divArroba', 'destino'] },
  desp:  { aba: 'Despesas',      cols: ['id', 'data', 'tipo', 'arrId', 'loteId', 'desc', 'valor', 'status', 'ref', 'pagoEm', 'criadoEm'] },
  props: { aba: 'Proprietarios', cols: ['id', 'nome'] }
};

function planilha_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('Gestão de Arrendamentos FJK - Dados');
  p.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function aba_(t) {
  const ss = planilha_();
  const def = TABELAS[t];
  let sh = ss.getSheetByName(def.aba);
  if (!sh) {
    sh = ss.insertSheet(def.aba);
    sh.getRange('A:Z').setNumberFormat('@'); // tudo como texto: datas ficam aaaa-mm-dd
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold').setBackground('#6d4321').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    ['Sheet1', 'Página1', 'Planilha1'].forEach(n => { const s = ss.getSheetByName(n); if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });
  }
  const n = def.cols.length, cab = sh.getRange(1, 1, 1, n).getValues()[0];
  if (cab.join('|') !== def.cols.join('|')) sh.getRange(1, 1, 1, n).setValues([def.cols]).setFontWeight('bold').setBackground('#6d4321').setFontColor('#ffffff');
  return sh;
}

function ler_(t) {
  const sh = aba_(t), cols = TABELAS[t].cols, n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => { const o = {}; cols.forEach((c, i) => { const v = r[i]; o[c] = v === null ? '' : (v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v)); }); return o; });
}

function linhaDoId_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function pastaFotos_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('FOTOS_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('Arrendamentos FJK - Fotos das mortes');
  p.setProperty('FOTOS_ID', f.getId());
  return f;
}

function salvarFoto_(reg) {
  const m = String(reg.foto).match(/^data:(image\/\w+);base64,(.+)$/);
  if (!m) return reg.foto;
  const nome = (reg.data || '') + '_brinco_' + (reg.brinco || 'sem') + '_' + reg.id + '.jpg';
  const arq = pastaFotos_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nome));
  arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + arq.getId() + '&sz=w800';
}

function salvar_(t, reg) {
  if (!TABELAS[t]) throw new Error('Tabela inválida: ' + t);
  if (t === 'movs' && reg.foto && String(reg.foto).indexOf('data:') === 0) reg.foto = salvarFoto_(reg);
  const sh = aba_(t), cols = TABELAS[t].cols;
  const linha = cols.map(c => reg[c] === undefined || reg[c] === null ? '' : String(reg[c]));
  const r = linhaDoId_(sh, reg.id);
  const lin = r > 0 ? r : sh.getLastRow() + 1;
  sh.getRange(lin, 1, 1, cols.length).setNumberFormat('@').setValues([linha]);
  return reg;
}

function excluir_(t, id) {
  const sh = aba_(t), r = linhaDoId_(sh, id);
  if (r > 0) sh.deleteRow(r);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const acao = (e && e.parameter && e.parameter.acao) || 'ping';
    if (acao === 'tudo') {
      const dados = {};
      Object.keys(TABELAS).forEach(t => dados[t] = ler_(t));
      return json_({ ok: true, dados, planilha: planilha_().getUrl() });
    }
    return json_({ ok: true, msg: 'Gestão de Arrendamentos FJK no ar' });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const corpo = JSON.parse(e.postData.contents);
    const resultados = (corpo.ops || []).map(op => {
      if (op.acao === 'salvar') return { tabela: op.tabela, reg: salvar_(op.tabela, op.reg) };
      if (op.acao === 'excluir') { excluir_(op.tabela, op.id); return { tabela: op.tabela, id: op.id }; }
      return null;
    });
    return json_({ ok: true, resultados });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Rode uma vez pelo editor (botão Executar) para criar a planilha, as abas e a pasta de fotos e autorizar o acesso. */
function configurar() {
  Object.keys(TABELAS).forEach(aba_);
  pastaFotos_();
  Logger.log('Planilha: ' + planilha_().getUrl());
}
