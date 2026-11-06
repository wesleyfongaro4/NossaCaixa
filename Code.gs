/**
 * Nosso Caixa — Backend Google Apps Script
 * Versão: 2026-11-05-v4  (transporte JSONP / ContentService)
 *
 * Mudança de arquitetura:
 *   NÃO usa mais google.script.run.
 *   O frontend chama a URL do Web App com ?action=...&payload=...&callback=...
 *   O servidor responde SEMPRE com  callback( {json} );  (ContentService)
 *
 * Vantagens:
 *   - Funciona dentro do Web App e também com o HTML aberto localmente.
 *   - Não depende de serialização do google.script.run.
 *   - Qualquer erro volta como texto legível no próprio corpo da resposta.
 */

var NC_VERSION = '2026-11-05-v4';
var SHEET_NAME = 'Nosso Caixa';
var SHEET_ID = '1Gg8bc1IbbWhAXzh8hn8Hk7FS7rYtuf00blnRuVeEXA0'; // planilha "Nosso Caixa" fixada
var TZ = 'America/Sao_Paulo';
var NC_PROP = 'NC_SHEET_ID';

var COLLECTIONS = ['tx','goals','contas','pagamentos','avisos','recebimentos','cartoes','compras','faturas','projetos','membros','perfis','cfg'];

var SCHEMA = {
  tx: ['id','esp','type','amount','cat','who','shared','date','desc','comp','ts'],
  goals: ['id','esp','name','target','saved','deadline'],
  contas: ['id','esp','name','cat','amount','who','pay','notes','shared','due','rec','n','every','unit','canceledFrom','created'],
  pagamentos: ['id','esp','billId','comp','paidOn','amount','who','cat','name','pay','shared'],
  avisos: ['id','esp','ts'],
  recebimentos: ['id','esp','name','amount','due','rec','who','canceledFrom'],
  cartoes: ['id','esp','name','limit','closeDay','dueDay'],
  compras: ['id','esp','cardId','name','amount','date','n','cat','who'],
  faturas: ['id','esp','paidOn','total','dueDate'],
  projetos: ['id','esp','nome','descricao','justificativa','valor','categoria','prioridade','prazo','forma','extras','cenarios','versao','status','criado'],
  membros: ['id','esp','uid','nome','papel','entrou','ativo','cor'],
  perfis: ['id','esp','uid','nome','email','metodo','criado'],
  cfg: ['id','esp','names','budget','remind','ch','saldo','reservaEmergencia','familia']
};

/* ------------------------------------------------------------------ */
/* Utilitários                                                         */
/* ------------------------------------------------------------------ */

function ncErr_(err) {
  if (!err) return 'erro desconhecido';
  if (typeof err === 'string') return err;
  return String(err.message || err);
}

function ncJson_(value) {
  try {
    return JSON.stringify(ncClean_(value));
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'falha ao serializar: ' + ncErr_(e) });
  }
}

function ncClean_(value, depth, seen) {
  depth = depth || 0;
  seen = seen || [];
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return Utilities.formatDate(value, TZ, "yyyy-MM-dd'T'HH:mm:ss");
  var t = typeof value;
  if (t === 'string' || t === 'number' || t === 'boolean') return value;
  if (t === 'function') return '[funcao]';
  if (depth > 20) return '[profundidade]';
  for (var s = 0; s < seen.length; s++) if (seen[s] === value) return '[ciclo]';
  seen = seen.concat([value]);
  if (Array.isArray(value)) {
    var arr = [];
    for (var i = 0; i < value.length; i++) arr.push(ncClean_(value[i], depth + 1, seen));
    return arr;
  }
  if (t === 'object') {
    var out = {};
    for (var k in value) {
      if (Object.prototype.hasOwnProperty.call(value, k)) out[k] = ncClean_(value[k], depth + 1, seen);
    }
    return out;
  }
  return String(value);
}

function nowIso_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function getUser_() {
  try {
    var email = (Session.getActiveUser().getEmail() || 'local@demo').trim().toLowerCase();
    var name = email.split('@')[0].replace(/[._]/g, ' ').replace(/\b\w/g, function (c) { return c.toUpperCase(); });
    return { uid: email, name: name, email: email };
  } catch (e) {
    return { uid: 'local@demo', name: 'Usuario', email: 'local@demo' };
  }
}

/* ------------------------------------------------------------------ */
/* Planilha                                                            */
/* ------------------------------------------------------------------ */

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();

  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);

  var cached = props.getProperty(NC_PROP);
  if (cached) {
    try { return SpreadsheetApp.openById(cached); }
    catch (e) { props.deleteProperty(NC_PROP); }
  }

  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) { props.setProperty(NC_PROP, active.getId()); return active; }
  } catch (e) { /* standalone */ }

  var found = null;
  try {
    var files = DriveApp.getFilesByName(SHEET_NAME);
    while (files.hasNext()) {
      var f = files.next();
      if (f.getMimeType() === MimeType.GOOGLE_SHEETS ||
          f.getMimeType() === 'application/vnd.google-apps.spreadsheet') { found = f; break; }
    }
  } catch (e) { /* sem DriveApp */ }

  var ss = found ? SpreadsheetApp.openById(found.getId()) : SpreadsheetApp.create(SHEET_NAME);
  props.setProperty(NC_PROP, ss.getId());
  return ss;
}

function ensureSheets_() {
  var ss = getSpreadsheet_();
  var existing = {};
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) existing[sheets[i].getName()] = true;

  for (var c = 0; c < COLLECTIONS.length; c++) {
    var name = COLLECTIONS[c];
    if (!existing[name]) {
      var sh = ss.insertSheet(name);
      var header = SCHEMA[name];
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.setFrozenRows(1);
    }
  }
  return ss;
}

/* ------------------------------------------------------------------ */
/* Células                                                             */
/* ------------------------------------------------------------------ */

function stringify_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss");
  if (typeof v === 'object') return JSON.stringify(v);
  return v;
}

function parseCell_(v) {
  if (typeof v !== 'string') return v;
  var t = v.trim();
  if ((t.charAt(0) === '{' && t.charAt(t.length - 1) === '}') ||
      (t.charAt(0) === '[' && t.charAt(t.length - 1) === ']')) {
    try { return JSON.parse(t); } catch (e) { return v; }
  }
  return v;
}

function readSheet_(name) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var hdr = vals[0].map(function (h) { return String(h); });
  var out = [];

  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var empty = true;
    for (var j = 0; j < row.length; j++) {
      if (row[j] !== '' && row[j] !== null && row[j] !== undefined) { empty = false; break; }
    }
    if (empty) continue;
    var o = {};
    for (var k = 0; k < hdr.length; k++) o[hdr[k]] = parseCell_(row[k]);
    out.push(o);
  }
  return out;
}

function findRow_(sh, id) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return null;
  var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var hdr = vals[0].map(function (h) { return String(h); });
  var idIdx = hdr.indexOf('id');
  if (idIdx < 0) return null;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][idIdx]) === String(id)) return { row: i + 1, hdr: hdr };
  }
  return null;
}

function upsertRow_(collection, id, record) {
  var ss = ensureSheets_();
  var sh = ss.getSheetByName(collection);
  if (!sh) throw new Error('Aba nao encontrada: ' + collection);

  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h); });
  var idIdx = hdr.indexOf('id');
  if (idIdx < 0) throw new Error('Coluna id ausente na aba: ' + collection);

  var rowIndex = -1;
  if (sh.getLastRow() >= 2) {
    var vals = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][idIdx]) === String(id)) { rowIndex = i + 1; break; }
    }
  }

  var previous = rowIndex > 0 ? sh.getRange(rowIndex, 1, 1, lastCol).getValues()[0] : [];
  var row = [];
  for (var k = 0; k < hdr.length; k++) {
    var key = hdr[k];
    if (record[key] !== undefined) row.push(stringify_(record[key]));
    else row.push(rowIndex > 0 ? previous[k] : '');
  }

  if (rowIndex > 0) sh.getRange(rowIndex, 1, 1, hdr.length).setValues([row]);
  else sh.appendRow(row);

  return { ok: true, id: id };
}

function deleteRow_(collection, id) {
  var ss = ensureSheets_();
  var sh = ss.getSheetByName(collection);
  if (!sh) return { ok: false, error: 'aba_nao_encontrada' };

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'not_found' };
  var vals = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  var hdr = vals[0].map(function (h) { return String(h); });
  var idIdx = hdr.indexOf('id');
  if (idIdx < 0) return { ok: false, error: 'coluna_id_ausente' };

  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][idIdx]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, error: 'not_found' };
}

/* ------------------------------------------------------------------ */
/* Config                                                              */
/* ------------------------------------------------------------------ */

function parseCfg_(row) {
  if (!row) return null;
  var names = row.names;
  if (typeof names === 'string' && names) {
    try { names = JSON.parse(names); } catch (e) { names = names.split(','); }
  }
  if (!Array.isArray(names) || !names.length) names = ['Eu', 'Esposo'];

  return {
    names: names,
    budget: (row.budget && typeof row.budget === 'object') ? row.budget : {},
    remind: (row.remind && Array.isArray(row.remind)) ? row.remind : [3, 1, 0, -1],
    ch: (row.ch && Array.isArray(row.ch)) ? row.ch : ['interno'],
    saldo: (row.saldo === '' || row.saldo === null || row.saldo === undefined) ? undefined : Number(row.saldo),
    reservaEmergencia: (row.reservaEmergencia === '' || row.reservaEmergencia === null || row.reservaEmergencia === undefined) ? 0 : Number(row.reservaEmergencia),
    familia: (row.familia && typeof row.familia === 'object') ? row.familia : undefined
  };
}

/* ------------------------------------------------------------------ */
/* Bootstrap                                                           */
/* ------------------------------------------------------------------ */

function ncBootstrapPayload_() {
  var warnings = [];
  var collections = {};

  try { ensureSheets_(); }
  catch (e) { warnings.push('planilha: ' + ncErr_(e)); }

  for (var i = 0; i < COLLECTIONS.length; i++) {
    var c = COLLECTIONS[i];
    try { collections[c] = readSheet_(c) || []; }
    catch (e2) { collections[c] = []; warnings.push(c + ': ' + ncErr_(e2)); }
  }

  var cfgRows = collections.cfg || [];
  var user = getUser_();
  var cfgMain = null, cfgP = null;

  for (var a = 0; a < cfgRows.length; a++) {
    var r = cfgRows[a];
    if (!r) continue;
    if (!cfgMain && (r.esp === 'f' || r.esp === '' || r.esp === null || r.esp === undefined)) cfgMain = parseCfg_(r);
    if (!cfgP && r.esp === ('p_' + user.uid)) cfgP = parseCfg_(r);
  }

  collections.membros = (collections.membros || []).map(function (m) { m.esp = m.esp || 'f'; return m; });
  collections.perfis = (collections.perfis || []).map(function (p) { p.esp = p.esp || 'f'; return p; });

  var payload = {
    ok: true,
    version: NC_VERSION,
    user: user,
    today: nowIso_(),
    collections: collections,
    cfg_main: cfgMain,
    cfg_p: cfgP
  };
  if (warnings.length) payload.warning = warnings.join(' | ');
  return payload;
}

/* ------------------------------------------------------------------ */
/* CRUD                                                                */
/* ------------------------------------------------------------------ */

function adicionar(collection, record) {
  try {
    if (COLLECTIONS.indexOf(collection) < 0) throw new Error('Colecao invalida: ' + collection);
    record = record || {};
    if (!record.id) record.id = Utilities.getUuid();

    // Idempotencia: mesma requisicao repetida (ex.: duplo clique, retry)
    // nao cria uma segunda linha para o mesmo id.
    try {
      var ss = ensureSheets_();
      var sh = ss.getSheetByName(collection);
      if (sh && findRow_(sh, record.id)) return { ok: true, id: record.id, existed: true };
    } catch (e) { /* se a verificacao falhar, segue para inserir */ }

    upsertRow_(collection, record.id, record);
    return { ok: true, id: record.id };
  } catch (err) { return { ok: false, error: ncErr_(err) }; }
}

function salvarDoc(collection, id, record) {
  try {
    if (COLLECTIONS.indexOf(collection) < 0) throw new Error('Colecao invalida: ' + collection);
    record = record || {};
    record.id = id;
    upsertRow_(collection, id, record);
    return { ok: true, id: id };
  } catch (err) { return { ok: false, error: ncErr_(err) }; }
}

function atualizar(collection, id, patch) {
  try {
    if (COLLECTIONS.indexOf(collection) < 0) throw new Error('Colecao invalida: ' + collection);
    var ss = ensureSheets_();
    var sh = ss.getSheetByName(collection);
    if (!sh) return { ok: false, error: 'aba_nao_encontrada' };

    var info = findRow_(sh, id);
    if (!info) return { ok: false, error: 'not_found' };

    var hdr = info.hdr;
    var vals = sh.getRange(info.row, 1, 1, hdr.length).getValues()[0];
    var merged = {};
    for (var i = 0; i < hdr.length; i++) merged[hdr[i]] = parseCell_(vals[i]);

    patch = patch || {};
    for (var b in patch) if (Object.prototype.hasOwnProperty.call(patch, b)) merged[b] = patch[b];
    merged.id = id;

    var row = [];
    for (var k = 0; k < hdr.length; k++) row.push(merged[hdr[k]] === undefined ? '' : stringify_(merged[hdr[k]]));
    sh.getRange(info.row, 1, 1, hdr.length).setValues([row]);
    return { ok: true };
  } catch (err) { return { ok: false, error: ncErr_(err) }; }
}

function excluir(collection, id) {
  try {
    if (COLLECTIONS.indexOf(collection) < 0) throw new Error('Colecao invalida: ' + collection);
    return deleteRow_(collection, id);
  } catch (err) { return { ok: false, error: ncErr_(err) }; }
}

function salvarCfg(esp, cfg) {
  try {
    esp = esp || 'f';
    cfg = cfg || {};
    var record = {};
    for (var k in cfg) if (Object.prototype.hasOwnProperty.call(cfg, k)) record[k] = cfg[k];
    record.id = esp;
    record.esp = esp;

    var keys = ['names', 'budget', 'remind', 'ch', 'familia'];
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (record[key] !== undefined && record[key] !== null && typeof record[key] === 'object') {
        record[key] = JSON.stringify(record[key]);
      }
    }
    if (record.saldo === undefined || record.saldo === null) record.saldo = '';

    upsertRow_('cfg', esp, record);
    return { ok: true };
  } catch (err) { return { ok: false, error: ncErr_(err) }; }
}

/* ------------------------------------------------------------------ */
/* Despacho                                                            */
/* ------------------------------------------------------------------ */

function ncDispatch_(action, args) {
  args = Array.isArray(args) ? args : [];
  switch (action) {
    case 'ping':
      return { ok: true, version: NC_VERSION, now: new Date().getTime(), user: getUser_() };
    case 'bootstrap':
      return ncBootstrapPayload_();
    case 'selftest':
      return { ok: true, version: NC_VERSION, text: ncSelftest() };
    case 'add':
      return adicionar(args[0], args[1]);
    case 'save':
      return salvarDoc(args[0], args[1], args[2]);
    case 'update':
      return atualizar(args[0], args[1], args[2]);
    case 'delete':
      return excluir(args[0], args[1]);
    case 'cfg':
      return salvarCfg(args[0], args[1]);
    default:
      return { ok: false, error: 'acao desconhecida: ' + action };
  }
}

/** Compatibilidade: retorna string JSON */
function ncApi(action, payloadJson) {
  var args = [];
  try { args = payloadJson ? JSON.parse(payloadJson) : []; }
  catch (e) { return ncJson_({ ok: false, error: 'payload invalido: ' + ncErr_(e) }); }
  try { return ncJson_(ncDispatch_(action, args)); }
  catch (err) { return ncJson_({ ok: false, error: ncErr_(err) }); }
}

/* ------------------------------------------------------------------ */
/* HTTP                                                               */
/* ------------------------------------------------------------------ */

function ncApiUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; }
  catch (e) { return ''; }
}

function ncRespond_(action, payloadJson, callback) {
  var args = [];
  try { args = payloadJson ? JSON.parse(payloadJson) : []; }
  catch (e) { return ncEmit_({ ok: false, error: 'payload invalido: ' + ncErr_(e) }, callback); }

  var result;
  try { result = ncDispatch_(action, args); }
  catch (err) { result = { ok: false, error: ncErr_(err), stack: err && err.stack ? String(err.stack) : '' }; }

  return ncEmit_(result, callback);
}

function ncEmit_(result, callback) {
  var json = ncJson_(result);
  if (callback) {
    var safe = String(callback).replace(/[^A-Za-z0-9_$.]/g, '');
    return ContentService.createTextOutput(safe + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Diagnóstico                                                        */
/* ------------------------------------------------------------------ */

function ncSelftest() {
  var linhas = [];
  linhas.push('NC_VERSION=' + NC_VERSION);
  try { linhas.push('usuario=' + (Session.getActiveUser().getEmail() || '(vazio)')); }
  catch (e) { linhas.push('usuario=ERRO ' + ncErr_(e)); }

  try {
    var ss = ensureSheets_();
    linhas.push('planilha=' + ss.getName());
    linhas.push('planilha_id=' + ss.getId());
    var sheets = ss.getSheets();
    linhas.push('abas=' + sheets.map(function (s) { return s.getName(); }).join(','));
    for (var i = 0; i < COLLECTIONS.length; i++) {
      var c = COLLECTIONS[i];
      try { linhas.push('linhas_' + c + '=' + (readSheet_(c) || []).length); }
      catch (e2) { linhas.push('linhas_' + c + '=ERRO ' + ncErr_(e2)); }
    }
    var payload = ncBootstrapPayload_();
    var json = ncJson_(payload);
    linhas.push('bootstrap_ok=' + (payload && payload.ok));
    linhas.push('bootstrap_bytes=' + json.length);
    linhas.push('bootstrap_inicio=' + json.slice(0, 120));
  } catch (err) {
    linhas.push('ERRO_GERAL=' + ncErr_(err));
    if (err && err.stack) linhas.push('STACK=' + String(err.stack).slice(0, 400));
  }
  return linhas.join('\n');
}

/* ------------------------------------------------------------------ */
/* Web App                                                             */
/* ------------------------------------------------------------------ */

function doGet(e) {
  e = e || {};
  var p = e.parameter || {};

  if (p.action && (p.callback || p.json === '1')) {
    return ncRespond_(p.action, p.payload || '', p.callback || '');
  }

  try {
    var t = HtmlService.createTemplateFromFile('Index');
    return t.evaluate()
      .setTitle('Nosso Caixa')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput('<h1>Erro ao carregar a interface</h1><pre>' + ncErr_(err) + '</pre>');
  }
}

function doPost(e) {
  try {
    e = e || {};
    var body = {};

    // Formato 1: corpo JSON puro (postData.contents = '{"action":..., "payload":[...]}')
    if (e.postData && e.postData.contents && e.postData.contents.trim().charAt(0) === '{') {
      try { body = JSON.parse(e.postData.contents) || {}; }
      catch (err) { return ncEmit_({ ok: false, error: 'payload invalido: ' + ncErr_(err) }, ''); }
    } else {
      // Formato 2: formulário application/x-www-form-urlencoded
      // (ncFormPost no frontend envia action/payload como campos ocultos via POST,
      //  quando o payload estoura o limite de URL do GET/JSONP)
      var p = e.parameter || {};
      if (p.action) {
        body.action = p.action;
        body.payload = p.payload || '';
        body.callback = p.callback || '';
      }
    }

    if (!body.action) return ncEmit_({ ok: false, error: 'acao ausente' }, '');

    return ncRespond_(body.action, body.payload || '', body.callback || '');
  } catch (err) {
    return ncEmit_({ ok: false, error: ncErr_(err) }, '');
  }
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
