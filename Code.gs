/**
 * Signature Scents — Google Sheets + Apps Script web app
 * Data tabs: App_Perfumes, App_Mixes (same spreadsheet)
 */
const SPREADSHEET_ID = '1VhzH0hJvAiKtT4xlyHesqBG76MvpJYKbHASr7uzbviY';
const TABS = {
  perfume: { name: 'App_Perfumes', fields: ['id','n','b','s','t','h','ba','c','o','q','i','bought','date','note','fav','rate'] },
  mix:     { name: 'App_Mixes',    fields: ['id','n','d','a','b','ra','pct','g','s','o','i','tip','bought','date','note','fav'] }
};
const NUMERIC = ['q','i','ra','pct','rate'];
const BOOLEAN = ['bought','fav'];

/** JSON API (GitHub Pages front-end). GET ?fn=getData, POST {fn,args} as text/plain */
const API = {
  getData: getData, savePerfume: savePerfume, saveMix: saveMix,
  deletePerfume: deletePerfume, deleteMix: deleteMix, setStatus: setStatus
};

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function api_(fn, args) {
  if (!Object.prototype.hasOwnProperty.call(API, fn)) return { ok: false, error: 'Invalid function' };
  try { return { ok: true, result: API[fn].apply(null, args || []) }; }
  catch (err) { return { ok: false, error: String(err && err.message || err) }; }
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Invalid request' }); }
  return json_(api_(req.fn, req.args));
}

function doGet(e) {
  if (e && e.parameter && e.parameter.fn) return json_(e.parameter.fn === 'getData' ? api_('getData') : { ok: false, error: 'Only getData is allowed via GET' });
  return page_();
}

function page_() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Mehboob's Signature Scent")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function sheet_(kind) {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TABS[kind].name);
  if (!sh) throw new Error('Tab not found: ' + TABS[kind].name);
  return sh;
}

function readAll_(kind) {
  const sh = sheet_(kind), f = TABS[kind].fields;
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, f.length).getValues();
  return rows.filter(r => String(r[0]).trim() !== '').map(r => {
    const o = {};
    f.forEach((k, i) => {
      let v = r[i];
      if (NUMERIC.includes(k)) v = Number(v) || 0;
      else if (BOOLEAN.includes(k)) v = v === true || String(v).toLowerCase() === 'true';
      else if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'd MMM yyyy');
      else v = String(v);
      o[k] = v;
    });
    return o;
  });
}

function getData() {
  return { perfumes: readAll_('perfume'), mixes: readAll_('mix') };
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function withLock_(fn) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function clean_(kind, obj) {
  const f = TABS[kind].fields;
  return f.map(k => {
    let v = obj[k];
    if (NUMERIC.includes(k)) return Number(v) || 0;
    if (BOOLEAN.includes(k)) return v === true;
    v = v == null ? '' : String(v).slice(0, 500);
    return /^[=+\-@]/.test(v) ? "'" + v : v; // guard against formula injection
  });
}

function save_(kind, obj) {
  return withLock_(() => {
    const sh = sheet_(kind);
    if (!obj.id) obj.id = (kind === 'mix' ? 'm' : 'p') + Date.now().toString(36);
    const row = findRow_(sh, obj.id);
    const values = [clean_(kind, obj)];
    if (row > 0) sh.getRange(row, 1, 1, values[0].length).setValues(values);
    else sh.appendRow(values[0]);
    return { id: obj.id, data: getData() };
  });
}

function delete_(kind, id) {
  return withLock_(() => {
    const sh = sheet_(kind);
    const row = findRow_(sh, id);
    if (row > 0) sh.deleteRow(row);
    return getData();
  });
}

function savePerfume(obj) { return save_('perfume', obj); }
function saveMix(obj)     { return save_('mix', obj); }
function deletePerfume(id){ return delete_('perfume', id); }
function deleteMix(id)    { return delete_('mix', id); }

/** updates bought / fav / note */
function setStatus(kind, id, field, value) {
  if (['bought','fav','note','rate'].indexOf(field) < 0) throw new Error('Invalid field');
  return withLock_(() => {
    const sh = sheet_(kind), f = TABS[kind].fields;
    const row = findRow_(sh, id);
    if (row < 0) throw new Error('Item not found');
    if (field === 'rate') {
      if (f.indexOf('rate') < 0) throw new Error('Invalid field');
      sh.getRange(row, f.indexOf('rate') + 1).setValue(Math.max(0, Math.min(5, Math.round(Number(value)) || 0)));
    } else if (field === 'note') {
      sh.getRange(row, f.indexOf('note') + 1).setValue(clean_(kind, { note: value })[f.indexOf('note')]);
    } else {
      sh.getRange(row, f.indexOf(field) + 1).setValue(value === true);
      if (field === 'bought') {
        const d = value === true ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy') : '';
        sh.getRange(row, f.indexOf('date') + 1).setValue(d);
      }
    }
    return getData();
  });
}
