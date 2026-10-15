/**
 * =====================================================================
 *  HUB — private launcher for your apps · Apps Script backend (Code.gs)   API_VERSION 1.0-hub
 * =====================================================================
 *  SETUP: new Google Sheet → Extensions → Apps Script → paste this file →
 *   1. change HUB_PIN below  2. run setupSheets once  3. Deploy → Web app (Execute as Me, Anyone) → URL into site/index.html (API_URL).
 *  Tabs:  Cards  (Title, Note, Icon, URL, Group, Sort, Status)   Groups (Group, Icon, Sort — optional, for section order / icon)
 *  Add a row = new card. Status નિષ્ક્રિય hides it. Cards are sent only after the PIN is correct.
 * =====================================================================
 */
var HUB_PIN = '1234';                       // ← CHANGE THIS (4–12 characters)

var SHEET_CARDS = 'Cards', SHEET_GROUPS = 'Groups';
var HEADERS = { Cards: ['Title', 'Note', 'Icon', 'URL', 'Group', 'Sort', 'Status'], Groups: ['Group', 'Icon', 'Sort'] };
var SESSION_DAYS = 90, MAX_ATTEMPTS = 5, LOCKOUT_SECONDS = 900;
var API_VERSION = '1.0-hub';

function nowMs_() { return Date.now(); }
function doGet() { return out_({ success: true, version: API_VERSION }); }

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!body || typeof body !== 'object' || Array.isArray(body)) return fail_('bad_request', 'અમાન્ય વિનંતી.');
    if (body.action === 'login') return login_(body);
    if (body.action === 'cards') { checkToken_(body.token); return ok_(cards_()); }
    return fail_('bad_request', 'અજાણી વિનંતી.');
  } catch (err) {
    if (err && err.code) return fail_(err.code, err.message);
    console.error(err && err.stack ? err.stack : err);
    return fail_('server_error', 'સર્વર ભૂલ. ફરી પ્રયત્ન કરો.');
  }
}

function login_(body) {
  var cache = CacheService.getScriptCache(), key = 'fail_hub';
  if (Number(cache.get(key) || 0) >= MAX_ATTEMPTS) return fail_('too_many_attempts', 'ઘણા ખોટા પ્રયત્નો. ૧૫ મિનિટ પછી પ્રયત્ન કરો.');
  if (!safeEqual_(String(body.pin == null ? '' : body.pin), String(HUB_PIN))) {
    cache.put(key, String(Number(cache.get(key) || 0) + 1), LOCKOUT_SECONDS);
    return fail_('invalid_credentials', 'ખોટો PIN.');
  }
  cache.remove(key);
  var exp = nowMs_() + SESSION_DAYS * 86400000;
  return ok_({ token: exp + '.' + sign_('hub.' + exp + '.' + HUB_PIN), expires: exp });
}
/** the signature includes HUB_PIN, so changing the PIN logs every phone out */
function checkToken_(token) {
  var bad = { code: 'session_expired', message: 'સત્ર પૂરું થયું. ફરી PIN નાખો.' };
  token = String(token || ''); var dot = token.indexOf('.');
  if (dot < 1) throw bad;
  var exp = Number(token.slice(0, dot));
  if (!isFinite(exp) || exp < nowMs_() || !safeEqual_(sign_('hub.' + exp + '.' + HUB_PIN), token.slice(dot + 1))) throw bad;
}

function str_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }
/** only real https links are ever sent (never javascript: etc.) */
function cleanUrl_(v) { var u = str_(v, 300); return /^https:\/\/[^\s]+$/i.test(u) ? u : ''; }
function isInactive_(v) { var s = String(v == null ? '' : v).trim().toLowerCase(); return s === 'inactive' || s === 'નિષ્ક્રિય'; }

function table_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw { code: 'server_error', message: 'ટેબ "' + name + '" નથી. setupSheets ચલાવો.' };
  var need = HEADERS[name], width = Math.max(sh.getLastColumn(), need.length);
  var head = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); }), idx = {};
  head.forEach(function (h, i) { if (h) idx[h] = i; });
  need.forEach(function (h) { if (idx[h] === undefined) throw { code: 'server_error', message: 'કોલમ "' + h + '" ' + name + ' માં નથી.' }; });
  var last = sh.getLastRow();
  return { sh: sh, idx: idx, width: width, rows: last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [] };
}

/** groups in their Sort order (Groups tab), then groups only named on cards (alphabetical); each with its sorted cards */
function cards_() {
  var c = table_(SHEET_CARDS), g = table_(SHEET_GROUPS), meta = {}, order = [];
  g.rows.forEach(function (r) {
    var n = str_(r[g.idx.Group], 60); if (!n) return;
    meta[n] = { icon: str_(r[g.idx.Icon], 8), sort: Number(r[g.idx.Sort]) || 0 };
  });
  var by = {};
  c.rows.forEach(function (r) {
    var title = str_(r[c.idx.Title], 80), url = cleanUrl_(r[c.idx.URL]);
    if (!title || !url || isInactive_(r[c.idx.Status])) return;
    var grp = str_(r[c.idx.Group], 60) || 'અન્ય';
    (by[grp] = by[grp] || []).push({ title: title, note: str_(r[c.idx.Note], 120), icon: str_(r[c.idx.Icon], 8) || '🔗', url: url, sort: Number(r[c.idx.Sort]) || 0 });
  });
  var names = Object.keys(by);
  names.sort(function (a, b) {
    var ma = meta[a], mb = meta[b];
    if (ma && mb) return ma.sort - mb.sort || (a < b ? -1 : 1);
    if (ma) return -1; if (mb) return 1;
    return a < b ? -1 : 1;
  });
  return { groups: names.map(function (n) {
    by[n].sort(function (a, b) { return a.sort - b.sort || (a.title < b.title ? -1 : 1); });
    return { name: n, icon: meta[n] ? meta[n].icon : '', cards: by[n].map(function (x) { return { title: x.title, note: x.note, icon: x.icon, url: x.url }; }) };
  }) };
}

function sign_(m) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(m, secret_())); }
function secret_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('HUB_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('HUB_SECRET', s); }
  return s;
}
function safeEqual_(a, b) {
  a = String(a); b = String(b);
  var d = a.length ^ b.length, n = Math.max(a.length, b.length);
  for (var i = 0; i < n; i++) d |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return d === 0;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ok_(d) { d.success = true; return out_(d); }
function fail_(code, message) { return out_({ success: false, error: code, message: message }); }

/** Creates both tabs + example rows (only when empty). Safe to re-run. */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (String(HUB_PIN).length < 4 || String(HUB_PIN).length > 12) throw new Error('HUB_PIN must be 4–12 characters.');
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var have = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    HEADERS[name].forEach(function (h) { if (have.indexOf(h) < 0) { sh.getRange(1, have.length + 1).setValue(h); have.push(h); } });
    sh.getRange(1, 1, 1, have.length).setFontWeight('bold').setBackground('#FFE0B2'); sh.setFrozenRows(1);
  });
  var first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  var c = table_(SHEET_CARDS);
  c.sh.getRange(2, c.idx.Status + 1, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['સક્રિય', 'નિષ્ક્રિય'], true).setAllowInvalid(false).build());
  if (!c.rows.length) {
    var B = 'https://bhaveshpp.github.io/';
    var ex = [
      ['ટોકન', 'દુકાનના ટોકન', '🧾', B + 'token/', 'દુકાન', 10], ['દુકાન હિસાબ', 'રોજનો હિસાબ', '🏪', B + 'hisab/', 'દુકાન', 20],
      ['Finance', 'અંગત પૈસા', '💰', B + 'finance/', 'ઘર', 10], ['ભાડું', 'મકાન અને ભાડુઆત', '🏠', B + 'rent/', 'ઘર', 20],
      ['અભરામપરા મંડળ', 'વ્યાસ પરિવાર', '🙏', B + 'mandal/', 'ગામ', 10]];
    ex.forEach(function (d, i) {
      var row = []; for (var k = 0; k < c.width; k++) row.push('');
      row[c.idx.Title] = d[0]; row[c.idx.Note] = d[1]; row[c.idx.Icon] = d[2]; row[c.idx.URL] = d[3]; row[c.idx.Group] = d[4]; row[c.idx.Sort] = d[5]; row[c.idx.Status] = 'સક્રિય';
      c.sh.getRange(i + 2, 1, 1, c.width).setValues([row]);
    });
  }
  var g = table_(SHEET_GROUPS);
  if (!g.rows.length) {
    [['દુકાન', '🏪', 10], ['ઘર', '🏠', 20], ['ગામ', '🙏', 30]].forEach(function (d, i) { g.sh.getRange(i + 2, 1, 1, 3).setValues([d]); });
  }
  SpreadsheetApp.flush();
}
