/**
 * INCOLORSHOP — DISTRIBUTOR ERP PORTAL (Google Sheets backend)
 * File map:
 *   Code.gs        → config, routing, includes, login/session, user preferences
 *   IMS.gs         → IMS sync (every 5 min) + IMS data API
 *   index.html     → shell that includes the HTML modules below
 *   Styles.html, Login.html, Layout.html, Scripts.html
 *   Page_Dashboard.html, Page_IMS.html, Page_Profile.html, Page_Appearance.html  ← one file per ERP page
 *
 * Sheets in THIS spreadsheet:
 *   LOGIN PAGE  → A: PERSON NAME | B: DIST ID | C: DIST PASSWORD | D: IMS DIST (YES/NO) | E: PARTY NAME (optional)
 *                 D1 = URL of the linked MIS MEETING sheet.
 *                 PARTY NAME = the distributor's header name on DISTRIBUTOR CLOSING BALANCE, if it differs from PERSON NAME.
 *   SETTINGS    → optional. A: KEY | B: VALUE  (COMPANY NAME, TAGLINE, DEFAULT THEME, LOW STOCK LIMIT)
 *   IMS_SNAPSHOT → created automatically
 */

var APP_VERSION = 'v1.5';

/* Module permissions: sidebar button shows only if the user's LOGIN PAGE column says YES.
   Add a module = add it here + a Page_X.html + a nav button with data-perm="x" in Layout.html */
var MODULES = {
  ims: ['IMS DIST', 'IMS']
  // orders: ['ORDERS'], reports: ['REPORTS']  ← future
};

var CONFIG = {
  companyName: 'INCOLORSHOP',
  tagline: 'Distributor ERP',
  defaultTheme: 'ruby',
  lowStockLimit: 5,          // stock at or below this = low
  tz: 'Asia/Kolkata',
  sessionHours: 8
};

function doGet() {
  var t = HtmlService.createTemplateFromFile('index');
  t.config = getConfig();
  return t.evaluate()
    .setTitle(t.config.companyName + ' · ERP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lets index.html pull in other HTML files: <?!= include('Login') ?> */
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

/* ---------- shared helpers ---------- */
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function norm(v) { return String(v || '').trim().toUpperCase().replace(/\s+/g, ' '); }
function num(v) { var n = Number(String(v == null ? '' : v).replace(/[,₹\s]/g, '')); return isNaN(n) ? 0 : n; }
function nowStr(fmt) { return Utilities.formatDate(new Date(), CONFIG.tz, fmt || 'dd-MMM-yyyy hh:mm a'); }

function getConfig() {
  var c = { companyName: CONFIG.companyName, tagline: CONFIG.tagline, defaultTheme: CONFIG.defaultTheme, lowStockLimit: CONFIG.lowStockLimit, version: APP_VERSION };
  var sh = ss().getSheetByName('SETTINGS');
  if (sh) sh.getDataRange().getDisplayValues().forEach(function (r) {
    var k = norm(r[0]), v = String(r[1] || '').trim();
    if (!v) return;
    if (k === 'COMPANY NAME') c.companyName = v;
    if (k === 'TAGLINE') c.tagline = v;
    if (k === 'DEFAULT THEME') c.defaultTheme = v.toLowerCase();
    if (k === 'LOW STOCK LIMIT') c.lowStockLimit = v;
  });
  return c;
}

/* ---------- users & sessions ---------- */
function loginSheet() {
  var sh = ss().getSheetByName('LOGIN PAGE');
  if (!sh) throw new Error('LOGIN PAGE sheet not found');
  return sh;
}

function findUser(id, pw) {
  var rows = loginSheet().getDataRange().getDisplayValues();
  var hdr = -1, col = {};
  for (var r = 0; r < rows.length; r++) {
    var up = rows[r].map(norm);
    if (up.indexOf('DIST ID') > -1) { hdr = r; up.forEach(function (h, i) { col[h] = i; }); break; }
  }
  if (hdr < 0) throw new Error('Header row with "DIST ID" not found in LOGIN PAGE');
  var cName = col['PERSON NAME'], cId = col['DIST ID'],
      cPw = col['DIST PASSWORD'] !== undefined ? col['DIST PASSWORD'] : col['DIST PASSOWRD'],
      cRole = col['ROLE'], cParty = col['PARTY NAME'];
  for (var i = hdr + 1; i < rows.length; i++) {
    var row = rows[i];
    if (norm(row[cId]) === norm(id) && String(row[cPw] || '').trim() === String(pw || '').trim()) {
      var perms = {};
      Object.keys(MODULES).forEach(function (m) {
        perms[m] = MODULES[m].some(function (h) { return col[h] !== undefined && norm(row[col[h]]) === 'YES'; });
      });
      return {
        name: String(row[cName] || '').trim(), id: String(row[cId]).trim(),
        perms: perms, ims: !!perms.ims,
        role: cRole !== undefined && row[cRole] ? String(row[cRole]).trim() : 'Distributor',
        party: cParty !== undefined && row[cParty] ? String(row[cParty]).trim() : ''
      };
    }
  }
  return null;
}

function makeToken(user) {
  var tok = Utilities.getUuid();
  CacheService.getScriptCache().put('tok_' + tok, JSON.stringify(user), CONFIG.sessionHours * 3600);
  return tok;
}
function userFromToken(tok) {
  var raw = CacheService.getScriptCache().get('tok_' + tok);
  if (!raw) throw new Error('Session expired. Please log in again.');
  return JSON.parse(raw);
}

/* ---------- public: auth ---------- */
function login(id, pw) {
  var user = findUser(id, pw);
  if (!user) return { ok: false, error: 'ID or password is wrong.' };
  return { ok: true, token: makeToken(user), user: user, prefs: getPrefs(user.id) };
}
function logout(token) { CacheService.getScriptCache().remove('tok_' + token); return { ok: true }; }

/* ---------- public: appearance preferences (theme / mode / font) ---------- */
function getPrefs(id) {
  var raw = PropertiesService.getScriptProperties().getProperty('prefs_' + norm(id));
  return raw ? JSON.parse(raw) : { theme: CONFIG.defaultTheme, mode: 'light', font: 'm' };
}
function savePrefs(token, prefs) {
  var user = userFromToken(token);
  var clean = {
    theme: /^(ruby|ocean|forest|violet|amber|slate|custom:#[0-9a-f]{6})$/i.test(prefs.theme || '') ? String(prefs.theme).toLowerCase() : CONFIG.defaultTheme,
    mode: prefs.mode === 'dark' ? 'dark' : 'light',
    font: ['s', 'm', 'l'].indexOf(prefs.font) > -1 ? prefs.font : 'm'
  };
  PropertiesService.getScriptProperties().setProperty('prefs_' + norm(user.id), JSON.stringify(clean));
  return { ok: true, prefs: clean };
}
