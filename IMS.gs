/**
 * IMS module — reads multiple tabs from the linked MIS MEETING spreadsheet into
 * a hidden IMS_SNAPSHOT tab.
 *
 * DISTRIBUTOR CLOSING BALANCE layout:
 *   Row 1 : party names above their column blocks
 *   Row 2 : CATEGORY | ITEM CODE | ITEM NAME | SKU | (party blocks: 2 cols each)
 *   Row 3 : (blanks) | Opening Date | Stock | Opening Date | Stock | ...
 *   Row 4+: item data
 *
 * Pagination: getIMSPage() reads from IMS_SNAPSHOT with SERVER-SIDE filters.
 *
 * FIELDS = ['party','cat','code','name','sku','max','rate','open','today','close','inward','outward','pct','lead']
 */

var SYNC_MINUTES = 5;
var SNAP = 'IMS_SNAPSHOT';

var IMS_TAB        = 'DISTRIBUTOR CLOSING BALANCE';
var FG_IMS_TAB     = 'FG IMS';
var RATE_TAB       = 'PURCHASE RATE';
var LEAD_TIME_TAB  = 'DISTRIBUTOR LEAD TIME';

var FIELDS = ['party','cat','code','name','sku','max','rate','open','today','close','inward','outward','pct','lead'];

function linkedBook() {
  var url = String(loginSheet().getRange('D1').getValue() || '').trim();
  if (!url) throw new Error('Put the MIS MEETING sheet URL in LOGIN PAGE cell D1.');
  return SpreadsheetApp.openByUrl(url);
}

function openSourceTab(book) {
  var tab = book.getSheetByName(IMS_TAB);
  if (!tab) throw new Error('Tab "' + IMS_TAB + '" not found in the linked sheet.');
  return tab;
}

/* ------------------------------------------------------------------
 * LOADERS
 * ------------------------------------------------------------------ */

function loadMaxLevels(book) {
  var map = {};
  var tab = book.getSheetByName(FG_IMS_TAB);
  if (!tab) return map;
  var rows = tab.getDataRange().getDisplayValues();
  if (!rows.length) return map;

  var H = rows[0].map(norm);
  var cCode = H.indexOf('ITEM CODE');
  var cSku  = H.indexOf('SHADE NO / SKU');
  if (cSku < 0) cSku = H.indexOf('SKU');
  var cMax  = H.indexOf('MAX LEVEL');

  if (cCode < 0 || cSku < 0 || cMax < 0) return map;

  for (var i = 1; i < rows.length; i++) {
    var code = norm(rows[i][cCode]);
    var sku  = norm(rows[i][cSku]);
    var max  = String(rows[i][cMax] || '').trim();
    if (!code || !sku) continue;
    var key = code + '|' + sku;
    if (map[key] === undefined) map[key] = max;
  }
  return map;
}

/** RATE from PURCHASE RATE — keyed by ITEM CODE (multiple code/price blocks scanned) */
function loadRates(book) {
  var map = {};
  var tab = book.getSheetByName(RATE_TAB);
  if (!tab) return map;
  var rows = tab.getDataRange().getDisplayValues();
  if (!rows.length) return map;

  var PAIRS = [
    { code: 2,  price: 4 },
    { code: 8,  price: 9 },
    { code: 12, price: 15 }
  ];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    for (var p = 0; p < PAIRS.length; p++) {
      var code  = norm(row[PAIRS[p].code]);
      var price = String(row[PAIRS[p].price] || '').trim();
      if (code && price && map[code] === undefined) map[code] = price;
    }
  }
  return map;
}

/** LEAD TIME — returns nested map: { partyNameNorm: { "ITEMNAME|SKU": value } } */
function loadLeadTimes(book) {
  var map = {};

  var url1 = String(loginSheet().getRange('D1').getValue() || '').trim();
  if (!url1) { Logger.log('❌ D1 URL missing'); return map; }

  var leadBook;
  try {
    leadBook = SpreadsheetApp.openByUrl(url1);
  } catch (e) {
    Logger.log('❌ D1 open failed: ' + e.message);
    return map;
  }
  Logger.log('Lead book: ' + leadBook.getName());

  var tab = leadBook.getSheetByName(LEAD_TIME_TAB);
  if (!tab) { Logger.log('❌ Tab "' + LEAD_TIME_TAB + '" not found in D1'); return map; }

  var rows = tab.getDataRange().getDisplayValues();
  if (rows.length < 3) return map;

  var partyRow = rows[0];
  var H = rows[1].map(norm);
  var cName = H.indexOf('ITEM NAME');
  var cSku  = H.indexOf('SKU');
  if (cName < 0 || cSku < 0) return map;

  var parties = [];
  for (var j = 0; j < H.length; j++) {
    if (H[j] !== 'LEAD TIME') continue;
    var name = String(partyRow[j] || '').trim();
    if (!name) continue;
    parties.push({ name: norm(name), col: j });
  }
  Logger.log('Lead parties detected: ' + parties.length);

  for (var i = 2; i < rows.length; i++) {
    var itemName = norm(rows[i][cName]);
    var sku      = norm(rows[i][cSku]);
    if (!itemName && !sku) continue;
    var key = itemName + '|' + sku;

    for (var k = 0; k < parties.length; k++) {
      var p = parties[k];
      var val = String(rows[i][p.col] || '').trim();
      if (!val) continue;
      if (!map[p.name]) map[p.name] = {};
      if (map[p.name][key] === undefined) map[p.name][key] = val;
    }
  }
  return map;
}
/* ------------------------------------------------------------------
 * MAIN PARSER
 * ------------------------------------------------------------------ */
function parseClosingBalance(rows, rates, maxLevels, leadTimes) {
  rates = rates || {};
  maxLevels = maxLevels || {};
  leadTimes = leadTimes || {};

  var hdr = -1;
  for (var r = 0; r < Math.min(rows.length, 15); r++) {
    var up = rows[r].map(norm);
    if (up.indexOf('ITEM CODE') > -1 && up.indexOf('SKU') > -1) { hdr = r; break; }
  }
  if (hdr < 0) throw new Error('Header row with ITEM CODE and SKU not found in "' + IMS_TAB + '".');

  var H = rows[hdr].map(norm);
  var cCat  = H.indexOf('CATEGORY');
  var cCode = H.indexOf('ITEM CODE');
  var cName = H.indexOf('ITEM NAME');
  var cSku  = H.indexOf('SKU');

  var out = [FIELDS];
  var metaOut = [['ALL', '', '']];

  // data starts 2 rows below header row
  for (var i = hdr + 2; i < rows.length; i++) {
    var row = rows[i];
    var code = String(row[cCode] || '').trim();
    var nm   = String(row[cName] || '').trim();
    if (!code && !nm) continue;

    var codeKey = norm(code);
    var skuRaw  = String(row[cSku] || '').trim();
    var skuKey  = norm(skuRaw);
    var joinKey = codeKey + '|' + skuKey;

    var max = maxLevels[joinKey] !== undefined ? maxLevels[joinKey] : '';
    var rate = rates[codeKey] !== undefined ? rates[codeKey] : '';

    // ONE row per item — party column stores 'ALL' (not used for filtering)
    // Lead will be resolved at read time from the user's party.
    out.push([
      'ALL',
      row[cCat] || '',
      code,
      nm,
      skuRaw,
      max,
      rate,
      '', '', '', '', '', '',
      ''   // lead — filled at read time
    ]);
  }

  return { rows: out, meta: metaOut };
}
/* ------------------------------------------------------------------
 * SYNC + TRIGGERS
 * ------------------------------------------------------------------ */
function syncIMS() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;
  try {
    var book = linkedBook();   // D1

        var rates        = loadRates(book);
    var maxLevels    = loadMaxLevels(book);
    var leadTimes    = loadLeadTimes(book);
    var stockBoth    = loadStockBoth();      // { closing,opening }
    var parsed = parseClosingBalance(
      openSourceTab(book).getDataRange().getDisplayValues(),
      rates, maxLevels, leadTimes
    );

    // ---- IMS_SNAPSHOT ----
    var sh = ss().getSheetByName(SNAP) || ss().insertSheet(SNAP);
    sh.clearContents();
    sh.getRange(1, 1, parsed.rows.length, FIELDS.length).setValues(parsed.rows);
    sh.hideSheet();

    // ---- IMS_LEAD (only ONE block) ----
    var leadSh = ss().getSheetByName('IMS_LEAD') || ss().insertSheet('IMS_LEAD');
    leadSh.clearContents();
    var leadRows = [['party', 'itemName', 'sku', 'lead']];
    Object.keys(leadTimes).forEach(function (party) {
      var sub = leadTimes[party];
      Object.keys(sub).forEach(function (k) {
        var parts = k.split('|');
        leadRows.push([party, parts[0] || '', parts[1] || '', sub[k]]);
      });
    });
    if (leadRows.length > 1) leadSh.getRange(1, 1, leadRows.length, 4).setValues(leadRows);
    leadSh.hideSheet();
    Logger.log('IMS_LEAD: ' + (leadRows.length - 1) + ' rows');

       // ---- IMS_CLOSING ----
    var csSh = ss().getSheetByName('IMS_CLOSING') || ss().insertSheet('IMS_CLOSING');
    csSh.clearContents();
    var csRows = [['party', 'itemName', 'sku', 'closingStock']];
    Object.keys(stockBoth.closing).forEach(function (k) {
      var parts = k.split('|');
      csRows.push([parts[0] || '', parts[1] || '', parts[2] || '', stockBoth.closing[k]]);
    });
    if (csRows.length > 1) csSh.getRange(1, 1, csRows.length, 4).setValues(csRows);
    csSh.hideSheet();
    Logger.log('IMS_CLOSING: ' + (csRows.length - 1) + ' rows');

    // ---- IMS_OPENING ----
    var opSh = ss().getSheetByName('IMS_OPENING') || ss().insertSheet('IMS_OPENING');
    opSh.clearContents();
    var opRows = [['party', 'itemName', 'sku', 'openingStock']];
    Object.keys(stockBoth.opening).forEach(function (k) {
      var parts = k.split('|');
      opRows.push([parts[0] || '', parts[1] || '', parts[2] || '', stockBoth.opening[k]]);
    });
    if (opRows.length > 1) opSh.getRange(1, 1, opRows.length, 4).setValues(opRows);
    opSh.hideSheet();
    Logger.log('IMS_OPENING: ' + (opRows.length - 1) + ' rows');

    var stamp = nowStr();
    var props = PropertiesService.getScriptProperties();
    props.setProperty('ims_synced', stamp);
    props.setProperty('ims_meta', JSON.stringify(parsed.meta));
  } finally {
    lock.releaseLock();
  }
}
function installTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncIMS') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('syncIMS').timeBased().everyMinutes(SYNC_MINUTES).create();
  syncIMS();
}

/* ------------------------------------------------------------------
 * FRONT-END ADAPTER
 * ------------------------------------------------------------------ */
function rowsToItems(rows) {
  return rows.map(function (r) {
    return {
      cat:      r[1] || '',
      code:     r[2] || '',
      name:     r[3] || '',
      sku:      r[4] || '',
      max:      num(r[5]),
      maxRaw:   r[5] || '',
      rate:     num(r[6]),
      rateRaw:  r[6] || '',
      open:     r[7] || '',
      today:    r[8] || '',
      close:    r[9] || '',
      inward:   r[10] || '',
      outward:  r[11] || '',
      pct:      r[12] || '',
      lead:     r[13] || ''
    };
  });
}

function partyKeyFor(user) {
  return 'ALL';
}

/* ------------------------------------------------------------------
 * PUBLIC API — paginated, with SERVER-SIDE filters
 * ------------------------------------------------------------------ */

/**
 * filters shape:
 *   {
 *     q:    "search string",
 *     cat:  { "LIPSTICK": 1, ... },
 *     code: { "IC027": 1, ... },
 *     name: { "...": 1 },
 *     sku:  { "801": 1, ... },
 *     rate: { op: "eq"|"gt"|"lt"|"bt"|"blank"|"nblank", v1, v2 }
 *   }
 */
function getIMSPage(token, page, pageSize, since, filters) {
  var user = userFromToken(token);
  if (!user.ims) return { ok: true, allowed: false, items: [] };

  page = Math.max(1, num(page) || 1);
  pageSize = Math.min(500, Math.max(10, num(pageSize) || 100));

  var pk = norm(user.party || user.name || user.id);
  var sh = ss().getSheetByName(SNAP);
  if (!sh) { syncIMS(); sh = ss().getSheetByName(SNAP); }

  var synced = PropertiesService.getScriptProperties().getProperty('ims_synced') || nowStr();
  var all = sh.getDataRange().getDisplayValues();

  // ---- load lead times for this user's party from IMS_LEAD sheet ----
  var partyLead = {};
  var leadSh = ss().getSheetByName('IMS_LEAD');
  if (leadSh) {
    var leadRows = leadSh.getDataRange().getDisplayValues();
    for (var li = 1; li < leadRows.length; li++) {
      if (norm(leadRows[li][0]) !== pk) continue;
      var k = norm(leadRows[li][1]) + '|' + norm(leadRows[li][2]);
      partyLead[k] = leadRows[li][3];
    }
  }
    // ---- load closing stock for this user's party from IMS_CLOSING sheet ----
  var partyClosing = {};
  var csSh = ss().getSheetByName('IMS_CLOSING');
  if (csSh) {
    var csRows = csSh.getDataRange().getDisplayValues();
    for (var ci = 1; ci < csRows.length; ci++) {
      if (norm(csRows[ci][0]) !== pk) continue;
      var ck = norm(csRows[ci][1]) + '|' + norm(csRows[ci][2]);
      partyClosing[ck] = csRows[ci][3];
    }
  }
    // ---- load opening stock for this user's party from IMS_OPENING sheet ----
  var partyOpening = {};
  var opSh = ss().getSheetByName('IMS_OPENING');
  if (opSh) {
    var opRows = opSh.getDataRange().getDisplayValues();
    for (var oi = 1; oi < opRows.length; oi++) {
      if (norm(opRows[oi][0]) !== pk) continue;
      var ok = norm(opRows[oi][1]) + '|' + norm(opRows[oi][2]);
      partyOpening[ok] = opRows[oi][3];
    }
  }

  // ---- filters ----
  filters = filters || {};
  var setMap = { cat: 1, code: 2, name: 3, sku: 4 };
  var q = (filters.q || '').toString().toLowerCase().trim();
  var rateF = filters.rate || null;

  var filtered = all.slice(1).filter(function (r) {
    if (q) {
      var hay = (String(r[1]) + ' ' + String(r[2]) + ' ' + String(r[3]) + ' ' + String(r[4])).toLowerCase();
      if (hay.indexOf(q) === -1) return false;
    }
    for (var key in setMap) {
      var set = filters[key];
      if (!set || typeof set !== 'object') continue;
      var val = String(r[setMap[key]] || '');
      if (!set[val]) return false;
    }
    if (rateF && rateF.op) {
      var raw = String(r[6] || '').trim();
      var nv = raw === '' ? null : Number(raw);
      var has = nv !== null && !isNaN(nv);
      if (rateF.op === 'blank') { if (has) return false; }
      else if (rateF.op === 'nblank') { if (!has) return false; }
      else if (rateF.op === 'eq') { if (!has || nv !== Number(rateF.v1)) return false; }
      else if (rateF.op === 'gt') { if (!has || nv <= Number(rateF.v1)) return false; }
      else if (rateF.op === 'lt') { if (!has || nv >= Number(rateF.v1)) return false; }
      else if (rateF.op === 'bt') {
        if (!has) return false;
        if (rateF.v1 != null && rateF.v1 !== '' && nv < Number(rateF.v1)) return false;
        if (rateF.v2 != null && rateF.v2 !== '' && nv > Number(rateF.v2)) return false;
      }
    }
    return true;
  });

    // ---- attach per-user lead time AND closing stock ----
  filtered = filtered.map(function (r) {
    var itemName = norm(r[3]);
    var sku      = norm(r[4]);
    var key = itemName + '|' + sku;

    var lead  = partyLead[key]  !== undefined ? partyLead[key]  : '';
    var close = partyClosing[key] !== undefined ? partyClosing[key] : '';
    var open  = partyOpening[key] !== undefined ? partyOpening[key] : '';

    var out = r.slice();
    out[7]  = open;    // OPENING STOCK (column index 7)
    out[9]  = close;   // closing stock (column index 9 in FIELDS)
    out[13] = lead;
    return out;
  });

  var total = filtered.length;
  var pages = Math.max(1, Math.ceil(total / pageSize));
  if (page > pages) page = pages;
  var start = (page - 1) * pageSize;
  var end = Math.min(filtered.length, start + pageSize);
  var slice = filtered.slice(start, end);

  return {
    ok: true,
    allowed: true,
    synced: synced,
    total: total,
    page: page,
    pages: pages,
    pageSize: pageSize,
    items: rowsToItems(slice),
    meta: { openDate: '', total: total }
  };
}

function getIMS(token, since) {
  var user = userFromToken(token);
  if (!user.ims) return { ok: true, allowed: false, items: [] };

  var sh = ss().getSheetByName(SNAP);
  if (!sh) { syncIMS(); sh = ss().getSheetByName(SNAP); }

  var synced = PropertiesService.getScriptProperties().getProperty('ims_synced') || nowStr();
  var all = sh.getDataRange().getDisplayValues();
  var items = [];
  for (var i = 1; i < all.length; i++) items.push(all[i]);

  return {
    ok: true,
    allowed: true,
    synced: synced,
    total: items.length,
    items: rowsToItems(items),
    meta: { openDate: '', total: items.length }
  };
}

function forceSync(token) {
  userFromToken(token);
  syncIMS();
  return getIMSPage(token, 1, 100, null, null);
}

/* ------------------------------------------------------------------
 * UNIQUE VALUES PER COLUMN (for search dropdown)
 * ------------------------------------------------------------------ */
function getColValues(token, column) {
  var user = userFromToken(token);
  if (!user.ims) return { ok: false, values: [] };

  var sh = ss().getSheetByName(SNAP);
  if (!sh) return { ok: false, values: [] };

  var colIndex = { cat: 1, code: 2, name: 3, sku: 4 }[column];
  if (colIndex === undefined) return { ok: false, values: [] };

  var rows = sh.getDataRange().getDisplayValues();
  var seen = {};
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var v = String(rows[i][colIndex] || '').trim();
    if (!v) continue;
    if (seen[v]) continue;
    seen[v] = 1;
    out.push(v);
  }
  out.sort();
  return { ok: true, values: out };
}

/* ------------------------------------------------------------------
 * ADMIN HELPERS
 * ------------------------------------------------------------------ */
function testParse() {
  var book = linkedBook();
  var rows = openSourceTab(book).getDataRange().getDisplayValues();
  var parsed = parseClosingBalance(
    rows, loadRates(book), loadMaxLevels(book), loadLeadTimes(book)
  );
  Logger.log('Header: ' + parsed.rows[0].join(' | '));
  if (parsed.rows[1]) Logger.log('Row 1: ' + parsed.rows[1].join(' | '));
  if (parsed.rows[2]) Logger.log('Row 2: ' + parsed.rows[2].join(' | '));
  Logger.log('Parties: ' + parsed.meta.map(function(m){return m[0]}).join(', '));
  Logger.log('Total item rows: ' + (parsed.rows.length - 1));
}

function listPartiesFound() {
  Logger.log('Party filter is disabled — all items shown for every user.');
}

function clearOldSnapshot() {
  var names = ['IMS_SNAPSHOT', 'IMS_LEAD', 'IMS_CLOSING', 'IMS_OPENING'];
  names.forEach(function (n) {
    var s = ss().getSheetByName(n);
    if (s) s.clearContents();
  });
  PropertiesService.getScriptProperties().deleteProperty('ims_synced');
  PropertiesService.getScriptProperties().deleteProperty('ims_meta');
  Logger.log('All sheets cleared.');
}
/** Loads BOTH closing (latest date) and opening (second-latest date) stock
 *  per party from D2 sheet → DISTRIBUTOR CLOSING BALANCE tab.
 *
 *  Rules:
 *    - Har party ke saare columns collect karo (jaha "Stock" header hai)
 *    - Date ke hisab se sort karo (newest first)
 *    - Latest = CLOSING STOCK
 *    - Second-latest = OPENING STOCK
 */
function loadStockBoth() {
  var result = { closing: {}, opening: {} };

  var url = String(loginSheet().getRange('D2').getValue() || '').trim();
  if (!url) { Logger.log('❌ D2 URL missing'); return result; }

  var closingBook;
  try {
    closingBook = SpreadsheetApp.openByUrl(url);
  } catch (e) {
    Logger.log('❌ D2 open failed: ' + e.message);
    return result;
  }
  Logger.log('D2 book: ' + closingBook.getName());

  var tab = closingBook.getSheetByName('DISTRIBUTOR CLOSING BALANCE');
  if (!tab) { Logger.log('❌ Tab not found'); return result; }

  var rows = tab.getDataRange().getDisplayValues();
  if (rows.length < 5) return result;

  // find header row (ITEM CODE + SKU)
  var hdr = -1;
  for (var r = 0; r < Math.min(rows.length, 15); r++) {
    var up = rows[r].map(norm);
    if (up.indexOf('ITEM CODE') > -1 && up.indexOf('SKU') > -1) { hdr = r; break; }
  }
  if (hdr < 0) return result;

  var H = rows[hdr].map(norm);
  var cName = H.indexOf('ITEM NAME');
  var cSku  = H.indexOf('SKU');
  if (cName < 0 || cSku < 0) return result;

  // ---- ACTUAL LAYOUT ----
  // Row (hdr - 1) = party names (row 1)
  // Row (hdr)     = CATEGORY|ITEM CODE|ITEM NAME|SKU + DATES in party columns  ← dates YAHAN
  // Row (hdr + 1) = totals in party columns
  // Row (hdr + 2) = "Stock" headers in party columns
  // Row (hdr + 3)+ = item data
  var partyRow = rows[hdr - 1] || [];
  var dateRow  = rows[hdr]     || [];     // ← DATES yahan
  var stockRow = rows[hdr + 2] || [];     // "Stock" header row
  // Data starts at hdr + 3

  // ---- collect parties with their columns ----
  var partiesMap = {};
  for (var j = 0; j < stockRow.length; j++) {
    if (norm(stockRow[j]) !== 'STOCK') continue;
    var pname = String(partyRow[j] || partyRow[j - 1] || '').trim();
    if (!pname) continue;

    // Date is in SAME column (col j) of dateRow
    var dstr = String(dateRow[j] || '').trim();
    var dt = parseDate(dstr);

    var pk = norm(pname);
    if (!partiesMap[pk]) partiesMap[pk] = { name: pname, columns: [] };
    partiesMap[pk].columns.push({ col: j, date: dstr, dateObj: dt });
  }

  // ---- sort columns by date DESC (latest first), tie → rightmost first ----
  Object.keys(partiesMap).forEach(function (pk) {
    var p = partiesMap[pk];
    p.columns.sort(function (a, b) {
      var da = a.dateObj ? a.dateObj.getTime() : 0;
      var db = b.dateObj ? b.dateObj.getTime() : 0;
      if (db !== da) return db - da;
      return b.col - a.col;
    });
  });

  // ---- log ----
  Object.keys(partiesMap).forEach(function (pk) {
    var cols = partiesMap[pk].columns;
    var log = pk + ' → ' + cols.length + ' cols: ';
    cols.forEach(function (c, i) {
      log += '[#' + i + ' col ' + (c.col + 1) + ' ' + c.date + '] ';
    });
    Logger.log(log);
  });

  // ---- read item data ----
  for (var i = hdr + 3; i < rows.length; i++) {
    var row = rows[i];
    var nm  = String(row[cName] || '').trim();
    var sku = String(row[cSku] || '').trim();
    if (!nm && !sku) continue;
    var itemKey = norm(nm) + '|' + norm(sku);

    Object.keys(partiesMap).forEach(function (pk) {
      var cols = partiesMap[pk].columns;
      if (cols[0]) {
        result.closing[pk + '|' + itemKey] = String(row[cols[0].col] || '').trim();
      }
      if (cols[1]) {
        result.opening[pk + '|' + itemKey] = String(row[cols[1].col] || '').trim();
      }
    });
  }

  Logger.log('Closing keys: ' + Object.keys(result.closing).length);
  Logger.log('Opening keys: ' + Object.keys(result.opening).length);

  ['RAJAT', 'VICKY BINDI'].forEach(function (name) {
    var p = partiesMap[name];
    if (p && p.columns.length > 0) {
      Logger.log(name + ' latest: col ' + (p.columns[0].col + 1) + ' date=' + p.columns[0].date);
      if (p.columns[1]) {
        Logger.log(name + ' second-latest: col ' + (p.columns[1].col + 1) + ' date=' + p.columns[1].date);
      }
    }
  });

  return result;
}
/** CLOSING STOCK from DISTRIBUTOR CLOSING BALANCE.
 *  Layout:
 *    hdr - 1 : party names (VICKY BINDI, ..., RAJAT, ...)
 *    hdr     : CATEGORY | ITEM CODE | ITEM NAME | SKU | ...
 *    hdr + 1 : dates + totals (30 July 2026 | 1229 | ...)
 *    hdr + 2 : Opening Date | Stock | Opening Date | Stock | ...
 *    hdr + 3+: item data
 *
 *  Returns { "PARTYNAME|ITEMNAME|SKU": closingStockValue }
 *  Picks the LATEST date's column per party.
 */
function loadClosingStock(book) {
  var map = {};

  var url = String(loginSheet().getRange('D2').getValue() || '').trim();
  if (!url) { Logger.log('❌ D2 URL missing'); return map; }

  var closingBook;
  try {
    closingBook = SpreadsheetApp.openByUrl(url);
  } catch (e) {
    Logger.log('❌ D2 open failed: ' + e.message);
    return map;
  }
  Logger.log('D2 book: ' + closingBook.getName());

  // Use DISTRIBUTOR CLOSING BALANCE tab (jisme RAJAT Col 46 hai)
  var tab = closingBook.getSheetByName('DISTRIBUTOR CLOSING BALANCE');
  if (!tab) { Logger.log('❌ Tab "DISTRIBUTOR CLOSING BALANCE" not found'); return map; }
  Logger.log('Tab "DISTRIBUTOR CLOSING BALANCE" found');

  var rows = tab.getDataRange().getDisplayValues();
  if (rows.length < 4) return map;

  // find header row (contains ITEM CODE + SKU)
  var hdr = -1;
  for (var r = 0; r < Math.min(rows.length, 15); r++) {
    var up = rows[r].map(norm);
    if (up.indexOf('ITEM CODE') > -1 && up.indexOf('SKU') > -1) { hdr = r; break; }
  }
  if (hdr < 0) { Logger.log('❌ Header not found'); return map; }
  Logger.log('Header row: ' + hdr);

  var H = rows[hdr].map(norm);
  var cName = H.indexOf('ITEM NAME');
  var cSku  = H.indexOf('SKU');
  if (cName < 0 || cSku < 0) return map;

  // layout:
  //   hdr - 1 = party names (row 1)
  //   hdr     = headers (row 2) — CATEGORY | ITEM CODE | ITEM NAME | SKU | ...
  //   hdr + 1 = dates (row 3)
  //   hdr + 2 = Opening Date | Stock (row 4)
  //   hdr + 3+ = item data
  var partyRow = rows[hdr - 1] || [];
  var stockRow = rows[hdr + 1] || [];      // dates + totals row
  var metaRow  = rows[hdr + 2] || [];      // Opening Date | Stock row

  // Detect parties: for every STOCK in metaRow, party name in partyRow same or prev col
  var parties = [];
  for (var j = 0; j < metaRow.length; j++) {
    if (norm(metaRow[j]) !== 'STOCK') continue;

    var pname = String(partyRow[j] || partyRow[j - 1] || '').trim();
    if (!pname) continue;

    parties.push({ name: pname, col: j });
  }

  Logger.log('Parties detected: ' + parties.length);
  parties.forEach(function (p) {
    Logger.log('  ' + p.name + ' → col ' + (p.col + 1));
  });

  // read data — starts at row 5 (idx hdr + 3)
  for (var i = hdr + 3; i < rows.length; i++) {
    var row = rows[i];
    var nm  = String(row[cName] || '').trim();
    var sku = String(row[cSku] || '').trim();
    if (!nm && !sku) continue;
    var itemKey = norm(nm) + '|' + norm(sku);

    parties.forEach(function (p) {
      var val = String(row[p.col] || '').trim();
      map[norm(p.name) + '|' + itemKey] = val;
    });
  }

  Logger.log('Total map keys: ' + Object.keys(map).length);
  return map;
}

/** Parse "31 Aug 2026" → Date */
function parseDate(str) {
  if (!str) return null;
  var s = String(str).trim();
  var m1 = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{2,4})$/);
  if (m1) {
    var months = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
    var mn = months[m1[2].toLowerCase().substring(0,3)];
    if (mn === undefined) return null;
    var yr = Number(m1[3]); if (yr < 100) yr += 2000;
    return new Date(yr, mn, Number(m1[1]));
  }
  var m2 = s.match(/^(\d{1,2})[-\/]([A-Za-z]+)[-\/](\d{2,4})$/);
  if (m2) {
    var months2 = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
    var mn2 = months2[m2[2].toLowerCase().substring(0,3)];
    if (mn2 === undefined) return null;
    var yr2 = Number(m2[3]); if (yr2 < 100) yr2 += 2000;
    return new Date(yr2, mn2, Number(m2[1]));
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function testClosingStockDirect() {
  try {
    Logger.log('=== START ===');
    var book = linkedBook();
    Logger.log('Book: ' + book.getName());

    Logger.log('=== loadClosingStock ===');
    var cs = loadClosingStock(book);
    var keys = Object.keys(cs);
    Logger.log('Total keys: ' + keys.length);

    if (keys.length > 0) {
      Logger.log('Sample keys: ' + keys.slice(0, 5).join(' | '));
      Logger.log('Sample values: ' + keys.slice(0, 5).map(function(k){return cs[k]}).join(', '));

      var rajatKeys = keys.filter(function(k){ return k.indexOf('RAJAT|') === 0; });
      Logger.log('RAJAT items: ' + rajatKeys.length);
      rajatKeys.slice(0, 5).forEach(function(k){
        Logger.log('  ' + k + ' = ' + cs[k]);
      });
    }

    Logger.log('');
    Logger.log('=== Writing IMS_CLOSING sheet ===');
    var csSh = ss().getSheetByName('IMS_CLOSING') || ss().insertSheet('IMS_CLOSING');
    csSh.clearContents();
    var csRows = [['party', 'itemName', 'sku', 'closingStock']];
    keys.forEach(function (k) {
      var parts = k.split('|');
      csRows.push([parts[0] || '', parts[1] || '', parts[2] || '', cs[k]]);
    });
    if (csRows.length > 1) {
      csSh.getRange(1, 1, csRows.length, 4).setValues(csRows);
      csSh.hideSheet();
      Logger.log('✓ IMS_CLOSING created: ' + (csRows.length - 1) + ' rows');
    } else {
      Logger.log('❌ No rows — loadClosingStock returned empty');
    }
    Logger.log('=== DONE ===');
  } catch (e) {
    Logger.log('❌ ERROR: ' + e.message);
    Logger.log(e.stack);
  }
}
function testStockBoth() {
  Logger.log('=== START ===');
  var stockBoth = loadStockBoth();
  Logger.log('Closing keys: ' + Object.keys(stockBoth.closing).length);
  Logger.log('Opening keys: ' + Object.keys(stockBoth.opening).length);

  var rajatClosing = Object.keys(stockBoth.closing).filter(function(k){ return k.indexOf('RAJAT|') === 0; });
  var rajatOpening = Object.keys(stockBoth.opening).filter(function(k){ return k.indexOf('RAJAT|') === 0; });

  Logger.log('');
  Logger.log('RAJAT closing items: ' + rajatClosing.length);
  rajatClosing.slice(0, 5).forEach(function(k){
    Logger.log('  ' + k + ' = ' + stockBoth.closing[k]);
  });

  Logger.log('');
  Logger.log('RAJAT opening items: ' + rajatOpening.length);
  rajatOpening.slice(0, 5).forEach(function(k){
    Logger.log('  ' + k + ' = ' + stockBoth.opening[k]);
  });
  Logger.log('=== DONE ===');
}
