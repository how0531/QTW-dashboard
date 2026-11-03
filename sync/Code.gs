/**
 * QTW 試算表同步接收端（Google Apps Script，綁在「策略管理」試算表上）
 *
 * 電腦上的 sheet_sync.py 每天 14:00 報表跑完後，把 C:\XS\data 的資料 POST 到這裡：
 *   - trips：一趟一列，用 trip_id 判斷是新增還是更新（mode=replace 時整份覆蓋）
 *   - positions：收盤持倉快照，同一天的資料整批換掉（mode=replace 時整份覆蓋）
 *   - daily：daily_pnl.csv，每次整份覆蓋
 * 每次同步在 _sync_log 留一列紀錄。
 *
 * 安裝：見 sync/README.md。第一次要先執行 setup() 產生密碼（SYNC_TOKEN）。
 */

var VERSION = '2026-11-03';

var TABS = { trips: '_trips', positions: '_positions', daily: '_daily', log: '_sync_log' };

var COLS = {
  trips: ['trip_id', 'exit_date', 'strategy', 'strategy_log', 'code', 'name', 'side', 'shares', 'entry_date', 'entry_time',
    'entry_price', 'exit_time', 'exit_price', 'hold_minutes', 'trade_type', 'entry_amount', 'gross_pnl', 'fee', 'tax',
    'net_pnl', 'return_pct', 'synced_at'],
  positions: ['date', 'strategy', 'code', 'name', 'side', 'shares', 'avg_cost', 'entry_date', 'close_price',
    'unrealized_pnl', 'unrealized_pct', 'synced_at'],
  daily: ['date', 'strategy', 'trips', 'wins', 'entry_amount', 'realized_pnl', 'realized_pct', 'unrealized_pnl',
    'open_positions', 'cum_realized_pnl', 'day_pct', 'cum_pct', 'gross_wins', 'gross_losses'],
};

// 這些欄位一律存成文字（例如 00878 這種代號才不會掉 0）
var TEXT = {
  trips: ['trip_id', 'strategy', 'strategy_log', 'code', 'name', 'side', 'trade_type', 'synced_at'],
  positions: ['strategy', 'code', 'name', 'side', 'synced_at'],
  daily: ['strategy'],
};
var DATE = { trips: ['exit_date', 'entry_date'], positions: ['date', 'entry_date'], daily: ['date'] };
var TIME = { trips: ['entry_time', 'exit_time'], positions: [], daily: [] };
var REQUIRED = { trips: ['trip_id', 'exit_date', 'strategy'], positions: ['date', 'strategy', 'code'], daily: ['date', 'strategy'] };

/** 第一次安裝時在編輯器裡執行一次：產生 SYNC_TOKEN，抄到電腦的 sheet_sync.json */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var t = props.getProperty('SYNC_TOKEN');
  if (!t) {
    t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('SYNC_TOKEN', t);
  }
  Logger.log('SYNC_TOKEN（抄到電腦 C:\\XS\\tools\\sheet_sync.json 的 token，不要給別人）：' + t);
  return t;
}

function doGet() {
  return json_({ ok: true, service: 'qtw-sheet-sync', version: VERSION });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e && e.postData ? e.postData.contents : '');
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  var token = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
  if (!token || !body || body.token !== token) return json_({ ok: false, error: 'unauthorized' });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return json_({ ok: false, error: 'busy' });
  try {
    return json_(applySync_(SpreadsheetApp.getActiveSpreadsheet(), body));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function applySync_(ss, body) {
  var tz = ss.getSpreadsheetTimeZone ? ss.getSpreadsheetTimeZone() : 'Asia/Taipei';
  var now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var mode = body.mode === 'replace' ? 'replace' : 'upsert';
  var res = { ok: true, mode: mode, version: VERSION };

  // 先全部檢查過，有一筆不對就整批不寫，避免寫到一半
  var sets = {};
  ['trips', 'positions', 'daily'].forEach(function (kind) {
    if (body[kind] == null) return;
    if (!Array.isArray(body[kind])) throw new Error(kind + ' 不是陣列');
    sets[kind] = body[kind].map(function (r, i) { return normalize_(kind, r, i, now); });
  });
  if (!sets.trips && !sets.positions && !sets.daily) throw new Error('沒有資料');

  if (sets.trips) res.trips = writeTrips_(ss, sets.trips, mode);
  if (sets.positions) res.positions = writePositions_(ss, sets.positions, mode, body.positions_as_of);
  if (sets.daily) res.daily = writeTable_(sheet_(ss, 'daily'), 'daily', sets.daily);

  var allCum = '';
  if (sets.daily) {
    var alls = sets.daily.filter(function (r) { return r.strategy === 'ALL'; });
    alls.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    if (alls.length) allCum = alls[alls.length - 1].cum_pct;
  }
  var log = ss.getSheetByName(TABS.log);
  if (log) {
    var dataDate = body.data_date || '';
    log.getRange(log.getLastRow() + 1, 1, 1, 8).setValues([[now, dataDate, (body.source || 'sheet_sync.py') + (mode === 'replace' ? '（全量）' : ''),
      res.trips ? res.trips.added : '', res.trips ? res.trips.total : '', res.positions ? res.positions.written : '', allCum, body.note || '']]);
  }
  return res;
}

function normalize_(kind, r, i, now) {
  if (!r || typeof r !== 'object') throw new Error(kind + ' 第 ' + (i + 1) + ' 筆格式不對');
  REQUIRED[kind].forEach(function (c) {
    if (r[c] == null || String(r[c]).trim() === '') throw new Error(kind + ' 第 ' + (i + 1) + ' 筆缺少 ' + c);
  });
  var o = {};
  COLS[kind].forEach(function (c) {
    var v = r[c];
    if (c === 'synced_at') v = now;
    if (v == null) v = '';
    if (DATE[kind].indexOf(c) >= 0 && v !== '') {
      var m = /^(\d{4})\D?(\d{1,2})\D?(\d{1,2})/.exec(String(v));
      if (!m) throw new Error(kind + ' 第 ' + (i + 1) + ' 筆 ' + c + ' 日期看不懂：' + v);
      v = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    } else if (TIME[kind].indexOf(c) >= 0 && v !== '') {
      var t = /^(\d{1,2}):?(\d{2}):?(\d{2})?/.exec(String(v));
      if (!t) throw new Error(kind + ' 第 ' + (i + 1) + ' 筆 ' + c + ' 時間看不懂：' + v);
      v = ('0' + t[1]).slice(-2) + ':' + t[2] + ':' + (t[3] || '00');
    } else if (TEXT[kind].indexOf(c) < 0 && c !== 'synced_at' && v !== '') {
      var n = Number(String(v).replace(/,/g, ''));
      if (isNaN(n)) throw new Error(kind + ' 第 ' + (i + 1) + ' 筆 ' + c + ' 不是數字：' + v);
      v = n;
    } else {
      v = String(v);
    }
    o[c] = v;
  });
  return o;
}

function writeTrips_(ss, rows, mode) {
  var sh = sheet_(ss, 'trips');
  var existing = mode === 'replace' ? [] : readTable_(sh, 'trips');
  var index = {};
  existing.forEach(function (r, i) { index[r.trip_id] = i; });
  var added = 0, updated = 0;
  rows.forEach(function (r) {
    if (index.hasOwnProperty(r.trip_id)) { existing[index[r.trip_id]] = r; updated++; }
    else { index[r.trip_id] = existing.length; existing.push(r); added++; }
  });
  existing.sort(function (a, b) {
    var ka = a.exit_date + ' ' + a.exit_time + ' ' + a.trip_id, kb = b.exit_date + ' ' + b.exit_time + ' ' + b.trip_id;
    return ka < kb ? -1 : ka > kb ? 1 : 0;
  });
  writeTable_(sh, 'trips', existing);
  return { added: added, updated: updated, total: existing.length };
}

function writePositions_(ss, rows, mode, asOf) {
  var sh = sheet_(ss, 'positions');
  var dates = {};
  rows.forEach(function (r) { dates[r.date] = true; });
  var keep = mode === 'replace' ? [] : readTable_(sh, 'positions').filter(function (r) { return !dates[r.date]; });
  // 電腦說「這一天收盤沒有持倉」：比這天舊的快照都清掉，網頁和總覽才不會停在前一天的持倉
  var cleared = 0;
  var m = asOf ? /^(\d{4})\D?(\d{1,2})\D?(\d{1,2})/.exec(String(asOf)) : null;
  if (m) {
    asOf = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    if (!dates[asOf]) {
      var newer = keep.concat(rows).some(function (r) { return r.date >= asOf; });
      if (!newer) { cleared = keep.length; keep = []; rows = rows.filter(function (r) { return r.date >= asOf; }); }
    }
  }
  var all = keep.concat(rows);
  all.sort(function (a, b) { var ka = a.date + a.strategy + a.code, kb = b.date + b.strategy + b.code; return ka < kb ? -1 : ka > kb ? 1 : 0; });
  writeTable_(sh, 'positions', all);
  return { written: rows.length, dates: Object.keys(dates).sort(), total: all.length, cleared: cleared, as_of: asOf || '' };
}

function sheet_(ss, kind) {
  var sh = ss.getSheetByName(TABS[kind]);
  if (!sh) throw new Error('找不到分頁 ' + TABS[kind]);
  var head = sh.getRange(1, 1, 1, COLS[kind].length).getValues()[0].map(function (v) { return String(v).trim(); });
  if (head.join(',') !== COLS[kind].join(',')) throw new Error(TABS[kind] + ' 第一列欄名和程式不一樣，先不寫入：' + head.join(','));
  return sh;
}

function readTable_(sh, kind) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  // 日期和時間讀「畫面上顯示的字」，不讀 Date 物件，才不會因為時區換算差幾分鐘
  var range = sh.getRange(2, 1, n, COLS[kind].length);
  var values = range.getValues(), shown = range.getDisplayValues();
  var out = [];
  values.forEach(function (row, ri) {
    if (!row.some(function (v) { return v !== '' && v != null; })) return;
    var o = {};
    COLS[kind].forEach(function (c, i) {
      var v = row[i];
      if (DATE[kind].indexOf(c) >= 0 || TIME[kind].indexOf(c) >= 0) v = shown[ri][i];
      if (TEXT[kind].indexOf(c) >= 0 && v !== '') v = String(v);
      o[c] = v;
    });
    out.push(normalize_(kind, o, ri, o.synced_at));
  });
  return out;
}

function writeTable_(sh, kind, rows) {
  var cols = COLS[kind];
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (!rows.length) return { written: 0 };
  var range = sh.getRange(2, 1, rows.length, cols.length);
  // 文字欄先設成純文字，日期、時間給固定格式，其他照試算表原本的格式
  cols.forEach(function (c, i) {
    var col = sh.getRange(2, i + 1, rows.length, 1);
    if (TEXT[kind].indexOf(c) >= 0) col.setNumberFormat('@');
    else if (DATE[kind].indexOf(c) >= 0) col.setNumberFormat('yyyy-mm-dd');
    else if (TIME[kind].indexOf(c) >= 0) col.setNumberFormat('hh:mm:ss');
  });
  range.setValues(rows.map(function (r) { return cols.map(function (c) { return r[c]; }); }));
  return { written: rows.length };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
