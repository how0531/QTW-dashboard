/**
 * QTW 試算表同步接收端（Google Apps Script，綁在「策略管理」試算表上）
 *
 * 電腦上的 sheet_sync.py 每天 14:00 報表跑完後，把 C:\XS\data 的資料 POST 到這裡：
 *   - trips：一趟一列，用 trip_id 判斷是新增還是更新（mode=replace 時整份覆蓋）
 *   - positions：收盤持倉快照，同一天的資料整批換掉（mode=replace 時整份覆蓋）
 *   - daily：daily_pnl.csv，每次整份覆蓋
 * 每次同步在 _sync_log 留一列紀錄。
 *
 * 網頁（GitHub Pages）用 GET 向這裡拿資料：只回傳 WEB_TABS 這 4 個分頁，試算表本身不用公開。
 *
 * 每個交易日 15:00 左右 checkSync() 會檢查今天有沒有同步，沒有就寄提醒信（setupReminder() 裝一次排程）。
 *
 * 安裝：見 sync/README.md。第一次要先執行 setup() 產生密碼（SYNC_TOKEN）。
 */

var VERSION = '2026-11-26';

// 沒同步的提醒信寄給誰。空白＝寄給這個 Apps Script 的擁有者（就是部署它的 Google 帳號）
var REMIND_TO = '';

var TABS = { trips: '_trips', positions: '_positions', daily: '_daily', log: '_sync_log' };
// 用公式整欄展開的呈現頁：列數不夠時陣列公式會出錯，所以跟著原始資料一起加列。值是標題佔掉的列數
var VIEWS = { trips: { '交易明細': 1 }, daily: { '每日': 3 } };
// 網頁讀得到的分頁（公開）。其他分頁、公式、試算表擁有者都不會經過這裡
var WEB_TABS = { strategies: '策略對照', trips: '_trips', positions: '_positions', sync: '_sync_log' };
// 網頁只顯示 %（使用者 2026-11-05 決定）：金額和張數不給網頁，只留在試算表裡
var WEB_HIDE = ['shares', 'entry_amount', 'gross_pnl', 'fee', 'tax', 'net_pnl', 'unrealized_pnl'];

var COLS = {
  trips: ['trip_id', 'exit_date', 'strategy', 'strategy_log', 'code', 'name', 'side', 'shares', 'entry_date', 'entry_time',
    'entry_price', 'exit_time', 'exit_price', 'hold_minutes', 'trade_type', 'entry_amount', 'gross_pnl', 'fee', 'tax',
    'net_pnl', 'return_pct', 'synced_at'],
  positions: ['date', 'strategy', 'code', 'name', 'side', 'shares', 'avg_cost', 'entry_date', 'close_price',
    'unrealized_pnl', 'unrealized_pct', 'synced_at', 'flag'],
  daily: ['date', 'strategy', 'trips', 'wins', 'entry_amount', 'realized_pnl', 'realized_pct', 'unrealized_pnl',
    'open_positions', 'cum_realized_pnl', 'day_pct', 'cum_pct', 'gross_wins', 'gross_losses'],
};

// 這些欄位一律存成文字（例如 00878 這種代號才不會掉 0）
var TEXT = {
  trips: ['trip_id', 'strategy', 'strategy_log', 'code', 'name', 'side', 'trade_type', 'synced_at'],
  positions: ['strategy', 'code', 'name', 'side', 'synced_at', 'flag'],
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

/** 網頁讀資料用：回傳 WEB_TABS 每頁的內容（第一列是欄名），拿掉 WEB_HIDE 的金額欄。不用密碼，所以只放本來就要公開的東西 */
function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = { ok: true, service: 'qtw-sheet-sync', version: VERSION };
  Object.keys(WEB_TABS).forEach(function (k) { out[k] = webTable_(ss.getSheetByName(WEB_TABS[k])); });
  return json_(out);
}

// 數字照原值給（不會被顯示格式四捨五入）；日期、時間給畫面上顯示的字（不會因為時區換算差幾分鐘）
function webTable_(sh) {
  if (!sh) return [];
  var nr = sh.getLastRow(), nc = sh.getLastColumn();
  if (nr < 1 || nc < 1) return [];
  var range = sh.getRange(1, 1, nr, nc), values = range.getValues(), shown = range.getDisplayValues();
  var keep = values[0].map(function (h, j) { return WEB_HIDE.indexOf(String(h).trim()) < 0 ? j : -1; }).filter(function (j) { return j >= 0; });
  return values.map(function (row, i) {
    return keep.map(function (j) { var v = row[j]; return v instanceof Date ? shown[i][j] : v; });
  }).filter(function (row, i) { return i === 0 || row.some(function (v) { return v !== '' && v != null; }); });
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

  // 全量覆蓋的防呆：新資料比試算表現有的少一半以上（例如電腦資料夾不完整），先不寫，除非明確帶 force
  if (mode === 'replace' && sets.trips && !body.force) {
    var have = readTable_(sheet_(ss, 'trips'), 'trips', true).length;
    if ((sets.trips.length === 0 && have > 0) || (have >= 20 && sets.trips.length < have / 2)) {
      throw new Error('全量覆蓋會讓 trips 從 ' + have + ' 筆變成 ' + sets.trips.length + ' 筆，先不寫入。確定沒問題再加 --force 重跑');
    }
  }

  if (sets.trips) res.trips = writeTrips_(ss, sets.trips, mode);
  if (sets.positions) res.positions = writePositions_(ss, sets.positions, mode, body.positions_as_of);
  if (sets.daily) { res.daily = writeTable_(sheet_(ss, 'daily'), 'daily', sets.daily); growViews_(ss, 'daily', sets.daily.length); }

  var allCum = '';
  if (sets.daily) {
    var alls = sets.daily.filter(function (r) { return r.strategy === 'ALL'; });
    alls.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    if (alls.length) allCum = alls[alls.length - 1].cum_pct;
  }
  var log = ss.getSheetByName(TABS.log);
  if (log) {
    var dataDate = body.data_date || '';
    // appendRow 會自己加列，不會撞到分頁列數上限
    log.appendRow([now, dataDate, (body.source || 'sheet_sync.py') + (mode === 'replace' ? '（全量）' : ''),
      res.trips ? res.trips.added : '', res.trips ? res.trips.total : '', res.positions ? res.positions.written : '', allCum, body.note || '']);
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
      // 接受 9:08、09:08:03、090803、上午 9:08:03；整串都要是時間，不能是「2026-11-03 09:08」這種
      var ts = String(v).trim(), pm = /下午|PM/i.test(ts), am = /上午|AM/i.test(ts);
      ts = ts.replace(/上午|下午|AM|PM/ig, '').trim();
      var t = /^(\d{1,2}):?(\d{2})(?::?(\d{2}))?$/.exec(ts);
      var hh = t ? Number(t[1]) : NaN;
      if (t && pm && hh < 12) hh += 12;
      if (t && am && hh === 12) hh = 0;
      if (!t || hh > 23 || Number(t[2]) > 59 || Number(t[3] || 0) > 59) throw new Error(kind + ' 第 ' + (i + 1) + ' 筆 ' + c + ' 時間看不懂：' + v);
      v = ('0' + hh).slice(-2) + ':' + t[2] + ':' + (t[3] || '00');
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
  growViews_(ss, 'trips', existing.length);
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

// 後來才加在最後面的欄：試算表還沒有這欄標題就自己補上（舊資料這欄是空的）
var ADDED = { positions: ['flag'] };

function sheet_(ss, kind) {
  var sh = ss.getSheetByName(TABS[kind]);
  if (!sh) throw new Error('找不到分頁 ' + TABS[kind]);
  var add = ADDED[kind] || [], base = COLS[kind].length - add.length;
  if (sh.getMaxColumns && sh.getMaxColumns() < COLS[kind].length) sh.insertColumnsAfter(sh.getMaxColumns(), COLS[kind].length - sh.getMaxColumns());
  var cur = sh.getRange(1, 1, 1, COLS[kind].length).getValues()[0].map(function (v) { return String(v).trim(); });
  if (cur.slice(0, base).join(',') === COLS[kind].slice(0, base).join(',')) {
    add.forEach(function (c, i) { if (cur[base + i] === '') sh.getRange(1, base + i + 1).setValue(c); });
  }
  var head = sh.getRange(1, 1, 1, COLS[kind].length).getValues()[0].map(function (v) { return String(v).trim(); });
  if (head.join(',') !== COLS[kind].join(',')) throw new Error(TABS[kind] + ' 第一列欄名和程式不一樣，先不寫入：' + head.join(','));
  return sh;
}

function readTable_(sh, kind, countOnly) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  // 日期和時間讀「畫面上顯示的字」，不讀 Date 物件，才不會因為時區換算差幾分鐘
  var range = sh.getRange(2, 1, n, COLS[kind].length);
  var values = range.getValues();
  if (countOnly) return values.filter(function (row) { return row.some(function (v) { return v !== '' && v != null; }); });
  var shown = range.getDisplayValues();
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

// 分頁列數不夠就先加列（新試算表每頁只有 1000 列，超過時 getRange 會直接出錯）
function ensureRows_(sh, need) {
  var max = sh.getMaxRows();
  if (max < need) sh.insertRowsAfter(max, need - max + 500);
}

function growViews_(ss, kind, n) {
  var views = VIEWS[kind] || {};
  Object.keys(views).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) ensureRows_(sh, views[name] + n + 1);
  });
}

function writeTable_(sh, kind, rows) {
  var cols = COLS[kind];
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (!rows.length) return { written: 0 };
  ensureRows_(sh, rows.length + 1);
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

// ───────────── 沒同步就寄提醒 ─────────────
// 電腦每個交易日 14:00 跑報表後推送。15:00 左右檢查：今天是交易日、_sync_log 卻沒有今天的推送（或推了但資料日期不是今天）→ 寄信。
// 交易日：週一到週五，扣掉證交所公告的休市日（openapi.twse.com.tw，和報表程式同一個來源）。抓不到休市日時平日一律當交易日。
// 同一天只寄一次。

/** 裝排程：在編輯器裡執行一次。每天 15 點那一小時跑 checkSync（Google 的排程誤差約 15 分鐘） */
function setupReminder() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'checkSync') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('checkSync').timeBased().atHour(15).nearMinute(0).everyDays(1).inTimezone('Asia/Taipei').create();
  Logger.log('已設定：每天 15:00 左右檢查同步，沒同步會寄到 ' + remindTo_());
}

/** 排程呼叫。手動執行也可以（會照今天的狀況判斷，今天已經寄過就不再寄） */
function checkSync() {
  var now = new Date();
  var today = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
  var r = syncStatus_(SpreadsheetApp.getActiveSpreadsheet(), today);
  if (r.ok) { Logger.log(today + ' 已同步：' + r.detail); return r; }
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('REMIND_SENT') === today) { Logger.log(today + ' 今天已經寄過提醒'); return r; }
  var md = Number(today.slice(5, 7)) + '/' + Number(today.slice(8, 10));
  var subject = '【QTW同步提醒 ' + md + '】' + r.title;
  var body = [
    r.title + '。',
    '',
    r.detail,
    '',
    '可能的原因：電腦睡眠或關機、14:00 排程沒跑、報表程式出錯、網路斷線。',
    '可以先看：',
    '1. 電腦是不是醒著、XQ 有沒有開。',
    '2. C:\\XS\\report\\sheet_sync.log 最後幾行的錯誤訊息。',
    '3. 手動補跑：雙擊 C:\\XS\\tools\\每日報表.bat（會重跑報表並同步）；只想補同步就在 C:\\XS\\tools 執行 python sheet_sync.py。',
    '',
    '網站會顯示最後一次推送的資料，補跑後 5 分鐘內自動更新。',
    '（這封信由「策略管理」試算表的 Apps Script 自動寄出，每個交易日最多一封。）'
  ].join('\n');
  MailApp.sendEmail(remindTo_(), subject, body);
  props.setProperty('REMIND_SENT', today);
  Logger.log('已寄提醒：' + subject);
  r.sent = true;
  return r;
}

function remindTo_() {
  return REMIND_TO || Session.getEffectiveUser().getEmail();
}

// 今天該不該有同步、有沒有同步。回傳 { ok, title, detail }
function syncStatus_(ss, today) {
  var td = tradingDay_(today);
  if (!td.trading) return { ok: true, detail: '今天不是交易日（' + td.why + '），不用檢查' };
  var log = ss.getSheetByName(TABS.log);
  var rows = [];
  if (log && log.getLastRow() > 1) rows = log.getRange(2, 1, log.getLastRow() - 1, 2).getDisplayValues();
  var dayOf = function (s) { var m = /^(\d{4})\D(\d{1,2})\D(\d{1,2})/.exec(String(s || '').trim()); return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : ''; };
  var last = null, todays = [];
  rows.forEach(function (r) { var d = dayOf(r[0]); if (!d) return; if (!last || d >= dayOf(last[0])) last = r; if (d === today) todays.push(r); });
  var lastText = last ? '最後一次推送是 ' + String(last[0]).trim() + '（資料日期 ' + (dayOf(last[1]) || '空白') + '）。' : '_sync_log 沒有任何推送紀錄。';
  if (!todays.length) return { ok: false, title: '今天還沒有同步到試算表', detail: lastText };
  var fresh = todays.filter(function (r) { return dayOf(r[1]) === today; });
  if (!fresh.length) return { ok: false, title: '今天有推送，但資料停在 ' + dayOf(todays[todays.length - 1][1]), detail: '今天 ' + todays.length + ' 次推送的資料日期都不是今天，報表可能沒產生今天的檔案。' + lastText };
  return { ok: true, detail: '今天推送 ' + todays.length + ' 次，最後一次 ' + String(todays[todays.length - 1][0]).trim() };
}

// 證交所休市日，快取在 Script Properties（每年一份，30 天更新一次）
function tradingDay_(iso) {
  var wd = new Date(iso + 'T12:00:00Z').getUTCDay();
  if (wd === 0 || wd === 6) return { trading: false, why: '週末' };
  var hol = holidays_(iso.slice(0, 4));
  if (hol && hol[iso]) return { trading: false, why: hol[iso] };
  return { trading: true };
}

function holidays_(year) {
  var props = PropertiesService.getScriptProperties(), key = 'TWSE_HOLIDAYS_' + year;
  var cached = null;
  try { cached = JSON.parse(props.getProperty(key) || 'null'); } catch (e) { cached = null; }
  if (cached && Date.now() - cached.at < 30 * 864e5) return cached.days;
  try {
    var res = UrlFetchApp.fetch('https://openapi.twse.com.tw/v1/holidaySchedule/holidaySchedule', { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
    var days = {};
    JSON.parse(res.getContentText()).forEach(function (r) {
      var m = /^(\d{2,3})(\d{2})(\d{2})$/.exec(String(r.Date || '').trim());
      if (!m) return;
      var text = String(r.Name || '') + ' ' + String(r.Description || '');
      if (/開始交易|最後交易/.test(text)) return;   // 清單裡也有開盤日，例如「春節前最後交易日」
      days[(Number(m[1]) + 1911) + '-' + m[2] + '-' + m[3]] = String(r.Name || '休市').trim();
    });
    props.setProperty(key, JSON.stringify({ at: Date.now(), days: days }));
    return days;
  } catch (e) {
    Logger.log('抓不到證交所休市日（' + e.message + '），' + (cached ? '用舊的快取' : '平日一律當交易日'));
    return cached ? cached.days : null;
  }
}
