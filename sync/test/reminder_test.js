// checkSync()（15:00 沒同步就寄提醒）的單元測試：假的試算表、信箱、證交所 API、時間。用法：node reminder_test.js
const fs = require('fs'), path = require('path'), vm = require('vm');
const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
let fails = 0;
const check = (name, ok, detail) => { console.log((ok ? 'PASS ' : 'FAIL ') + name + (ok || detail === undefined ? '' : '  (' + JSON.stringify(detail) + ')')); if (!ok) fails++; };

function make({ now, log, holidays, fetchFails }) {
  const sent = [], props = {}, fetched = [];
  const sheet = { getLastRow: () => log.length + 1, getRange: (r, c, nr, nc) => ({ getDisplayValues: () => log.slice(r - 2, r - 2 + nr).map((row) => row.slice(c - 1, c - 1 + nc)) }) };
  const RealDate = Date;
  class FakeDate extends RealDate { constructor(...a) { super(...(a.length ? a : [now])); } static now() { return new RealDate(now).getTime(); } }
  const ctx = {
    Date: FakeDate, JSON, Math, Number, String, Object, Array, RegExp, Error,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: (n) => (n === '_sync_log' ? sheet : null) }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    MailApp: { sendEmail: (to, subject, body) => sent.push({ to, subject, body }) },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) },
    UrlFetchApp: { fetch: (u) => { fetched.push(u); if (fetchFails) throw new Error('down'); return { getResponseCode: () => 200, getContentText: () => JSON.stringify(holidays || []) }; } },
    Utilities: { formatDate: (d, tz, f) => new Intl.DateTimeFormat('sv-SE', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d) },
    Logger: { log: () => {} },
  };
  vm.createContext(ctx);
  vm.runInContext(code, ctx);
  return { ctx, sent, props, fetched };
}
const HOL = [{ Name: '中華民國開國紀念日', Date: '1150101', Description: '依規定放假1日' }, { Name: '國曆新年開始交易日', Date: '1150102', Description: '國曆新年開始交易' }, { Name: '測試休市', Date: '1151127', Description: '市場無交易，僅辦理結算交割作業' }];
const at = (iso) => iso + 'T15:00:00+08:00';
const okLog = [['2026-11-24 14:00', '2026-11-24'], ['2026-11-25 14:00', '2026-11-25']];

// 1. 今天有同步：不寄
let t = make({ now: at('2026-11-25'), log: okLog, holidays: HOL });
let r = t.ctx.checkSync();
check('今天有同步就不寄', r.ok && !t.sent.length, r);

// 2. 今天沒同步：寄一封，主旨有日期，信裡有最後一次推送
t = make({ now: at('2026-11-25'), log: okLog.slice(0, 1), holidays: HOL });
r = t.ctx.checkSync();
check('今天沒同步就寄提醒', t.sent.length === 1 && t.sent[0].subject === '【QTW同步提醒 11/25】今天還沒有同步到試算表' && /2026-11-24 14:00/.test(t.sent[0].body), t.sent);
check('沒設 REMIND_TO 時寄給擁有者', t.sent[0].to === 'owner@example.com', t.sent[0].to);
t.ctx.checkSync();
check('同一天第二次檢查不重寄', t.sent.length === 1, t.sent.length);

// 3. 有推送但資料日期是前一天
t = make({ now: at('2026-11-25'), log: [['2026-11-25 14:00', '2026-11-24']], holidays: HOL });
t.ctx.checkSync();
check('推了但資料日期不是今天也要提醒', t.sent.length === 1 && /資料停在 2026-11-24/.test(t.sent[0].subject), t.sent);

// 4. 週末、休市日不檢查
t = make({ now: at('2026-11-28'), log: [], holidays: HOL });
check('週六不寄', t.ctx.checkSync().ok && !t.sent.length && !t.fetched.length);
t = make({ now: at('2026-11-27'), log: [], holidays: HOL });
check('證交所休市日不寄（含「僅辦理結算交割」）', t.ctx.checkSync().ok && !t.sent.length);
t = make({ now: at('2026-01-02'), log: [], holidays: HOL });
check('「開始交易」那天算交易日，照樣檢查', !t.ctx.checkSync().ok && t.sent.length === 1);

// 5. 證交所 API 抓不到：平日照常檢查
t = make({ now: at('2026-11-27'), log: [], fetchFails: true });
check('抓不到休市日時平日照常檢查', !t.ctx.checkSync().ok && t.sent.length === 1);

// 6. 試算表顯示的時間沒補 0（2026-11-05 0:38）也讀得懂
t = make({ now: at('2026-11-05'), log: [['2026-11-05 0:38', '2026-11-5']], holidays: HOL });
check('沒補 0 的日期時間也讀得懂', t.ctx.checkSync().ok && !t.sent.length);

console.log(fails ? fails + ' 項沒過' : '全部通過');
process.exit(fails ? 1 : 0);
