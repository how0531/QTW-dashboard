// 在本機模擬 Apps Script：用假的 SpreadsheetApp 執行 Code.gs，並模仿 /exec 的 302 轉址。
// 用法：node fake_gas_server.js <port> <token>
const fs = require('fs');
const http = require('http');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

const COLS = {
  _trips: 'trip_id,exit_date,strategy,strategy_log,code,name,side,shares,entry_date,entry_time,entry_price,exit_time,exit_price,hold_minutes,trade_type,entry_amount,gross_pnl,fee,tax,net_pnl,return_pct,synced_at'.split(','),
  _positions: 'date,strategy,code,name,side,shares,avg_cost,entry_date,close_price,unrealized_pnl,unrealized_pct,synced_at'.split(','),
  _daily: 'date,strategy,trips,wins,entry_amount,realized_pnl,realized_pct,unrealized_pnl,open_positions,cum_realized_pnl,day_pct,cum_pct,gross_wins,gross_losses'.split(','),
  _sync_log: '同步時間,資料日期,方式,trips 新增,trips 總數,positions 筆數,ALL cum_pct,備註'.split(','),
};
// 假試算表：值存成「像 Sheets 會顯示的樣子」，日期時間依格式轉字串，數字照存。
// 和真的一樣有列數上限（新分頁 1000 列 × 26 欄），範圍超出就丟錯
const MAX_ROWS = Number(process.env.FAKE_MAX_ROWS || 1000);
function makeSheet(name, header) {
  const grid = [(header || COLS[name]).slice()];
  const fmt = {};
  let maxRows = MAX_ROWS;
  const sh = {
    name, grid, fmt,
    getLastRow: () => { for (let r = grid.length; r > 0; r--) if (grid[r - 1].some((v) => v !== '' && v != null)) return r; return 0; },
    getMaxRows: () => maxRows,
    getLastColumn: () => Math.max(0, ...grid.map((r) => { for (let c = r.length; c > 0; c--) if (r[c - 1] !== '' && r[c - 1] != null) return c; return 0; })),
    insertRowsAfter: (after, n) => { if (after > maxRows) throw new Error('insertRowsAfter 超出範圍'); maxRows += n; },
    appendRow: (row) => { const r = sh.getLastRow() + 1; if (r > maxRows) maxRows = r; grid[r - 1] = row.slice(); },
    getParent: () => SS,
    getRange: (r, c, nr = 1, nc = 1) => {
      if (r < 1 || c < 1 || r - 1 + nr > maxRows || c - 1 + nc > 26) throw new Error(`範圍座標超出工作表大小（${name} 第 ${r}～${r + nr - 1} 列，上限 ${maxRows}）`);
      return {
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => ((grid[r - 1 + i] || [])[c - 1 + j] ?? ''))),
      getDisplayValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => String((grid[r - 1 + i] || [])[c - 1 + j] ?? ''))),
      setValues: (vals) => {
        if (vals.length !== nr || vals.some((row) => row.length !== nc)) throw new Error('setValues 大小不符');
        vals.forEach((row, i) => { grid[r - 1 + i] = grid[r - 1 + i] || []; row.forEach((v, j) => {
          const f = fmt[c + j];
          if (f === '@') v = String(v);
          else if (typeof v === 'string' && /^\d+(\.\d+)?$/.test(v) && f !== '@') v = Number(v); // Sheets 會把數字字串轉成數字
          grid[r - 1 + i][c - 1 + j] = v; }); });
      },
      clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (grid[r - 1 + i]) grid[r - 1 + i][c - 1 + j] = ''; },
      setNumberFormat: (f) => { for (let j = 0; j < nc; j++) fmt[c + j] = f; },
      };
    },
  };
  return sh;
}
const sheets = Object.fromEntries(Object.keys(COLS).map((n) => [n, makeSheet(n)]));
// 用公式展開的呈現頁，只檢查列數有沒有跟著加
sheets['交易明細'] = makeSheet('交易明細', ['出場日']);
sheets['每日'] = makeSheet('每日', ['策略']);
sheets['策略對照'] = makeSheet('策略對照', ['策略', '狀態', '上線日', '備註', '分類']);
sheets['總覽'] = makeSheet('總覽', ['不該被網頁讀到']);
const SS = { getSheetByName: (n) => sheets[n] || null, getSpreadsheetTimeZone: () => 'Asia/Taipei' };
const props = { SYNC_TOKEN: process.argv[3] };
const ctx = {
  SpreadsheetApp: { getActiveSpreadsheet: () => SS },
  PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: (s) => ({ s, setMimeType() { return this; } }) },
  Utilities: { getUuid: () => crypto.randomUUID(), formatDate: (d, tz, f) => new Intl.DateTimeFormat('sv-SE', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).format(d).replace(',', '') },
  Logger: { log: () => {} },
  console,
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), ctx);

const pending = new Map();
const server = http.createServer((req, res) => {
  const u = new URL(req.url, 'http://x');
  if (req.method === 'POST' && u.pathname === '/exec') {
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => {
      const out = ctx.doPost({ postData: { contents: body } });
      const id = crypto.randomUUID();
      pending.set(id, out.s);
      res.writeHead(302, { Location: `/echo?id=${id}` }); res.end();
    });
    return;
  }
  if (req.method === 'GET' && u.pathname === '/exec') {
    const out = ctx.doGet({ parameter: Object.fromEntries(u.searchParams) });
    const id = crypto.randomUUID();
    pending.set(id, out.s);
    res.writeHead(302, { Location: `/echo?id=${id}` }); res.end();
    return;
  }
  if (req.method === 'GET' && u.pathname === '/echo') { res.writeHead(200, { 'Content-Type': 'application/json' }); res.end(pending.get(u.searchParams.get('id')) || '{}'); return; }
  if (req.method === 'GET' && u.pathname === '/dump') { res.writeHead(200, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(Object.fromEntries(Object.entries(sheets).filter(([k]) => k.startsWith('_')).map(([k, s]) => [k, s.grid])))); return; }
  if (req.method === 'GET' && u.pathname === '/rows') { res.writeHead(200, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(Object.fromEntries(Object.entries(sheets).map(([k, s]) => [k, s.getMaxRows()])))); return; }
  if (req.method === 'POST' && u.pathname === '/seed') {
    let body = ''; req.on('data', (c) => (body += c));
    req.on('end', () => { const d = JSON.parse(body); Object.entries(d).forEach(([k, rows]) => rows.forEach((r) => sheets[k].grid.push(r))); res.end('ok'); });
    return;
  }
  res.writeHead(404); res.end();
});
server.listen(Number(process.argv[2]), '127.0.0.1', () => console.log('ready'));
