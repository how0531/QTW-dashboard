"""sheet_sync.py + Code.gs 的整條流程測試（由 run_test.sh 呼叫）。"""
import csv
import json
import subprocess
import sys
import urllib.request
from pathlib import Path

SYNC_DIR, WORK, BASE, TOKEN = Path(sys.argv[1]), Path(sys.argv[2]), sys.argv[3], sys.argv[4]
SCRIPT = SYNC_DIR / "sheet_sync.py"
TRIP_H = "trip_id,exit_date,strategy,strategy_log,code,name,side,shares,entry_date,entry_time,entry_price,exit_time,exit_price,hold_minutes,trade_type,entry_amount,gross_pnl,fee,tax,net_pnl,return_pct".split(",")
POS_H = "date,strategy,code,name,side,shares,avg_cost,entry_date,close_price,unrealized_pnl,unrealized_pct".split(",")
DAILY_H = "date,strategy,trips,wins,entry_amount,realized_pnl,realized_pct,unrealized_pnl,open_positions,cum_realized_pnl,day_pct,cum_pct,gross_wins,gross_losses".split(",")
fails = []


def check(name, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + name + (f"  ({detail})" if detail and not cond else ""))
    if not cond:
        fails.append(name)


def write(path, header, rows, bom=True):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig" if bom else "utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def trip(sid, code, name, day, t_in, t_out, p_in, p_out, ret, net):
    return [f"{sid}|{code}|{day.replace('-', '')}T{t_in.replace(':', '')}", day, sid, "QTW" + sid, code, name, "short", 7000, day, t_in, p_in, t_out, p_out, 1, "daytrade", 433300, -14000, 350, 649, net, ret]


def run(*args, config=None):
    cmd = [sys.executable, "-I", str(SCRIPT), "--data-dir", str(WORK / "data"), "--log", str(WORK / "sync.log")]
    cmd += ["--config", str(config or WORK / "sheet_sync.json")] + list(args)
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    return p.returncode, p.stdout + p.stderr


def dump():
    with urllib.request.urlopen(BASE + "/dump") as r:
        d = json.loads(r.read().decode("utf-8"))
    return {k: [row for row in v[1:] if any(x not in ("", None) for x in row)] for k, v in d.items()}


data = WORK / "data"
S1, S2 = "沖空_轉強失敗_4損4移", "投信卡位_15損10日多"
write(data / "trips/2026/trips_20261030.csv", TRIP_H, [trip(S1, "2468", "華經", "2026-10-30", "09:05:12", "09:15:40", 61.9, 60.1, 0.0251, 10870)])
write(data / "trips/2026/trips_20261103.csv", TRIP_H, [trip(S1, "2468", "華經", "2026-11-03", "09:08:03", "09:09:41", 61.9, 63.9, -0.0346, -14999),
                                                      trip(S1, "00878", "國泰永續高股息", "2026-11-03", "10:01:00", "10:30:00", 22.1, 21.9, 0.009, 1200)], bom=False)
write(data / "positions/2026/positions_20261030.csv", POS_H, [["2026-10-30", S2, "1802", "台玻", "long", 27000, 67.8, "2026-10-27", 66.0, -48600, -0.0265]])
write(data / "positions/2026/positions_20261103.csv", POS_H, [["2026-11-03", S2, "1802", "台玻", "long", 27000, 67.8, "2026-10-27", 67.9, 2700, 0.0015],
                                                              ["2026-11-03", S2, "4938", "和碩", "long", 16000, 108, "2026-10-29", 102, -96000, -0.0556]])
write(data / "daily_pnl.csv", DAILY_H, [["2026-10-30", "ALL", 1, 1, 433300, 10870, 0.0251, -48600, 1, 10870, 0.0251, 0.0251, 1, 0],
                                        ["2026-11-03", "ALL", 2, 1, 588000, -13799, -0.0256, -93300, 2, -2929, -0.0256, -0.0005, 1, 1]])

# 試算表上現在有一列 11/3 手動補的資料（trip_id 沒有秒數），全量同步要把它蓋掉
req = urllib.request.Request(BASE + "/seed", data=json.dumps({"_trips": [[f"{S1}|2468|20261103T0908", "2026-11-03", S1, "", 2468, "華經", "short", 7000, "2026-11-03", "9:08:00", 61.9, "9:09:00", 63.9, 1, "daytrade", 433300, -14000, 350, 649, -14999, -0.0346, "2026-11-03 報表信補入"]]}).encode(), method="POST")
urllib.request.urlopen(req).read()

# 1. 沒有設定檔
code, out = run()
check("沒有設定檔時結束代碼 2，不送資料", code == 2, out)

# 2. dry-run
code, out = run("--dry-run")
check("dry-run 只讀檔：daily 模式讀到最近 trips 3 筆、最新持倉 2 筆", code == 0 and "trips 3 筆" in out and "持倉 2 筆" in out, out)

# 3. 密碼錯
(WORK / "bad.json").write_text(json.dumps({"url": BASE + "/exec", "token": "wrong"}), encoding="utf-8")
code, out = run(config=WORK / "bad.json")
check("密碼錯時結束代碼 1，而且不重試三次", code == 1 and "unauthorized" in out and "試了" not in out, out)
check("密碼錯時試算表沒有被改", len(dump()["_trips"]) == 1)

# 4. 全量同步
(WORK / "sheet_sync.json").write_text(json.dumps({"url": BASE + "/exec", "token": TOKEN}), encoding="utf-8")
code, out = run("--full")
d = dump()
check("全量同步成功", code == 0 and "同步成功（replace）" in out, out)
ids = [r[0] for r in d["_trips"]]
check("全量同步後 trips 3 筆，手動補的那列被蓋掉", len(ids) == 3 and f"{S1}|2468|20261103T0908" not in ids, ids)
check("trips 依出場日、時間排序", [r[1] for r in d["_trips"]] == ["2026-10-30", "2026-11-03", "2026-11-03"], [r[1] for r in d["_trips"]])
check("代號 00878 保留開頭的 0", any(r[4] == "00878" for r in d["_trips"]), [r[4] for r in d["_trips"]])
check("數字欄是數字不是文字", all(isinstance(r[20], (int, float)) for r in d["_trips"]), [r[20] for r in d["_trips"]])
check("時間補成 hh:mm:ss", all(len(r[9]) == 8 for r in d["_trips"]), [r[9] for r in d["_trips"]])
check("持倉兩天共 3 筆都在（全量）", len(d["_positions"]) == 3, d["_positions"])
check("daily 2 列", len(d["_daily"]) == 2)
check("同步紀錄寫了一列", len(d["_sync_log"]) == 1 and d["_sync_log"][0][1] == "2026-11-03", d["_sync_log"])

# 5. 隔天：每日同步（新增一趟、改一趟、持倉換新）
write(data / "trips/2026/trips_20261104.csv", TRIP_H, [trip(S1, "3017", "奇鋐", "2026-11-04", "09:30:00", "09:45:00", 512, 500, 0.0234, 9000)])
rows = list(csv.reader(open(data / "trips/2026/trips_20261103.csv", encoding="utf-8")))
rows[1][19] = "-15100"  # 修正一筆淨損益
write(data / "trips/2026/trips_20261103.csv", rows[0], rows[1:])
write(data / "positions/2026/positions_20261104.csv", POS_H, [["2026-11-04", S2, "1802", "台玻", "long", 27000, 67.8, "2026-10-27", 68.5, 18900, 0.0103]])
daily = list(csv.reader(open(data / "daily_pnl.csv", encoding="utf-8-sig")))
daily.append(["2026-11-04", "ALL", 1, 1, 512000, 9000, 0.0234, 18900, 1, 6071, 0.0234, 0.0229, 1, 0])
write(data / "daily_pnl.csv", daily[0], daily[1:])
code, out = run()
d = dump()
check("每日同步成功", code == 0 and "同步成功（upsert）" in out, out)
check("每日同步後 trips 4 筆，沒有重複", len(d["_trips"]) == 4 and len({r[0] for r in d["_trips"]}) == 4, [r[0] for r in d["_trips"]])
fixed = [r for r in d["_trips"] if r[0].endswith("T090803")]
check("改過的那趟被更新成 -15100", fixed and fixed[0][19] == -15100, fixed)
check("10/30 那趟還在（不在最近 5 個檔也不會被刪）", any(r[1] == "2026-10-30" for r in d["_trips"]))
check("持倉保留 10/30、11/3，新增 11/4", sorted({r[0] for r in d["_positions"]}) == ["2026-10-30", "2026-11-03", "2026-11-04"], d["_positions"])
check("daily 整份換成 3 列", len(d["_daily"]) == 3)
check("同步紀錄共 2 列，最新的 cum_pct 是 0.0229", len(d["_sync_log"]) == 2 and abs(float(d["_sync_log"][1][6]) - 0.0229) < 1e-9, d["_sync_log"])

# 6. 同一天再跑一次不會多出資料
code, out = run()
d2 = dump()
check("重跑一次結果一樣（不會重複寫）", code == 0 and d2["_trips"] == d["_trips"] and len(d2["_positions"]) == len(d["_positions"]))

# 7. 壞資料整批不寫
bad = data / "trips/2026/trips_20261105.csv"
write(bad, TRIP_H, [["", "2026-11-05", S1] + [""] * 18])
before = dump()["_trips"]
code, out = run()
check("有一筆缺 trip_id 時整批不寫、結束代碼 1", code == 1 and dump()["_trips"] == before, out)
bad.unlink()

# 8. daily_pnl 比持倉檔新、但 open_positions 不是 0（例如持倉檔沒寫出來）：不能清掉持倉
daily = list(csv.reader(open(data / "daily_pnl.csv", encoding="utf-8-sig")))
write(data / "daily_pnl.csv", daily[0], daily[1:] + [["2026-11-05", "ALL", 0, 0, 0, 0, 0, 18900, 1, 6071, 0, 0.0229, 0, 0]])
code, out = run()
check("daily 較新但還有持倉時，不清掉持倉", code == 0 and len(dump()["_positions"]) == 4, out)

# 9. 全部出場、當天沒有持倉檔：daily_pnl 的 ALL open_positions 是 0，要把舊持倉清掉
daily = list(csv.reader(open(data / "daily_pnl.csv", encoding="utf-8-sig")))
write(data / "daily_pnl.csv", daily[0], daily[1:] + [["2026-11-06", "ALL", 1, 1, 512000, 3000, 0.006, 0, 0, 9071, 0.006, 0.0289, 1, 0]])
code, out = run()
check("全部出場（沒有持倉檔、daily 說 0 檔）後，持倉清空", code == 0 and dump()["_positions"] == [] and "全部出場" in out, out)

# 10. 隔天又有持倉，再隔天的持倉檔只有標題列
write(data / "positions/2026/positions_20261109.csv", POS_H, [["2026-11-09", S2, "2330", "台積電", "long", 1000, 1000, "2026-11-09", 1010, 10000, 0.01]])
code, out = run()
check("新的持倉寫進去", code == 0 and [r[0] for r in dump()["_positions"]] == ["2026-11-09"], dump()["_positions"])
write(data / "positions/2026/positions_20261110.csv", POS_H, [])
code, out = run()
check("持倉檔只有標題列時，持倉清空", code == 0 and dump()["_positions"] == [], out)

print()
print("全部通過" if not fails else f"{len(fails)} 項沒過：{fails}")
sys.exit(1 if fails else 0)
