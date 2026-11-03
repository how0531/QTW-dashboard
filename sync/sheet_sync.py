"""把 C:\\XS\\data 的交易資料推到 Google 試算表「策略管理」，網頁 dashboard 就會跟著更新。

用法（在 C:\\XS\\tools 執行）：
    python sheet_sync.py            每天用：推最近 5 天的 trips、最新一天的持倉、整份 daily_pnl
    python sheet_sync.py --full     第一次用：全部歷史資料整份覆蓋試算表
    python sheet_sync.py --dry-run  只讀檔案、印出會送幾筆，不連網路

設定檔 sheet_sync.json（放在這個程式旁邊，裡面有密碼，不要給別人）：
    {"url": "Apps Script 網頁應用程式網址", "token": "setup() 產生的 SYNC_TOKEN"}

只用 Python 標準函式庫。結果會附加寫進 C:\\XS\\report\\sheet_sync.log。
結束代碼：0 成功；1 推送失敗；2 沒有設定檔或資料夾。
"""
import argparse
import csv
import datetime as dt
import json
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_DATA = Path(r"C:\XS\data")
DEFAULT_LOG = Path(r"C:\XS\report\sheet_sync.log")
RECENT_DAYS = 5

TRIP_NUM = ["shares", "entry_price", "exit_price", "hold_minutes", "entry_amount", "gross_pnl", "fee", "tax", "net_pnl", "return_pct"]
POS_NUM = ["shares", "avg_cost", "close_price", "unrealized_pnl", "unrealized_pct"]
DAILY_NUM = ["trips", "wins", "entry_amount", "realized_pnl", "realized_pct", "unrealized_pnl", "open_positions",
             "cum_realized_pnl", "day_pct", "cum_pct", "gross_wins", "gross_losses"]


def log(msg, log_path):
    line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    print(line)
    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def read_csv(path, numeric):
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            if not any((v or "").strip() for v in r.values()):
                continue
            o = {}
            for k, v in r.items():
                if k is None:
                    continue
                k = k.strip()
                v = (v or "").strip()
                if k in numeric and v != "":
                    try:
                        n = float(v.replace(",", ""))
                        v = int(n) if n.is_integer() and k not in ("return_pct", "unrealized_pct", "realized_pct", "day_pct", "cum_pct") else n
                    except ValueError:
                        pass
                o[k] = v
            rows.append(o)
    return rows


def day_of(path, prefix):
    try:
        return dt.datetime.strptime(path.stem.replace(prefix, ""), "%Y%m%d").date()
    except ValueError:
        return None


def collect(data_dir, full):
    trip_files = sorted(p for p in (data_dir / "trips").glob("*/trips_*.csv") if day_of(p, "trips_"))
    pos_files = sorted(p for p in (data_dir / "positions").glob("*/positions_*.csv") if day_of(p, "positions_"))
    if not full:
        trip_files = trip_files[-RECENT_DAYS:]
        pos_files = pos_files[-1:]
    trips = [r for p in trip_files for r in read_csv(p, TRIP_NUM)]
    positions = [r for p in pos_files for r in read_csv(p, POS_NUM)]
    daily_path = data_dir / "daily_pnl.csv"
    daily = read_csv(daily_path, DAILY_NUM) if daily_path.exists() else None
    # 同一個 trip_id 出現在兩個檔案時，以較新的檔案為準
    # 沒有 trip_id 的列照樣送出，讓試算表那邊整批擋下來並回報，不在這裡悄悄丟掉
    seen, broken = {}, []
    for r in trips:
        if r.get("trip_id"):
            seen[r["trip_id"]] = r
        else:
            broken.append(r)
    days = [day_of(p, "trips_") for p in trip_files] + [day_of(p, "positions_") for p in pos_files]
    latest = max(days).isoformat() if days else ""
    # 持倉以哪一天為準：最新的持倉檔（只有標題列也算，代表當天收盤沒有持倉）；
    # 如果 daily_pnl 有更新的一天、而且那天 ALL 的 open_positions 是 0，代表全部出場了，以那天為準
    as_of = day_of(pos_files[-1], "positions_").isoformat() if pos_files else ""
    for r in daily or []:
        d = str(r.get("date", ""))[:10].replace("/", "-")
        if r.get("strategy") == "ALL" and d > as_of and r.get("open_positions") in (0, 0.0, "0"):
            as_of = d
    return list(seen.values()) + broken, positions, daily, latest, len(trip_files), len(pos_files), as_of


def post(url, payload, retries=3):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    last = None
    for i in range(retries):
        try:
            req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"}, method="POST")
            # Apps Script 收到 POST 後會 302 轉址到結果頁；urllib 會自動改用 GET 去拿結果
            with urllib.request.urlopen(req, timeout=180) as res:
                body = res.read().decode("utf-8", "replace")
            try:
                out = json.loads(body)
            except ValueError:
                raise RuntimeError("回應不是 JSON（網址可能不對，或部署沒有開放「所有人」）：" + body[:200])
            if not out.get("ok"):
                err = out.get("error", "unknown")
                if err in ("unauthorized", "bad_json") or "欄" in str(err) or "缺少" in str(err) or "看不懂" in str(err):
                    raise PermissionError(err)  # 重試也沒用
                raise RuntimeError(err)
            return out
        except PermissionError:
            raise
        except (urllib.error.URLError, OSError, RuntimeError) as e:
            last = e
            if i < retries - 1:
                time.sleep(5 * (i + 1) ** 2)
    raise RuntimeError(f"試了 {retries} 次都失敗：{last}")


def main(argv=None):
    ap = argparse.ArgumentParser(description="把 XQ 交易資料推到 Google 試算表")
    ap.add_argument("--full", action="store_true", help="全部歷史資料整份覆蓋")
    ap.add_argument("--dry-run", action="store_true", help="只讀檔案，不送出")
    ap.add_argument("--data-dir", default=str(DEFAULT_DATA))
    ap.add_argument("--config", default=str(HERE / "sheet_sync.json"))
    ap.add_argument("--log", default=str(DEFAULT_LOG))
    a = ap.parse_args(argv)
    log_path = Path(a.log)
    data_dir = Path(a.data_dir)

    if not data_dir.exists():
        log(f"找不到資料夾 {data_dir}，略過同步", log_path)
        return 2
    trips, positions, daily, latest, nt, npos, as_of = collect(data_dir, a.full)
    summary = f"trips {len(trips)} 筆（{nt} 個檔）、持倉 {len(positions)} 筆（{npos} 個檔）、daily {len(daily) if daily is not None else '無'} 列、資料日期 {latest or '無'}"
    if a.dry_run:
        log("[dry-run] " + summary, log_path)
        return 0

    cfg_path = Path(a.config)
    if not cfg_path.exists():
        log(f"沒有設定檔 {cfg_path}，略過同步（照 sync/README.md 建一個）", log_path)
        return 2
    cfg = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
    if not cfg.get("url") or not cfg.get("token"):
        log("sheet_sync.json 缺 url 或 token，略過同步", log_path)
        return 2

    payload = {"token": cfg["token"], "mode": "replace" if a.full else "upsert", "source": "sheet_sync.py",
               "data_date": latest, "trips": trips, "positions": positions, "positions_as_of": as_of}
    if daily is not None:
        payload["daily"] = daily
    try:
        out = post(cfg["url"], payload)
    except Exception as e:  # noqa: BLE001 要把任何錯誤都記進 log
        log(f"同步失敗：{e}｜{summary}", log_path)
        return 1
    t = out.get("trips") or {}
    p = out.get("positions") or {}
    cleared = f"，清掉舊持倉 {p['cleared']} 筆（{p.get('as_of')} 全部出場）" if p.get("cleared") else ""
    log(f"同步成功（{out.get('mode')}）：trips 新增 {t.get('added', 0)}、更新 {t.get('updated', 0)}、共 {t.get('total', '?')}；"
        f"持倉寫入 {p.get('written', 0)}（{','.join(p.get('dates', []))}）{cleared}；daily {(out.get('daily') or {}).get('written', '-')} 列", log_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
