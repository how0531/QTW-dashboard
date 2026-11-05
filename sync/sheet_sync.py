"""把 C:\\XS\\data 的交易資料推到 Google 試算表「策略管理」，網頁 dashboard 就會跟著更新。

用法（在 C:\\XS\\tools 執行）：
    python sheet_sync.py            每天用：推最近 5 天的 trips（上次同步失敗的話，從上次成功那天起全部補送）、
                                    最新一天的持倉、整份 daily_pnl
    python sheet_sync.py --full     第一次用：全部歷史資料整份覆蓋試算表
    python sheet_sync.py --dry-run  只讀檔案、印出會送幾筆，不連網路
    python sheet_sync.py --full --force  全量覆蓋的筆數比試算表少一半以上時，試算表會拒絕；確定沒問題才加 --force

設定檔 sheet_sync.json（放在這個程式旁邊，裡面有密碼，不要給別人）：
    {"url": "Apps Script 網頁應用程式網址", "token": "setup() 產生的 SYNC_TOKEN"}

只用 Python 標準函式庫。結果（包括任何錯誤）都會附加寫進 C:\\XS\\report\\sheet_sync.log。
最後一次成功的資料日期記在設定檔旁邊的 sheet_sync_state.json，下次從那天起補送。
結束代碼：0 成功；1 推送失敗或資料有問題；2 沒有設定檔或資料夾、設定檔寫錯。
"""
import argparse
import csv
import datetime as dt
import json
import sys
import time
import traceback
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
# 和 Code.gs 的 REQUIRED 一樣；在這裡先檢查，錯誤訊息才能指出是哪個檔案第幾行
REQUIRED = {"trips": ["trip_id", "exit_date", "strategy"], "positions": ["date", "strategy", "code"], "daily": ["date", "strategy"]}
SRC = "_src"  # 每列記住來自哪個檔案第幾行，送出前拿掉


class DataError(Exception):
    """電腦上的資料檔有問題（編碼、缺欄位），整批不送。"""


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
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for r in reader:
                if not any((v or "").strip() for v in r.values()):
                    continue
                o = {SRC: f"{path.name} 第 {reader.line_num} 行"}
                rows.append(_convert(r, o, numeric))
    except UnicodeDecodeError as e:
        raise DataError(f"{path} 不是 UTF-8 編碼，讀不懂（位置 {e.start}）") from None
    except csv.Error as e:
        raise DataError(f"{path} 格式壞掉：{e}") from None
    return rows


def _convert(r, o, numeric):
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
    return o


def day_of(path, prefix):
    try:
        return dt.datetime.strptime(path.stem.replace(prefix, ""), "%Y%m%d").date()
    except ValueError:
        return None


def collect(data_dir, full, since=None):
    trip_files = sorted(p for p in (data_dir / "trips").glob("*/trips_*.csv") if day_of(p, "trips_"))
    pos_files = sorted(p for p in (data_dir / "positions").glob("*/positions_*.csv") if day_of(p, "positions_"))
    if not full:
        # 平常送最近 5 個檔；如果上次成功是更早以前（中間同步失敗過），從那天起全部補送
        recent = set(trip_files[-RECENT_DAYS:])
        if since:
            recent |= {p for p in trip_files if day_of(p, "trips_") >= since}
        trip_files = sorted(recent)
        pos_files = pos_files[-1:]
    trips = [r for p in trip_files for r in read_csv(p, TRIP_NUM)]
    positions = [r for p in pos_files for r in read_csv(p, POS_NUM)]
    daily_path = data_dir / "daily_pnl.csv"
    daily = read_csv(daily_path, DAILY_NUM) if daily_path.exists() else None
    # 同一個 trip_id 出現在兩個檔案時，以較新的檔案為準
    # 沒有 trip_id 的列不在這裡悄悄丟掉：留著讓 problems() 指出是哪個檔案第幾行，整批不送
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


def problems(kind, rows):
    out = []
    for r in rows or []:
        miss = [c for c in REQUIRED[kind] if str(r.get(c, "")).strip() == ""]
        if miss:
            out.append(f"{r.get(SRC, kind)}缺 {'、'.join(miss)}")
    return out


def strip_src(rows):
    return [{k: v for k, v in r.items() if k != SRC} for r in rows] if rows is not None else None


def read_state(path):
    try:
        d = json.loads(path.read_text(encoding="utf-8"))
        return dt.date.fromisoformat(d["last_ok_date"])
    except (OSError, ValueError, KeyError, TypeError):
        return None


def write_state(path, latest):
    try:
        path.write_text(json.dumps({"last_ok_date": latest, "at": f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}"}), encoding="utf-8")
    except OSError:
        pass


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
                if err in ("unauthorized", "bad_json") or any(w in str(err) for w in ("欄", "缺少", "看不懂", "不是數字", "先不寫入")):
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
    ap.add_argument("--force", action="store_true", help="全量覆蓋筆數變少很多時也照樣覆蓋")
    ap.add_argument("--dry-run", action="store_true", help="只讀檔案，不送出")
    ap.add_argument("--data-dir", default=str(DEFAULT_DATA))
    ap.add_argument("--config", default=str(HERE / "sheet_sync.json"))
    ap.add_argument("--log", default=str(DEFAULT_LOG))
    a = ap.parse_args(argv)
    log_path = Path(a.log)
    try:
        return run(a, log_path)
    except DataError as e:
        log(f"資料有問題，整批不送：{e}", log_path)
        return 1
    except Exception as e:  # noqa: BLE001 排程執行時看不到畫面，任何錯誤都要留在 log
        where = traceback.extract_tb(e.__traceback__)[-1]
        log(f"同步失敗（程式錯誤）：{type(e).__name__}: {e}（{Path(where.filename).name} 第 {where.lineno} 行）", log_path)
        return 1


def run(a, log_path):
    data_dir = Path(a.data_dir)
    cfg_path = Path(a.config)
    state_path = cfg_path.parent / "sheet_sync_state.json"

    if not data_dir.exists():
        log(f"找不到資料夾 {data_dir}，略過同步", log_path)
        return 2
    since = None if a.full else read_state(state_path)
    trips, positions, daily, latest, nt, npos, as_of = collect(data_dir, a.full, since)
    summary = f"trips {len(trips)} 筆（{nt} 個檔）、持倉 {len(positions)} 筆（{npos} 個檔）、daily {len(daily) if daily is not None else '無'} 列、資料日期 {latest or '無'}"
    bad = problems("trips", trips) + problems("positions", positions) + problems("daily", daily)
    if bad:
        more = f"，另有 {len(bad) - 5} 筆" if len(bad) > 5 else ""
        raise DataError("；".join(bad[:5]) + more)
    if a.dry_run:
        log("[dry-run] " + summary, log_path)
        return 0

    if not cfg_path.exists():
        log(f"沒有設定檔 {cfg_path}，略過同步（照 sync/README.md 建一個）", log_path)
        return 2
    try:
        cfg = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
    except ValueError as e:
        log(f"設定檔 {cfg_path.name} 寫錯了，不是正確的 JSON：{e}", log_path)
        return 2
    if not isinstance(cfg, dict) or not cfg.get("url") or not cfg.get("token"):
        log("sheet_sync.json 缺 url 或 token，略過同步", log_path)
        return 2

    payload = {"token": cfg["token"], "mode": "replace" if a.full else "upsert", "source": "sheet_sync.py",
               "data_date": latest, "trips": strip_src(trips), "positions": strip_src(positions), "positions_as_of": as_of}
    if daily is not None:
        payload["daily"] = strip_src(daily)
    if a.force:
        payload["force"] = True
    if since and nt > RECENT_DAYS:
        summary += f"（從上次成功的 {since.isoformat()} 起補送）"
    try:
        out = post(cfg["url"], payload)
    except Exception as e:  # noqa: BLE001 要把任何錯誤都記進 log
        log(f"同步失敗：{e}｜{summary}", log_path)
        return 1
    if latest:
        write_state(state_path, latest)
    t = out.get("trips") or {}
    p = out.get("positions") or {}
    cleared = f"，清掉舊持倉 {p['cleared']} 筆（{p.get('as_of')} 全部出場）" if p.get("cleared") else ""
    log(f"同步成功（{out.get('mode')}）：trips 新增 {t.get('added', 0)}、更新 {t.get('updated', 0)}、共 {t.get('total', '?')}；"
        f"持倉寫入 {p.get('written', 0)}（{','.join(p.get('dates', []))}）{cleared}；daily {(out.get('daily') or {}).get('written', '-')} 列｜{summary}", log_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
