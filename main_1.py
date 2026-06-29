"""
實價登錄資料串接系統（月報版）

資料來源說明：
  - 內政部實價登錄每月 1、11、21 日更新
  - 預售屋：簽約後 30 天申報、審核後約 40 天上線 → 落差約 2~3 個月
  - 建議每月 1 日執行，下載最新 XLS 滾動資料
  - 保留最近 3 個月，搭配 sheets_writer.py 產出月報
"""

import requests
import zipfile
import os
import sqlite3
import pandas as pd
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta

DB_PATH = "real_estate.db"
DATA_FOLDER = "downloaded_data"

TAICHUNG_DISTRICTS = [
    "中區", "東區", "南區", "西區", "北區",
    "西屯區", "南屯區", "北屯區",
    "豐原區", "東勢區", "大甲區", "清水區", "沙鹿區",
    "梧棲區", "后里區", "神岡區", "潭子區", "大雅區",
    "新社區", "石岡區", "外埔區", "大安區",
    "烏日區", "大肚區", "龍井區", "霧峰區",
    "太平區", "大里區", "和平區"
]

PRE_COLS = {
    "鄉鎮市區": "鄉鎮市區",
    "交易標的": "交易標的",
    "土地區段位置或建物門牌": "門牌",
    "建物型態": "建物型態",
    "總價元": "總價元",
    "單價元平方公尺": "單價元平方公尺",
    "建物移轉總面積平方公尺": "建物移轉總面積平方公尺",
    "屋齡": "屋齡",
    "交易年月日": "交易年月日",
    "建案名稱": "建案名稱"
}


def log(msg):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] {msg}")


# ── 資料庫 ───────────────────────────────────────────────

def init_database():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS presale (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            年月 TEXT,
            資料來源 TEXT DEFAULT 'XLS',
            縣市 TEXT DEFAULT '臺中市',
            鄉鎮市區 TEXT, 交易標的 TEXT,
            門牌 TEXT, 建物型態 TEXT, 總價元 REAL,
            單價元平方公尺 REAL, 建物移轉總面積平方公尺 REAL,
            屋齡 REAL, 交易年月日 TEXT, 建案名稱 TEXT,
            匯入時間 TEXT,
            UNIQUE(門牌, 交易年月日, 總價元)
        )
    """)
    for col, dtype in [("年月", "TEXT"), ("資料來源", "TEXT")]:
        try:
            c.execute(f"ALTER TABLE presale ADD COLUMN {col} {dtype}")
            conn.commit()
            log(f"已新增欄位：{col}")
        except Exception:
            pass
    c.execute("""
        CREATE TABLE IF NOT EXISTS month_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            年月 TEXT UNIQUE,
            資料來源 TEXT,
            季別 TEXT,
            匯入時間 TEXT,
            新增筆數 INTEGER
        )
    """)
    conn.commit()
    conn.close()
    log("資料庫初始化完成")


# ── 日期工具 ─────────────────────────────────────────────

def get_season_code(year_ad, month):
    roc_y = year_ad - 1911
    if month <= 3:   return f"{roc_y}S1"
    elif month <= 6: return f"{roc_y}S2"
    elif month <= 9: return f"{roc_y}S3"
    else:            return f"{roc_y}S4"


def roc_date_range(year_ad, month):
    roc_y = year_ad - 1911
    last_day = calendar.monthrange(year_ad, month)[1]
    return (roc_y * 10000 + month * 100 + 1,
            roc_y * 10000 + month * 100 + last_day)


# ── 下載 ─────────────────────────────────────────────────

def download_xls():
    """下載最新 XLS 滾動資料（每月 1、11、21 日更新）"""
    log("[XLS] 下載最新滾動資料...")
    try:
        r = requests.get(
            "https://plvr.land.moi.gov.tw/Download",
            params={"type": "zip", "fileName": "lvr_landxls.zip"},
            timeout=180
        )
        if r.status_code == 200 and len(r.content) > 50000:
            os.makedirs(DATA_FOLDER, exist_ok=True)
            path = os.path.join(DATA_FOLDER, "lvr_landxls.zip")
            with open(path, "wb") as f:
                f.write(r.content)
            log(f"[XLS] 下載完成（{len(r.content)//1024//1024:.1f} MB）")
            return path
        else:
            log(f"[XLS] 下載失敗，HTTP {r.status_code}")
            return None
    except Exception as e:
        log(f"[XLS] 下載錯誤：{e}")
        return None


def download_season_csv(season_code):
    """備用：下載季度歸檔 CSV（季度結束後才有）"""
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Referer": "https://plvr.land.moi.gov.tw/DownloadOpenData",
    }
    log(f"[CSV] 下載季度歸檔 {season_code}...")
    try:
        session.get("https://plvr.land.moi.gov.tw/DownloadOpenData",
                    headers=headers, timeout=30)
        r = session.get(
            "https://plvr.land.moi.gov.tw/DownloadSeason",
            params={"season": season_code, "type": "zip", "fileName": "lvr_landcsv.zip"},
            headers=headers, timeout=180
        )
        if r.status_code == 200 and len(r.content) > 50000:
            os.makedirs(DATA_FOLDER, exist_ok=True)
            path = os.path.join(DATA_FOLDER, f"{season_code}.zip")
            with open(path, "wb") as f:
                f.write(r.content)
            log(f"[CSV] 下載完成（{len(r.content)//1024} KB）")
            return path
        else:
            log(f"[CSV] 下載失敗")
            return None
    except Exception as e:
        log(f"[CSV] 下載錯誤：{e}")
        return None


def extract_zip(zip_path, folder_name):
    out = os.path.join(DATA_FOLDER, folder_name)
    os.makedirs(out, exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(out)
    log(f"解壓縮：{out}")
    return out


# ── 讀取與篩選 ───────────────────────────────────────────

def _parse_date_int(val):
    try:
        if hasattr(val, 'year'):
            roc_y = val.year - 1911
            return roc_y * 10000 + val.month * 100 + val.day
        s = str(val).strip().replace("/", "").replace(".", "").replace("-", "")
        return int(s.split(".")[0])
    except Exception:
        return 0


def _filter_taichung(df, date_start, date_end):
    cleaned = pd.DataFrame()
    for orig, new in PRE_COLS.items():
        cleaned[new] = df[orig] if orig in df.columns else ""

    before = len(cleaned)
    cleaned = cleaned[pd.to_numeric(cleaned["總價元"], errors="coerce").notna()].copy()
    if before != len(cleaned):
        log(f"  過濾無效列 {before - len(cleaned)} 筆")

    cleaned = cleaned[cleaned["鄉鎮市區"].astype(str).str.contains(
        "|".join(TAICHUNG_DISTRICTS), na=False
    )].copy()
    log(f"  大台中：{len(cleaned)} 筆")

    cleaned["_d"] = cleaned["交易年月日"].apply(_parse_date_int)
    cleaned = cleaned[(cleaned["_d"] >= date_start) & (cleaned["_d"] <= date_end)].copy()
    cleaned = cleaned.drop(columns=["_d"])
    log(f"  日期 {date_start}～{date_end}：{len(cleaned)} 筆")
    return cleaned


def read_xls(folder, date_start, date_end):
    xls_path = os.path.join(folder, "b_lvr_land_b.xls")
    if not os.path.exists(xls_path):
        log(f"[XLS] 找不到 b_lvr_land_b.xls")
        return pd.DataFrame()
    try:
        df = pd.read_excel(xls_path, sheet_name="預售屋買賣",
                           header=0, skiprows=[1], engine="xlrd")
        log(f"[XLS] 共 {len(df)} 筆，篩選中...")
        return _filter_taichung(df, date_start, date_end)
    except Exception as e:
        log(f"[XLS] 讀取失敗：{e}")
        return pd.DataFrame()


def read_csv(folder, date_start, date_end):
    target = None
    for fname in os.listdir(folder):
        if fname.lower() == "b_lvr_land_b.csv":
            target = os.path.join(folder, fname)
            break
    if not target:
        log("[CSV] 找不到 b_lvr_land_b.csv")
        return pd.DataFrame()
    for skip in [0, 1, 2]:
        try:
            df = pd.read_csv(target, encoding="utf-8-sig",
                             skiprows=skip, low_memory=False)
            if any(c in df.columns for c in ["鄉鎮市區", "總價元"]) and len(df) > 0:
                log(f"[CSV] 共 {len(df)} 筆，篩選中...")
                return _filter_taichung(df, date_start, date_end)
        except Exception:
            pass
    log("[CSV] 無法讀取")
    return pd.DataFrame()


# ── 寫入資料庫 ───────────────────────────────────────────

def save_to_db(df, ym, season_code, source):
    if df.empty:
        log(f"  {ym}：無資料")
        return 0

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    inserted = 0

    for _, row in df.iterrows():
        try:
            c.execute("""
                INSERT OR IGNORE INTO presale
                (年月, 資料來源, 縣市, 鄉鎮市區, 交易標的, 門牌, 建物型態,
                 總價元, 單價元平方公尺, 建物移轉總面積平方公尺,
                 屋齡, 交易年月日, 建案名稱, 匯入時間)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                ym, source, "臺中市",
                str(row.get("鄉鎮市區", "")),
                str(row.get("交易標的", "")),
                str(row.get("門牌", "")),
                str(row.get("建物型態", "")),
                row.get("總價元", None),
                row.get("單價元平方公尺", None),
                row.get("建物移轉總面積平方公尺", None),
                row.get("屋齡", None),
                str(row.get("交易年月日", "")),
                str(row.get("建案名稱", "")),
                now_str
            ))
            if c.rowcount > 0:
                inserted += 1
        except Exception:
            pass

    total = c.execute("SELECT COUNT(*) FROM presale").fetchone()[0]
    c.execute("""
        INSERT OR REPLACE INTO month_log (年月, 資料來源, 季別, 匯入時間, 新增筆數)
        VALUES (?,?,?,?,?)
    """, (ym, source, season_code, now_str, inserted))

    conn.commit()
    conn.close()
    log(f"  {ym}：新增 {inserted} 筆，資料庫累積 {total} 筆")
    return inserted


# ── 清理舊資料 ───────────────────────────────────────────

def cleanup_old_data(months_keep=3):
    """只保留最近 N 個月"""
    cutoff = (datetime.now() - relativedelta(months=months_keep)).strftime("%Y-%m")
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    try:
        before = c.execute("SELECT COUNT(*) FROM presale").fetchone()[0]
        c.execute("DELETE FROM presale WHERE 年月 IS NOT NULL AND 年月 != '' AND 年月 < ?",
                  (cutoff,))
        c.execute("DELETE FROM month_log WHERE 年月 < ?", (cutoff,))
        after = c.execute("SELECT COUNT(*) FROM presale").fetchone()[0]
        if before - after > 0:
            log(f"清除 {before - after} 筆（{cutoff} 以前），剩 {after} 筆")
        else:
            log(f"無需清除（共 {after} 筆）")
    except Exception as e:
        log(f"清除失敗：{e}")
    conn.commit()
    conn.close()


# ── 主流程 ───────────────────────────────────────────────

def run_import(months_back=3):
    """
    下載最新 XLS，匯入最近 N 個月的成交紀錄。
    每次執行都會補齊前幾個月的新增案件（INSERT OR IGNORE 防重複）。
    建議每月 1 日執行。
    """
    now = datetime.now()
    end = now - relativedelta(months=1)

    months = [end - relativedelta(months=i) for i in range(months_back)]
    log(f"=== 匯入最近 {months_back} 個月："
        f"{months[-1].strftime('%Y-%m')} ~ {months[0].strftime('%Y-%m')} ===")

    zip_path = download_xls()
    if not zip_path:
        log("XLS 下載失敗，終止")
        return

    folder = extract_zip(zip_path, "opendata")

    for t in months:
        ym = t.strftime("%Y-%m")
        d_start, d_end = roc_date_range(t.year, t.month)
        season_code = get_season_code(t.year, t.month)
        log(f"--- {ym} ---")
        df = read_xls(folder, d_start, d_end)
        save_to_db(df, ym, season_code, "XLS")

    cleanup_old_data(months_back)
    log("=== 匯入完成 ===")


def run_backfill(season_code):
    """
    補匯歷史季度 CSV（季度結束後才可用）
    例：python main_1.py backfill 115S1
    """
    log(f"=== 補匯季度歸檔：{season_code} ===")
    zip_path = download_season_csv(season_code)
    if not zip_path:
        log("CSV 下載失敗，終止")
        return

    folder = extract_zip(zip_path, season_code)
    roc_y = int(season_code[:3])
    year_ad = roc_y + 1911
    ranges = {"S1": (1, 3), "S2": (4, 6), "S3": (7, 9), "S4": (10, 12)}
    q_start, q_end = ranges.get(season_code[3:], (1, 3))

    for mo in range(q_start, q_end + 1):
        ym = f"{year_ad}-{mo:02d}"
        d_start, d_end = roc_date_range(year_ad, mo)
        log(f"--- {ym} ---")
        df = read_csv(folder, d_start, d_end)
        save_to_db(df, ym, season_code, "CSV")

    log(f"=== {season_code} 補匯完成 ===")


def get_stats():
    conn = sqlite3.connect(DB_PATH)
    print("\n資料庫統計")
    print("=" * 55)
    try:
        n = conn.execute("SELECT COUNT(*) FROM presale").fetchone()[0]
        print(f"  預售屋累積：{int(n):,} 筆")
    except Exception:
        print("  預售屋：0 筆")
    try:
        logs = pd.read_sql(
            "SELECT 年月, 資料來源, 季別, 新增筆數 FROM month_log ORDER BY 年月 DESC",
            conn
        )
        print("\n月度記錄：")
        print(logs.to_string(index=False))
    except Exception:
        print("  尚無記錄")
    conn.close()


if __name__ == "__main__":
    import sys
    try:
        from dateutil.relativedelta import relativedelta
    except ImportError:
        import subprocess
        subprocess.run(["pip", "install", "python-dateutil", "--break-system-packages"])
        from dateutil.relativedelta import relativedelta

    init_database()
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""

    if cmd == "import":
        n = int(sys.argv[2]) if len(sys.argv) > 2 else 3
        run_import(n)
    elif cmd == "backfill":
        season = sys.argv[2] if len(sys.argv) > 2 else None
        if not season:
            print("請指定季別，例：python main_1.py backfill 115S1")
        else:
            run_backfill(season)
    elif cmd == "cleanup":
        cleanup_old_data()
    elif cmd == "stats":
        get_stats()
    else:
        print("用法：")
        print("  python main_1.py import       ← 下載 XLS，匯入最近 3 個月")
        print("  python main_1.py import 6     ← 下載 XLS，匯入最近 6 個月")
        print("  python main_1.py backfill 115S1  ← 補匯歷史季度 CSV")
        print("  python main_1.py cleanup      ← 手動清除 3 個月前資料")
        print("  python main_1.py stats        ← 查看統計")
        print()
        print("建議執行時機：每月 1 日（配合官網 1/11/21 日更新）")
