"""
歷史預售屋資料匯入（2018-2023，ROC 107S1～112S4）
- 下載季度 CSV，篩選台中市，存入 presale_history 表
- 用途：建立地址→建案名稱對照表，供成屋建案名稱比對使用
- 執行：python3 import_history.py
"""

import requests
import zipfile
import os
import io
import sqlite3
import pandas as pd
from datetime import datetime

DB_PATH = "real_estate.db"
DATA_FOLDER = "downloaded_data"

TAICHUNG_DISTRICTS = [
    "中區","東區","南區","西區","北區","西屯區","南屯區","北屯區",
    "豐原區","東勢區","大甲區","清水區","沙鹿區","梧棲區","后里區",
    "神岡區","潭子區","大雅區","新社區","石岡區","外埔區","大安區",
    "烏日區","大肚區","龍井區","霧峰區","太平區","大里區","和平區"
]

PRE_COLS = {
    "鄉鎮市區": "鄉鎮市區",
    "交易標的": "交易標的",
    "土地位置建物門牌": "門牌",
    "建物型態": "建物型態",
    "總價元": "總價元",
    "單價元平方公尺": "單價元平方公尺",
    "建物移轉總面積平方公尺": "建物移轉總面積平方公尺",
    "屋齡": "屋齡",
    "交易年月日": "交易年月日",
    "建案名稱": "建案名稱",
}

SEASONS = [
    f"{roc}S{q}"
    for roc in range(107, 115)   # 107~114（2018~2025）
    for q in range(1, 5)         # S1~S4
]


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def download_season(season_code):
    path = os.path.join(DATA_FOLDER, f"{season_code}.zip")
    if os.path.exists(path) and os.path.getsize(path) > 50000:
        return path

    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Referer": "https://plvr.land.moi.gov.tw/DownloadOpenData",
    }
    try:
        session.get("https://plvr.land.moi.gov.tw/DownloadOpenData",
                    headers=headers, timeout=30)
    except Exception:
        pass

    try:
        r = session.get(
            "https://plvr.land.moi.gov.tw/DownloadSeason",
            params={"season": season_code, "type": "zip", "fileName": "lvr_landcsv.zip"},
            headers=headers, timeout=300
        )
        if r.status_code == 200 and len(r.content) > 50000:
            os.makedirs(DATA_FOLDER, exist_ok=True)
            with open(path, "wb") as f:
                f.write(r.content)
            return path
        else:
            log(f"  下載失敗 HTTP {r.status_code}")
            return None
    except Exception as e:
        log(f"  下載錯誤：{e}")
        return None


def read_presale_csv(zip_path):
    try:
        with zipfile.ZipFile(zip_path) as z:
            fnames = [n for n in z.namelist() if n.lower() == "b_lvr_land_b.csv"]
            if not fnames:
                return pd.DataFrame()
            with z.open(fnames[0]) as f:
                raw = f.read()
        for skip in [0, 1]:
            try:
                df = pd.read_csv(
                    io.BytesIO(raw), encoding="utf-8-sig",
                    skiprows=skip, low_memory=False
                )
                if "鄉鎮市區" in df.columns and "總價元" in df.columns:
                    return df
            except Exception:
                pass
    except Exception as e:
        log(f"  解壓失敗：{e}")
    return pd.DataFrame()


def filter_taichung(df):
    cleaned = pd.DataFrame()
    for orig, new in PRE_COLS.items():
        cleaned[new] = df[orig].values if orig in df.columns else ""
    cleaned = cleaned[pd.to_numeric(cleaned["總價元"], errors="coerce").notna()].copy()
    cleaned = cleaned[
        cleaned["鄉鎮市區"].astype(str).str.contains(
            "|".join(TAICHUNG_DISTRICTS), na=False)
    ].copy()
    return cleaned


def get_year_month(roc_date_str):
    try:
        s = str(roc_date_str).strip().split(".")[0]
        roc_y = int(s[:3])
        m = int(s[3:5])
        return f"{roc_y + 1911}-{m:02d}"
    except Exception:
        return ""


def save_to_db(df, season_code):
    if df.empty:
        return 0
    conn = sqlite3.connect(DB_PATH)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    inserted = 0
    for _, row in df.iterrows():
        ym = get_year_month(row.get("交易年月日", ""))
        addr = str(row.get("門牌", "")).strip()
        date = str(row.get("交易年月日", "")).strip().split(".")[0]
        price = row.get("總價元", None)
        name = str(row.get("建案名稱", "")).replace("?", "").strip()
        try:
            pd.to_numeric(price)
        except Exception:
            continue
        try:
            conn.execute(
                """INSERT OR IGNORE INTO presale_history
                   (年月, 季別, 鄉鎮市區, 交易標的, 門牌, 建物型態,
                    總價元, 單價元平方公尺, 建物移轉總面積平方公尺,
                    屋齡, 交易年月日, 建案名稱, 匯入時間)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (ym, season_code,
                 str(row.get("鄉鎮市區", "")),
                 str(row.get("交易標的", "")),
                 addr,
                 str(row.get("建物型態", "")),
                 float(price) if price not in (None, "") else None,
                 float(row["單價元平方公尺"]) if pd.notna(pd.to_numeric(row.get("單價元平方公尺"), errors="coerce")) else None,
                 float(row["建物移轉總面積平方公尺"]) if pd.notna(pd.to_numeric(row.get("建物移轉總面積平方公尺"), errors="coerce")) else None,
                 float(row["屋齡"]) if pd.notna(pd.to_numeric(row.get("屋齡"), errors="coerce")) else None,
                 date,
                 name,
                 now)
            )
            if conn.execute("SELECT changes()").fetchone()[0]:
                inserted += 1
        except Exception:
            pass
    conn.commit()
    conn.close()
    return inserted


def main():
    # 建立 presale_history 表
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS presale_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            年月 TEXT, 季別 TEXT,
            縣市 TEXT DEFAULT '臺中市',
            鄉鎮市區 TEXT, 交易標的 TEXT,
            門牌 TEXT, 建物型態 TEXT, 總價元 REAL,
            單價元平方公尺 REAL, 建物移轉總面積平方公尺 REAL,
            屋齡 REAL, 交易年月日 TEXT, 建案名稱 TEXT,
            匯入時間 TEXT,
            UNIQUE(門牌, 交易年月日, 總價元)
        )
    """)
    conn.commit()

    # 已匯入的季別
    done = {r[0] for r in conn.execute("SELECT DISTINCT 季別 FROM presale_history").fetchall()}
    conn.close()

    log(f"共 {len(SEASONS)} 季，已匯入 {len(done)} 季")

    total_inserted = 0
    for season in SEASONS:
        if season in done:
            log(f"{season}: 已匯入，略過")
            continue

        log(f"{season}: 下載中...")
        zip_path = download_season(season)
        if not zip_path:
            log(f"{season}: 下載失敗，略過")
            continue

        df_raw = read_presale_csv(zip_path)
        if df_raw.empty:
            log(f"{season}: 無預售屋資料")
            continue

        df = filter_taichung(df_raw)
        n = save_to_db(df, season)
        total_inserted += n
        log(f"{season}: 台中 {len(df)} 筆 → 新增 {n} 筆")

    # 最終統計
    conn = sqlite3.connect(DB_PATH)
    total = conn.execute("SELECT COUNT(*) FROM presale_history").fetchone()[0]
    with_addr = conn.execute("SELECT COUNT(*) FROM presale_history WHERE 門牌 != '' AND 門牌 IS NOT NULL").fetchone()[0]
    conn.close()
    log(f"\n完成！presale_history 共 {total} 筆，有地址 {with_addr} 筆")
    log(f"本次新增：{total_inserted} 筆")


if __name__ == "__main__":
    main()
