"""
Google Sheet 寫入器（月報版）
- 所有統計改用「成交年月」（YYYY-MM）分組
- 搭配 main_1.py 使用
"""

import re
import sqlite3
import json
import os
import time
import math
import pandas as pd
from datetime import datetime

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
except ImportError:
    print("請安裝：pip install google-auth google-api-python-client")
    exit(1)

DB_PATH = "real_estate.db"
SPREADSHEET_ID = "1pN9_h5Pqe6CewXs8WPULSNpW8tXUKj1h8nZgMneu4HE"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

TAICHUNG_DISTRICTS = [
    "中區", "東區", "南區", "西區", "北區",
    "西屯區", "南屯區", "北屯區",
    "豐原區", "東勢區", "大甲區", "清水區", "沙鹿區",
    "梧棲區", "后里區", "神岡區", "潭子區", "大雅區",
    "新社區", "石岡區", "外埔區", "大安區",
    "烏日區", "大肚區", "龍井區", "霧峰區",
    "太平區", "大里區", "和平區"
]

FIXED_SHEETS = ["總覽摘要", "預售屋總表", "月度統計摘要", "各區建案統計摘要", "月度趨勢",
                "成屋總表", "成屋月度統計", "2018-2025成交資料", "Threads輿情"]


def log(msg):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] {msg}")


def clean_val(v):
    if v is None:
        return ""
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""
    return v


def clean_rows(rows):
    return [[clean_val(c) for c in row] for row in rows]


def get_service():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise ValueError("找不到 GOOGLE_CREDENTIALS 環境變數")
    creds_info = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(
        creds_info, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    return service.spreadsheets()


def api_call(fn, *args, **kwargs):
    for attempt in range(5):
        try:
            result = fn(*args, **kwargs)
            time.sleep(1.5)
            return result
        except Exception as e:
            if "429" in str(e) or "Quota" in str(e):
                wait = (attempt + 1) * 20
                log(f"  限流，等待 {wait} 秒（第 {attempt+1} 次）")
                time.sleep(wait)
            else:
                raise
    raise Exception("超過重試次數")


def get_existing_sheets(sheets):
    meta = api_call(sheets.get(spreadsheetId=SPREADSHEET_ID).execute)
    return {s["properties"]["title"]: s["properties"]["sheetId"]
            for s in meta["sheets"]}


def ensure_all_sheets(sheets, titles, existing):
    to_create = [t for t in titles if t not in existing]
    if not to_create:
        log(f"所有 {len(titles)} 個 sheet 已存在")
        return existing
    requests_body = [{"addSheet": {"properties": {"title": t}}} for t in to_create]
    res = api_call(sheets.batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests_body}
    ).execute)
    new_sheets = {}
    for reply in res.get("replies", []):
        if "addSheet" in reply:
            props = reply["addSheet"]["properties"]
            new_sheets[props["title"]] = props["sheetId"]
    log(f"新建 {len(to_create)} 個 sheet：{to_create}")
    return {**existing, **new_sheets}


def batch_clear(sheets, titles):
    ranges = [f"'{t}'!A:ZZ" for t in titles]
    api_call(sheets.values().batchClear(
        spreadsheetId=SPREADSHEET_ID, body={"ranges": ranges}).execute)
    log(f"批次清空 {len(titles)} 個 sheet")


def batch_write(sheets, data_map):
    items = list(data_map.items())
    for i in range(0, len(items), 5):
        chunk = items[i:i+5]
        value_ranges = [
            {"range": f"'{title}'!A1", "values": clean_rows(rows)}
            for title, rows in chunk if rows
        ]
        if not value_ranges:
            continue
        api_call(sheets.values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": value_ranges}
        ).execute)
        log(f"  批次寫入 {len(value_ranges)} 個 sheet（第 {i//5+1} 批）")


# ── 資料載入 ─────────────────────────────────────────────

def _roc_to_ym(val):
    """民國日期 YYYMMDD → 西元年月 YYYY-MM"""
    try:
        if hasattr(val, 'year'):
            return f"{val.year}-{val.month:02d}"
        s = str(val).strip().replace("/", "").replace(".", "").replace("-", "")
        s = s.split(".")[0].zfill(7)
        roc_y = int(s[:3])
        m = int(s[3:5])
        return f"{roc_y + 1911}-{m:02d}"
    except Exception:
        return ""


def load_data():
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql(
            "SELECT * FROM presale ORDER BY 年月 DESC, 交易年月日 DESC", conn)
    except Exception:
        try:
            df = pd.read_sql(
                "SELECT * FROM presale ORDER BY 交易年月日 DESC", conn)
        except Exception as e:
            log(f"讀取失敗：{e}")
            df = pd.DataFrame()

    try:
        log_df = pd.read_sql(
            "SELECT 年月, 資料來源, 新增筆數, 匯入時間 FROM month_log ORDER BY 年月 DESC",
            conn)
    except Exception:
        log_df = pd.DataFrame()

    conn.close()

    if not df.empty:
        df["總價元"] = pd.to_numeric(df["總價元"], errors="coerce")
        df["單價元平方公尺"] = pd.to_numeric(df["單價元平方公尺"], errors="coerce")
        df["建物移轉總面積平方公尺"] = pd.to_numeric(
            df["建物移轉總面積平方公尺"], errors="coerce")
        df["總價萬"] = (df["總價元"] / 10000).round(1)
        df["單價萬坪"] = (df["單價元平方公尺"] * 3.3058 / 10000).round(2)
        df["面積坪"] = (df["建物移轉總面積平方公尺"] * 0.3025).round(1)

        if "年月" not in df.columns or df["年月"].isna().all():
            df["年月"] = df["交易年月日"].apply(_roc_to_ym)

        # 建案名稱含 ? 代表編碼問題，移除問號保留其餘文字
        if "建案名稱" in df.columns:
            df["建案名稱"] = df["建案名稱"].apply(
                lambda x: re.sub(r'[?-�]', '', str(x)).strip())

    return df, log_df


def _normalize_addr(addr):
    """地址標準化：全形轉半形、取到門牌號碼止，去除樓層後綴"""
    s = str(addr).strip()
    s = s.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    m = re.search(r'\d+號', s)
    return s[:m.end()] if m else s


def _street_key(addr):
    """取路段名稱（第一個數字前），例：臺中市北屯區光復路三段"""
    s = str(addr).strip()
    s = s.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    m = re.search(r'\d', s)
    return s[:m.start()].rstrip() if m else s


def build_address_name_lookup(df_presale):
    """從預售屋資料＋歷史資料建立兩層地址→建案名稱對照表
    回傳 (exact_lookup, street_lookup)：
      exact_lookup：精確比對（取到號碼）
      street_lookup：路名模糊比對（該路只有一個建案才納入）
    """
    exact_lookup = {}
    street_index = {}   # street_key → set(建案名稱)

    def _add_from_rows(rows):
        for addr, name in rows:
            name = str(name).strip()
            if not name or name == "nan":
                continue
            key = _normalize_addr(addr)
            if key:
                exact_lookup[key] = name
            sk = _street_key(addr)
            if sk:
                street_index.setdefault(sk, set()).add(name)

    if not df_presale.empty and "門牌" in df_presale.columns and "建案名稱" in df_presale.columns:
        _add_from_rows(df_presale[["門牌", "建案名稱"]].values.tolist())

    # 從 presale_history 補充歷史地址對照
    try:
        conn = sqlite3.connect(DB_PATH)
        hist_rows = conn.execute(
            "SELECT 門牌, 建案名稱 FROM presale_history WHERE 門牌 != '' AND 建案名稱 != ''"
        ).fetchall()
        conn.close()
        _add_from_rows(hist_rows)
        log(f"歷史地址對照：{len(hist_rows)} 筆，精確 lookup {len(exact_lookup)} 組")
    except Exception as e:
        log(f"presale_history 讀取失敗（略過）：{e}")

    # 路名唯一建案才納入模糊 lookup（多個建案的路名跳過，避免貼錯名）
    street_lookup = {sk: list(names)[0] for sk, names in street_index.items() if len(names) == 1}
    log(f"路名模糊 lookup：{len(street_lookup)} 條路（唯一建案）")

    return exact_lookup, street_lookup


def _lookup_name(addr, exact_lookup, street_lookup):
    """兩段式地址查詢：先精確比對，再路名模糊比對"""
    key = _normalize_addr(addr)
    if key in exact_lookup:
        return exact_lookup[key]
    sk = _street_key(addr)
    return street_lookup.get(sk, "")


def load_sale_data():
    """從 DB 讀取成屋資料，計算屋齡"""
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql(
            "SELECT 年月, 鄉鎮市區, 交易標的, 建物型態, 主要用途, 建築完成年月, "
            "門牌, 建物移轉總面積平方公尺, 總價元, 單價元平方公尺, 交易年月日, "
            "移轉層次, 總樓層數 FROM sale ORDER BY 年月, 鄉鎮市區",
            conn
        )
    except Exception:
        df = pd.DataFrame()
    conn.close()

    if df.empty:
        return df

    # 計算屋齡（建築完成年月為民國格式 YYYMMDD）
    current_roc_year = datetime.now().year - 1911
    def calc_age(val):
        try:
            s = str(val).strip().replace("/", "").replace("-", "").split(".")[0]
            if s in ("", "nan", "0"):
                return ""
            # 民國 YYYMMDD：7 位 = 3 位年（100 年後）；6 位 = 2 位年（99 年以前，前導零已去除）
            if len(s) == 7:
                build_roc_y = int(s[:3])
            elif len(s) == 6:
                build_roc_y = int(s[:2])
            else:
                return ""
            age = current_roc_year - build_roc_y
            return age if age >= 0 else ""
        except Exception:
            pass
        return ""
    df["屋齡"] = df["建築完成年月"].apply(calc_age)

    # 萬元換算
    df["總價萬"] = pd.to_numeric(df["總價元"], errors="coerce").div(10000).round(1)
    df["單價萬坪"] = pd.to_numeric(df["單價元平方公尺"], errors="coerce").mul(3.3058).div(10000).round(1)

    return df


# ── 各 Sheet 資料準備 ────────────────────────────────────

def build_summary(df, log_df):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    months_str = "、".join(log_df["年月"].tolist()) if not log_df.empty else "-"
    rows = [
        ["大台中預售屋實價登錄月報"],
        [f"資料月份：{months_str}"],
        [f"更新時間：{now}"],
        [f"累積總筆數：{len(df):,}"],
        [],
        ["各月匯入記錄"],
        ["成交年月", "資料來源", "新增筆數", "匯入時間"],
    ]
    for _, r in log_df.iterrows():
        rows.append([str(r.get("年月", "")), str(r.get("資料來源", "")),
                     int(r.get("新增筆數", 0)), str(r.get("匯入時間", ""))])
    rows.append([])

    if not df.empty:
        rows += [
            ["各行政區彙總（全期）"],
            ["行政區", "筆數", "均單（萬/坪）", "中位單（萬/坪）",
             "均總（萬）", "最高總（萬）", "最低總（萬）"],
        ]
        grp = df.groupby("鄉鎮市區").agg(
            筆數=("id", "count"),
            均單=("單價萬坪", "mean"),
            中位單=("單價萬坪", "median"),
            均總=("總價萬", "mean"),
            最高總=("總價萬", "max"),
            最低總=("總價萬", "min"),
        ).reset_index().sort_values("筆數", ascending=False)
        for _, r in grp.iterrows():
            rows.append([str(r["鄉鎮市區"]), int(r["筆數"]),
                         round(r["均單"], 2), round(r["中位單"], 2),
                         round(r["均總"], 1), round(r["最高總"], 1),
                         round(r["最低總"], 1)])
    return rows


def build_raw_data(df):
    if df.empty:
        return []
    cols_src = ["id", "年月", "縣市", "鄉鎮市區", "交易標的",
                "建案名稱", "門牌", "建物型態",
                "總價萬", "單價萬坪", "面積坪", "屋齡", "交易年月日", "匯入時間"]
    cols_dst = ["id", "成交年月", "縣市", "鄉鎮市區", "交易標的",
                "建案名稱", "門牌", "建物型態",
                "總價（萬）", "單價（萬/坪）", "面積（坪）", "屋齡", "交易年月日", "匯入時間"]
    available = [c for c in cols_src if c in df.columns]
    out = df[available].copy()
    out.columns = cols_dst[:len(available)]
    out = out.fillna("").astype(str)
    return [out.columns.tolist()] + out.values.tolist()


def build_month_summary(df):
    rows = [["成交年月", "行政區", "成交筆數",
             "均單（萬/坪）", "中位單（萬/坪）",
             "最高單（萬/坪）", "最低單（萬/坪）",
             "均總（萬）", "最高總（萬）", "最低總（萬）"]]
    if not df.empty and "年月" in df.columns:
        grp = df.groupby(["年月", "鄉鎮市區"]).agg(
            筆數=("id", "count"),
            均單=("單價萬坪", "mean"),
            中位單=("單價萬坪", "median"),
            最高單=("單價萬坪", "max"),
            最低單=("單價萬坪", "min"),
            均總=("總價萬", "mean"),
            最高總=("總價萬", "max"),
            最低總=("總價萬", "min"),
        ).reset_index().sort_values(["年月", "筆數"], ascending=[False, False])
        for _, r in grp.iterrows():
            rows.append([str(r["年月"]), str(r["鄉鎮市區"]), int(r["筆數"]),
                         round(r["均單"], 2), round(r["中位單"], 2),
                         round(r["最高單"], 2), round(r["最低單"], 2),
                         round(r["均總"], 1), round(r["最高總"], 1),
                         round(r["最低總"], 1)])
    return rows


def build_case_summary(df):
    rows = [["成交年月", "行政區", "建案名稱", "成交筆數",
             "均單（萬/坪）", "中位單（萬/坪）",
             "最高單（萬/坪）", "最低單（萬/坪）",
             "均總（萬）", "最高總（萬）", "最低總（萬）"]]
    if not df.empty and "年月" in df.columns:
        has_name = df[df["建案名稱"].fillna("").str.strip() != ""]
        if not has_name.empty:
            grp = has_name.groupby(["年月", "鄉鎮市區", "建案名稱"]).agg(
                筆數=("id", "count"),
                均單=("單價萬坪", "mean"),
                中位單=("單價萬坪", "median"),
                最高單=("單價萬坪", "max"),
                最低單=("單價萬坪", "min"),
                均總=("總價萬", "mean"),
                最高總=("總價萬", "max"),
                最低總=("總價萬", "min"),
            ).reset_index().sort_values(
                ["年月", "鄉鎮市區", "筆數"], ascending=[False, True, False])
            for _, r in grp.iterrows():
                rows.append([str(r["年月"]), str(r["鄉鎮市區"]),
                             str(r["建案名稱"]), int(r["筆數"]),
                             round(r["均單"], 2), round(r["中位單"], 2),
                             round(r["最高單"], 2), round(r["最低單"], 2),
                             round(r["均總"], 1), round(r["最高總"], 1),
                             round(r["最低總"], 1)])
    return rows


def build_monthly_trend(df):
    if df.empty or "年月" not in df.columns:
        return []
    valid = df[df["年月"].fillna("").str.match(r"\d{4}-\d{2}")].copy()
    if valid.empty:
        return []

    grp = valid.groupby("年月").agg(
        成交筆數=("id", "count"),
        均單=("單價萬坪", "mean"),
        中位單=("單價萬坪", "median"),
        最高單=("單價萬坪", "max"),
        最低單=("單價萬坪", "min"),
        均總=("總價萬", "mean"),
        中位總=("總價萬", "median"),
        最高總=("總價萬", "max"),
        最低總=("總價萬", "min"),
    ).reset_index().sort_values("年月")

    grp["筆數差"] = grp["成交筆數"].diff()
    grp["筆數變化率"] = (grp["成交筆數"].pct_change() * 100).round(1)

    def trend_label(diff, pct):
        try:
            if math.isnan(float(diff)):
                return "-"
            d = int(diff)
            p = round(float(pct), 1)
            if d > 0:   return f"↑ +{d} 筆（+{p}%）"
            elif d < 0: return f"↓ {d} 筆（{p}%）"
            else:       return "→ 持平"
        except Exception:
            return "-"

    rows = [
        ["大台中預售屋月度成交趨勢"],
        [f"資料期間：{grp['年月'].min()} ～ {grp['年月'].max()}，共 {len(grp)} 個月"],
        [],
        ["成交年月", "成交筆數", "較上月變化",
         "均單（萬/坪）", "中位單（萬/坪）", "最高單（萬/坪）", "最低單（萬/坪）",
         "均總（萬）", "中位總（萬）", "最高總（萬）", "最低總（萬）"],
    ]
    for _, r in grp.iterrows():
        rows.append([
            str(r["年月"]), int(r["成交筆數"]),
            trend_label(r["筆數差"], r["筆數變化率"]),
            round(r["均單"], 2), round(r["中位單"], 2),
            round(r["最高單"], 2), round(r["最低單"], 2),
            round(r["均總"], 1), round(r["中位總"], 1),
            round(r["最高總"], 1), round(r["最低總"], 1),
        ])
    return rows


def build_district(df, dist):
    ddf = df[df["鄉鎮市區"] == dist].copy()
    if ddf.empty:
        return []

    rows = []

    # ── 月度彙總 ──
    rows.append([f"【{dist}】月度成交統計"])
    rows.append(["成交年月", "成交筆數",
                 "均單（萬/坪）", "中位單（萬/坪）",
                 "最高單（萬/坪）", "最低單（萬/坪）",
                 "均總（萬）", "最高總（萬）", "最低總（萬）"])
    if "年月" in ddf.columns:
        mgrp = ddf.groupby("年月").agg(
            筆數=("id", "count"),
            均單=("單價萬坪", "mean"),
            中位單=("單價萬坪", "median"),
            最高單=("單價萬坪", "max"),
            最低單=("單價萬坪", "min"),
            均總=("總價萬", "mean"),
            最高總=("總價萬", "max"),
            最低總=("總價萬", "min"),
        ).reset_index().sort_values("年月", ascending=False)
        for _, r in mgrp.iterrows():
            rows.append([str(r["年月"]), int(r["筆數"]),
                         round(r["均單"], 2), round(r["中位單"], 2),
                         round(r["最高單"], 2), round(r["最低單"], 2),
                         round(r["均總"], 1), round(r["最高總"], 1),
                         round(r["最低總"], 1)])
    rows.append([])
    rows.append([])

    # ── 建案明細 ──
    rows.append([f"【{dist}】建案統計摘要（按成交月份）"])
    rows.append(["成交年月", "建案名稱", "成交筆數",
                 "均單（萬/坪）", "中位單（萬/坪）",
                 "最高單（萬/坪）", "最低單（萬/坪）",
                 "均總（萬）", "最高總（萬）", "最低總（萬）"])

    has_name = ddf[ddf["建案名稱"].fillna("").str.strip() != ""]
    if not has_name.empty and "年月" in has_name.columns:
        grp = has_name.groupby(["年月", "建案名稱"]).agg(
            筆數=("id", "count"),
            均單=("單價萬坪", "mean"),
            中位單=("單價萬坪", "median"),
            最高單=("單價萬坪", "max"),
            最低單=("單價萬坪", "min"),
            均總=("總價萬", "mean"),
            最高總=("總價萬", "max"),
            最低總=("總價萬", "min"),
        ).reset_index().sort_values(["年月", "筆數"], ascending=[False, False])
        for _, r in grp.iterrows():
            rows.append([str(r["年月"]), str(r["建案名稱"]), int(r["筆數"]),
                         round(r["均單"], 2), round(r["中位單"], 2),
                         round(r["最高單"], 2), round(r["最低單"], 2),
                         round(r["均總"], 1), round(r["最高總"], 1),
                         round(r["最低總"], 1)])
    else:
        rows.append(["（本區無建案名稱資料）"])

    rows.append([])
    rows.append([])
    rows.append([f"【{dist}】原始交易資料（共 {len(ddf)} 筆）"])
    rows.append(["成交年月", "建案名稱", "門牌", "建物型態",
                 "總價（萬）", "單價（萬/坪）", "面積（坪）", "屋齡", "交易年月日"])

    cols = ["年月", "建案名稱", "門牌", "建物型態",
            "總價萬", "單價萬坪", "面積坪", "屋齡", "交易年月日"]
    available = [c for c in cols if c in ddf.columns]
    out = ddf[available].copy().fillna("").astype(str)
    for _, r in out.iterrows():
        rows.append(r.tolist())

    return rows


def load_history_data():
    """從 presale_history 讀取 2018-2023 歷史預售屋資料"""
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql(
            "SELECT 年月, 季別, 鄉鎮市區, 交易標的, 建案名稱, 門牌, 建物型態, "
            "總價元, 單價元平方公尺, 建物移轉總面積平方公尺, 屋齡, 交易年月日 "
            "FROM presale_history ORDER BY 年月, 鄉鎮市區",
            conn
        )
    except Exception:
        df = pd.DataFrame()
    conn.close()
    if df.empty:
        return df
    df["總價萬"] = pd.to_numeric(df["總價元"], errors="coerce").div(10000).round(1)
    df["單價萬坪"] = pd.to_numeric(df["單價元平方公尺"], errors="coerce").mul(3.3058).div(10000).round(1)
    df["面積坪"] = pd.to_numeric(df["建物移轉總面積平方公尺"], errors="coerce").mul(0.3025).round(1)
    return df


def build_history_sheet(df_hist):
    """2018-2025成交資料 分頁"""
    rows = [["大台中預售屋歷史成交資料（2018-2023）"],
            ["年月", "季別", "鄉鎮市區", "交易標的", "建案名稱", "門牌", "建物型態",
             "總價（萬）", "單價（萬/坪）", "面積（坪）", "屋齡", "交易年月日"]]
    if df_hist.empty:
        rows.append(["（尚無歷史資料，請先執行 import_history.py）"])
        return rows
    cols = ["年月", "季別", "鄉鎮市區", "交易標的", "建案名稱", "門牌", "建物型態",
            "總價萬", "單價萬坪", "面積坪", "屋齡", "交易年月日"]
    available = [c for c in cols if c in df_hist.columns]
    out = df_hist[available].fillna("").astype(str)
    for _, r in out.iterrows():
        rows.append(r.tolist())
    return rows


def build_sale_raw(df_sale):
    """成屋總表：原始資料（只含屋齡 5 年以內）"""
    rows = [["大台中成屋（不動產買賣）總表（屋齡 5 年以內）"],
            ["年月", "鄉鎮市區", "交易標的", "建案名稱", "門牌", "建物型態", "主要用途",
             "屋齡", "坪數", "總價（萬）", "單價（萬/坪）", "移轉層次", "總樓層數", "交易年月日"]]
    if df_sale.empty:
        rows.append(["（尚無成屋資料）"])
        return rows
    df_sale = df_sale.copy()
    # 篩選屋齡 5 年以內（屋齡為空的排除）
    age = pd.to_numeric(df_sale["屋齡"], errors="coerce")
    df_sale = df_sale[age.notna() & (age >= 0) & (age <= 5)].copy()
    cols = ["年月", "鄉鎮市區", "交易標的", "建案名稱", "門牌", "建物型態", "主要用途",
            "屋齡", "面積坪", "總價萬", "單價萬坪", "移轉層次", "總樓層數", "交易年月日"]
    if "建物移轉總面積平方公尺" in df_sale.columns:
        df_sale["面積坪"] = pd.to_numeric(
            df_sale["建物移轉總面積平方公尺"], errors="coerce").mul(0.3025).round(1)
    if "建案名稱" not in df_sale.columns:
        df_sale["建案名稱"] = ""
    available = [c for c in cols if c in df_sale.columns]
    out = df_sale[available].fillna("").astype(str)
    for _, r in out.iterrows():
        rows.append(r.tolist())
    return rows


def build_sale_month_summary(df_sale):
    """成屋月度統計：按年月 × 行政區彙整"""
    rows = [["成屋月度統計摘要"],
            ["年月", "鄉鎮市區", "成交筆數", "均單（萬/坪）", "中位單（萬/坪）",
             "均總（萬）", "均屋齡（年）"]]
    if df_sale.empty:
        rows.append(["（尚無成屋資料）"])
        return rows
    df = df_sale.copy()
    df["單價萬坪_n"] = pd.to_numeric(df.get("單價萬坪", pd.Series(dtype=float)), errors="coerce")
    df["總價萬_n"]   = pd.to_numeric(df.get("總價萬",   pd.Series(dtype=float)), errors="coerce")
    df["屋齡_n"]     = pd.to_numeric(df.get("屋齡",     pd.Series(dtype=float)), errors="coerce")
    grp = df.groupby(["年月", "鄉鎮市區"]).agg(
        筆數=("總價萬_n", "count"),
        均單=("單價萬坪_n", "mean"),
        中位單=("單價萬坪_n", "median"),
        均總=("總價萬_n", "mean"),
        均屋齡=("屋齡_n", "mean"),
    ).reset_index().sort_values(["年月", "鄉鎮市區"], ascending=[False, True])
    for _, r in grp.iterrows():
        rows.append([
            str(r["年月"]), str(r["鄉鎮市區"]), int(r["筆數"]),
            round(r["均單"], 1) if pd.notna(r["均單"]) else "",
            round(r["中位單"], 1) if pd.notna(r["中位單"]) else "",
            round(r["均總"], 0) if pd.notna(r["均總"]) else "",
            round(r["均屋齡"], 1) if pd.notna(r["均屋齡"]) else "",
        ])
    return rows


# ── 主程式 ────────────────────────────────────────────────

def load_threads_data():
    """從 DB 載入 Threads 貼文（最近 90 天）"""
    try:
        conn = sqlite3.connect(DB_PATH)
        df = pd.read_sql("""
            SELECT 貼文時間, 來源值 AS 關鍵字, 帳號, 顯示名稱, 內容,
                   觀看數, 按讚數, 回覆數, 轉發數, 貼文連結, 來源類型
            FROM threads_posts
            WHERE 貼文時間 >= date('now', '-90 days')
            ORDER BY 觀看數 DESC, 按讚數 DESC, 貼文時間 DESC
        """, conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()


def build_threads_sheet(df_threads):
    rows = [["Threads 輿情監控（近 90 天）"],
            ["貼文時間", "來源關鍵字／帳號", "帳號", "顯示名稱",
             "內容", "觀看數", "按讚數", "回覆數", "轉發數", "貼文連結"]]
    if df_threads.empty:
        rows.append(["（尚無資料，請先執行 threads_crawler.py）"])
        return rows
    for _, r in df_threads.iterrows():
        rows.append([
            str(r.get("貼文時間", ""))[:16],
            str(r.get("關鍵字", "")),
            str(r.get("帳號", "")),
            str(r.get("顯示名稱", "")),
            str(r.get("內容", ""))[:500],
            int(r.get("觀看數", 0) or 0),
            int(r.get("按讚數", 0) or 0),
            int(r.get("回覆數", 0) or 0),
            int(r.get("轉發數", 0) or 0),
            str(r.get("貼文連結", "")),
        ])
    return rows


def run():
    log("=== 開始寫入 Google Sheet（月報版）===")
    sheets = get_service()
    df, log_df = load_data()
    log(f"載入 {len(df):,} 筆，{len(log_df)} 個月記錄")

    actual_districts = sorted(
        df["鄉鎮市區"].dropna().unique().tolist()) if not df.empty else []
    all_titles = FIXED_SHEETS + TAICHUNG_DISTRICTS
    log(f"共需 {len(all_titles)} 個 sheet（{len(TAICHUNG_DISTRICTS)} 個行政區）")

    existing = get_existing_sheets(sheets)
    existing = ensure_all_sheets(sheets, all_titles, existing)
    batch_clear(sheets, all_titles)  # 所有區分頁一律清空，無資料的留白

    df_sale = load_sale_data()
    log(f"成屋資料：{len(df_sale):,} 筆")

    # 用預售屋地址→建案名稱對照表補齊成屋建案名稱（含歷史資料，兩段式比對）
    exact_lookup, street_lookup = build_address_name_lookup(df)
    if not df_sale.empty and "門牌" in df_sale.columns:
        df_sale["建案名稱"] = df_sale["門牌"].apply(
            lambda x: _lookup_name(x, exact_lookup, street_lookup))
        matched = (df_sale["建案名稱"] != "").sum()
        log(f"成屋建案名稱比對：{matched}/{len(df_sale)} 筆成功")

    df_hist = load_history_data()
    log(f"歷史預售屋資料：{len(df_hist):,} 筆")

    log("準備資料中...")
    data_map = {
        "總覽摘要":           build_summary(df, log_df),
        "預售屋總表":         build_raw_data(df),
        "月度統計摘要":       build_month_summary(df),
        "各區建案統計摘要":   build_case_summary(df),
        "月度趨勢":           build_monthly_trend(df),
        "成屋總表":           build_sale_raw(df_sale),
        "成屋月度統計":       build_sale_month_summary(df_sale),
        "2018-2025成交資料":  build_history_sheet(df_hist),
        "Threads輿情":        build_threads_sheet(load_threads_data()),
    }
    for dist in actual_districts:
        data_map[dist] = build_district(df, dist)
        log(f"  準備 {dist}：{len(df[df['鄉鎮市區']==dist])} 筆")

    log("批次寫入中...")
    batch_write(sheets, data_map)

    log("=== Google Sheet 更新完成 ===")
    log(f"👉 https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")


if __name__ == "__main__":
    run()
