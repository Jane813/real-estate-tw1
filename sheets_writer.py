"""
Google Sheet 寫入器（月報版）
- 所有統計改用「成交年月」（YYYY-MM）分組
- 搭配 main_1.py 使用
"""

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

FIXED_SHEETS = ["總覽摘要", "預售屋總表", "月度統計摘要", "各區建案統計摘要", "月度趨勢"]


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
    for i in range(0, len(items), 10):
        chunk = items[i:i+10]
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
        log(f"  批次寫入 {len(value_ranges)} 個 sheet（第 {i//10+1} 批）")


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

    return df, log_df


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


# ── 主程式 ────────────────────────────────────────────────

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

    log("準備資料中...")
    data_map = {
        "總覽摘要":         build_summary(df, log_df),
        "預售屋總表":       build_raw_data(df),
        "月度統計摘要":     build_month_summary(df),
        "各區建案統計摘要": build_case_summary(df),
        "月度趨勢":         build_monthly_trend(df),
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
