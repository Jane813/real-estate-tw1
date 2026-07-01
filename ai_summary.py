"""
實價登錄 AI 摘要產生器
讀取 Google Sheet 統計，透過 Gemini 2.0 Flash 產出每頁 PPT 分析摘要，
寫回 Google Sheet「AI摘要」頁
"""

import os, io, json, time
import urllib.parse
import requests
import pandas as pd
import google.generativeai as genai
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

SHEET_ID = "1pN9_h5Pqe6CewXs8WPULSNpW8tXUKj1h8nZgMneu4HE"
SCOPES   = ["https://www.googleapis.com/auth/spreadsheets"]


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def fetch_sheet(name, nrows=25):
    url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
           f"/export?format=csv&sheet={urllib.parse.quote(name)}")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return pd.read_csv(io.BytesIO(r.content), encoding="utf-8-sig").head(nrows)


def ask(model, prompt):
    try:
        resp = model.generate_content(prompt)
        time.sleep(2)
        return resp.text.strip()
    except Exception as e:
        log(f"Gemini 錯誤：{e}")
        return ""


def run():
    log("=== AI 摘要開始 ===")

    genai.configure(api_key=os.environ["GEMINI_API_KEY"])
    model = genai.GenerativeModel("gemini-2.0-flash")

    log("讀取 Google Sheet 資料...")
    try:
        df_overview = fetch_sheet("總覽摘要",      15)
        df_month    = fetch_sheet("月度統計摘要",  30)
        df_case     = fetch_sheet("各區建案統計摘要", 20)
        df_trend    = fetch_sheet("月度趨勢",       12)
    except Exception as e:
        log(f"讀取失敗：{e}")
        return

    INSTRUCTION = "請用繁體中文寫2句話（50字以內），語氣專業簡潔。"
    summaries = {}

    log("產生 Slide 1（整體概況）摘要...")
    summaries["slide1"] = ask(model,
        f"以下是大台中預售屋最新成交統計：\n{df_overview.to_string(index=False)}\n\n"
        f"{INSTRUCTION}說明整體市場量能與均價水準。"
    )

    log("產生 Slide 2（三大排行榜）摘要...")
    summaries["slide2"] = ask(model,
        f"以下是本期成交量、均單價、均總價前10名建案：\n{df_case.head(10).to_string(index=False)}\n\n"
        f"{INSTRUCTION}點出市場熱點區域與最活躍建案。"
    )

    log("產生 Slide 3（行政區分析）摘要...")
    summaries["slide3"] = ask(model,
        f"以下是各行政區成交筆數與單價統計：\n{df_month.to_string(index=False)}\n\n"
        f"{INSTRUCTION}分析各區冷熱差異與單價高低分布。"
    )

    log("產生 Slide 4（建物型態 & 量價）摘要...")
    summaries["slide4"] = ask(model,
        f"根據以下大台中預售屋成交資料：\n{df_case.head(10).to_string(index=False)}\n\n"
        f"{INSTRUCTION}說明主流建物型態與量價之間的關係。"
    )

    log("產生 Slide 5（月度趨勢）摘要...")
    summaries["slide5"] = ask(model,
        f"以下是大台中預售屋月度成交趨勢：\n{df_trend.to_string(index=False)}\n\n"
        f"{INSTRUCTION}解讀量能趨勢方向與均價走勢。"
    )

    for k, v in summaries.items():
        log(f"  {k}：{v[:60]}...")

    # ── 寫回 Google Sheet ────────────────────────────────────
    log("寫回 Google Sheet...")
    creds = service_account.Credentials.from_service_account_info(
        json.loads(os.environ["GOOGLE_CREDENTIALS"]), scopes=SCOPES)
    svc = build("sheets", "v4", credentials=creds).spreadsheets()

    # 確保 AI摘要 sheet 存在
    meta     = svc.get(spreadsheetId=SHEET_ID).execute()
    existing = {s["properties"]["title"] for s in meta["sheets"]}
    if "AI摘要" not in existing:
        svc.batchUpdate(spreadsheetId=SHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": "AI摘要"}}}]}
        ).execute()
        log("已建立 AI摘要 sheet")

    now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [
        ["AI摘要（Gemini 自動產生）"],
        [f"更新時間：{now}"],
        [],
        ["頁面",              "AI分析摘要"],
        ["slide1｜整體概況",  summaries.get("slide1", "")],
        ["slide2｜三大排行榜", summaries.get("slide2", "")],
        ["slide3｜行政區分析", summaries.get("slide3", "")],
        ["slide4｜建物型態量價", summaries.get("slide4", "")],
        ["slide5｜月度趨勢",  summaries.get("slide5", "")],
    ]
    svc.values().update(
        spreadsheetId=SHEET_ID,
        range="'AI摘要'!A1",
        valueInputOption="RAW",
        body={"values": rows}
    ).execute()

    log("=== AI 摘要已寫入 Google Sheet ===")


if __name__ == "__main__":
    run()
