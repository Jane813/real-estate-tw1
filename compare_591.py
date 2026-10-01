"""
591 市場行情 vs 本地 DB 建案名稱比對工具
- 在瀏覽器內呼叫 bff-market.591.com.tw/v1/search/match（帶 session cookie）
- 比對名稱差異 + 成交筆數差異
- 發現不同時自動建議寫入 name_corrections.json
執行：python3 compare_591.py [--diff-only] [--limit N]
"""

import argparse
import asyncio
import sqlite3
import json
import time
import urllib.parse
from playwright.async_api import async_playwright

DB_PATH = "real_estate.db"
CORRECTIONS_FILE = "name_corrections.json"
REGION_ID = 8   # 台中市

# ── DB ───────────────────────────────────────────────────────

def get_buildings_from_db():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("""
        SELECT 建案名稱, 鄉鎮市區, COUNT(*) as 筆數
        FROM presale
        WHERE 建案名稱 IS NOT NULL AND 建案名稱 != ''
        GROUP BY 建案名稱, 鄉鎮市區
        ORDER BY 筆數 DESC
    """).fetchall()
    conn.close()
    return [{"name": r[0], "district": r[1], "count": r[2]} for r in rows]

# ── 對照表 ───────────────────────────────────────────────────

def load_corrections():
    try:
        with open(CORRECTIONS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_corrections(entries):
    with open(CORRECTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)
    print(f"已寫入 {CORRECTIONS_FILE}（共 {len(entries)} 筆）")

# ── 591 API（透過瀏覽器 fetch）───────────────────────────────

_device_id = ""
_captured_results: dict = {}   # keyword → items

def _setup_request_capture(page):
    """攔截 591 真實 API request，取得 deviceid 並儲存搜尋結果"""
    async def req_handler(request):
        global _device_id
        if "search/match" in request.url and not _device_id:
            _device_id = request.headers.get("deviceid", "")

    async def resp_handler(response):
        try:
            if "search/match" not in response.url:
                return
            body = await response.json()
            if body.get("status") == 1:
                from urllib.parse import urlparse, parse_qs
                qs = parse_qs(urlparse(response.url).query)
                kw_decoded = qs.get("keyword", [""])[0]   # parse_qs 自動 URL-decode
                if kw_decoded:
                    _captured_results[kw_decoded] = body.get("data", {}).get("items", [])
        except Exception:
            pass

    page.on("request", req_handler)
    page.on("response", resp_handler)

async def search_591_via_browser(page, keyword, district=None):
    """觸發 591 頁面搜尋，從攔截的 API 回應取結果"""
    _captured_results.pop(keyword, None)

    # 找搜尋框，清空後輸入關鍵字
    inp = await page.query_selector('input[placeholder], input[type="text"], input[type="search"]')
    if inp:
        await inp.click(click_count=3)
        await inp.fill(keyword)
        await asyncio.sleep(1.5)   # 等 API 回應
    else:
        # fallback：直接跳轉搜尋頁
        encoded_kw = urllib.parse.quote(keyword)
        url = f"https://market.591.com.tw/?regionId={REGION_ID}&keyword={encoded_kw}"
        await page.goto(url, wait_until="domcontentloaded", timeout=20000)
        await asyncio.sleep(2)

    items = _captured_results.get(keyword, [])
    if district:
        items = [i for i in items if i.get("section") == district]
    return items

# ── 名稱匹配邏輯 ─────────────────────────────────────────────

def find_best_match(db_name, items):
    """
    DB 名稱是 591 名稱的子字串（政府縮寫 ⊆ 591 全名），且同地區
    排除只靠建商名稱相同的誤配（如「遠雄XX」配「遠雄YY」）
    """
    for item in items:
        s591 = item.get("name", "")
        if not s591:
            continue
        if db_name == s591:
            return item, "exact"
        if db_name in s591:
            # 確認不只是建商前綴相符（要求 DB 名稱至少 3 字 or 全都在 591 名中）
            return item, "db_in_591"
        if s591 in db_name:
            return item, "591_in_db"
    return None, None

# ── 主程式 ───────────────────────────────────────────────────

async def main(diff_only=False, limit=None, auto_write=False):
    buildings = get_buildings_from_db()
    if limit:
        buildings = buildings[:limit]

    corrections = load_corrections()
    existing_wrong = {(c["wrong"], c.get("district", "")) for c in corrections}
    new_suggestions = []

    print(f"DB 共 {len(buildings)} 個建案，開始與 591 比對...\n")
    matched = 0
    not_found = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            locale="zh-TW",
        )
        page = await context.new_page()
        _setup_request_capture(page)

        # 先開主頁並觸發一次搜尋，讓 591 JS 初始化並產生 deviceid
        print("初始化 591 session...")
        await page.goto("https://market.591.com.tw/?regionId=8", wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(3)
        # 觸發一次搜尋以取得 deviceid
        inp = await page.query_selector('input[placeholder], input[type="text"]')
        if inp:
            await inp.fill("台中")
            await asyncio.sleep(2)
        print(f"deviceid: {_device_id[:8]}..." if _device_id else "⚠️ 未取得 deviceid")
        print("開始比對\n")

        for bld in buildings:
            name = bld["name"]
            district = bld["district"]
            db_count = bld["count"]

            items = await search_591_via_browser(page, name, district)
            best, match_type = find_best_match(name, items)

            if not best:
                not_found.append(f"{district} {name}（DB:{db_count}）")
                if not diff_only:
                    print(f"  ✗ {district} {name}（DB:{db_count}）→ 591 查無")
                await asyncio.sleep(0.4)
                continue

            s591_name = best["name"]
            s591_count = best.get("price_num", "?")
            matched += 1

            if s591_name == name:
                if not diff_only:
                    print(f"  ✓ {district} {name}  DB:{db_count} / 591:{s591_count}")
            else:
                key = (name, district)
                already = key in existing_wrong
                tag = "（已在對照表）" if already else "⚠️ 新發現"
                print(f"  {tag}  {district} | {name} → {s591_name}  DB:{db_count} / 591:{s591_count}  [{match_type}]")
                if not already:
                    new_suggestions.append({
                        "wrong": name,
                        "correct": s591_name,
                        "district": district,
                        "note": f"591比對 DB:{db_count}筆 591:{s591_count}筆"
                    })

            await asyncio.sleep(0.4)

        await browser.close()

    print(f"\n{'='*55}")
    print(f"已比對：{matched}  查無：{len(not_found)}  新差異：{len(new_suggestions)}")

    if not_found:
        print(f"\n查無（591 上找不到）：")
        for s in not_found[:20]:
            print(f"  {s}")
        if len(not_found) > 20:
            print(f"  ...（共 {len(not_found)} 個）")

    if new_suggestions:
        print(f"\n建議新增到 {CORRECTIONS_FILE}：")
        for s in new_suggestions:
            print(f"  {s['district']} | {s['wrong']} → {s['correct']}")

        if auto_write:
            ans = "y"
        else:
            try:
                ans = input("\n要自動寫入嗎？(y/n): ").strip().lower()
            except EOFError:
                ans = "n"
        if ans == "y":
            corrections.extend(new_suggestions)
            save_corrections(corrections)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--diff-only", action="store_true", help="只顯示有差異的建案")
    parser.add_argument("--limit", type=int, help="只比對前 N 個建案（測試用）")
    parser.add_argument("--yes", action="store_true", help="自動寫入 name_corrections.json 不詢問")
    args = parser.parse_args()
    asyncio.run(main(diff_only=args.diff_only, limit=args.limit, auto_write=args.yes))
