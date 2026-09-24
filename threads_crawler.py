"""
Threads 輿情爬蟲
- 關鍵字搜尋：每日抓取熱門貼文
- 帳號監控：追蹤指定帳號最新貼文
- 執行：python threads_crawler.py
- 環境變數：THREADS_USERNAME / THREADS_PASSWORD
"""

import asyncio
import sqlite3
import os
import json
import re
import base64
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

DB_PATH = "real_estate.db"
COOKIES_FILE = "threads_cookies.json"

# ── 設定區（可自訂） ─────────────────────────────────────────

KEYWORDS = [
    # 全市關鍵字
    "台中房地產",
    "台中預售屋",
    "台中建案",
    "台中買房",
    "台中新成屋",
    # 各區關鍵字
    "台中西屯區",
    "台中南屯區",
    "台中北屯區",
    "台中豐原區",
    "台中大里區",
    "台中太平區",
    "台中烏日區",
    "台中北區房",
    "台中西區房",
]

WATCH_ACCOUNTS = [
    "ben_lin0621",
    "walkingwithmycorgi",
    "ensonhsu",
    "new_home666",
    "allen_newhouse",
    "pennybuyhouse",
]

MAX_POSTS_PER_KEYWORD = 30   # 每個關鍵字最多抓幾則
MAX_POSTS_PER_ACCOUNT = 20   # 每個帳號最多抓幾則

# ── 資料庫 ───────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS threads_posts (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            post_id     TEXT UNIQUE,
            帳號        TEXT,
            顯示名稱    TEXT,
            內容        TEXT,
            按讚數      INTEGER DEFAULT 0,
            回覆數      INTEGER DEFAULT 0,
            轉發數      INTEGER DEFAULT 0,
            觀看數      INTEGER DEFAULT 0,
            來源類型    TEXT,
            來源值      TEXT,
            貼文時間    TEXT,
            貼文連結    TEXT,
            爬取時間    TEXT
        )
    """)
    # 舊資料庫補欄位（若不存在）
    for col in ("轉發數", "觀看數"):
        try:
            conn.execute(f"ALTER TABLE threads_posts ADD COLUMN {col} INTEGER DEFAULT 0")
        except Exception:
            pass
    conn.commit()
    conn.close()


def save_posts(posts):
    if not posts:
        return 0
    conn = sqlite3.connect(DB_PATH)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    inserted = 0
    for p in posts:
        try:
            conn.execute("""
                INSERT OR IGNORE INTO threads_posts
                (post_id, 帳號, 顯示名稱, 內容, 按讚數, 回覆數, 轉發數, 觀看數,
                 來源類型, 來源值, 貼文時間, 貼文連結, 爬取時間)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                p.get("post_id"), p.get("帳號"), p.get("顯示名稱"),
                p.get("內容"), p.get("按讚數", 0), p.get("回覆數", 0),
                p.get("轉發數", 0), p.get("觀看數", 0),
                p.get("來源類型"), p.get("來源值"),
                p.get("貼文時間"), p.get("貼文連結"), now
            ))
            if conn.execute("SELECT changes()").fetchone()[0]:
                inserted += 1
        except Exception as e:
            print(f"  存入失敗：{e}")
    conn.commit()
    conn.close()
    return inserted


# ── 時間解析 ─────────────────────────────────────────────────

def parse_relative_time(text):
    """把 '2h', '3d', '1w' 等轉成 ISO 時間字串"""
    now = datetime.now()
    text = str(text).strip().lower()
    m = re.match(r"(\d+)\s*(s|m|h|d|w)", text)
    if not m:
        return now.strftime("%Y-%m-%d %H:%M:%S")
    val, unit = int(m.group(1)), m.group(2)
    delta = {
        "s": timedelta(seconds=val),
        "m": timedelta(minutes=val),
        "h": timedelta(hours=val),
        "d": timedelta(days=val),
        "w": timedelta(weeks=val),
    }.get(unit, timedelta())
    return (now - delta).strftime("%Y-%m-%d %H:%M:%S")


# ── Playwright ───────────────────────────────────────────────

async def load_cookies(context):
    # 優先從環境變數（base64 編碼的 THREADS_COOKIES secret）載入
    encoded = os.environ.get("THREADS_COOKIES", "")
    if encoded:
        try:
            cookies = json.loads(base64.b64decode(encoded).decode())
            await context.add_cookies(cookies)
            print("  使用 THREADS_COOKIES secret 登入")
            return True
        except Exception as e:
            print(f"  THREADS_COOKIES 解碼失敗：{e}")

    # fallback：本地 cookie 檔案
    if os.path.exists(COOKIES_FILE):
        with open(COOKIES_FILE) as f:
            cookies = json.load(f)
        await context.add_cookies(cookies)
        print("  使用本地 cookie 檔案登入")
        return True
    return False


async def save_cookies(context):
    cookies = await context.cookies()
    with open(COOKIES_FILE, "w") as f:
        json.dump(cookies, f)


async def is_logged_in(page):
    try:
        await page.goto("https://www.threads.net/", wait_until="domcontentloaded", timeout=20000)
        await asyncio.sleep(2)
        url = page.url
        return "login" not in url and "threads.net" in url
    except Exception:
        return False


async def login(page, context):
    print("  登入 Threads（透過 Instagram）...")
    username = os.environ.get("THREADS_USERNAME", "")
    password = os.environ.get("THREADS_PASSWORD", "")
    if not username or not password:
        raise ValueError("請設定環境變數 THREADS_USERNAME 和 THREADS_PASSWORD")

    # 直接走 Instagram 登入，登入後自動套用 Threads session
    await page.goto(
        "https://www.instagram.com/accounts/login/?next=https%3A%2F%2Fwww.threads.net%2F",
        wait_until="domcontentloaded", timeout=30000
    )
    await asyncio.sleep(4)

    # 等待 Instagram 登入表單出現
    try:
        await page.wait_for_selector('input[name="username"]', timeout=15000)
    except Exception:
        # 可能已登入或頁面結構不同
        current = page.url
        print(f"  頁面：{current[:60]}")
        if "instagram.com" not in current and "login" not in current:
            print("  可能已登入")
            await save_cookies(context)
            return

    await page.fill('input[name="username"]', username)
    await asyncio.sleep(1)
    await page.fill('input[name="password"]', password)
    await asyncio.sleep(1)
    await page.click('button[type="submit"]')

    await asyncio.sleep(8)
    await page.wait_for_load_state("domcontentloaded")

    current_url = page.url
    print(f"  登入後頁面：{current_url[:60]}")

    if "login" in current_url:
        raise RuntimeError("登入失敗，請確認帳號密碼")

    # 導回 Threads
    await page.goto("https://www.threads.net/", wait_until="domcontentloaded", timeout=20000)
    await asyncio.sleep(3)

    await save_cookies(context)
    print("  登入成功")


async def extract_posts_from_page(page):
    """從目前頁面擷取所有貼文資料"""
    posts = []
    try:
        # 等待貼文出現
        await page.wait_for_selector('article, [data-pressable-container]', timeout=10000)
    except Exception:
        return posts

    await asyncio.sleep(2)

    # 透過 JS 擷取貼文內容
    posts = await page.evaluate("""() => {
        const results = [];
        const seen = new Set();

        // 找所有可能是貼文的容器
        const articles = document.querySelectorAll('article');
        const containers = articles.length > 0 ? articles :
            document.querySelectorAll('[data-pressable-container="true"]');

        containers.forEach(el => {
            try {
                // 貼文連結（從 a href 找含 /post/ 的）
                const links = el.querySelectorAll('a[href*="/post/"]');
                if (links.length === 0) return;
                const link = 'https://www.threads.net' + links[0].getAttribute('href');
                const postId = link.split('/post/')[1]?.split('?')[0];
                if (!postId || seen.has(postId)) return;
                seen.add(postId);

                // 帳號（找 @username 格式的 a href）
                let account = '';
                const profileLinks = el.querySelectorAll('a[href^="/@"]');
                if (profileLinks.length > 0) {
                    account = profileLinks[0].getAttribute('href').replace('/@', '');
                }

                // 顯示名稱（通常是第一個粗體文字或 span）
                let displayName = '';
                const nameEl = el.querySelector('span[class*="username"], strong, h1, h2, [dir="auto"] span');
                if (nameEl) displayName = nameEl.textContent.trim();
                if (!displayName) displayName = account;

                // 內容（找最長的文字 div）
                let content = '';
                const textEls = el.querySelectorAll('[dir="auto"]');
                textEls.forEach(t => {
                    const txt = t.textContent.trim();
                    if (txt.length > content.length && txt.length > 10) {
                        content = txt;
                    }
                });

                // 時間（找 time 標籤）
                let timeText = '';
                const timeEl = el.querySelector('time');
                if (timeEl) {
                    timeText = timeEl.getAttribute('datetime') || timeEl.textContent.trim();
                }

                // 按讚數、回覆數、轉發數（依序取前三個非零數字）
                let likes = 0, replies = 0, reposts = 0;
                const btns = el.querySelectorAll('span[class*="count"], button span, svg + span');
                btns.forEach(b => {
                    const n = parseInt(b.textContent.replace(/[^0-9]/g, ''), 10);
                    if (!isNaN(n) && n > 0) {
                        if (likes === 0) likes = n;
                        else if (replies === 0) replies = n;
                        else if (reposts === 0) reposts = n;
                    }
                });

                // 觀看數（抓 "X人看過" 或 "X 人看過" 或 "X views"）
                let views = 0;
                const fullText = el.textContent || '';
                const viewMatch = fullText.match(/([0-9,，]+)\s*人看過/) ||
                                  fullText.match(/([0-9,，]+)\s*views?/i);
                if (viewMatch) {
                    views = parseInt(viewMatch[1].replace(/[,，]/g, ''), 10) || 0;
                }

                if (content) {
                    results.push({ postId, account, displayName, content, likes, replies, reposts, views, timeText, link });
                }
            } catch (e) {}
        });
        return results;
    }""")

    return posts


async def scroll_and_collect(page, max_posts):
    """滾動頁面並收集貼文，直到達到上限"""
    all_posts = []
    seen_ids = set()
    no_new_count = 0

    for _ in range(10):  # 最多滾動 10 次
        posts = await extract_posts_from_page(page)
        new = 0
        for p in posts:
            pid = p.get("postId", "")
            if pid and pid not in seen_ids:
                seen_ids.add(pid)
                all_posts.append(p)
                new += 1

        if new == 0:
            no_new_count += 1
            if no_new_count >= 2:
                break
        else:
            no_new_count = 0

        if len(all_posts) >= max_posts:
            break

        await page.evaluate("window.scrollBy(0, window.innerHeight * 2)")
        await asyncio.sleep(2)

    return all_posts[:max_posts]


async def search_keyword(page, keyword):
    """搜尋關鍵字，取熱門貼文"""
    print(f"  搜尋：{keyword}")
    url = f"https://www.threads.net/search?q={keyword}&serp_type=default"
    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(3)

    raw = await scroll_and_collect(page, MAX_POSTS_PER_KEYWORD)
    posts = []
    for p in raw:
        time_str = p.get("timeText", "")
        if "T" in time_str:  # ISO format
            post_time = time_str[:19].replace("T", " ")
        else:
            post_time = parse_relative_time(time_str)

        posts.append({
            "post_id":   p.get("postId", ""),
            "帳號":      p.get("account", ""),
            "顯示名稱":  p.get("displayName", ""),
            "內容":      p.get("content", ""),
            "按讚數":    p.get("likes", 0),
            "回覆數":    p.get("replies", 0),
            "轉發數":    p.get("reposts", 0),
            "觀看數":    p.get("views", 0),
            "來源類型":  "keyword",
            "來源值":    keyword,
            "貼文時間":  post_time,
            "貼文連結":  p.get("link", ""),
        })
    return posts


async def get_account_posts(page, account):
    """取得指定帳號的最新貼文"""
    print(f"  帳號：@{account}")
    await page.goto(f"https://www.threads.net/@{account}", wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(3)

    raw = await scroll_and_collect(page, MAX_POSTS_PER_ACCOUNT)
    posts = []
    for p in raw:
        time_str = p.get("timeText", "")
        if "T" in time_str:
            post_time = time_str[:19].replace("T", " ")
        else:
            post_time = parse_relative_time(time_str)

        posts.append({
            "post_id":   p.get("postId", ""),
            "帳號":      account,
            "顯示名稱":  p.get("displayName", ""),
            "內容":      p.get("content", ""),
            "按讚數":    p.get("likes", 0),
            "回覆數":    p.get("replies", 0),
            "轉發數":    p.get("reposts", 0),
            "觀看數":    p.get("views", 0),
            "來源類型":  "account",
            "來源值":    account,
            "貼文時間":  post_time,
            "貼文連結":  p.get("link", ""),
        })
    return posts


# ── 主程式 ───────────────────────────────────────────────────

async def main():
    init_db()
    print(f"[{datetime.now().strftime('%H:%M:%S')}] === Threads 輿情爬蟲開始 ===")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 900},
            locale="zh-TW",
        )
        page = await context.new_page()

        # 登入（優先用 cookie；有 cookie 就直接繼續，不 fallback 密碼登入）
        cookie_loaded = await load_cookies(context)
        if cookie_loaded:
            logged_in = await is_logged_in(page)
            if logged_in:
                print("  使用已儲存的登入狀態")
            else:
                print("  Cookie 已載入，繼續執行（headless 可能無法驗證登入狀態）")
        else:
            await login(page, context)

        total_new = 0

        # 關鍵字搜尋
        if KEYWORDS:
            print(f"\n── 關鍵字搜尋（{len(KEYWORDS)} 組）──")
            for kw in KEYWORDS:
                try:
                    posts = await search_keyword(page, kw)
                    n = save_posts(posts)
                    print(f"    {kw}：抓到 {len(posts)} 則，新增 {n} 筆")
                    total_new += n
                    await asyncio.sleep(3)
                except Exception as e:
                    print(f"    {kw} 失敗：{e}")

        # 帳號監控
        if WATCH_ACCOUNTS:
            print(f"\n── 帳號監控（{len(WATCH_ACCOUNTS)} 個）──")
            for acc in WATCH_ACCOUNTS:
                try:
                    posts = await get_account_posts(page, acc)
                    n = save_posts(posts)
                    print(f"    @{acc}：抓到 {len(posts)} 則，新增 {n} 筆")
                    total_new += n
                    await asyncio.sleep(3)
                except Exception as e:
                    print(f"    @{acc} 失敗：{e}")

        await browser.close()

    # 統計
    conn = sqlite3.connect(DB_PATH)
    total = conn.execute("SELECT COUNT(*) FROM threads_posts").fetchone()[0]
    conn.close()
    print(f"\n本次新增：{total_new} 筆，資料庫累計：{total} 筆")
    print(f"[{datetime.now().strftime('%H:%M:%S')}] === 完成 ===")


if __name__ == "__main__":
    asyncio.run(main())
