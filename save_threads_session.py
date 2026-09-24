"""
本地執行一次，用有頭瀏覽器登入 Threads，儲存 Cookie 供 GitHub Actions 使用
執行：python save_threads_session.py
"""

import asyncio
import json
import base64
from playwright.async_api import async_playwright


async def main():
    print("開啟瀏覽器，請手動登入 Threads...")
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False)  # 有頭，讓你手動操作
        context = await browser.new_context(
            viewport={"width": 1280, "height": 900},
            locale="zh-TW",
        )
        page = await context.new_page()
        await page.goto("https://www.threads.net/login")

        print("\n請在瀏覽器中完成登入（包含手機驗證碼），登入成功看到首頁後按 Enter...")
        input()

        cookies = await context.cookies()
        with open("threads_cookies.json", "w") as f:
            json.dump(cookies, f)

        encoded = base64.b64encode(json.dumps(cookies).encode()).decode()
        print(f"\n✅ Cookie 已儲存到 threads_cookies.json")
        print(f"\n複製以下內容，到 GitHub → Settings → Secrets → New secret")
        print(f"名稱：THREADS_COOKIES")
        print(f"值（已 base64 編碼）：\n{encoded}")

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
