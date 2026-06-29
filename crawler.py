"""
預售屋建案備查爬蟲
從 lvr.land.moi.gov.tw 爬取建案名稱和地址
建立對照表供實價登錄資料比對使用
"""

import requests
import json
import sqlite3
import time
import os
from datetime import datetime

DB_PATH = "real_estate.db"

# 縣市代碼對應表
CITY_CODES = {
    "A": "臺北市",
    "B": "臺中市",
    "C": "基隆市",
    "D": "臺南市",
    "E": "高雄市",
    "F": "新北市",
    "G": "宜蘭縣",
    "H": "桃園市",
    "I": "嘉義市",
    "J": "新竹縣",
    "K": "苗栗縣",
    "M": "南投縣",
    "N": "彰化縣",
    "O": "新竹市",
    "P": "雲林縣",
    "Q": "嘉義縣",
    "R": "臺南市",
    "S": "高雄市",
    "T": "屏東縣",
    "U": "花蓮縣",
    "V": "臺東縣",
    "W": "金門縣",
    "X": "澎湖縣",
    "Z": "連江縣"
}

BASE_URL = "https://lvr.land.moi.gov.tw/SERVICE"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://lvr.land.moi.gov.tw/jsp/list.jsp",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-TW,zh;q=0.9",
}


def log(msg):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] {msg}")


def init_lookup_table():
    """建立建案名稱對照表"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS building_lookup (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            建案名稱 TEXT,
            地址 TEXT,
            縣市 TEXT,
            行政區 TEXT,
            建商 TEXT,
            更新時間 TEXT,
            UNIQUE(建案名稱, 地址)
        )
    """)
    conn.commit()
    conn.close()
    log("建案對照表初始化完成")


def get_districts(city_code):
    """取得縣市下的所有行政區"""
    try:
        url = f"{BASE_URL}/CITY/{city_code}/"
        resp = requests.get(url, headers=HEADERS, timeout=30)
        if resp.status_code == 200:
            data = resp.json()
            districts = []
            for item in data:
                if isinstance(item, dict) and item.get("use", False):
                    districts.append(item.get("code", ""))
            return districts
    except Exception as e:
        log(f"  取得行政區失敗：{e}")
    return []


def fetch_presale_cases(city_code, district_code):
    """
    爬取預售屋建案備查資料
    使用政府網站的建案查詢 API
    """
    cases = []
    try:
        # 建案備查查詢 API
        url = f"{BASE_URL}/PRESALE/{city_code}/{district_code}/"
        resp = requests.get(url, headers=HEADERS, timeout=30)

        if resp.status_code == 200:
            data = resp.json()
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, dict):
                        name = item.get("name", "").strip()
                        addr = item.get("addr", item.get("address", "")).strip()
                        mark = item.get("mark", "").strip()  # 建商
                        if name:
                            cases.append({
                                "建案名稱": name,
                                "地址": addr,
                                "建商": mark
                            })
    except Exception as e:
        log(f"    API 查詢失敗，嘗試備用方式：{e}")

    # 如果第一種方式失敗，嘗試備用 API
    if not cases:
        try:
            url2 = f"{BASE_URL}/MgrHtml/PRESALE/{city_code}/{district_code}/"
            resp2 = requests.get(url2, headers=HEADERS, timeout=30)
            if resp2.status_code == 200:
                data2 = resp2.json()
                if isinstance(data2, list):
                    for item in data2:
                        if isinstance(item, dict):
                            name = item.get("name", "").strip()
                            addr = item.get("addr", "").strip()
                            mark = item.get("mark", "").strip()
                            if name:
                                cases.append({
                                    "建案名稱": name,
                                    "地址": addr,
                                    "建商": mark
                                })
        except Exception as e2:
            log(f"    備用方式也失敗：{e2}")

    return cases


def save_cases(cases, city_name, district_code):
    """將建案資料存入資料庫"""
    if not cases:
        return 0

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    count = 0
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for case in cases:
        try:
            c.execute("""
                INSERT OR REPLACE INTO building_lookup
                (建案名稱, 地址, 縣市, 行政區, 建商, 更新時間)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                case["建案名稱"],
                case["地址"],
                city_name,
                district_code,
                case["建商"],
                now
            ))
            count += 1
        except Exception as e:
            pass

    conn.commit()
    conn.close()
    return count


def match_building_name_from_db(addr, city):
    """從資料庫查詢建案名稱"""
    if not addr:
        return ""
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()

        # 嘗試不同長度的地址前綴比對
        addr = str(addr).strip()
        for length in [12, 10, 8, 6]:
            if len(addr) >= length:
                prefix = addr[:length]
                c.execute("""
                    SELECT 建案名稱 FROM building_lookup
                    WHERE 地址 LIKE ? AND 縣市 = ?
                    LIMIT 1
                """, (f"{prefix}%", city))
                row = c.fetchone()
                if row:
                    conn.close()
                    return row[0]

        conn.close()
    except Exception as e:
        pass
    return ""


def update_presale_building_names():
    """更新預售屋資料表的建案名稱"""
    log("開始更新預售屋建案名稱...")
    conn = sqlite3.connect(DB_PATH)

    try:
        df_check = conn.execute("SELECT COUNT(*) FROM building_lookup").fetchone()
        total_lookup = df_check[0] if df_check else 0
        log(f"建案對照表共 {total_lookup} 筆")

        if total_lookup == 0:
            log("建案對照表是空的，請先執行爬蟲")
            conn.close()
            return

        # 取得所有預售屋資料
        rows = conn.execute("""
            SELECT id, 門牌, 縣市 FROM presale
            WHERE 建案名稱 IS NULL OR 建案名稱 = ''
        """).fetchall()

        log(f"需要比對建案名稱的預售屋：{len(rows)} 筆")
        updated = 0

        for row_id, addr, city in rows:
            name = match_building_name_from_db(addr, city)
            if name:
                conn.execute(
                    "UPDATE presale SET 建案名稱 = ? WHERE id = ?",
                    (name, row_id)
                )
                updated += 1

        conn.commit()
        log(f"成功比對建案名稱：{updated}/{len(rows)} 筆")

    except Exception as e:
        log(f"更新失敗：{e}")
    finally:
        conn.close()


def run_crawler():
    """執行完整爬蟲"""
    log("=== 開始爬取預售屋建案備查資料 ===")
    init_lookup_table()

    total = 0

    for city_code, city_name in CITY_CODES.items():
        log(f"處理縣市：{city_name}（{city_code}）")

        # 取得行政區列表
        districts = get_districts(city_code)
        if not districts:
            log(f"  {city_name} 無行政區資料，跳過")
            continue

        log(f"  共 {len(districts)} 個行政區")
        city_total = 0

        for district in districts:
            cases = fetch_presale_cases(city_code, district)
            if cases:
                count = save_cases(cases, city_name, district)
                city_total += count
                log(f"  {district}：{count} 筆建案")
            time.sleep(0.5)  # 避免請求過快

        log(f"  {city_name} 共 {city_total} 筆建案")
        total += city_total
        time.sleep(1)

    log(f"\n=== 爬蟲完成，共 {total} 筆建案 ===")

    # 爬完後自動更新預售屋建案名稱
    update_presale_building_names()


if __name__ == "__main__":
    import sys
    cmd = sys.argv[1] if len(sys.argv) > 1 else "crawl"

    if cmd == "crawl":
        run_crawler()
    elif cmd == "update":
        update_presale_building_names()
    elif cmd == "stats":
        conn = sqlite3.connect(DB_PATH)
        try:
            n = conn.execute("SELECT COUNT(*) FROM building_lookup").fetchone()[0]
            print(f"建案對照表：{n} 筆")
        except:
            print("建案對照表：0 筆")
        conn.close()
