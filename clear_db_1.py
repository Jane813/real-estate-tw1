"""清除資料庫，供強制重新匯入使用"""
import sqlite3

conn = sqlite3.connect("real_estate.db")

tables = ["presale", "month_log", "building_lookup"]
for table in tables:
    try:
        conn.execute(f"DELETE FROM {table}")
        print(f"已清除：{table}")
    except Exception:
        pass

conn.commit()
conn.close()
print("資料庫清除完成，準備重新匯入")
