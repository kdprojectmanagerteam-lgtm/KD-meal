"""一次性腳本：從本地 SQLite 搬移資料到 Cloud SQL MySQL"""
import sqlite3
import pymysql

# Cloud SQL 連線資訊
MYSQL_HOST = "34.81.156.151"
MYSQL_USER = "mealuser"
MYSQL_PASS = "MealUser2026!"
MYSQL_DB = "meal_order"

# 本地 SQLite
SQLITE_PATH = "meal_order.db"


def migrate():
    # 連接 SQLite
    lite = sqlite3.connect(SQLITE_PATH)
    lite.row_factory = sqlite3.Row

    # 連接 MySQL
    my = pymysql.connect(
        host=MYSQL_HOST, user=MYSQL_USER, password=MYSQL_PASS,
        database=MYSQL_DB, charset="utf8mb4"
    )
    cur = my.cursor()

    tables = {
        "menu_items": [
            "id", "date", "type", "vendor", "phone", "name", "price",
            "category", "remarks", "sold_out", "start_time", "deadline"
        ],
        "orders": [
            "id", "emp_id", "emp_name", "date", "type", "floor",
            "item_name", "vendor", "qty", "price", "remarks",
            "company_name", "company_paid", "guest_dept", "self_paid",
            "proxy_by", "created_at"
        ],
        "drink_options": [
            "id", "date", "vendor", "option_type", "name", "price"
        ],
        "extra_fees": [
            "id", "date", "type", "description", "amount"
        ],
        "holidays": [
            "id", "date", "name"
        ],
    }

    for table, cols in tables.items():
        rows = lite.execute(f"SELECT * FROM {table}").fetchall()
        if not rows:
            print(f"  {table}: 0 筆，跳過")
            continue

        placeholders = ", ".join(["%s"] * len(cols))
        col_names = ", ".join([f"`{c}`" for c in cols])
        sql = f"INSERT INTO `{table}` ({col_names}) VALUES ({placeholders})"

        batch = []
        for row in rows:
            values = [row[c] for c in cols]
            batch.append(values)

        cur.executemany(sql, batch)
        my.commit()
        print(f"  {table}: {len(batch)} 筆完成")

    cur.close()
    my.close()
    lite.close()
    print("\n搬移完成！")


if __name__ == "__main__":
    migrate()
