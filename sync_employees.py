"""從 HRD MySQL 同步員工資料到本地 SQLite"""
import configparser
import pymysql
from database import SessionLocal, engine, Base
from models import Employee

# 確保表存在
Base.metadata.create_all(bind=engine)


def sync():
    config = configparser.ConfigParser()
    config.read("config.ini", encoding="utf-8")

    mysql_cfg = config["mysql"]
    conn = pymysql.connect(
        host=mysql_cfg.get("host", "127.0.0.1"),
        port=int(mysql_cfg.get("port", "3307")),
        user=mysql_cfg["user"],
        password=mysql_cfg["password"],
        database=mysql_cfg["database"],
        charset="utf8mb4",
    )

    try:
        with conn.cursor() as cur:
            # 從 mdl_local_ct_employees + mdl_user 取工號和姓名
            cur.execute("""
                SELECT e.account_id,
                       CONCAT(u.lastname, u.firstname) AS name,
                       e.dept_code,
                       e.is_active
                FROM mdl_local_ct_employees e
                JOIN mdl_user u ON u.id = e.userid
                WHERE e.account_id IS NOT NULL
                  AND e.account_id != ''
            """)
            rows = cur.fetchall()
    finally:
        conn.close()

    db = SessionLocal()
    try:
        count = 0
        for account_id, name, dept_code, is_active in rows:
            existing = db.query(Employee).filter(Employee.account_id == account_id).first()
            if existing:
                existing.name = name
                existing.dept_code = dept_code
                existing.is_active = 1 if is_active else 0
            else:
                db.add(Employee(
                    account_id=account_id,
                    name=name,
                    dept_code=dept_code,
                    is_active=1 if is_active else 0,
                ))
            count += 1
        db.commit()
        print(f"同步完成：共 {count} 筆員工資料")
    finally:
        db.close()


if __name__ == "__main__":
    sync()
