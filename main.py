from datetime import datetime, timezone, timedelta
from io import BytesIO
import threading
import time
import logging
import asyncio
import json

TW_TZ = timezone(timedelta(hours=8))

# 工號→企微帳號對照（帳號不同的人）
WECHAT_ID_MAP = {
    "11020008": "chiahaohsu", "11020034": "CaiJiaWei-KeDingkd", "11020035": "jiachilin",
    "11020092": "ssupinyang", "11020183": "hsiaoyinchen", "11020220": "shuyuchang",
    "11020383": "TWCS0000", "11020805": "twbs002", "11020890": "INTCS0005",
    "11020926": "TWCS0004", "11020929": "TWCS0015", "11020931": "TWCS0008",
    "11021371": "INCS998", "11021382": "A0000056", "11021395": "11021395_6",
    "11110018": "shiyanghuang", "11110040": "yuchunchien", "11110078": "polongyu",
    "11110110": "chinweiwu", "11120010": "chienmaochiu", "11130014": "11130045",
    "12040003": "elemar", "12040011": "solekhan1", "12040012": "imamsapii",
    "12040018": "garry", "12040024": "valentino", "12040032": "kenn",
    "12040038": "alex", "12040052": "topher", "12040061": "melchor",
    "12040063": "ariel", "12040078": "juliu", "12040114": "vanlocnguyen",
    "12040124": "vanchinhpham", "12040136": "agussusilo", "12040137": "sugianto",
    "12040150": "oseas", "12040151": "ronell", "12040156": "jarwin",
    "12040169": "patrick", "12040179": "gyver", "12040182": "siswandi",
    "12040183": "teguh", "12040184": "nyono", "12040188": "sucipto",
    "12040208": "asep", "12040212": "junedi", "12040214": "eko",
    "12040218": "yahya", "12040220": "bayu", "12040377": "12040331",
    "12040395": "12040393", "12040407": "12040400", "20000009": "XuWeiZhen",
    "20100009": "yunanzhang", "20300003": "kd-XieMin", "20300021": "wangqingqing",
    "20500005": "mengcao", "20800023": "24000041", "20800049": "raoshihua",
    "20800132": "jx", "20800158": "20800152", "20800163": "20800018",
    "20800167": "zhangyan-z", "20800168": "shanshanwang", "21100002": "yanyongwang",
    "21100011": "20300037", "21200001": "wanlinwang", "21200010": "linjiqian",
    "22000001": "qinwen", "22000002": "20800151", "22000113": "dawang",
    "24000001": "suqin", "24090001": "KeDing-kd-XieLiRong", "36010018": "36010018_3",
    "A0000104": "a0000104", "CN0001": "20800164", "TW0001": "itadd",
}

import os
from pathlib import Path

from fastapi import FastAPI, Depends, HTTPException, Query, UploadFile, File, Form, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse, HTMLResponse
from sqlalchemy.orm import Session
from sqlalchemy import func
import re
import openpyxl

from database import Base, engine, get_db
from models import MenuItem, Order, Employee, Holiday, DrinkOption, ExtraFee, CoffeeItem, CoffeeOrder, DeadlineOverride, VendorNote, TreatNotification, OrderReview, MeetingEvent, MeetingAttendee, UniformApplication, UniformSize, UniformInventory, UniformInventoryLog, UniformRole
from schemas import (
    MenuItemCreate, MenuItemBatchImport, MenuItemOut,
    OrderCreate, OrderOut, OrderUpdate,
    ExtraFeeCreate, ExtraFeeOut,
    LoginRequest, ChangePasswordRequest, ForgotPasswordRequest,
    CoffeeItemCreate, CoffeeItemOut, CoffeeOrderCreate, CoffeeOrderOut,
    SetAdminRequest,
    UniformApplicationCreate, UniformApplicationOut, UniformApplicationUpdate, UniformSizeOut,
    UniformInventoryOut, UniformInventoryUpdate, UniformInventoryLogOut,
    UniformRoleCreate, UniformRoleOut,
)
import bcrypt

Base.metadata.create_all(bind=engine)

# Migration: 為既有 orders 表加上欄位（Cloud SQL 新建時表已包含所有欄位，此段僅供舊資料庫升級）
try:
    with engine.connect() as conn:
        from sqlalchemy import text, inspect
        columns = [c["name"] for c in inspect(engine).get_columns("orders")]
        if "vendor" not in columns:
            conn.execute(text("ALTER TABLE orders ADD COLUMN vendor VARCHAR(100)"))
            conn.commit()
        if "company_paid" not in columns:
            conn.execute(text("ALTER TABLE orders ADD COLUMN company_paid INTEGER NOT NULL DEFAULT 0"))
            conn.commit()
        if "self_paid" not in columns:
            conn.execute(text("ALTER TABLE orders ADD COLUMN self_paid INTEGER NOT NULL DEFAULT 0"))
            conn.commit()
        if "company_name" not in columns:
            conn.execute(text("ALTER TABLE orders ADD COLUMN company_name VARCHAR(100)"))
            conn.commit()
        if "guest_dept" not in columns:
            conn.execute(text("ALTER TABLE orders ADD COLUMN guest_dept VARCHAR(100)"))
            conn.commit()
        if "guest_qty" not in columns:
            conn.execute(text("ALTER TABLE orders ADD COLUMN guest_qty INTEGER"))
            conn.commit()
        emp_columns = [c["name"] for c in inspect(engine).get_columns("employees")]
        if "password_hash" not in emp_columns:
            conn.execute(text("ALTER TABLE employees ADD COLUMN password_hash VARCHAR(255)"))
            conn.commit()
        if "must_change_pwd" not in emp_columns:
            conn.execute(text("ALTER TABLE employees ADD COLUMN must_change_pwd INTEGER NOT NULL DEFAULT 1"))
            conn.commit()
        if "is_admin" not in emp_columns:
            conn.execute(text("ALTER TABLE employees ADD COLUMN is_admin INTEGER NOT NULL DEFAULT 0"))
            conn.commit()
            # 將現有管理員設為 is_admin=1
            for admin_id in ['TW0283', 'TW0498', '11020035', '11021038']:
                conn.execute(text("UPDATE employees SET is_admin=1 WHERE account_id=:id"), {"id": admin_id})
            conn.commit()
        menu_columns = [c["name"] for c in inspect(engine).get_columns("menu_items")]
        if "sold_out" not in menu_columns:
            conn.execute(text("ALTER TABLE menu_items ADD COLUMN sold_out INTEGER NOT NULL DEFAULT 0"))
            conn.commit()
        # 回填 vendor
        conn.execute(text("""
            UPDATE orders SET vendor = (
                SELECT m.vendor FROM menu_items m
                WHERE m.date = orders.date AND m.type = orders.type AND m.name = orders.item_name
                LIMIT 1
            ) WHERE vendor IS NULL OR vendor = ''
        """))
        conn.commit()
        # vendor_notes 加 date 欄位
        vn_columns = [c["name"] for c in inspect(engine).get_columns("vendor_notes")]
        if "date" not in vn_columns:
            conn.execute(text("ALTER TABLE vendor_notes ADD COLUMN date VARCHAR(10) NOT NULL DEFAULT ''"))
            conn.commit()
        if "confirmed_by" not in vn_columns:
            conn.execute(text("ALTER TABLE vendor_notes ADD COLUMN confirmed_by VARCHAR(50) NOT NULL DEFAULT ''"))
            conn.commit()
        # meeting_events 加 meal_type / deadline_time / check_orders
        me_columns = [c["name"] for c in inspect(engine).get_columns("meeting_events")]
        if "meal_type" not in me_columns:
            conn.execute(text("ALTER TABLE meeting_events ADD COLUMN meal_type VARCHAR(20)"))
            conn.commit()
        if "deadline_time" not in me_columns:
            conn.execute(text("ALTER TABLE meeting_events ADD COLUMN deadline_time VARCHAR(10)"))
            conn.commit()
        if "check_orders" not in me_columns:
            conn.execute(text("ALTER TABLE meeting_events ADD COLUMN check_orders INTEGER NOT NULL DEFAULT 1"))
            conn.commit()
        # meeting_attendees 新增 notified 欄位
        ma_columns = [c["name"] for c in inspect(engine).get_columns("meeting_attendees")]
        if "notified" not in ma_columns:
            conn.execute(text("ALTER TABLE meeting_attendees ADD COLUMN notified INTEGER NOT NULL DEFAULT 0"))
            conn.commit()
except Exception as e:
    logging.warning(f"Migration skipped or failed: {e}")

# 插入預設咖啡品項
try:
    from sqlalchemy.orm import Session as _Sess
    with _Sess(engine) as _s:
        if _s.query(CoffeeItem).count() == 0:
            _s.add_all([
                CoffeeItem(name="經典美式 Americano", price=25),
                CoffeeItem(name="濃郁拿鐵 Latté", price=35),
            ])
            _s.commit()
            logging.info("預設咖啡品項已插入")
except Exception as e:
    logging.warning(f"Coffee items init skipped: {e}")

# Migration: uniform_sizes 加新欄位
try:
    with engine.connect() as conn:
        from sqlalchemy import text, inspect
        if 'uniform_sizes' in inspect(engine).get_table_names():
            cols = [c["name"] for c in inspect(engine).get_columns("uniform_sizes")]
            if "size_code" not in cols:
                conn.execute(text("ALTER TABLE uniform_sizes ADD COLUMN size_code VARCHAR(10) NOT NULL DEFAULT ''"))
                conn.commit()
            for col_name in ['shoulder_width', 'chest_width', 'full_length', 'sleeve_length']:
                if col_name not in cols:
                    conn.execute(text(f"ALTER TABLE uniform_sizes ADD COLUMN {col_name} FLOAT"))
                    conn.commit()
except Exception as e:
    logging.warning(f"Uniform sizes migration skipped: {e}")

# 插入預設制服尺寸對照表
try:
    from sqlalchemy.orm import Session as _Sess2
    with _Sess2(engine) as _s:
        # 重建尺寸表（schema 有變更時）
        _s.query(UniformSize).delete()
        _s.commit()
        # 女版: (size_code, size_label, min_h, max_h, min_w, max_w, shoulder, chest, full_len, sleeve)
        female_sizes = [
            ('7','S', 148,158, 40,50, 35,43,58,55),
            ('9','M', 155,163, 48,58, 36,45,59,55),
            ('11','L', 160,168, 55,65, 38,49,59,55),
            ('13','XL', 165,173, 63,75, 41,54,60,56),
            ('15','2L', 168,178, 72,85, 41,55,60,57),
            ('18','3L', 175,999, 82,999, 47,58,60,60),
        ]
        male_sizes = [
            ('37','S', 160,168, 50,60, 38,48,71,61.5),
            ('38','M', 165,170, 58,65, 39,50,70,61),
            ('39','M', 168,173, 63,72, 41,53.5,73,61),
            ('40','L', 170,178, 68,78, 42,54,69.5,63),
            ('41','XL', 175,183, 75,88, 45,56,79,62),
            ('42','XL', 178,185, 85,95, 48,57.5,76.5,61.5),
            ('43','2L', 180,188, 90,105, 48,61,76.5,63),
            ('44','2L', 185,193, 100,115, 53,64,80,62),
            ('45','3L', 185,195, 108,125, 53,69,78,61),
            ('48','3L', 190,999, 120,999, 55,72,80,62),
        ]
        for code, label, minh, maxh, minw, maxw, sw, cw, fl, sl in female_sizes:
            _s.add(UniformSize(gender='female', size_code=code, size_label=label, min_height=minh, max_height=maxh, min_weight=minw, max_weight=maxw, shoulder_width=sw, chest_width=cw, full_length=fl, sleeve_length=sl))
        for code, label, minh, maxh, minw, maxw, sw, cw, fl, sl in male_sizes:
            _s.add(UniformSize(gender='male', size_code=code, size_label=label, min_height=minh, max_height=maxh, min_weight=minw, max_weight=maxw, shoulder_width=sw, chest_width=cw, full_length=fl, sleeve_length=sl))
        _s.commit()
        logging.info("制服尺寸對照表已更新")
except Exception as e:
    logging.warning(f"Uniform sizes init skipped: {e}")

# 每天自動同步員工資料
def _employee_sync_scheduler():
    logger = logging.getLogger("employee_sync")
    # 啟動時先同步一次
    try:
        from sync_employees import sync
        sync()
        logger.info("啟動時員工同步完成")
    except Exception as e:
        logger.warning(f"啟動時員工同步失敗：{e}")
    # 之後每天 06:00 同步
    while True:
        now = datetime.now(TW_TZ)
        # 計算到明天 06:00 的秒數
        tomorrow_6am = now.replace(hour=6, minute=0, second=0, microsecond=0)
        if now >= tomorrow_6am:
            from datetime import timedelta
            tomorrow_6am = tomorrow_6am + timedelta(days=1)
        wait_seconds = (tomorrow_6am - now).total_seconds()
        time.sleep(wait_seconds)
        try:
            from sync_employees import sync
            sync()
            logger.info("定時員工同步完成")
        except Exception as e:
            logger.warning(f"定時員工同步失敗：{e}")

_sync_thread = threading.Thread(target=_employee_sync_scheduler, daemon=True)
_sync_thread.start()

app = FastAPI(title="Meal Order System", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ==================== SSE for Coffee Notifications ====================

coffee_sse_clients: list[asyncio.Queue] = []

def broadcast_coffee_event(event_data: dict):
    """廣播咖啡狀態變更給所有 SSE 客戶端"""
    dead = []
    for q in coffee_sse_clients:
        try:
            q.put_nowait(event_data)
        except:
            dead.append(q)
    for q in dead:
        coffee_sse_clients.remove(q)

@app.get("/api/coffee/events")
async def coffee_sse(request: Request):
    """SSE 端點，推送咖啡訂單狀態變更"""
    q: asyncio.Queue = asyncio.Queue()
    coffee_sse_clients.append(q)

    async def event_stream():
        try:
            # 每 30 秒送一次 heartbeat 保持連線
            while True:
                try:
                    data = await asyncio.wait_for(q.get(), timeout=30.0)
                    yield f"data: {json.dumps(data, ensure_ascii=False)}\n\n"
                except asyncio.TimeoutError:
                    yield ": heartbeat\n\n"
                if await request.is_disconnected():
                    break
        finally:
            if q in coffee_sse_clients:
                coffee_sse_clients.remove(q)

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ==================== Auth ====================

DEFAULT_PASSWORD = "Kd123456"


def verify_password(plain: str, hashed: str | None) -> bool:
    if hashed is None:
        return plain == DEFAULT_PASSWORD
    return bcrypt.checkpw(plain.encode("utf-8"), hashed.encode("utf-8"))


def hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


@app.post("/api/login")
def login(req: LoginRequest, db: Session = Depends(get_db)):
    emp = db.query(Employee).filter(
        Employee.account_id == req.account_id,
        Employee.is_active == 1,
    ).first()
    if not emp or emp.account_id != req.account_id:
        return {"success": False, "message": "工號或密碼錯誤"}
    if not verify_password(req.password, emp.password_hash):
        return {"success": False, "message": "工號或密碼錯誤"}
    # 8/6 維護日：13:00 前只允許管理員登入
    now_tw = datetime.now(TW_TZ)
    if now_tw.strftime("%Y-%m-%d") == "2026-08-06" and now_tw.hour < 13 and not emp.is_admin:
        return {"success": False, "message": "訂餐系統即將於8/6(四)下午13:00開放使用\nThe meal ordering system will be open for use starting at 13:00 on Thursday, August 6."}
    return {
        "success": True,
        "account_id": emp.account_id,
        "name": emp.name,
        "dept_code": emp.dept_code,
        "must_change_pwd": bool(emp.must_change_pwd),
        "is_admin": bool(emp.is_admin),
    }


@app.post("/api/change-password")
def change_password(req: ChangePasswordRequest, db: Session = Depends(get_db)):
    emp = db.query(Employee).filter(
        Employee.account_id == req.account_id,
        Employee.is_active == 1,
    ).first()
    if not emp or not verify_password(req.old_password, emp.password_hash):
        return {"success": False, "message": "舊密碼錯誤"}
    if req.old_password == req.new_password:
        return {"success": False, "message": "新密碼不能與舊密碼相同"}
    if emp.password_hash and bcrypt.checkpw(req.new_password.encode("utf-8"), emp.password_hash.encode("utf-8")):
        return {"success": False, "message": "新密碼不能與舊密碼相同"}
    emp.password_hash = hash_password(req.new_password)
    emp.must_change_pwd = 0
    db.commit()
    return {"success": True, "message": "密碼已更新"}


@app.post("/api/forgot-password")
def forgot_password(req: ForgotPasswordRequest, db: Session = Depends(get_db)):
    emp = db.query(Employee).filter(
        Employee.account_id == req.account_id,
        Employee.name == req.name,
        Employee.is_active == 1,
    ).first()
    if not emp:
        return {"success": False, "message": "工號與姓名不符"}
    emp.password_hash = None
    emp.must_change_pwd = 1
    db.commit()
    return {"success": True, "message": "密碼已重設，請使用預設密碼登入"}


# ==================== Employee APIs ====================

@app.get("/api/employees/search")
def search_employee(q: str, db: Session = Depends(get_db)):
    """用工號查詢員工姓名"""
    emp = db.query(Employee).filter(
        Employee.account_id == q,
        Employee.is_active == 1,
    ).first()
    if emp and emp.account_id == q:
        return {"account_id": emp.account_id, "name": emp.name, "dept_code": emp.dept_code, "is_admin": bool(emp.is_admin)}
    # 員工表查不到時，從訂單紀錄找姓名
    order = db.query(Order).filter(Order.emp_id == q).order_by(Order.id.desc()).first()
    if order and order.emp_id == q:
        return {"account_id": q, "name": order.emp_name, "dept_code": ""}
    return {"account_id": q, "name": "", "dept_code": ""}


@app.get("/api/employees")
def list_employees(keyword: str = "", is_admin: int | None = None, db: Session = Depends(get_db)):
    """搜尋員工（支援工號或姓名模糊查詢，可篩選管理員）"""
    query = db.query(Employee).filter(Employee.is_active == 1)
    if is_admin is not None:
        query = query.filter(Employee.is_admin == is_admin)
    if keyword:
        query = query.filter(
            (Employee.account_id.like(f"%{keyword}%")) |
            (Employee.name.like(f"%{keyword}%"))
        )
    limit = 100
    if keyword and not is_admin:
        from sqlalchemy import case
        exact_first = case((Employee.name == keyword, 0), else_=1)
        query = query.order_by(exact_first, Employee.account_id)
    else:
        query = query.order_by(Employee.account_id)
    employees = query.limit(limit).all()
    result = []
    for e in employees:
        last_order = db.query(Order.floor).filter(
            Order.emp_id == e.account_id,
            Order.floor != None,
            Order.floor != '',
        ).order_by(Order.id.desc()).first()
        result.append({
            "account_id": e.account_id, "name": e.name,
            "dept_code": e.dept_code, "is_admin": bool(e.is_admin),
            "last_floor": last_order.floor if last_order else "",
        })
    return result


@app.post("/api/admin/set-admin")
def set_admin(req: SetAdminRequest, db: Session = Depends(get_db)):
    emp = db.query(Employee).filter(Employee.account_id == req.account_id, Employee.is_active == 1).first()
    if not emp:
        raise HTTPException(status_code=404, detail="員工不存在")
    emp.is_admin = 1 if req.is_admin else 0
    db.commit()
    return {"success": True}


SYNC_TOKEN = os.getenv("SYNC_TOKEN", "kd-meal-sync-2026")

@app.post("/api/sync-employees")
async def sync_employees(request: Request, db: Session = Depends(get_db)):
    """HRD 伺服器推送員工資料，帶 token 驗證"""
    auth = request.headers.get("Authorization", "")
    if auth != f"Bearer {SYNC_TOKEN}":
        raise HTTPException(status_code=401, detail="Unauthorized")

    body = await request.json()
    employees = body.get("employees", [])
    if not employees:
        raise HTTPException(status_code=400, detail="employees list is empty")

    updated = 0
    created = 0
    for emp in employees:
        account_id = emp.get("account_id", "").strip()
        name = emp.get("name", "").strip()
        if not account_id or not name:
            continue
        dept_code = emp.get("dept_code", "") or ""
        is_active = 1 if emp.get("is_active", True) else 0

        existing = db.query(Employee).filter(Employee.account_id == account_id).first()
        if existing:
            existing.name = name
            existing.dept_code = dept_code
            existing.is_active = is_active
            updated += 1
        else:
            db.add(Employee(
                account_id=account_id,
                name=name,
                dept_code=dept_code,
                is_active=is_active,
            ))
            created += 1

    db.commit()
    return {"ok": True, "created": created, "updated": updated, "total": created + updated}


# ==================== Menu APIs ====================

@app.get("/api/menu")
def get_menu(
    date: str | None = None,
    type: str | None = None,
    db: Session = Depends(get_db),
):
    q = db.query(MenuItem)
    if date:
        q = q.filter(MenuItem.date == date)
    if type:
        q = q.filter(MenuItem.type == type)
    items = q.order_by(MenuItem.date, MenuItem.type, MenuItem.category, MenuItem.name).all()
    return [{"id":m.id,"date":m.date,"type":m.type,"vendor":m.vendor,"phone":m.phone,"name":m.name,"price":m.price,"category":m.category,"remarks":m.remarks,"sold_out":m.sold_out or 0,"start_time":m.start_time,"deadline":m.deadline} for m in items]


@app.post("/api/menu", response_model=MenuItemOut)
def create_menu_item(item: MenuItemCreate, db: Session = Depends(get_db)):
    data = item.model_dump()
    data['name'] = data['name'].strip()
    if data.get('type') == '飲料' and data.get('deadline') and data['deadline'].endswith('T10:00'):
        data['deadline'] = data['deadline'].replace('T10:00', 'T10:10')
    db_item = MenuItem(**data)
    db.add(db_item)
    db.commit()
    db.refresh(db_item)
    return db_item


@app.post("/api/menu/batch", response_model=list[MenuItemOut])
def batch_import_menu(data: MenuItemBatchImport, db: Session = Depends(get_db)):
    created = []
    dl = data.deadline
    if data.type == '飲料' and dl and dl.endswith('T10:00'):
        dl = dl.replace('T10:00', 'T10:10')
    for item in data.items:
        db_item = MenuItem(
            date=data.date,
            type=data.type,
            start_time=data.start_time,
            deadline=dl,
            vendor=item.get("vendor", ""),
            phone=item.get("phone", ""),
            name=item.get("name", ""),
            price=int(item.get("price", 0)),
            category=item.get("category", ""),
        )
        db.add(db_item)
        created.append(db_item)
    db.commit()
    for item in created:
        db.refresh(item)
    return created


@app.post("/api/menu/upload")
async def upload_menu_excel(
    file: UploadFile = File(...),
    date: str = Form(...),
    meal_type: str = Form(..., alias="type"),
    start_time: str = Form(""),
    deadline: str = Form(""),
    db: Session = Depends(get_db),
):
    """上傳 Excel 匯入菜單。"""
    import re, logging
    from datetime import datetime as dt
    logger = logging.getLogger("uvicorn.error")
    try:
        return await _do_upload(file, date, meal_type, start_time, deadline, db)
    except Exception as e:
        logger.exception("Upload error")
        raise HTTPException(status_code=500, detail=str(e))


async def _do_upload(file, date, meal_type, start_time, deadline, db):
    import re
    from datetime import datetime as dt

    # 飲料截止時間強制 10:10
    if meal_type == '飲料' and deadline and deadline.endswith('T10:00'):
        deadline = deadline.replace('T10:00', 'T10:10')

    content = await file.read()
    try:
        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
    except Exception:
        raise HTTPException(status_code=400, detail="檔案格式錯誤，請上傳 Excel (.xlsx) 檔案")
    created = []
    skipped = 0
    errors = []

    # 預設截止時間對照
    deadline_defaults = {'早餐': 'T16:00', '午餐': 'T09:30', '飲料': 'T10:10', '晚餐': 'T14:30'}
    valid_types = {'早餐', '午餐', '飲料', '晚餐'}

    # 判斷要處理哪些 sheet
    sheets_to_process = []
    if len(wb.sheetnames) > 1:
        for sn in wb.sheetnames:
            sn_clean = sn.strip()
            if sn_clean in valid_types:
                sheets_to_process.append((wb[sn], sn_clean))
    if not sheets_to_process:
        sheets_to_process = [(wb.active, meal_type)]

    header_words = {"日期", "date", "分類", "category", "廠商", "店家", "vendor", "品項", "品項名稱", "name", "單價", "price", "備註", "remarks"}

    for ws, sheet_type in sheets_to_process:
        # 飲料 sheet 特殊處理
        if sheet_type == '飲料':
            import re as re2
            rows_list = list(ws.iter_rows(min_row=1, values_only=True))
            if not rows_list:
                continue
            # 第一列是標題，跳過
            header = rows_list[0]
            last_date_d = ""
            last_vendor_d = ""
            added_options = set()  # 追蹤同批次已新增的 drink_options，避免重複
            for row in rows_list[1:]:
                if not row or len(row) < 3:
                    continue
                # 日期
                raw_date = row[0]
                if hasattr(raw_date, 'strftime'):
                    rd = raw_date.strftime('%Y-%m-%d')
                else:
                    rd = str(raw_date or "").strip()
                if rd and re2.match(r"\d{4}-\d{2}-\d{2}", rd):
                    last_date_d = rd
                else:
                    rd = last_date_d
                # 店家
                vendor_d = str(row[1] or "").strip()
                if vendor_d:
                    last_vendor_d = vendor_d
                else:
                    vendor_d = last_vendor_d
                if not rd:
                    if tea_raw:
                        errors.append(f"飲料日期格式錯誤：'{raw_date}'，請使用 YYYY-MM-DD 格式")
                    continue
                if not vendor_d:
                    if tea_raw:
                        errors.append(f"飲料缺少店家：日期 {rd} 品項 '{tea_raw}' 無店家資訊")
                    continue
                # 茶類（品項+價格）
                tea_raw = str(row[2] or "").strip()
                if tea_raw:
                    price_match = re2.search(r'\$(\d+)', tea_raw)
                    tea_price = int(price_match.group(1)) if price_match else 0
                    tea_name = tea_raw
                    # 檢查重複
                    clean = re2.sub(r"^\(\d+-\d+\)\s*", "", tea_name)
                    existing_drinks = db.query(MenuItem).filter(
                        MenuItem.date == rd, MenuItem.type == '飲料'
                    ).all()
                    is_dup = any(re2.sub(r"^\(\d+-\d+\)\s*", "", e.name) == clean for e in existing_drinks)
                    if not is_dup:
                        item = MenuItem(
                            date=rd, type='飲料', vendor=vendor_d, name=tea_name,
                            price=tea_price, category='飲品',
                            start_time=start_time or None,
                            deadline=f"{rd}T10:10",
                        )
                        db.add(item)
                        created.append(item)
                # 加料
                topping_raw = str(row[3] or "").strip() if len(row) > 3 else ""
                if topping_raw:
                    tp_key = (rd, vendor_d, 'topping', topping_raw)
                    tp_match = re2.search(r'\$(\d+)', topping_raw)
                    tp_price = int(tp_match.group(1)) if tp_match else 0
                    if tp_key not in added_options:
                        existing_tp = db.query(DrinkOption).filter(
                            DrinkOption.date == rd, DrinkOption.vendor == vendor_d,
                            DrinkOption.option_type == 'topping', DrinkOption.name == topping_raw
                        ).first()
                        if not existing_tp:
                            db.add(DrinkOption(date=rd, vendor=vendor_d, option_type='topping', name=topping_raw, price=tp_price))
                        added_options.add(tp_key)
                # 甜度
                sweet_raw = str(row[4] or "").strip() if len(row) > 4 else ""
                if sweet_raw:
                    sw_key = (rd, vendor_d, 'sweetness', sweet_raw)
                    if sw_key not in added_options:
                        existing_sw = db.query(DrinkOption).filter(
                            DrinkOption.date == rd, DrinkOption.vendor == vendor_d,
                            DrinkOption.option_type == 'sweetness', DrinkOption.name == sweet_raw
                        ).first()
                        if not existing_sw:
                            db.add(DrinkOption(date=rd, vendor=vendor_d, option_type='sweetness', name=sweet_raw, price=0))
                        added_options.add(sw_key)
                # 冰塊
                ice_raw = str(row[5] or "").strip() if len(row) > 5 else ""
                if ice_raw:
                    ice_key = (rd, vendor_d, 'ice', ice_raw)
                    if ice_key not in added_options:
                        existing_ice = db.query(DrinkOption).filter(
                            DrinkOption.date == rd, DrinkOption.vendor == vendor_d,
                            DrinkOption.option_type == 'ice', DrinkOption.name == ice_raw
                        ).first()
                        if not existing_ice:
                            db.add(DrinkOption(date=rd, vendor=vendor_d, option_type='ice', name=ice_raw, price=0))
                        added_options.add(ice_key)
            continue

        # 判斷是否為多日格式
        is_weekly = False
        for row in ws.iter_rows(min_row=1, max_row=2, values_only=True):
            if not row:
                continue
            first_col = str(row[0] or "").strip()
            if first_col in ("日期", "date"):
                is_weekly = True
                break
            if hasattr(row[0], 'strftime'):
                is_weekly = True
                break
            if re.match(r"\d{4}-\d{2}-\d{2}", first_col):
                is_weekly = True
                break

        last_date = ""
        last_category = ""
        last_phone = ""
        last_vendor = ""

        # 動態偵測標題列欄位位置
        col_map = {"date": -1, "category": -1, "vendor": -1, "name": -1, "phone": -1}
        first_row = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        if first_row and first_row[0]:
            header = [str(c or "").strip() for c in first_row[0]]
            for i, h in enumerate(header):
                if h in ("日期", "date"):
                    col_map["date"] = i
                elif h in ("分類", "category"):
                    col_map["category"] = i
                elif h in ("電話", "phone", "tel"):
                    col_map["phone"] = i
                elif h in ("店家", "廠商", "vendor"):
                    col_map["vendor"] = i
                elif h in ("品項", "品項名稱", "name"):
                    col_map["name"] = i
            # 如果沒偵測到，用預設順序
            if col_map["date"] == -1: col_map["date"] = 0
            if col_map["name"] == -1: col_map["name"] = len(header) - 1
            if col_map["vendor"] == -1 and col_map["category"] == -1: col_map["vendor"] = 1

        for row in ws.iter_rows(min_row=1, values_only=True):
            if not row:
                continue
            db_item = None

            if is_weekly:
                if len(row) < 3:
                    continue
                raw_date = row[col_map["date"]] if col_map["date"] < len(row) else None
                if hasattr(raw_date, 'strftime'):
                    row_date = raw_date.strftime('%Y-%m-%d')
                else:
                    row_date = str(raw_date or "").strip()
                category = str(row[col_map["category"]] or "").strip() if col_map["category"] >= 0 and col_map["category"] < len(row) else ""
                phone_val = str(row[col_map["phone"]] or "").strip() if col_map["phone"] >= 0 and col_map["phone"] < len(row) else ""
                vendor = str(row[col_map["vendor"]] or "").strip() if col_map["vendor"] < len(row) else ""
                name = str(row[col_map["name"]] or "").strip() if col_map["name"] < len(row) else ""
                # 從品名提取價格 $XXX
                price = 0
                price_match = re.search(r'\$(\d+)', name)
                if price_match:
                    price = int(price_match.group(1))
                # 其他欄位：如果有數字當價格（覆蓋），其他當備註
                item_remarks = ""
                used_cols = set(col_map.values())
                for ci in range(len(row)):
                    if ci in used_cols:
                        continue
                    val = row[ci]
                    if val is None:
                        continue
                    try:
                        price = int(float(val))
                    except (ValueError, TypeError):
                        if str(val).strip() and str(val).strip() not in header_words:
                            item_remarks = str(val).strip()

                if row_date in header_words or category in header_words or vendor in header_words or name in header_words:
                    continue
                if row_date and re.match(r"\d{4}-\d{2}-\d{2}", row_date):
                    last_date = row_date
                elif row_date and row_date not in header_words and not last_date:
                    errors.append(f"日期格式錯誤：'{row_date}'，請使用 YYYY-MM-DD 格式")
                    continue
                else:
                    row_date = last_date
                if category:
                    last_category = category
                else:
                    category = last_category
                if phone_val:
                    last_phone = phone_val
                else:
                    phone_val = last_phone
                if vendor:
                    last_vendor = vendor
                else:
                    vendor = last_vendor
                if not row_date:
                    errors.append(f"缺少日期：品項 '{name}' 無法判斷日期")
                    continue
                if not name:
                    continue
                if not vendor:
                    errors.append(f"缺少店家：日期 {row_date} 品項 '{name}' 無店家資訊")
                    continue
                dl_suffix = deadline_defaults.get(sheet_type, 'T09:30')
                # 早餐截止是前一天
                if sheet_type == '早餐':
                    from datetime import timedelta
                    prev = dt.strptime(row_date, '%Y-%m-%d') - timedelta(days=1)
                    while prev.weekday() >= 5:
                        prev -= timedelta(days=1)
                    row_deadline = prev.strftime('%Y-%m-%d') + dl_suffix
                else:
                    row_deadline = row_date + dl_suffix
                # 早餐飲品/湯自動補 temp:cold,hot
                if sheet_type == '早餐' and category in ('飲品', '飲品/湯') and not item_remarks:
                    item_remarks = 'temp:cold,hot|upgrade:0'
                db_item = MenuItem(
                    date=row_date, type=sheet_type, start_time=start_time or None,
                    deadline=row_deadline, category=category, vendor=vendor, phone=phone_val or None,
                    name=name, price=price, remarks=item_remarks or None,
                )
            else:
                if len(row) < 3:
                    continue
                category = str(row[0] or "").strip()
                vendor = str(row[1] or "").strip()
                name = str(row[2] or "").strip()
                if category in header_words or vendor in header_words or name in header_words:
                    continue
                if category:
                    last_category = category
                else:
                    category = last_category
                if vendor:
                    last_vendor = vendor
                else:
                    vendor = last_vendor
                if not name:
                    continue
                # 從品名提取價格 $XXX
                price = 0
                price_match2 = re.search(r'\$(\d+)', name)
                if price_match2:
                    price = int(price_match2.group(1))
                item_remarks = ""
                for ci in range(3, len(row)):
                    val = row[ci]
                    if val is None:
                        continue
                    try:
                        price = int(float(val))
                    except (ValueError, TypeError):
                        if str(val).strip():
                            item_remarks = str(val).strip()
                db_item = MenuItem(
                    date=date, type=sheet_type, start_time=start_time or None,
                    deadline=deadline or None, category=category, vendor=vendor, phone=None,
                    name=name, price=price, remarks=item_remarks or None,
                )

            if db_item is None:
                continue
            # 檢查重複：同日期+同類型+同品名（忽略前面的編號和emoji比對）
            clean_name = re.sub(r"^\(\d+-\d+\)\s*", "", db_item.name)
            clean_name = re.sub(r"^[\U0001F96C\U0001F95A]\s*", "", clean_name)
            existing = db.query(MenuItem).filter(
                MenuItem.date == db_item.date,
                MenuItem.type == db_item.type,
            ).all()
            is_dup = False
            for e in existing:
                e_clean = re.sub(r"^\(\d+-\d+\)\s*", "", e.name)
                e_clean = re.sub(r"^[\U0001F96C\U0001F95A]\s*", "", e_clean)
                if e_clean == clean_name:
                    is_dup = True
                    break
            if is_dup:
                skipped += 1
                continue

            db.add(db_item)
            created.append(db_item)

    # 自動編號：午餐按分類（飯=1, 麵=2, 特=3, 素=4），晚餐按店家順序（第一家=1, 第二家=2）
    cat_prefix = {'飯': 1, '麵': 2, '特': 3, '素': 4}
    from collections import defaultdict

    # 先對晚餐品項按日期+店家分組，確定店家順序
    dinner_vendor_order = defaultdict(dict)  # date -> {vendor: vendor_num}
    for item in created:
        if item.type == '晚餐':
            if item.vendor not in dinner_vendor_order[item.date]:
                # 查已存在的該日期晚餐有幾個店家
                if not dinner_vendor_order[item.date]:
                    existing_vendors = db.query(MenuItem.vendor).filter(
                        MenuItem.date == item.date,
                        MenuItem.type == '晚餐',
                        ~MenuItem.id.in_([i.id for i in created]),
                    ).distinct().all()
                    for idx, (v,) in enumerate(existing_vendors, 1):
                        dinner_vendor_order[item.date][v] = idx
                next_num = len(dinner_vendor_order[item.date]) + 1
                dinner_vendor_order[item.date][item.vendor] = next_num

    date_type_cat_counter = defaultdict(lambda: defaultdict(int))
    dinner_vendor_counter = defaultdict(lambda: defaultdict(int))  # date -> {vendor: seq}

    for item in created:
        # 先移除品名中已有的編號
        item.name = re.sub(r"^\(\d+-\d+\)\s*", "", item.name)

        if item.type == '晚餐':
            # 晚餐：按店家順序編號
            vendor_num = dinner_vendor_order[item.date].get(item.vendor, 1)
            vkey = (item.date, item.vendor)
            if dinner_vendor_counter[item.date][item.vendor] == 0:
                existing_count = db.query(MenuItem).filter(
                    MenuItem.date == item.date,
                    MenuItem.type == '晚餐',
                    MenuItem.vendor == item.vendor,
                    ~MenuItem.id.in_([i.id for i in created]),
                ).count()
                dinner_vendor_counter[item.date][item.vendor] = existing_count
            dinner_vendor_counter[item.date][item.vendor] += 1
            seq = dinner_vendor_counter[item.date][item.vendor]
            item.name = f"({vendor_num}-{seq}){item.name}"
        else:
            # 午餐等：按分類編號
            cat_num = cat_prefix.get(item.category, 0)
            if cat_num:
                key = item.date + item.type
                if date_type_cat_counter[key][item.category] == 0:
                    existing_count = db.query(MenuItem).filter(
                        MenuItem.date == item.date,
                        MenuItem.type == item.type,
                        MenuItem.category == item.category,
                        MenuItem.id != item.id,
                    ).count()
                    date_type_cat_counter[key][item.category] = existing_count
                date_type_cat_counter[key][item.category] += 1
                seq = date_type_cat_counter[key][item.category]
                item.name = f"({cat_num}-{seq}){item.name}"

    # 去重 errors
    errors = list(dict.fromkeys(errors))

    if not created and errors:
        raise HTTPException(status_code=400, detail="匯入失敗：\n" + "\n".join(errors))

    if not created and not errors:
        raise HTTPException(status_code=400, detail="匯入 0 筆：檔案中無有效資料，請確認格式是否正確")

    db.commit()
    for item in created:
        db.refresh(item)
    msg = f"新增 {len(created)} 筆"
    if skipped > 0:
        msg += f"，跳過 {skipped} 筆重複項目"
    if errors:
        msg += f"\n⚠ 部分資料有問題：\n" + "\n".join(errors)
    return {"count": len(created), "skipped": skipped, "message": msg, "errors": errors, "items": [{"date": i.date, "category": i.category, "vendor": i.vendor, "name": i.name, "price": i.price} for i in created]}


@app.put("/api/menu/{item_id}")
def update_menu_item(item_id: int, data: dict, db: Session = Depends(get_db)):
    item = db.query(MenuItem).filter(MenuItem.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Menu item not found")
    old_name = item.name
    old_price = item.price
    if "name" in data and data["name"] and str(data["name"]).strip():
        item.name = str(data["name"]).strip()
    if "price" in data:
        item.price = int(data["price"])
    if "vendor" in data:
        item.vendor = data["vendor"]
    if "sold_out" in data:
        item.sold_out = int(data["sold_out"])
        # 飲料店家全部 sold_out 時，自動清除該店家的 drink_options
        if int(data["sold_out"]) == 1 and item.type == "飲料":
            remaining = db.query(MenuItem).filter(
                MenuItem.date == item.date,
                MenuItem.vendor == item.vendor,
                MenuItem.type == "飲料",
                MenuItem.sold_out == 0,
                MenuItem.id != item.id,
            ).count()
            if remaining == 0:
                db.query(DrinkOption).filter(
                    DrinkOption.date == item.date,
                    DrinkOption.vendor == item.vendor,
                ).delete()

    if "category" in data:
        item.category = data["category"]
    if "remarks" in data:
        item.remarks = data["remarks"]
    if "phone" in data:
        item.phone = data["phone"]
    # 同步更新對應的訂單
    if "name" in data or "price" in data:
        orders = db.query(Order).filter(
            Order.date == item.date,
            Order.type == item.type,
            Order.item_name == old_name,
        ).all()
        for o in orders:
            if "name" in data:
                o.item_name = item.name
            if "price" in data:
                o.price = item.price
            if "vendor" in data:
                o.vendor = item.vendor
    db.commit()
    db.refresh(item)
    return {"id": item.id, "name": item.name, "price": item.price, "vendor": item.vendor, "sold_out": item.sold_out, "category": item.category, "remarks": item.remarks}


@app.delete("/api/menu/{item_id}")
def delete_menu_item(item_id: int, db: Session = Depends(get_db)):
    item = db.query(MenuItem).filter(MenuItem.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Menu item not found")
    db.delete(item)
    db.commit()
    return {"ok": True}


# ==================== Order APIs ====================

@app.get("/api/orders", response_model=list[OrderOut])
def get_orders(
    emp_id: str | None = None,
    proxy_by: str | None = None,
    date: str | None = None,
    type: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    db: Session = Depends(get_db),
):
    q = db.query(Order)
    if emp_id and proxy_by:
        q = q.filter(
            (Order.emp_id == emp_id) |
            (Order.paid_by == emp_id) |
            ((Order.proxy_by == proxy_by) & (Order.company_paid == 1))
        )
    elif emp_id:
        q = q.filter((Order.emp_id == emp_id) | (Order.paid_by == emp_id))
    if date:
        q = q.filter(Order.date == date)
    if type:
        q = q.filter(Order.type == type)
    if start_date:
        q = q.filter(Order.date >= start_date)
    if end_date:
        q = q.filter(Order.date <= end_date)
    return q.order_by(Order.date.desc(), Order.id.desc()).all()


@app.post("/api/orders", response_model=list[OrderOut])
def create_orders(data: OrderCreate, db: Session = Depends(get_db)):
    now = datetime.now(TW_TZ)
    created = []

    # 管理員代訂不限制截止時間
    is_proxy = bool(data.proxy_by)

    for item in data.items:
        menu_items = db.query(MenuItem).filter(
            MenuItem.date == item["date"],
            MenuItem.type == item["type"],
        ).all()

        # 檢查截止時間（override 優先），代訂跳過
        if not is_proxy:
            override = db.query(DeadlineOverride).filter(
                DeadlineOverride.date == item["date"],
                DeadlineOverride.type == item["type"],
            ).first()
            if override:
                deadline_str = f'{item["date"]}T{override.deadline_time}'
                deadline = datetime.fromisoformat(deadline_str).replace(tzinfo=TW_TZ)
                if now > deadline:
                    raise HTTPException(
                        status_code=400,
                        detail=f'{item["date"]} {item["type"]} 已超過截止時間（{override.deadline_time}），無法點餐',
                    )
            elif menu_items and menu_items[0].deadline:
                deadline = datetime.fromisoformat(menu_items[0].deadline)
                if deadline.tzinfo is None:
                    deadline = deadline.replace(tzinfo=TW_TZ)
                if now > deadline:
                    raise HTTPException(
                        status_code=400,
                        detail=f'{item["date"]} {item["type"]} 已超過截止時間，無法點餐',
                    )


        # 從 menu_items 查出店家名稱
        vendor = item.get("vendor", "")
        if not vendor and menu_items:
            matched = next((m for m in menu_items if m.name == item["item_name"]), None)
            if matched:
                vendor = matched.vendor or ""

        # 晚餐同天只能一份公司付：檢查 DB 已有 + 本次批次已加的（廠商不限制）
        self_paid_val = item.get("self_paid", 0)
        is_guest = bool(data.company_name)
        if item["type"] == "晚餐" and self_paid_val == 0 and not is_guest:
            # 查 DB 是否已有同員工同天公司付晚餐
            existing_company = db.query(Order).filter(
                Order.emp_id == data.emp_id,
                Order.date == item["date"],
                Order.type == "晚餐",
                Order.self_paid == 0,
            ).first()
            # 查本次批次是否已有公司付晚餐（同天）
            batch_company = any(
                o.date == item["date"] and o.type == "晚餐" and o.self_paid == 0
                for o in created
            )
            if existing_company or batch_company:
                self_paid_val = 1

        # 飲料加料費後端驗證：從 remarks 比對 drink_options 加料，確保 price 包含加料費
        final_price = item.get("price", 0)
        if item["type"] == "飲料" and item.get("remarks", ""):
            import re as _re2
            base_match = _re2.search(r'\$(\d+)', item["item_name"])
            if base_match:
                base_price = int(base_match.group(1))
                toppings = db.query(DrinkOption).filter(
                    DrinkOption.date == item["date"],
                    DrinkOption.option_type == "topping",
                ).all()
                topping_total = sum(t.price for t in toppings if t.name in item["remarks"])
                expected = base_price + topping_total
                if topping_total > 0 and final_price < expected:
                    final_price = expected

        order = Order(
            emp_id=data.emp_id,
            emp_name=data.emp_name,
            floor=data.floor,
            date=item["date"],
            type=item["type"],
            item_name=item["item_name"],
            vendor=vendor,
            qty=item.get("qty", 1),
            price=final_price,
            remarks=item.get("remarks", ""),
            company_name=data.company_name or None,
            company_paid=1 if (data.company_paid or data.company_name) else 0,
            guest_dept=data.guest_dept or None,
            guest_qty=data.guest_qty if data.guest_qty and data.guest_qty > 0 else None,
            self_paid=self_paid_val,
            proxy_by=data.proxy_by or None,
            paid_by=data.paid_by or None,
            paid_by_name=data.paid_by_name or None,
            treat_reason=data.treat_reason or None,
            created_at=now.strftime("%Y-%m-%dT%H:%M:%S+08:00"),
        )
        db.add(order)
        created.append(order)
    db.commit()
    for order in created:
        db.refresh(order)

    # 請客通知：如果有 paid_by（非公司請客），存 DB + 推送 SSE
    if data.paid_by and data.paid_by != "COMPANY":
        treat_items = [
            {"item_name": o.item_name, "price": o.price, "qty": o.qty, "type": o.type, "date": o.date}
            for o in created if o.paid_by
        ]
        total = sum(i["price"] * i["qty"] for i in treat_items)
        items_summary = "、".join(i["item_name"] for i in treat_items)
        treat_date = treat_items[0]["date"] if treat_items else ""
        notif = TreatNotification(
            paid_by=data.paid_by,
            paid_by_name=data.paid_by_name or "",
            orderer_id=data.emp_id,
            orderer_name=data.emp_name,
            date=treat_date,
            items_summary=items_summary,
            total=total,
            created_at=now.strftime("%Y-%m-%dT%H:%M:%S+08:00"),
        )
        db.add(notif)
        db.commit()
        db.refresh(notif)
        broadcast_coffee_event({
            "type": "treat_notification",
            "id": notif.id,
            "paid_by": data.paid_by,
            "paid_by_name": data.paid_by_name,
            "orderer_name": data.emp_name,
            "date": treat_date,
            "items": treat_items,
            "total": total,
        })

    return created


# ==================== 請客通知 API ====================

@app.get("/api/treat-notifications")
def get_treat_notifications(emp_id: str = Query(...), db: Session = Depends(get_db)):
    """取得未讀請客通知"""
    notifs = db.query(TreatNotification).filter(
        TreatNotification.paid_by == emp_id,
        TreatNotification.is_read == 0,
    ).order_by(TreatNotification.id.desc()).all()
    return [
        {
            "id": n.id,
            "orderer_id": n.orderer_id,
            "orderer_name": n.orderer_name,
            "date": n.date,
            "items_summary": n.items_summary,
            "total": n.total,
            "created_at": n.created_at,
        }
        for n in notifs
    ]


@app.put("/api/treat-notifications/{notif_id}/read")
def mark_treat_notification_read(notif_id: int, db: Session = Depends(get_db)):
    """標記請客通知為已讀"""
    notif = db.query(TreatNotification).filter(TreatNotification.id == notif_id).first()
    if notif:
        notif.is_read = 1
        db.commit()
    return {"success": True}


@app.put("/api/orders/{order_id}", response_model=OrderOut)
def update_order(order_id: int, data: OrderUpdate, db: Session = Depends(get_db)):
    order = db.query(Order).filter(Order.id == order_id).first()
    if not order:
        raise HTTPException(status_code=404, detail="Order not found")
    if data.item_name is not None:
        order.item_name = data.item_name
    if data.qty is not None:
        order.qty = data.qty
    if data.price is not None:
        order.price = data.price
    if data.remarks is not None:
        order.remarks = data.remarks
    if data.self_paid is not None:
        order.self_paid = data.self_paid
    db.commit()
    db.refresh(order)
    return order


@app.post("/api/orders/batch-replace")
def batch_replace_orders(data: dict, db: Session = Depends(get_db)):
    """批次更換品項"""
    date = data["date"]
    old_item = data["old_item"]
    new_item = data["new_item"]
    new_price = int(data.get("new_price", 0))
    new_vendor = data.get("new_vendor", "")
    orders = db.query(Order).filter(
        Order.date == date,
        Order.item_name == old_item,
    ).all()
    count = 0
    for o in orders:
        o.item_name = new_item
        o.price = new_price
        if new_vendor:
            o.vendor = new_vendor
        count += 1
    db.commit()
    return {"count": count, "message": f"已將 {count} 筆「{old_item}」更換為「{new_item}」"}


@app.post("/api/orders/batch-replace-vendor")
def batch_replace_vendor(data: dict, db: Session = Depends(get_db)):
    """整間店家換掉：將某日某店家的所有訂單，依對應表替換為新店家品項"""
    date = data["date"]
    old_vendor = data["old_vendor"]
    replacements = data["replacements"]  # [{order_id, new_item, new_price}]
    new_vendor = data.get("new_vendor", "")
    count = 0
    for r in replacements:
        order = db.query(Order).filter(Order.id == r["order_id"]).first()
        if order:
            order.item_name = r["new_item"]
            order.price = r["new_price"]
            if new_vendor:
                order.vendor = new_vendor
            count += 1
    db.commit()
    return {"count": count, "message": f"已將 {count} 筆「{old_vendor}」的訂單更換完成"}


@app.delete("/api/orders/{order_id}")
def delete_order(order_id: int, db: Session = Depends(get_db)):
    order = db.query(Order).filter(Order.id == order_id).first()
    if not order:
        raise HTTPException(status_code=404, detail="Order not found")
    db.delete(order)
    db.commit()
    return {"ok": True}


# ==================== Stats API ====================

@app.get("/api/stats")
def get_stats(month: str, db: Session = Depends(get_db)):
    """month format: YYYY-MM, salary period: prev month 26 ~ this month 25"""
    year = int(month.split("-")[0])
    mon = int(month.split("-")[1])
    # 薪資區間：上月26 ~ 本月25
    if mon == 1:
        start_date = f"{year-1}-12-26"
    else:
        start_date = f"{year}-{str(mon-1).zfill(2)}-26"
    end_date = f"{year}-{str(mon).zfill(2)}-25"
    orders = db.query(Order).filter(
        Order.date >= start_date,
        Order.date <= end_date,
    ).all()

    coffee_orders = db.query(CoffeeOrder).filter(
        CoffeeOrder.date >= start_date,
        CoffeeOrder.date <= end_date,
    ).all()

    employees = {}
    emp_names = {}
    for o in orders:
        if o.company_paid:
            continue
        # 晚餐非自費（有加班）不計入個人統計
        if o.type == '晚餐' and not o.self_paid:
            continue
        # 公司請客：不扣任何人
        if o.paid_by == 'COMPANY':
            continue
        # 主管請客：金額算到付款人
        key = o.paid_by if o.paid_by else o.emp_id
        if o.paid_by and o.paid_by_name:
            emp_names[key] = o.paid_by_name
        else:
            emp_names[key] = o.emp_name
        if key not in employees:
            employees[key] = {}
        if o.date not in employees[key]:
            employees[key][o.date] = {}
        meal_type = o.type
        employees[key][o.date][meal_type] = employees[key][o.date].get(meal_type, 0) + o.price * o.qty

    for o in coffee_orders:
        if o.company_paid:
            continue
        key = o.emp_id
        if key not in emp_names:
            emp_names[key] = o.emp_name
        if key not in employees:
            employees[key] = {}
        if o.date not in employees[key]:
            employees[key][o.date] = {}
        employees[key][o.date]["咖啡"] = employees[key][o.date].get("咖啡", 0) + o.total_amount

    result = []
    for emp_id in sorted(employees.keys()):
        day_map = employees[emp_id]
        total = 0
        days = {}
        for date, type_map in day_map.items():
            days[date] = type_map
            total += sum(type_map.values())
        result.append({
            "emp_id": emp_id,
            "emp_name": emp_names.get(emp_id, ""),
            "days": days,
            "total": total,
        })
    return result


# ==================== Export APIs ====================

def _get_cjk_font_path() -> str:
    """取得中文字體路徑，Windows 用微軟正黑體，Linux 用 Noto Sans CJK"""
    import os
    # Windows
    win_font = os.path.join(os.environ.get("WINDIR", "C:/Windows"), "Fonts", "msjh.ttc")
    if os.path.exists(win_font):
        return win_font
    # Linux (Cloud Run)
    for p in ["/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
              "/usr/share/fonts/opentype/noto/NotoSansCJKtc-Regular.otf",
              "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"]:
        if os.path.exists(p):
            return p
    return "msjh.ttc"  # fallback


def _strip_eng(s: str) -> str:
    """移除品項名稱中的英文，只留中文+價格"""
    r = s.replace("◆", "")
    pm = re.search(r"\$\d+", r)
    price = pm.group(0) if pm else ""
    r = re.sub(r"\$\d+", "", r)
    r = re.sub(r"\s*\([A-Za-z0-9][A-Za-z0-9\s,&'.+-]*\)", "", r)
    r = re.sub(r"\s*[A-Za-z][A-Za-z\s,&'.+-]*", "", r)
    r = re.sub(r"\s{2,}", " ", r).strip()
    return f"{r} {price}" if price else r


def _is_drink(item: str) -> bool:
    """判斷早餐品項是否為飲品"""
    # 前綴偵測：(冰)/(溫)/冷-/熱-/(中)/(大)/(小)/(特大)
    if re.match(r"^[◆]?\s*(\(冰\)|\(溫\)|冷-|熱-|\(中\)|\(大\)|\(小\)|\(特大\))", item):
        return True
    # 關鍵字偵測：品項名稱包含飲品關鍵字（去掉◆後）
    name = re.sub(r"^◆\s*", "", item).strip()
    drink_keywords = ("紅茶", "奶茶", "綠茶", "豆漿", "鮮奶茶", "鮮奶", "奶綠", "果汁",
                      "古早紅茶", "古早奶茶", "錫蘭紅茶", "抹茶", "可可", "咖啡",
                      "拿鐵", "美式", "冷泡茶", "牛奶", "牛乳", "濃湯", "薏仁漿",
                      "Black Tea", "Milk Tea", "Soy Milk", "Green Tea",
                      "Latte", "Americano", "Cocoa", "Milk", "Juice")
    return any(kw in name for kw in drink_keywords)


def _strip_eng_keep_prefix(s: str) -> str:
    """去英文但保留開頭的 A/B/C 套餐字母"""
    original_prefix = ""
    m = re.match(r"^([A-C])([\u4e00-\u9fff])", s)
    if m:
        original_prefix = m.group(1)
    cleaned = _strip_eng(s)
    if original_prefix and not cleaned.startswith(original_prefix):
        cleaned = original_prefix + cleaned
    return cleaned


def _parse_breakfast_remarks(remarks: str, combo_map: dict = None):
    """解析早餐備註，回傳 (addons, extras)
    addons: 加蛋/加起司等，要跟著主品項顯示
    extras: 套餐(含飲品換購)等，獨立計算的品項
    combo_map: {letter: full_name}
    飲品換購會附加到前一個套餐項目上，例如：C麥克雞塊+中杯紅茶 $55 / 飲品換：(冰)鮮奶茶 $15
    """
    if not remarks:
        return [], []
    addons = []
    extras = []
    skip_words = {"冰", "熱", "溫", "去冰", "少冰", "微冰", "冰 Cold", "熱 Hot", "不要小黃瓜", "撒胡椒粉，醬油另包"}
    # 先以 / 分隔
    parts = [p.strip() for p in remarks.split("/")]
    for part in parts:
        if not part:
            continue
        # 再以 , 分隔（加蛋, 加起司）
        sub_parts = [s.strip() for s in part.split(",")]
        for sp in sub_parts:
            if not sp or sp in skip_words or sp.startswith("不要"):
                continue
            # 飲品換購：附加到前一個套餐項目
            dm = re.match(r"飲品換[：:](.+)", sp)
            if dm:
                cleaned = "飲品換：" + _strip_eng(dm.group(1).strip())
                if cleaned != "飲品換：" and extras:
                    extras[-1] = extras[-1] + " / " + cleaned
                elif cleaned != "飲品換：":
                    extras.append(cleaned)
                continue
            # 套餐：嘗試用 combo_map 替換為完整名稱
            cm = re.match(r"^([A-C])套餐\s*\$?\d*", sp)
            if cm and combo_map and cm.group(1) in combo_map:
                extras.append(combo_map[cm.group(1)])
                continue
            # 加蛋/加起司等加購 → 歸入 addons（保留金額）
            if re.match(r"^加", sp):
                cleaned = _strip_eng(sp)
                if cleaned:
                    addons.append(cleaned)
                continue
            # 其他（如套餐完整名稱 A果醬丹麥+...）
            cleaned = _strip_eng_keep_prefix(sp)
            if cleaned:
                extras.append(cleaned)
    return addons, extras


def _split_breakfast_items(item_name: str):
    """拆分 ◆ 分隔的早餐品項，回傳 (餐點列表, 飲品列表)"""
    if "◆" in item_name:
        parts = [p.strip() for p in item_name.split("◆") if p.strip()]
    else:
        parts = [item_name.strip()]
    foods, drinks = [], []
    for p in parts:
        if _is_drink(p):
            drinks.append(_strip_eng(p))
        else:
            foods.append(_strip_eng(p))
    return foods, drinks


@app.get("/api/export/breakfast-summary")
def export_breakfast_summary(date: str, db: Session = Depends(get_db)):
    """產生早餐品項總表 PNG 圖檔"""
    from PIL import Image, ImageDraw, ImageFont

    orders = db.query(Order).filter(
        Order.date == date, Order.type == "早餐"
    ).all()
    if not orders:
        raise HTTPException(status_code=404, detail="該日期無早餐訂單")

    # 查套餐定義，建立 combo_map {letter: 去英文完整名}
    combo_opts = db.query(DrinkOption).filter(
        DrinkOption.date == date, DrinkOption.option_type == "b_combo"
    ).all()
    combo_map = {}
    for co in combo_opts:
        m = re.match(r"^([A-C])", co.name)
        if m:
            combo_map[m.group(1)] = _strip_eng_keep_prefix(co.name)

    # 拆分品項並統計（含 remarks 裡的加購/套餐/飲品換購）
    item_count: dict[str, int] = {}
    for o in orders:
        if "◆" in o.item_name:
            parts = [p.strip() for p in o.item_name.split("◆") if p.strip()]
        else:
            parts = [o.item_name.strip()]
        # 解析 remarks：addons 跟著主品項，extras 獨立計算
        addons, extras = _parse_breakfast_remarks(o.remarks, combo_map)
        addon_suffix = " + " + " + ".join(addons) if addons else ""
        # 提取溫度前綴（獨立飲品用）
        temp_prefix = ""
        rm = o.remarks or ""
        if re.match(r"^(冰|熱)", rm) and not rm.startswith("加"):
            temp_prefix = "(" + rm[0] + ")"
        for p in parts:
            k = _strip_eng(p)
            if k:
                # 食物品項附加 addon
                if addon_suffix and not _is_drink(p):
                    k = k + addon_suffix
                    addon_suffix = ""
                # 飲品附加溫度前綴（沒有◆的獨立飲品才需要）
                elif _is_drink(p) and temp_prefix and "◆" not in o.item_name:
                    k = temp_prefix + k
                item_count[k] = item_count.get(k, 0) + o.qty
        for rk in extras:
            if rk:
                item_count[rk] = item_count.get(rk, 0) + o.qty

    sorted_items = sorted(item_count.items(), key=lambda x: x[0])
    total_qty = sum(v for _, v in sorted_items)
    vendor_name = orders[0].vendor or ""
    vendor_name_cn = _strip_eng(vendor_name) if vendor_name else ""

    # 繪圖參數
    _cjk = _get_cjk_font_path()
    font_title = ImageFont.truetype(_cjk, 24)
    font_sub = ImageFont.truetype(_cjk, 16)
    font_header = ImageFont.truetype(_cjk, 18)
    font_body = ImageFont.truetype(_cjk, 18)
    font_total = ImageFont.truetype(_cjk, 18)

    pad_x, pad_y = 30, 20
    row_h = 36
    header_h = 40
    title_h = 70
    col_name_w = 700
    col_qty_w = 80
    table_w = col_name_w + col_qty_w
    img_w = table_w + pad_x * 2
    img_h = title_h + pad_y + header_h + row_h * (len(sorted_items) + 1) + pad_y

    img = Image.new("RGB", (img_w, img_h), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    # 標題
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    from datetime import date as dt_date
    d = dt_date.fromisoformat(date)
    wd = weekdays[d.weekday()]
    title = "早餐品項總表"
    tw = draw.textlength(title, font=font_title)
    draw.text(((img_w - tw) / 2, pad_y), title, fill="#000", font=font_title)
    sub = f"{date} ({wd}) {vendor_name_cn}"
    sw = draw.textlength(sub, font=font_sub)
    draw.text(((img_w - sw) / 2, pad_y + 34), sub, fill="#666", font=font_sub)

    # 表格起始位置
    tx = pad_x
    ty = title_h + pad_y

    # 表頭
    draw.rectangle([tx, ty, tx + table_w, ty + header_h], fill="#333")
    draw.text((tx + 10, ty + 8), "品項", fill="#FFF", font=font_header)
    draw.text((tx + col_name_w + 20, ty + 8), "數量", fill="#FFF", font=font_header)
    ty += header_h

    # 資料行
    for i, (name, qty) in enumerate(sorted_items):
        bg = "#F8F8F8" if i % 2 == 1 else "#FFFFFF"
        draw.rectangle([tx, ty, tx + table_w, ty + row_h], fill=bg)
        # 邊框
        draw.line([tx, ty, tx + table_w, ty], fill="#CCC")
        draw.text((tx + 10, ty + 6), name, fill="#000", font=font_body)
        qty_str = str(qty)
        qw = draw.textlength(qty_str, font=font_body)
        draw.text((tx + col_name_w + (col_qty_w - qw) / 2, ty + 6), qty_str, fill="#000", font=font_body)
        ty += row_h

    # 合計行
    draw.rectangle([tx, ty, tx + table_w, ty + row_h], fill="#E8E8E8")
    draw.line([tx, ty, tx + table_w, ty], fill="#333")
    draw.text((tx + 10, ty + 6), "合計", fill="#000", font=font_total)
    total_str = str(total_qty)
    tw2 = draw.textlength(total_str, font=font_total)
    draw.text((tx + col_name_w + (col_qty_w - tw2) / 2, ty + 6), total_str, fill="#000", font=font_total)
    ty += row_h

    # 外框
    draw.rectangle([pad_x, title_h + pad_y, pad_x + table_w, ty], outline="#333", width=2)
    # 直線分隔品項/數量
    draw.line([pad_x + col_name_w, title_h + pad_y, pad_x + col_name_w, ty], fill="#333", width=1)

    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    from urllib.parse import quote
    filename = f"早餐品項總表_{date}.png"
    return StreamingResponse(
        buf,
        media_type="image/png",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


@app.get("/api/export/orders")
def export_orders(request: Request, date: str, db: Session = Depends(get_db)):
    """匯出某日某類型的訂單明細為 Excel，按樓層分 sheet"""
    meal_type = request.query_params.get("type", "午餐")
    orders = db.query(Order).filter(
        Order.date == date, Order.type == meal_type
    ).order_by(Order.emp_id).all()

    wb = openpyxl.Workbook()
    first_sheet = wb.active
    floor_order = {"3F": 0, "5F北": 1, "5F南": 2, "15F": 3, "16F": 4}

    if meal_type == "早餐":
        # --- 早餐專用：輸出 PDF，餐點/飲品分欄，按人合併，去英文 ---
        from collections import OrderedDict
        from fpdf import FPDF
        import os

        # 查套餐定義，建立 combo_map {letter: 去英文完整名}
        combo_opts = db.query(DrinkOption).filter(
            DrinkOption.date == date, DrinkOption.option_type == "b_combo"
        ).all()
        pdf_combo_map = {}
        for co in combo_opts:
            m = re.match(r"^([A-C])", co.name)
            if m:
                pdf_combo_map[m.group(1)] = _strip_eng_keep_prefix(co.name)

        person_map = OrderedDict()
        sorted_orders = sorted(orders, key=lambda o: (floor_order.get(o.floor or "", 99), o.emp_id or ""))
        for o in sorted_orders:
            key = (o.floor or "", o.emp_id or "", o.emp_name or "", o.company_name or "", o.guest_dept or "")
            if key not in person_map:
                person_map[key] = {"foods": [], "drinks": [], "remarks": [], "total": 0}
            foods, drinks = _split_breakfast_items(o.item_name)
            # 解析備註：加蛋/加起司附加到餐點，套餐/飲品換購放備註
            addon_parts = []
            remark_parts = []
            temp_prefix = ""
            rm = o.remarks or ""
            if o.remarks:
                for rp in o.remarks.split(" / "):
                    rp = rp.strip()
                    if not rp:
                        continue
                    # 溫度（冰/熱）→ 附加到飲品，不放備註
                    if re.match(r"^(冰|熱)(\s|$)", rp):
                        temp_prefix = "(" + rp[0] + ")"
                        continue
                    # 飲品換購
                    dm = re.match(r"飲品換[：:](.+)", rp)
                    if dm:
                        remark_parts.append("飲品換：" + _strip_eng(dm.group(1).strip()))
                        continue
                    # 套餐：替換為完整名稱
                    cm = re.match(r"^([A-C])套餐\s*\$?\d*", rp)
                    if cm and cm.group(1) in pdf_combo_map:
                        remark_parts.append(pdf_combo_map[cm.group(1)])
                        continue
                    # 逗號分隔的（加蛋, 加起司）
                    if "," in rp:
                        for sp in rp.split(","):
                            sp = sp.strip()
                            if not sp:
                                continue
                            if re.match(r"^加", sp):
                                addon_parts.append(_strip_eng(sp))
                            else:
                                remark_parts.append(_strip_eng_keep_prefix(sp) if re.search(r"[A-Za-z]", sp) else sp)
                    elif re.match(r"^加", rp):
                        addon_parts.append(_strip_eng(rp))
                    else:
                        remark_parts.append(_strip_eng_keep_prefix(rp) if re.search(r"[A-Za-z]", rp) else rp)
            # 加蛋/加起司附加到第一個餐點
            if addon_parts and foods:
                foods[0] = foods[0] + " + " + " + ".join(addon_parts)
            elif addon_parts:
                foods.extend(addon_parts)
            # 溫度前綴附加到飲品（非◆格式的獨立飲品）
            if temp_prefix and drinks and "◆" not in o.item_name:
                drinks = [temp_prefix + d for d in drinks]
            for _ in range(o.qty):
                person_map[key]["foods"].extend(foods)
                person_map[key]["drinks"].extend(drinks)
            if remark_parts:
                person_map[key]["remarks"].append(" / ".join(remark_parts))
            person_map[key]["total"] += o.price * o.qty

        vendor_name = _strip_eng(orders[0].vendor) if orders and orders[0].vendor else ""
        total_persons = len(person_map)
        total_amount = sum(o.price * o.qty for o in orders)

        weekdays = ["一", "二", "三", "四", "五", "六", "日"]
        from datetime import date as dt_date
        d = dt_date.fromisoformat(date)
        wd = weekdays[d.weekday()]

        # 建立 PDF
        pdf = FPDF(orientation="L", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=True, margin=5)
        pdf.set_margins(10, 8, 10)
        font_path = _get_cjk_font_path()
        pdf.add_font("msjh", "", font_path, uni=True)
        pdf.add_font("msjh", "B", font_path, uni=True)

        # 欄位寬度 (橫式 A4 = 297mm, 扣邊距)
        col_w = {"floor": 14, "name": 28, "count": 12, "food": 70, "drink": 55, "remark": 65, "amount": 22}
        col_w_no_floor = {"name": 28, "count": 12, "food": 75, "drink": 60, "remark": 70, "amount": 22}

        def _pdf_header(pdf, title_text, cols, sub_count, sub_amount):
            pdf.set_font("msjh", "B", 14)
            pdf.cell(0, 10, title_text, ln=True, align="C")
            pdf.set_font("msjh", "", 9)
            pdf.cell(0, 6, f"{date} ({wd})  {vendor_name}    品項小計 {sub_count}    金額小計 {sub_amount:,}", ln=True, align="C")
            pdf.ln(2)
            # 表頭
            pdf.set_fill_color(51, 51, 51)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("msjh", "B", 9)
            headers = list(cols.keys())
            labels = {"floor": "樓層", "name": "姓名", "count": "品項", "food": "餐點", "drink": "飲品", "remark": "其他備註", "amount": "金額"}
            for h in headers:
                pdf.cell(cols[h], 8, labels[h], border=1, align="C", fill=True)
            pdf.ln()
            pdf.set_text_color(0, 0, 0)

        def _pdf_table(pdf, data_map, cols, show_floor=False, sub_count=0, sub_amount=0):
            row_count = 0
            sheet_total = 0
            for (floor, emp_id, emp_name, company_name, guest_dept), data in data_map.items():
                name = f"{guest_dept or company_name} (來客)" if company_name else emp_name
                item_count = len(data["foods"]) + len(data["drinks"])
                food_str = " / ".join(data["foods"])
                drink_str = " / ".join(data["drinks"])
                remark_str = " / ".join(data["remarks"])
                amount = data["total"]

                # 準備每欄的值
                pdf.set_font("msjh", "", 9)
                line_h = 5.5  # 每行高度（加大避免重疊）
                col_keys = list(cols.keys())
                cell_vals = {}
                for ck in col_keys:
                    if ck == "floor":
                        cell_vals[ck] = floor
                    elif ck == "name":
                        cell_vals[ck] = name
                    elif ck == "count":
                        cell_vals[ck] = str(item_count)
                    elif ck == "food":
                        cell_vals[ck] = food_str
                    elif ck == "drink":
                        cell_vals[ck] = drink_str
                    elif ck == "remark":
                        cell_vals[ck] = remark_str
                    elif ck == "amount":
                        cell_vals[ck] = f"{amount:,}"
                    else:
                        cell_vals[ck] = ""

                # 用 dry_run 精確測量每欄實際渲染高度
                max_h = 8
                for ck in col_keys:
                    val = cell_vals[ck]
                    if not val:
                        continue
                    w = cols[ck]
                    pdf.set_font("msjh", "", 9)
                    # dry_run=True 不實際渲染，只回傳需要的高度
                    result = pdf.multi_cell(w, line_h, val, border=0, align="L", dry_run=True, output="HEIGHT")
                    if result > max_h:
                        max_h = result
                row_h = max_h + 2  # 加一點緩衝

                # 檢查是否需要換頁（預留合計行空間）
                if pdf.get_y() + row_h + 10 > 200:
                    pdf.add_page()
                    _pdf_header(pdf, "（續）", cols, sub_count, sub_amount)

                x_start = pdf.get_x()
                y_start = pdf.get_y()

                # 先畫所有欄位的邊框
                for ck in col_keys:
                    w = cols[ck]
                    pdf.rect(pdf.get_x(), y_start, w, row_h)
                    pdf.set_x(pdf.get_x() + w)
                pdf.set_xy(x_start, y_start)

                # 再填入文字（multi_cell 自動換行）
                for ck in col_keys:
                    w = cols[ck]
                    val = cell_vals[ck]
                    pdf.set_font("msjh", "", 9)
                    align = "C" if ck in ("count", "amount", "floor") else "L"
                    x_before = pdf.get_x()
                    pdf.multi_cell(w, line_h, val, border=0, align=align, new_x="RIGHT", new_y="TOP")
                    pdf.set_xy(x_before + w, y_start)

                pdf.set_y(y_start + row_h)
                pdf.set_x(x_start)
                row_count += 1
                sheet_total += amount

            # 合計行
            pdf.set_font("msjh", "B", 9)
            pdf.set_fill_color(230, 230, 230)
            total_cols = list(cols.keys())
            # 合計佔前面所有欄到 amount 前
            pre_w = sum(cols[ck] for ck in total_cols[:-1])
            pdf.cell(pre_w, 8, f"合計  {row_count} 份", border=1, align="C", fill=True)
            pdf.cell(cols["amount"], 8, f"{sheet_total:,}", border=1, align="C", fill=True)
            pdf.ln()

        # 第一頁：全部
        pdf.add_page()
        _pdf_header(pdf, "早餐訂餐明細 - 全部", col_w, total_persons, total_amount)
        _pdf_table(pdf, person_map, col_w, show_floor=True, sub_count=total_persons, sub_amount=total_amount)

        # 只保留全部頁，不分樓層

        buf = BytesIO()
        pdf.output(buf)
        buf.seek(0)
        from urllib.parse import quote
        filename = f"早餐訂餐明細_{date}.pdf"
        return StreamingResponse(
            buf,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
        )

    else:
        # --- 非早餐：原有格式 ---
        # 全部樓層 sheet
        ws_all = first_sheet
        ws_all.title = "全部"
        ws_all.append(["樓層", "員工工號", "姓名", "店家", "品項", "數量", "備註", "小計"])
        sorted_orders = sorted(orders, key=lambda o: floor_order.get(o.floor or "", 99))
        for o in sorted_orders:
            name = (o.guest_dept or o.company_name or "") + " (來客/廠商)" if o.company_name else o.emp_name
            ws_all.append([o.floor or "", o.emp_id, name, o.vendor or "", o.item_name, o.qty, o.remarks or "", o.price * o.qty])
        total_all = sum(o.price * o.qty for o in orders)
        ws_all.append(["", "", "", "", "", "合計", "", total_all])
        ws_all.column_dimensions['A'].width = 8
        ws_all.column_dimensions['B'].width = 14
        ws_all.column_dimensions['C'].width = 12
        ws_all.column_dimensions['D'].width = 30
        ws_all.column_dimensions['E'].width = 60
        ws_all.column_dimensions['F'].width = 8
        ws_all.column_dimensions['G'].width = 30
        ws_all.column_dimensions['H'].width = 10

        # 各樓層 sheet
        floors = ["3F", "5F北", "5F南", "15F", "16F"]
        for floor in floors:
            ws = wb.create_sheet(title=floor)
            ws.append(["員工工號", "姓名", "店家", "品項", "數量", "備註", "小計"])
            floor_orders = [o for o in orders if (o.floor or "") == floor]
            for o in floor_orders:
                name = (o.guest_dept or o.company_name or "") + " (來客/廠商)" if o.company_name else o.emp_name
                ws.append([o.emp_id, name, o.vendor or "", o.item_name, o.qty, o.remarks or "", o.price * o.qty])
            total = sum(o.price * o.qty for o in floor_orders)
            ws.append(["", "", "", "", "合計", "", total])
            ws.column_dimensions['A'].width = 14
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 30
            ws.column_dimensions['D'].width = 60
            ws.column_dimensions['E'].width = 8
            ws.column_dimensions['F'].width = 30
            ws.column_dimensions['G'].width = 10

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    from urllib.parse import quote
    filename = f"orders_{date}_{meal_type}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


@app.get("/api/export/expense")
def export_expense(date: str, db: Session = Depends(get_db)):
    """匯出某日訂餐費用統計 Excel"""
    orders = db.query(Order).filter(Order.date == date).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = date

    # 分類（排除 paid_by='COMPANY' 公司請客）
    self_paid = [o for o in orders if not o.company_paid and o.paid_by != 'COMPANY' and (o.type != '晚餐' or o.self_paid)]
    company_dinner = [o for o in orders if o.type == '晚餐' and not o.self_paid and not o.company_paid and o.paid_by != 'COMPANY']
    company_guest = [o for o in orders if o.company_paid]
    company_treat = [o for o in orders if not o.company_paid and o.paid_by == 'COMPANY']

    breakfast_total = sum(o.price * o.qty for o in self_paid if o.type == '早餐')
    lunch_total = sum(o.price * o.qty for o in self_paid if o.type == '午餐')
    beverage_total = sum(o.price * o.qty for o in self_paid if o.type == '飲料')
    dinner_self_total = sum(o.price * o.qty for o in self_paid if o.type == '晚餐')
    emp_self_total = breakfast_total + lunch_total + beverage_total + dinner_self_total

    company_dinner_total = sum(o.price * o.qty for o in company_dinner)
    company_guest_lunch = sum(o.price * o.qty for o in company_guest if o.type == '午餐')
    company_guest_dinner = sum(o.price * o.qty for o in company_guest if o.type == '晚餐')
    company_guest_other = sum(o.price * o.qty for o in company_guest if o.type not in ('午餐', '晚餐'))
    company_treat_total = sum(o.price * o.qty for o in company_treat)
    company_total = company_dinner_total + company_guest_lunch + company_guest_dinner + company_guest_other + company_treat_total

    # 費用摘要
    ws.append(["費用摘要", date])
    ws.append([])
    ws.append(["員工自費"])
    ws.append(["早餐", breakfast_total])
    ws.append(["午餐", lunch_total])
    ws.append(["飲料", beverage_total])
    ws.append(["晚餐自費", dinner_self_total])
    ws.append(["員工自費小計", emp_self_total])
    ws.append([])
    ws.append(["公司負擔"])
    ws.append(["晚餐（員工）", company_dinner_total, f"{len(company_dinner)}人"])
    ws.append(["廠商來訪（午餐）", company_guest_lunch])
    ws.append(["廠商來訪（晚餐）", company_guest_dinner])
    if company_guest_other:
        ws.append(["廠商來訪（其他）", company_guest_other])
    if company_treat_total:
        ws.append(["公司請客", company_treat_total, f"{len(company_treat)}筆"])
    ws.append(["公司負擔小計", company_total])
    ws.append([])
    ws.append(["合計", emp_self_total + company_total])
    ws.append([])

    # 員工自費明細
    ws.append(["員工自費明細"])
    ws.append(["姓名", "早餐", "午餐", "飲料", "晚餐", "合計"])
    emp_map = {}
    for o in self_paid:
        if o.emp_name not in emp_map:
            emp_map[o.emp_name] = {"早餐": 0, "午餐": 0, "飲料": 0, "晚餐": 0}
        emp_map[o.emp_name][o.type] = emp_map[o.emp_name].get(o.type, 0) + o.price * o.qty
    for name in sorted(emp_map.keys()):
        v = emp_map[name]
        ws.append([name, v["早餐"] or "", v["午餐"] or "", v["飲料"] or "", v["晚餐"] or "", sum(v.values())])
    ws.append([])

    # 訂單明細
    ws.append(["訂單明細"])
    ws.append(["類型", "姓名", "品項", "金額", "備註", "歸類"])
    for o in sorted(orders, key=lambda x: ({"早餐": 0, "午餐": 1, "飲料": 2, "晚餐": 3}.get(x.type, 9), x.emp_name)):
        category = "來客/廠商" if o.company_paid else ("公司請客" if o.paid_by == 'COMPANY' else ("自費" if o.self_paid else ("公司" if o.type == "晚餐" else "員工")))
        name = (o.guest_dept or o.company_name or "") + " (來客/廠商)" if o.company_name else o.emp_name
        ws.append([o.type, name, o.item_name, o.price * o.qty, o.remarks or "", category])

    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 10

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    from urllib.parse import quote
    filename = f"訂餐費用統計_{date}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


@app.get("/api/export/stats")
def export_stats(month: str, db: Session = Depends(get_db)):
    """匯出月度統計為 Excel"""
    orders = db.query(Order).filter(Order.date.like(f"{month}%")).all()

    year = int(month.split("-")[0])
    mon = int(month.split("-")[1])
    import calendar
    days_in_month = calendar.monthrange(year, mon)[1]
    date_list = [f"{month}-{str(d).zfill(2)}" for d in range(1, days_in_month + 1)]

    employees = {}
    for o in orders:
        key = (o.emp_id, o.emp_name)
        if key not in employees:
            employees[key] = {}
        employees[key][o.date] = employees[key].get(o.date, 0) + o.price * o.qty

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{month} 統計"
    header = ["工號", "姓名"] + [d.split("-")[2] for d in date_list] + ["當月總計"]
    ws.append(header)
    for (emp_id, emp_name), day_map in sorted(employees.items()):
        row = [emp_id, emp_name]
        row_total = 0
        for d in date_list:
            val = day_map.get(d, 0)
            row.append(val if val > 0 else "")
            row_total += val
        row.append(row_total)
        ws.append(row)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    from urllib.parse import quote
    filename = f"stats_{month}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


@app.get("/api/export/accounting-stats")
def export_accounting_stats(month: str, db: Session = Depends(get_db)):
    """匯出訂餐統計(會計) - 薪資週期：上月26~本月25，只列工作日，只算員工自費"""
    from datetime import date as date_cls, timedelta
    from urllib.parse import quote
    import calendar

    year = int(month.split("-")[0])
    mon = int(month.split("-")[1])

    # 薪資區間：上月26 ~ 本月25
    if mon == 1:
        start_date = f"{year-1}-12-26"
    else:
        start_date = f"{year}-{str(mon-1).zfill(2)}-26"
    end_date = f"{year}-{str(mon).zfill(2)}-25"

    # 取得假日清單
    holidays_set = set(
        h.date for h in db.query(Holiday).filter(
            Holiday.date >= start_date, Holiday.date <= end_date
        ).all()
    )

    # 產生工作日清單（排除六日和假日）
    d = date_cls.fromisoformat(start_date)
    end_d = date_cls.fromisoformat(end_date)
    work_days = []
    while d <= end_d:
        if d.weekday() < 5 and d.isoformat() not in holidays_set:
            work_days.append(d.isoformat())
        d += timedelta(days=1)

    # 查詢員工自費訂單（排除公司負擔、廠商來訪）
    orders = db.query(Order).filter(
        Order.date >= start_date,
        Order.date <= end_date,
        Order.company_paid == 0,
        (Order.company_name == None) | (Order.company_name == ""),
    ).all()

    # 查詢咖啡訂單（員工自費）
    coffee_orders = db.query(CoffeeOrder).filter(
        CoffeeOrder.date >= start_date,
        CoffeeOrder.date <= end_date,
        CoffeeOrder.company_paid == 0,
    ).all()

    # 按員工+日期彙總金額（晚餐加班 self_paid=0 不計入）
    emp_day_totals = {}
    for o in orders:
        if o.type == '晚餐' and o.self_paid == 0:
            continue
        # 公司請客：不扣任何人
        if o.paid_by == 'COMPANY':
            continue
        # 主管請客：金額算到付款人
        if o.paid_by:
            key = (o.paid_by, o.paid_by_name or o.paid_by)
        else:
            if not o.emp_id:
                continue
            key = (o.emp_id, o.emp_name)
        if key not in emp_day_totals:
            emp_day_totals[key] = {}
        emp_day_totals[key][o.date] = emp_day_totals[key].get(o.date, 0) + o.price * o.qty
    for o in coffee_orders:
        if not o.emp_id:
            continue
        key = (o.emp_id, o.emp_name)
        if key not in emp_day_totals:
            emp_day_totals[key] = {}
        emp_day_totals[key][o.date] = emp_day_totals[key].get(o.date, 0) + o.total_amount

    # 取得所有在職員工（沒訂餐也要列）
    all_employees = db.query(Employee).filter(Employee.is_active == 1).order_by(Employee.account_id).all()
    emp_list = []
    listed_keys = set()
    for emp in all_employees:
        key = (emp.account_id, emp.name)
        emp_list.append(key)
        listed_keys.add(key)
    # 補上有訂單但不在員工表的（以防資料不一致）
    for key in sorted(emp_day_totals.keys()):
        if key not in listed_keys:
            emp_list.append(key)

    # 建立 Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    start_mmdd = start_date[5:].replace("-", "")
    end_mmdd = end_date[5:].replace("-", "")
    ws.title = f"{start_mmdd}-{end_mmdd}訂餐統計"

    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    thin = Side(style="thin")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    sum_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    num_days = len(work_days)
    total_col = 3 + num_days  # 當月總合計欄 (1-indexed)

    # Row 1: 說明 + 日期標頭 + 當月總合計
    ws.cell(1, 1, "→更新日期，複製上次作業檔員編+姓名")
    for i, wd in enumerate(work_days):
        col = 3 + i
        mmdd = wd[5:].replace("-", "")
        cell = ws.cell(1, col, mmdd)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_all
    total_header = ws.cell(1, total_col, "當月總合計")
    total_header.font = header_font
    total_header.fill = header_fill
    total_header.alignment = Alignment(horizontal="center")
    total_header.border = border_all

    # Row 2: 每日全員合計
    ws.cell(2, 1, "每日更新▼")
    ws.cell(2, 1).font = Font(bold=True)
    data_start_row = 4
    data_end_row = data_start_row + len(emp_list) - 1
    for i in range(num_days):
        col = 3 + i
        col_letter = openpyxl.utils.get_column_letter(col)
        cell = ws.cell(2, col)
        cell.value = f"=SUM({col_letter}{data_start_row}:{col_letter}{data_end_row})"
        cell.font = Font(bold=True)
        cell.fill = sum_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_all
        cell.number_format = '#,##0'
    # 總合計列的合計
    first_col_letter = openpyxl.utils.get_column_letter(3)
    last_day_col_letter = openpyxl.utils.get_column_letter(2 + num_days)
    total_cell = ws.cell(2, total_col)
    total_cell.value = f"=SUM({first_col_letter}2:{last_day_col_letter}2)"
    total_cell.font = Font(bold=True)
    total_cell.fill = sum_fill
    total_cell.alignment = Alignment(horizontal="center")
    total_cell.border = border_all
    total_cell.number_format = '#,##0'

    # Row 3: 欄位標題
    ws.cell(3, 1, "員編").font = header_font
    ws.cell(3, 1).fill = header_fill
    ws.cell(3, 1).border = border_all
    ws.cell(3, 2, "姓名").font = header_font
    ws.cell(3, 2).fill = header_fill
    ws.cell(3, 2).border = border_all
    for i in range(num_days):
        cell = ws.cell(3, 3 + i, "出")
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_all
    tc = ws.cell(3, total_col)
    tc.fill = header_fill
    tc.border = border_all

    # Z欄 (col 26): 確認送單
    z_cell = ws.cell(3, 26, "確認送單")
    z_cell.font = header_font
    z_cell.fill = header_fill
    z_cell.alignment = Alignment(horizontal="center")
    z_cell.border = border_all

    # Row 4+: 員工資料
    for r_idx, (emp_id, emp_name) in enumerate(emp_list):
        row = data_start_row + r_idx
        ws.cell(row, 1, emp_id).border = border_all
        ws.cell(row, 2, emp_name).border = border_all
        day_map = emp_day_totals.get((emp_id, emp_name), {})
        for i, wd in enumerate(work_days):
            col = 3 + i
            val = day_map.get(wd, 0)
            cell = ws.cell(row, col)
            cell.value = val if val > 0 else ""
            cell.alignment = Alignment(horizontal="center")
            cell.border = border_all
            if val > 0:
                cell.number_format = '#,##0'
        # 當月總合計 = SUM
        fc = openpyxl.utils.get_column_letter(3)
        lc = openpyxl.utils.get_column_letter(2 + num_days)
        total_c = ws.cell(row, total_col)
        total_c.value = f"=IF(SUM({fc}{row}:{lc}{row})=0,\"\",SUM({fc}{row}:{lc}{row}))"
        total_c.font = Font(bold=True)
        total_c.alignment = Alignment(horizontal="center")
        total_c.border = border_all
        # Z欄: 確認送單文字
        z_data = ws.cell(row, 26, "本人確認以上申請，並同意由薪資扣款。")
        z_data.alignment = Alignment(horizontal="center")
        z_data.border = border_all
        total_c.number_format = '#,##0'

    # AC~AE 查詢區（從 row 1 開始）
    ac_col = total_col + 3  # 隔兩欄
    ac_letter = openpyxl.utils.get_column_letter(ac_col)
    ad_letter = openpyxl.utils.get_column_letter(ac_col + 1)
    ae_letter = openpyxl.utils.get_column_letter(ac_col + 2)
    total_col_letter = openpyxl.utils.get_column_letter(total_col)

    ws.cell(1, ac_col, "▼輸入工號或姓名").font = Font(bold=True, color="0070C0")
    ws.cell(2, ac_col, "查詢").font = header_font
    ws.cell(2, ac_col).fill = header_fill
    ws.cell(2, ac_col).border = border_all
    ws.cell(2, ac_col + 1, "目前合計").font = header_font
    ws.cell(2, ac_col + 1).fill = header_fill
    ws.cell(2, ac_col + 1).border = border_all

    # 預留 10 筆查詢列（支援工號或姓名查詢，TEXT處理數字工號）
    for qr in range(3, 13):
        ws.cell(qr, ac_col).border = border_all
        # 合計 = 先查工號(TEXT轉文字)，找不到再查姓名
        ws.cell(qr, ac_col + 1).value = f'=IF({ac_letter}{qr}="","",IF(ISNUMBER(MATCH(TEXT({ac_letter}{qr},"@"),A:A,0)),INDEX({total_col_letter}:{total_col_letter},MATCH(TEXT({ac_letter}{qr},"@"),A:A,0)),IF(ISNUMBER(MATCH({ac_letter}{qr},B:B,0)),INDEX({total_col_letter}:{total_col_letter},MATCH({ac_letter}{qr},B:B,0)),"查無此人")))'
        ws.cell(qr, ac_col + 1).border = border_all
        ws.cell(qr, ac_col + 1).number_format = '#,##0'

    # 欄寬調整
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 10
    for i in range(num_days):
        ws.column_dimensions[openpyxl.utils.get_column_letter(3 + i)].width = 7
    ws.column_dimensions[openpyxl.utils.get_column_letter(total_col)].width = 12
    ws.column_dimensions[openpyxl.utils.get_column_letter(ac_col)].width = 14
    ws.column_dimensions[openpyxl.utils.get_column_letter(ac_col + 1)].width = 12
    ws.column_dimensions["Z"].width = 42

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"訂餐統計(會計)_{start_mmdd}-{end_mmdd}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


@app.get("/api/export/dinner-detail")
def export_dinner_detail(month: str, db: Session = Depends(get_db)):
    """匯出晚餐明細(會計) - 薪資週期：上月26~本月25"""
    from urllib.parse import quote
    import re

    year = int(month.split("-")[0])
    mon = int(month.split("-")[1])

    # 薪資區間
    if mon == 1:
        start_date = f"{year-1}-12-26"
    else:
        start_date = f"{year}-{str(mon-1).zfill(2)}-26"
    end_date = f"{year}-{str(mon).zfill(2)}-25"

    # 查詢所有晚餐訂單
    orders = db.query(Order).filter(
        Order.date >= start_date,
        Order.date <= end_date,
        Order.type == "晚餐",
        (Order.company_name == None) | (Order.company_name == ""),
    ).order_by(Order.date, Order.emp_id).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工作表1"

    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    thin = Side(style="thin")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # 標頭
    headers = ["日期", "人名", "金額", "請款用", "員工編號", "類別", "確認送單"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_all

    self_paid_fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
    overtime_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")

    row_idx = 2
    for o in orders:
        # 日期轉中文格式: 2026-06-15 -> 6月15日
        d_parts = o.date.split("-")
        d_month = int(d_parts[1])
        d_day = int(d_parts[2])
        date_str = f"{d_month}月{d_day}日"

        # 請款用：從品項名稱取出編號+名稱，去掉價格
        item = o.item_name or ""
        # 取出 (X-X) 編號
        m = re.match(r"\(?(\d+-\d+)\)?", item)
        prefix = m.group(1) if m else ""
        # 去掉編號括號和價格 $XX
        claim_name = re.sub(r"^\(?\d+-\d+\)?\s*", "", item)
        claim_name = re.sub(r"\s*\$\d+\s*$", "", claim_name)
        claim_text = f"{prefix}  {claim_name}" if prefix else claim_name

        is_self_paid = o.self_paid == 1
        category = "自費" if is_self_paid else "加班"
        row_fill = self_paid_fill if is_self_paid else overtime_fill

        for q in range(o.qty):
            ws.cell(row_idx, 1, date_str).border = border_all
            ws.cell(row_idx, 2, o.emp_name).border = border_all
            c3 = ws.cell(row_idx, 3, o.price)
            c3.border = border_all
            c3.alignment = Alignment(horizontal="center")
            c3.number_format = '#,##0'
            ws.cell(row_idx, 4, claim_text).border = border_all
            ws.cell(row_idx, 5, o.emp_id).border = border_all
            c6 = ws.cell(row_idx, 6, category)
            c6.border = border_all
            c6.alignment = Alignment(horizontal="center")
            c7 = ws.cell(row_idx, 7, "本人確認以上申請，並同意由薪資扣款。")
            c7.border = border_all
            c7.alignment = Alignment(horizontal="center")
            for col in range(1, 8):
                ws.cell(row_idx, col).fill = row_fill
            row_idx += 1

    # 欄寬
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 70
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 8
    ws.column_dimensions["G"].width = 42

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    start_mmdd = start_date[5:].replace("-", "")
    end_mmdd = end_date[5:].replace("-", "")
    filename = f"晚餐明細(會計)_{start_mmdd}-{end_mmdd}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


@app.get("/api/export/guest-detail")
def export_guest_detail(month: str, db: Session = Depends(get_db)):
    """匯出來客/廠商明細 - 薪資週期：上月26~本月25"""
    from urllib.parse import quote
    import re

    year = int(month.split("-")[0])
    mon = int(month.split("-")[1])

    if mon == 1:
        start_date = f"{year-1}-12-21"
    else:
        start_date = f"{year}-{str(mon-1).zfill(2)}-21"
    end_date = f"{year}-{str(mon).zfill(2)}-20"

    orders = db.query(Order).filter(
        Order.date >= start_date,
        Order.date <= end_date,
        Order.company_name != None,
        Order.company_name != "",
    ).order_by(Order.date, Order.type, Order.guest_dept, Order.company_name).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "來客廠商明細"

    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    thin = Side(style="thin")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    type_fills = {
        "午餐": PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid"),
        "晚餐": PatternFill(start_color="F3E5F5", end_color="F3E5F5", fill_type="solid"),
        "飲料": PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid"),
    }

    headers = ["日期", "餐別", "部門", "廠商名稱", "品項", "數量", "單價", "小計", "樓層", "代訂人"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_all

    row_idx = 2
    total_amount = 0
    for o in orders:
        d_parts = o.date.split("-")
        date_str = f"{int(d_parts[1])}月{int(d_parts[2])}日"
        subtotal = o.price * o.qty
        total_amount += subtotal
        row_fill = type_fills.get(o.type)

        ws.cell(row_idx, 1, date_str).border = border_all
        ws.cell(row_idx, 2, o.type).border = border_all
        ws.cell(row_idx, 3, o.guest_dept or "").border = border_all
        ws.cell(row_idx, 4, o.company_name or "").border = border_all
        ws.cell(row_idx, 5, o.item_name or "").border = border_all
        c6 = ws.cell(row_idx, 6, o.qty)
        c6.border = border_all
        c6.alignment = Alignment(horizontal="center")
        c7 = ws.cell(row_idx, 7, o.price)
        c7.border = border_all
        c7.number_format = '#,##0'
        c8 = ws.cell(row_idx, 8, subtotal)
        c8.border = border_all
        c8.number_format = '#,##0'
        ws.cell(row_idx, 9, o.floor or "").border = border_all
        ws.cell(row_idx, 10, o.proxy_by or "").border = border_all

        if row_fill:
            for col in range(1, 11):
                ws.cell(row_idx, col).fill = row_fill
        row_idx += 1

    # 合計行
    if orders:
        ws.cell(row_idx, 7, "合計").border = border_all
        ws.cell(row_idx, 7).font = Font(bold=True)
        ws.cell(row_idx, 7).alignment = Alignment(horizontal="center")
        c_total = ws.cell(row_idx, 8, total_amount)
        c_total.border = border_all
        c_total.font = Font(bold=True)
        c_total.number_format = '#,##0'

    # 右側統計欄（L~O）：按日期+部門彙總
    from collections import OrderedDict
    stats = OrderedDict()
    for o in orders:
        d_parts = o.date.split("-")
        date_str = f"{int(d_parts[1])}月{int(d_parts[2])}日"
        dept = o.guest_dept or o.company_name or ""
        key = (o.date, date_str, dept)
        if key not in stats:
            stats[key] = {"qty": 0, "amount": 0}
        stats[key]["qty"] += o.qty
        stats[key]["amount"] += o.price * o.qty

    stat_col = 12  # L欄
    stat_headers = ["日期", "部門", "數量", "金額小計"]
    summary_fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
    for c, h in enumerate(stat_headers):
        cell = ws.cell(1, stat_col + c, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_all

    s_row = 2
    stat_total_qty = 0
    stat_total_amt = 0
    for (raw_date, date_str, dept), v in stats.items():
        ws.cell(s_row, stat_col, date_str).border = border_all
        ws.cell(s_row, stat_col + 1, dept).border = border_all
        c_qty = ws.cell(s_row, stat_col + 2, v["qty"])
        c_qty.border = border_all
        c_qty.alignment = Alignment(horizontal="center")
        c_amt = ws.cell(s_row, stat_col + 3, v["amount"])
        c_amt.border = border_all
        c_amt.number_format = '#,##0'
        for cc in range(stat_col, stat_col + 4):
            ws.cell(s_row, cc).fill = summary_fill
        stat_total_qty += v["qty"]
        stat_total_amt += v["amount"]
        s_row += 1

    if stats:
        ws.cell(s_row, stat_col, "合計").border = border_all
        ws.cell(s_row, stat_col).font = Font(bold=True)
        ws.cell(s_row, stat_col + 1).border = border_all
        c_tq = ws.cell(s_row, stat_col + 2, stat_total_qty)
        c_tq.border = border_all
        c_tq.font = Font(bold=True)
        c_tq.alignment = Alignment(horizontal="center")
        c_ta = ws.cell(s_row, stat_col + 3, stat_total_amt)
        c_ta.border = border_all
        c_ta.font = Font(bold=True)
        c_ta.number_format = '#,##0'

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 50
    ws.column_dimensions["F"].width = 6
    ws.column_dimensions["G"].width = 8
    ws.column_dimensions["H"].width = 10
    ws.column_dimensions["I"].width = 8
    ws.column_dimensions["J"].width = 10
    ws.column_dimensions["L"].width = 10
    ws.column_dimensions["M"].width = 14
    ws.column_dimensions["N"].width = 8
    ws.column_dimensions["O"].width = 12

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    start_mmdd = start_date[5:].replace("-", "")
    end_mmdd = end_date[5:].replace("-", "")
    filename = f"來客廠商明細_{start_mmdd}-{end_mmdd}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


# ==================== Drink Option APIs ====================

@app.get("/api/drink-options")
def get_drink_options(date: str | None = None, db: Session = Depends(get_db)):
    q = db.query(DrinkOption)
    if date:
        q = q.filter(DrinkOption.date == date)
    options = q.order_by(DrinkOption.date, DrinkOption.vendor, DrinkOption.option_type, DrinkOption.id).all()
    result = {}
    for o in options:
        key = f"{o.date}|{o.vendor}"
        if key not in result:
            result[key] = {"date": o.date, "vendor": o.vendor, "toppings": [], "sweetness": [], "ice": [], "b_addons": [], "b_combos": [], "b_drink_upgrade": None}
        if o.option_type == "topping":
            result[key]["toppings"].append({"id": o.id, "name": o.name, "price": o.price})
        elif o.option_type == "sweetness":
            result[key]["sweetness"].append(o.name)
        elif o.option_type == "ice":
            result[key]["ice"].append(o.name)
        elif o.option_type == "b_addon":
            result[key]["b_addons"].append({"name": o.name, "price": o.price})
        elif o.option_type == "b_combo":
            result[key]["b_combos"].append({"name": o.name, "price": o.price})
        elif o.option_type == "b_drink_upgrade":
            result[key]["b_drink_upgrade"] = {"name": o.name, "price": o.price}
    # 甜度按百分比降序排列（正常100→少糖70→半糖50→微糖30→一分糖10→無糖0）
    import re as _re
    def _sweet_order(name):
        m = _re.search(r'(\d+)%', name)
        return -(int(m.group(1))) if m else 0
    for v in result.values():
        v["sweetness"].sort(key=_sweet_order)
    return list(result.values())


@app.post("/api/drink-options")
def create_drink_option(data: dict, db: Session = Depends(get_db)):
    import re as _re
    name = data.get("name", "").strip()
    if not name:
        raise HTTPException(status_code=400, detail="name required")
    price_match = _re.search(r'\$(\d+)', name)
    price = int(data.get("price", 0)) or (int(price_match.group(1)) if price_match else 0)
    opt = DrinkOption(
        date=data["date"], vendor=data.get("vendor", ""),
        option_type=data.get("option_type", "topping"),
        name=name, price=price
    )
    db.add(opt)
    db.commit()
    db.refresh(opt)
    return {"id": opt.id, "name": opt.name, "price": opt.price, "option_type": opt.option_type, "date": opt.date, "vendor": opt.vendor}


@app.put("/api/drink-options/{option_id}")
def update_drink_option(option_id: int, data: dict, db: Session = Depends(get_db)):
    import re
    opt = db.query(DrinkOption).filter(DrinkOption.id == option_id).first()
    if not opt:
        raise HTTPException(status_code=404, detail="Not found")
    old_name = opt.name
    old_price = opt.price
    if "name" in data:
        opt.name = data["name"]
    if "price" in data:
        opt.price = int(data["price"])
    # 加料價格或名稱變動時，同步更新已下單的飲料訂單
    if opt.option_type == "topping" and (old_name != opt.name or old_price != opt.price):
        orders = db.query(Order).filter(
            Order.date == opt.date,
            Order.type == "飲料",
            Order.remarks.like(f"%{old_name}%"),
        ).all()
        price_diff = opt.price - old_price
        for o in orders:
            if old_name in (o.remarks or ""):
                o.remarks = o.remarks.replace(old_name, opt.name)
                o.price = o.price + price_diff
    db.commit()
    db.refresh(opt)
    return {"id": opt.id, "name": opt.name, "price": opt.price, "option_type": opt.option_type, "date": opt.date, "vendor": opt.vendor}


@app.delete("/api/drink-options/{option_id}")
def delete_drink_option(option_id: int, db: Session = Depends(get_db)):
    opt = db.query(DrinkOption).filter(DrinkOption.id == option_id).first()
    if not opt:
        raise HTTPException(status_code=404, detail="Not found")
    db.delete(opt)
    db.commit()
    return {"success": True}


# ==================== Vendor Note APIs ====================

@app.get("/api/vendor-notes")
def get_vendor_notes(db: Session = Depends(get_db)):
    return [{"id": v.id, "vendor": v.vendor, "date": v.date, "note": v.note, "confirmed_by": v.confirmed_by}
            for v in db.query(VendorNote).order_by(VendorNote.vendor).all()]


@app.post("/api/vendor-notes")
def upsert_vendor_note(data: dict, db: Session = Depends(get_db)):
    vendor = data["vendor"]
    date = data.get("date", "")
    existing = db.query(VendorNote).filter(VendorNote.vendor == vendor, VendorNote.date == date).first()
    if existing:
        existing.note = data["note"]
        db.commit()
        db.refresh(existing)
        return {"id": existing.id, "vendor": existing.vendor, "date": existing.date, "note": existing.note, "confirmed_by": existing.confirmed_by}
    v = VendorNote(vendor=vendor, date=date, note=data["note"])
    db.add(v)
    db.commit()
    db.refresh(v)
    return {"id": v.id, "vendor": v.vendor, "date": v.date, "note": v.note, "confirmed_by": v.confirmed_by}


@app.post("/api/vendor-notes/confirm")
def toggle_vendor_confirm(data: dict, db: Session = Depends(get_db)):
    vendor = data["vendor"]
    date = data.get("date", "")
    confirmed_by = data.get("confirmed_by", "")
    existing = db.query(VendorNote).filter(VendorNote.vendor == vendor, VendorNote.date == date).first()
    if existing:
        existing.confirmed_by = confirmed_by
        db.commit()
        db.refresh(existing)
        return {"id": existing.id, "vendor": existing.vendor, "date": existing.date, "note": existing.note, "confirmed_by": existing.confirmed_by}
    v = VendorNote(vendor=vendor, date=date, confirmed_by=confirmed_by)
    db.add(v)
    db.commit()
    db.refresh(v)
    return {"id": v.id, "vendor": v.vendor, "date": v.date, "note": v.note, "confirmed_by": v.confirmed_by}


@app.delete("/api/vendor-notes/{note_id}")
def delete_vendor_note(note_id: int, db: Session = Depends(get_db)):
    v = db.query(VendorNote).filter(VendorNote.id == note_id).first()
    if not v:
        raise HTTPException(status_code=404, detail="Not found")
    db.delete(v)
    db.commit()
    return {"ok": True}


# ==================== Holiday APIs ====================

@app.get("/api/holidays")
def get_holidays(db: Session = Depends(get_db)):
    return [{"id": h.id, "date": h.date, "name": h.name}
            for h in db.query(Holiday).order_by(Holiday.date).all()]


@app.post("/api/holidays")
def create_holiday(data: dict, db: Session = Depends(get_db)):
    h = Holiday(date=data["date"], name=data["name"])
    db.add(h)
    db.commit()
    db.refresh(h)
    return {"id": h.id, "date": h.date, "name": h.name}


@app.put("/api/holidays/{holiday_id}")
def update_holiday(holiday_id: int, data: dict, db: Session = Depends(get_db)):
    h = db.query(Holiday).filter(Holiday.id == holiday_id).first()
    if not h:
        raise HTTPException(status_code=404, detail="Holiday not found")
    if "date" in data:
        h.date = data["date"]
    if "name" in data:
        h.name = data["name"]
    db.commit()
    db.refresh(h)
    return {"id": h.id, "date": h.date, "name": h.name}


@app.delete("/api/holidays/{holiday_id}")
def delete_holiday(holiday_id: int, db: Session = Depends(get_db)):
    h = db.query(Holiday).filter(Holiday.id == holiday_id).first()
    if not h:
        raise HTTPException(status_code=404, detail="Holiday not found")
    db.delete(h)
    db.commit()
    return {"ok": True}


# ==================== Deadline Override APIs ====================

@app.get("/api/deadline-overrides")
def get_deadline_overrides(db: Session = Depends(get_db)):
    return [{"id": o.id, "date": o.date, "type": o.type,
             "deadline_time": o.deadline_time, "note": o.note}
            for o in db.query(DeadlineOverride).order_by(DeadlineOverride.date, DeadlineOverride.type).all()]


@app.post("/api/deadline-overrides")
def create_deadline_override(data: dict, db: Session = Depends(get_db)):
    if not data.get("note", "").strip():
        raise HTTPException(400, "備註為必填欄位")
    # 同一天同餐別只能一筆，先刪舊的
    db.query(DeadlineOverride).filter(
        DeadlineOverride.date == data["date"],
        DeadlineOverride.type == data["type"],
    ).delete()
    o = DeadlineOverride(
        date=data["date"], type=data["type"],
        deadline_time=data["deadline_time"], note=data.get("note", "")
    )
    db.add(o)
    db.commit()
    db.refresh(o)
    return {"id": o.id, "date": o.date, "type": o.type,
            "deadline_time": o.deadline_time, "note": o.note}


@app.delete("/api/deadline-overrides/{override_id}")
def delete_deadline_override(override_id: int, db: Session = Depends(get_db)):
    o = db.query(DeadlineOverride).filter(DeadlineOverride.id == override_id).first()
    if not o:
        raise HTTPException(status_code=404, detail="Not found")
    db.delete(o)
    db.commit()
    return {"ok": True}


# ==================== Extra Fees ====================

@app.get("/api/extra-fees", response_model=list[ExtraFeeOut])
def list_extra_fees(date: str | None = None, db: Session = Depends(get_db)):
    q = db.query(ExtraFee)
    if date:
        q = q.filter(ExtraFee.date == date)
    return q.order_by(ExtraFee.date).all()


@app.post("/api/extra-fees", response_model=ExtraFeeOut)
def create_extra_fee(data: ExtraFeeCreate, db: Session = Depends(get_db)):
    fee = ExtraFee(date=data.date, type=data.type, description=data.description, amount=data.amount)
    db.add(fee)
    db.commit()
    db.refresh(fee)
    return fee


@app.put("/api/extra-fees/{fee_id}", response_model=ExtraFeeOut)
def update_extra_fee(fee_id: int, data: ExtraFeeCreate, db: Session = Depends(get_db)):
    fee = db.query(ExtraFee).filter(ExtraFee.id == fee_id).first()
    if not fee:
        raise HTTPException(status_code=404, detail="ExtraFee not found")
    fee.date = data.date
    fee.type = data.type
    fee.description = data.description
    fee.amount = data.amount
    db.commit()
    db.refresh(fee)
    return fee


@app.delete("/api/extra-fees/{fee_id}")
def delete_extra_fee(fee_id: int, db: Session = Depends(get_db)):
    fee = db.query(ExtraFee).filter(ExtraFee.id == fee_id).first()
    if not fee:
        raise HTTPException(status_code=404, detail="ExtraFee not found")
    db.delete(fee)
    db.commit()
    return {"ok": True}


# ==================== Coffee APIs ====================

@app.get("/api/coffee/items")
def get_coffee_items(all: int = 0, db: Session = Depends(get_db)):
    q = db.query(CoffeeItem)
    if not all:
        q = q.filter(CoffeeItem.is_active == 1)
    return [CoffeeItemOut.model_validate(i) for i in q.order_by(CoffeeItem.id).all()]


@app.post("/api/coffee/items")
def create_coffee_item(item: CoffeeItemCreate, db: Session = Depends(get_db)):
    obj = CoffeeItem(name=item.name.strip(), price=item.price, is_active=item.is_active)
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return CoffeeItemOut.model_validate(obj)


@app.put("/api/coffee/items/{item_id}")
def update_coffee_item(item_id: int, item: CoffeeItemCreate, db: Session = Depends(get_db)):
    obj = db.query(CoffeeItem).get(item_id)
    if not obj:
        raise HTTPException(404, "品項不存在")
    obj.name = item.name.strip()
    obj.price = item.price
    obj.is_active = item.is_active
    db.commit()
    return CoffeeItemOut.model_validate(obj)


@app.delete("/api/coffee/items/{item_id}")
def delete_coffee_item(item_id: int, db: Session = Depends(get_db)):
    obj = db.query(CoffeeItem).get(item_id)
    if not obj:
        raise HTTPException(404, "品項不存在")
    db.delete(obj)
    db.commit()
    return {"ok": True}


@app.post("/api/coffee/orders")
def create_coffee_order(req: CoffeeOrderCreate, db: Session = Depends(get_db)):
    import json
    now = datetime.now(TW_TZ)
    order_date = req.date if req.date else now.strftime("%Y-%m-%d")
    ts = now.strftime("%Y-%m-%dT%H:%M:%S+08:00")
    created = []
    for it in req.items:
        qty = it.get("quantity", 1)
        price = it.get("price", 0)
        custs = it.get("customizations", [])
        # 濃縮加價 +5
        extra = 5 if any('濃縮' in c for c in custs) else 0
        obj = CoffeeOrder(
            emp_id=req.emp_id.strip(),
            emp_name=req.emp_name.strip(),
            floor=req.floor,
            date=order_date,
            coffee_name=it.get("coffee_name", ""),
            price=price,
            quantity=qty,
            total_amount=(price + extra) * qty,
            temp=it.get("temp", "冷"),
            time_pref=it.get("time_pref", "上午"),
            customizations=json.dumps(custs, ensure_ascii=False) if custs else None,
            note=it.get("note", "") or None,
            status="待製作",
            company_paid=1 if req.company_paid else 0,
            department=req.department if req.company_paid else None,
            created_at=ts,
        )
        db.add(obj)
        created.append(obj)
    db.commit()
    for o in created:
        db.refresh(o)
    return [CoffeeOrderOut.model_validate(o) for o in created]


@app.get("/api/coffee/orders")
def get_coffee_orders(
    date: str | None = None,
    emp_id: str | None = None,
    status: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    db: Session = Depends(get_db),
):
    q = db.query(CoffeeOrder)
    if date:
        q = q.filter(CoffeeOrder.date == date)
    if emp_id:
        q = q.filter(CoffeeOrder.emp_id == emp_id)
    if status:
        q = q.filter(CoffeeOrder.status == status)
    if start_date:
        q = q.filter(CoffeeOrder.date >= start_date)
    if end_date:
        q = q.filter(CoffeeOrder.date <= end_date)
    return [CoffeeOrderOut.model_validate(o) for o in q.order_by(CoffeeOrder.id.desc()).all()]


@app.put("/api/coffee/orders/{order_id}/status")
def update_coffee_order_status(order_id: int, body: dict, db: Session = Depends(get_db)):
    obj = db.query(CoffeeOrder).get(order_id)
    if not obj:
        raise HTTPException(404, "訂單不存在")
    new_status = body.get("status", "")
    if obj.status == new_status:
        return CoffeeOrderOut.model_validate(obj)
    obj.status = new_status
    if new_status == "已完成":
        obj.production_time = datetime.now(TW_TZ).strftime("%H:%M")
    else:
        obj.production_time = None
    db.commit()
    # 企業微信通知（同時段訂單都完成才發）
    if new_status == "已完成" and obj.emp_id:
        pending = db.query(CoffeeOrder).filter(
            CoffeeOrder.emp_id == obj.emp_id,
            CoffeeOrder.date == obj.date,
            CoffeeOrder.time_pref == obj.time_pref,
            CoffeeOrder.status.in_(["待製作", "製作中"]),
        ).count()
        if pending == 0:
            try:
                import requests as req
                now = datetime.now(TW_TZ)
                total_cups = db.query(CoffeeOrder).filter(
                    CoffeeOrder.emp_id == obj.emp_id,
                    CoffeeOrder.date == obj.date,
                    CoffeeOrder.time_pref == obj.time_pref,
                    CoffeeOrder.status == "已完成",
                ).count()
                date_cn = f"{now.month:02d}/{now.day:02d}"
                months_en = ['','January','February','March','April','May','June','July','August','September','October','November','December']
                date_en = f"{months_en[now.month]} {now.day}"
                msg = f"【{date_cn}】咖啡共 {total_cups} 杯，已製作完成，請至餐櫃領取。\n【{date_en}】{total_cups} cup(s) of coffee are ready for pickup at the counter."
                token_r = req.get("https://qyapi.weixin.qq.com/cgi-bin/gettoken", params={
                    "corpid": "ww5c54aa97ecd16ce0",
                    "corpsecret": "oP6tS_DMUNOKzoLShX-g0LD-sAlF19-Oc6HCObHiZPE"
                }, timeout=5)
                token = token_r.json().get("access_token")
                if token:
                    wechat_user = WECHAT_ID_MAP.get(obj.emp_id, obj.emp_id)
                    req.post("https://qyapi.weixin.qq.com/cgi-bin/message/send",
                        params={"access_token": token},
                        json={"msgtype": "text", "touser": wechat_user, "agentid": 1000162,
                              "text": {"content": msg}}, timeout=5)
            except Exception:
                pass  # 通知失敗不影響主流程
    # SSE 廣播
    broadcast_coffee_event({
        "type": "status_change",
        "order_id": obj.id,
        "emp_id": obj.emp_id,
        "coffee_name": obj.coffee_name,
        "status": new_status,
        "total_amount": obj.total_amount,
    })
    return CoffeeOrderOut.model_validate(obj)


@app.delete("/api/coffee/orders/{order_id}")
def delete_coffee_order(order_id: int, db: Session = Depends(get_db)):
    obj = db.query(CoffeeOrder).get(order_id)
    if not obj:
        raise HTTPException(404, "訂單不存在")
    if obj.status != '待製作':
        raise HTTPException(400, "該訂單狀態為「" + obj.status + "」，無法取消")
    db.delete(obj)
    db.commit()
    return {"ok": True}


@app.get("/api/coffee/stats")
def get_coffee_stats(month: str, db: Session = Depends(get_db)):
    year, mon = int(month[:4]), int(month[5:7])
    sm, sy = (mon - 1, year) if mon > 1 else (12, year - 1)
    start_date = f"{sy}-{sm:02d}-26"
    end_date = f"{year}-{mon:02d}-25"
    orders = db.query(CoffeeOrder).filter(
        CoffeeOrder.date >= start_date,
        CoffeeOrder.date <= end_date,
        CoffeeOrder.status == "已完成",
    ).all()
    summary = {}
    for o in orders:
        key = f"{o.date}_{o.emp_id}_{1 if o.company_paid else 0}"
        if key not in summary:
            summary[key] = {
                "date": o.date, "emp_id": o.emp_id, "emp_name": o.emp_name,
                "floor": o.floor, "company_paid": o.company_paid,
                "total_cups": 0, "total_amount": 0,
            }
        summary[key]["total_cups"] += o.quantity
        summary[key]["total_amount"] += o.total_amount
    result = sorted(summary.values(), key=lambda x: (x["date"], x["emp_id"]))
    return result


@app.get("/api/export/coffee")
def export_coffee_stats(month: str, db: Session = Depends(get_db)):
    data = get_coffee_stats(month, db)
    year, mon = int(month[:4]), int(month[5:7])
    sm, sy = (mon - 1, year) if mon > 1 else (12, year - 1)
    start_date = f"{sy}-{sm:02d}-26"
    end_date = f"{year}-{mon:02d}-25"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "咖啡費用統計"
    ws.append(["咖啡點餐月度費用統計"])
    ws.append([f"計費區間：{start_date} 至 {end_date}"])
    ws.append([])
    ws.append(["日期", "工號", "姓名", "付款類型", "樓層", "總杯數", "總金額"])
    for row in data:
        pay_type = "公司付" if row["company_paid"] else "個人"
        ws.append([row["date"], row["emp_id"], row["emp_name"], pay_type, row["floor"], row["total_cups"], row["total_amount"]])
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    from urllib.parse import quote
    fname = f"咖啡費用統計_{month}.xlsx"
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(fname)}"})


# ==================== Order Reviews ====================

@app.post("/api/reviews")
def create_review(data: dict, db: Session = Depends(get_db)):
    order_id = data.get("order_id")
    rating = data.get("rating")
    comment = data.get("comment", "")
    if not order_id or not rating:
        raise HTTPException(400, "order_id 和 rating 為必填")
    if rating < 1 or rating > 5:
        raise HTTPException(400, "rating 必須在 1-5 之間")
    order = db.query(Order).filter(Order.id == order_id).first()
    if not order:
        raise HTTPException(404, "訂單不存在")
    existing = db.query(OrderReview).filter(OrderReview.order_id == order_id).first()
    if existing:
        raise HTTPException(400, "此訂單已有評論，請使用修改功能")
    now = datetime.now(TW_TZ)
    review = OrderReview(
        order_id=order_id,
        emp_id=order.emp_id,
        emp_name=order.emp_name,
        vendor=order.vendor or "",
        item_name=order.item_name,
        order_date=order.date,
        meal_type=order.type,
        rating=rating,
        comment=comment.strip() if comment else "",
        created_at=now.strftime("%Y-%m-%dT%H:%M:%S+08:00"),
    )
    db.add(review)
    db.commit()
    db.refresh(review)
    return {"id": review.id, "order_id": review.order_id, "emp_id": review.emp_id, "emp_name": review.emp_name,
            "vendor": review.vendor, "item_name": review.item_name, "order_date": review.order_date,
            "meal_type": review.meal_type, "rating": review.rating, "comment": review.comment,
            "created_at": review.created_at}


@app.get("/api/reviews")
def get_reviews(emp_id: str | None = None, vendor: str | None = None,
                date_from: str | None = None, date_to: str | None = None,
                order_id: str | None = None,
                db: Session = Depends(get_db)):
    q = db.query(OrderReview)
    if emp_id:
        q = q.filter(OrderReview.emp_id == emp_id)
    if vendor:
        q = q.filter(OrderReview.vendor.contains(vendor))
    if date_from:
        q = q.filter(OrderReview.order_date >= date_from)
    if date_to:
        q = q.filter(OrderReview.order_date <= date_to)
    if order_id:
        order_ids = [int(x) for x in order_id.split(",") if x.strip()]
        q = q.filter(OrderReview.order_id.in_(order_ids))
    reviews = q.order_by(OrderReview.created_at.desc()).all()
    return [{"id": r.id, "order_id": r.order_id, "emp_id": r.emp_id, "emp_name": r.emp_name,
             "vendor": r.vendor, "item_name": r.item_name, "order_date": r.order_date,
             "meal_type": r.meal_type, "rating": r.rating, "comment": r.comment,
             "created_at": r.created_at} for r in reviews]


@app.put("/api/reviews/{review_id}")
def update_review(review_id: int, data: dict, db: Session = Depends(get_db)):
    review = db.query(OrderReview).filter(OrderReview.id == review_id).first()
    if not review:
        raise HTTPException(404, "評論不存在")
    if "rating" in data:
        if data["rating"] < 1 or data["rating"] > 5:
            raise HTTPException(400, "rating 必須在 1-5 之間")
        review.rating = data["rating"]
    if "comment" in data:
        review.comment = data["comment"].strip() if data["comment"] else ""
    db.commit()
    db.refresh(review)
    return {"id": review.id, "order_id": review.order_id, "emp_id": review.emp_id, "emp_name": review.emp_name,
            "vendor": review.vendor, "item_name": review.item_name, "order_date": review.order_date,
            "meal_type": review.meal_type, "rating": review.rating, "comment": review.comment,
            "created_at": review.created_at}


@app.delete("/api/reviews/{review_id}")
def delete_review(review_id: int, db: Session = Depends(get_db)):
    review = db.query(OrderReview).filter(OrderReview.id == review_id).first()
    if not review:
        raise HTTPException(404, "評論不存在")
    db.delete(review)
    db.commit()
    return {"ok": True}


# ==================== 會議訂餐檢查 ====================

@app.get("/api/meetings")
def get_meetings(month: str | None = None, db: Session = Depends(get_db)):
    q = db.query(MeetingEvent)
    if month:
        q = q.filter(MeetingEvent.date.like(month + '%'))
    meetings = q.order_by(MeetingEvent.date.asc()).all()
    result = []
    for m in meetings:
        atts = db.query(MeetingAttendee).filter(MeetingAttendee.meeting_id == m.id).all()
        count = len(atts)
        notified = 1 if count > 0 and all(a.notified for a in atts) else 0
        result.append({"id": m.id, "date": m.date, "title": m.title,
                       "meal_type": m.meal_type, "deadline_time": m.deadline_time,
                       "check_orders": m.check_orders, "notified": notified,
                       "created_at": m.created_at, "attendee_count": count})
    return result


@app.post("/api/meetings")
def create_meeting(data: dict, db: Session = Depends(get_db)):
    now = datetime.now(TW_TZ)
    meal_type = data.get("meal_type") or None
    deadline_time = data.get("deadline_time") or None
    meeting = MeetingEvent(
        date=data["date"], title=data["title"],
        meal_type=meal_type, deadline_time=deadline_time,
        check_orders=1 if data.get("check_orders", True) else 0,
        created_at=now.strftime("%Y-%m-%dT%H:%M:%S+08:00"),
    )
    db.add(meeting)
    db.flush()
    for att in data.get("attendees", []):
        db.add(MeetingAttendee(meeting_id=meeting.id, emp_id=att["emp_id"], emp_name=att["emp_name"]))
    # 自動建立 deadline_override
    if meal_type and deadline_time:
        db.query(DeadlineOverride).filter(
            DeadlineOverride.date == data["date"],
            DeadlineOverride.type == meal_type,
        ).delete()
        db.add(DeadlineOverride(
            date=data["date"], type=meal_type,
            deadline_time=deadline_time, note=data["title"],
        ))
    db.commit()
    db.refresh(meeting)
    return {"id": meeting.id, "date": meeting.date, "title": meeting.title}


@app.put("/api/meetings/{meeting_id}")
def update_meeting(meeting_id: int, data: dict, db: Session = Depends(get_db)):
    m = db.query(MeetingEvent).get(meeting_id)
    if not m:
        raise HTTPException(404, "活動不存在")
    old_date, old_meal_type = m.date, m.meal_type
    if "date" in data:
        m.date = data["date"]
    if "title" in data:
        m.title = data["title"]
    if "meal_type" in data:
        m.meal_type = data["meal_type"] or None
    if "deadline_time" in data:
        m.deadline_time = data["deadline_time"] or None
    if "check_orders" in data:
        m.check_orders = 1 if data["check_orders"] else 0
    # 同步 deadline_override：先刪舊的再建新的
    if old_meal_type:
        db.query(DeadlineOverride).filter(
            DeadlineOverride.date == old_date,
            DeadlineOverride.type == old_meal_type,
        ).delete()
    if m.meal_type and m.deadline_time:
        db.query(DeadlineOverride).filter(
            DeadlineOverride.date == m.date,
            DeadlineOverride.type == m.meal_type,
        ).delete()
        db.add(DeadlineOverride(
            date=m.date, type=m.meal_type,
            deadline_time=m.deadline_time, note=m.title,
        ))
    db.commit()
    return {"ok": True}


@app.delete("/api/meetings/{meeting_id}")
def delete_meeting(meeting_id: int, db: Session = Depends(get_db)):
    m = db.query(MeetingEvent).get(meeting_id)
    if not m:
        raise HTTPException(404, "活動不存在")
    # 自動刪除關聯的 deadline_override
    if m.meal_type and m.deadline_time:
        db.query(DeadlineOverride).filter(
            DeadlineOverride.date == m.date,
            DeadlineOverride.type == m.meal_type,
        ).delete()
    db.query(MeetingAttendee).filter(MeetingAttendee.meeting_id == meeting_id).delete()
    db.delete(m)
    db.commit()
    return {"ok": True}


@app.get("/api/meetings/{meeting_id}/attendees")
def get_meeting_attendees(meeting_id: int, db: Session = Depends(get_db)):
    atts = db.query(MeetingAttendee).filter(MeetingAttendee.meeting_id == meeting_id).all()
    return [{"id": a.id, "emp_id": a.emp_id, "emp_name": a.emp_name, "notified": a.notified} for a in atts]


@app.post("/api/meetings/{meeting_id}/attendees")
def add_meeting_attendees(meeting_id: int, data: dict, db: Session = Depends(get_db)):
    m = db.query(MeetingEvent).get(meeting_id)
    if not m:
        raise HTTPException(404, "會議不存在")
    added = 0
    for att in data.get("attendees", []):
        existing = db.query(MeetingAttendee).filter(
            MeetingAttendee.meeting_id == meeting_id,
            MeetingAttendee.emp_id == att["emp_id"],
        ).first()
        if not existing:
            db.add(MeetingAttendee(meeting_id=meeting_id, emp_id=att["emp_id"], emp_name=att["emp_name"]))
            added += 1
    db.commit()
    return {"added": added}


@app.get("/api/meetings/attendees-template")
def download_attendees_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "出席人員"
    ws.append(["工號", "姓名"])
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 15
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    from urllib.parse import quote
    filename = quote("出席人員範本.xlsx")
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename}"})


@app.post("/api/meetings/{meeting_id}/upload-attendees")
async def upload_meeting_attendees(meeting_id: int, file: UploadFile = File(...), db: Session = Depends(get_db)):
    m = db.query(MeetingEvent).get(meeting_id)
    if not m:
        raise HTTPException(404, "會議不存在")
    content = await file.read()
    wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise HTTPException(400, "Excel 檔案為空")
    # 找工號和姓名欄位
    header = [str(c).strip() if c else '' for c in rows[0]]
    emp_id_col = None
    name_col = None
    for i, h in enumerate(header):
        hl = h.lower()
        if '工號' in h or 'emp' in hl or 'id' in hl or 'account' in hl:
            emp_id_col = i
        if '姓名' in h or 'name' in hl:
            name_col = i
    # 如果找不到欄位，假設第一欄是工號、第二欄是姓名
    if emp_id_col is None:
        emp_id_col = 0
    if name_col is None:
        name_col = 1 if len(header) > 1 else 0
    added = 0
    skipped = 0
    not_found = []
    for row in rows[1:]:
        if not row or not row[emp_id_col]:
            continue
        raw_id = str(row[emp_id_col]).strip()
        raw_name = str(row[name_col]).strip() if name_col < len(row) and row[name_col] else ''
        # 用工號查員工表確認
        emp = db.query(Employee).filter(Employee.account_id == raw_id).first()
        if emp:
            emp_id_val = emp.account_id
            emp_name_val = emp.name
        elif raw_name:
            # 工號找不到但有姓名，用姓名查
            emp = db.query(Employee).filter(Employee.name == raw_name).first()
            if emp:
                emp_id_val = emp.account_id
                emp_name_val = emp.name
            else:
                not_found.append(raw_name or raw_id)
                continue
        else:
            not_found.append(raw_id)
            continue
        existing = db.query(MeetingAttendee).filter(
            MeetingAttendee.meeting_id == meeting_id,
            MeetingAttendee.emp_id == emp_id_val,
        ).first()
        if not existing:
            db.add(MeetingAttendee(meeting_id=meeting_id, emp_id=emp_id_val, emp_name=emp_name_val))
            added += 1
        else:
            skipped += 1
    db.commit()
    msg = f"已匯入 {added} 位"
    if skipped:
        msg += f"，跳過 {skipped} 位重複"
    if not_found:
        msg += f"，找不到 {len(not_found)} 位：{', '.join(not_found[:10])}"
    return {"added": added, "skipped": skipped, "not_found": not_found, "message": msg}


@app.delete("/api/meetings/{meeting_id}/attendees/{emp_id}")
def remove_meeting_attendee(meeting_id: int, emp_id: str, db: Session = Depends(get_db)):
    att = db.query(MeetingAttendee).filter(
        MeetingAttendee.meeting_id == meeting_id,
        MeetingAttendee.emp_id == emp_id,
    ).first()
    if att:
        db.delete(att)
        m = db.query(MeetingEvent).get(meeting_id)
        if m and m.notified:
            m.notified = 0
        db.commit()
    return {"ok": True}


@app.post("/api/meetings/{meeting_id}/copy-attendees/{source_id}")
def copy_meeting_attendees(meeting_id: int, source_id: int, db: Session = Depends(get_db)):
    m = db.query(MeetingEvent).get(meeting_id)
    if not m:
        raise HTTPException(404, "目標會議不存在")
    source_atts = db.query(MeetingAttendee).filter(MeetingAttendee.meeting_id == source_id).all()
    if not source_atts:
        raise HTTPException(404, "來源會議無出席人員")
    added = 0
    for att in source_atts:
        existing = db.query(MeetingAttendee).filter(
            MeetingAttendee.meeting_id == meeting_id,
            MeetingAttendee.emp_id == att.emp_id,
        ).first()
        if not existing:
            db.add(MeetingAttendee(meeting_id=meeting_id, emp_id=att.emp_id, emp_name=att.emp_name))
            added += 1
    db.commit()
    return {"added": added, "total_source": len(source_atts)}


@app.get("/api/meetings/{meeting_id}/check")
def check_meeting_dinner(meeting_id: int, db: Session = Depends(get_db)):
    m = db.query(MeetingEvent).get(meeting_id)
    if not m:
        raise HTTPException(404, "活動不存在")
    meal_type = m.meal_type or "晚餐"
    atts = db.query(MeetingAttendee).filter(MeetingAttendee.meeting_id == meeting_id).all()
    result = []
    for att in atts:
        orders = db.query(Order).filter(
            Order.emp_id == att.emp_id,
            Order.date == m.date,
            Order.type == meal_type,
        ).all()
        result.append({
            "emp_id": att.emp_id,
            "emp_name": att.emp_name,
            "has_ordered": len(orders) > 0,
            "orders": [{"item_name": o.item_name, "self_paid": o.self_paid} for o in orders],
        })
    return {"date": m.date, "title": m.title, "meal_type": meal_type, "attendees": result}


@app.post("/api/meetings/{meeting_id}/notify")
def notify_meeting_dinner(meeting_id: int, body: dict = None, db: Session = Depends(get_db)):
    m = db.query(MeetingEvent).get(meeting_id)
    if not m:
        raise HTTPException(404, "活動不存在")
    meal_type = m.meal_type or "晚餐"
    emp_ids = (body or {}).get("emp_ids", [])
    if emp_ids:
        # 只通知指定的人
        not_ordered = db.query(MeetingAttendee).filter(
            MeetingAttendee.meeting_id == meeting_id,
            MeetingAttendee.emp_id.in_(emp_ids),
        ).all()
    else:
        # 沒指定則通知所有未訂的人（向下相容）
        atts = db.query(MeetingAttendee).filter(MeetingAttendee.meeting_id == meeting_id).all()
        not_ordered = []
        for att in atts:
            has = db.query(Order).filter(
                Order.emp_id == att.emp_id, Order.date == m.date, Order.type == meal_type,
            ).first()
            if not has:
                not_ordered.append(att)
    if not not_ordered:
        return {"notified": 0, "message": "所有與會者都已訂餐"}
    # 發企微通知
    weekdays = ['一', '二', '三', '四', '五', '六', '日']
    from datetime import date as _date
    d = _date.fromisoformat(m.date)
    wd = weekdays[d.weekday()]
    date_str = f"{d.month}/{d.day}({wd})"
    months_en = ['','January','February','March','April','May','June','July','August','September','October','November','December']
    date_en = f"{months_en[d.month]} {d.day}"
    meal_en = {"午餐":"lunch","飲料":"beverages","晚餐":"dinner","早餐":"breakfast"}.get(meal_type, meal_type)
    msg = f"【訂餐提醒】\n{date_str} 有「{m.title}」，請記得訂購{meal_type}。\n【Meal Reminder】\nPlease remember to order {meal_en} for \"{m.title}\" on {date_en}."
    notified_list = []
    try:
        import requests as req
        token_r = req.get("https://qyapi.weixin.qq.com/cgi-bin/gettoken", params={
            "corpid": "ww5c54aa97ecd16ce0",
            "corpsecret": "2w4aYkNdzQp6BLXivTZ2y89NkDEjnKkqOdf3BLI8xcs"
        }, timeout=5)
        token = token_r.json().get("access_token")
        if token:
            for att in not_ordered:
                wechat_user = WECHAT_ID_MAP.get(att.emp_id, att.emp_id)
                req.post("https://qyapi.weixin.qq.com/cgi-bin/message/send",
                    params={"access_token": token},
                    json={"msgtype": "text", "touser": wechat_user, "agentid": 1000169,
                          "text": {"content": msg}}, timeout=5)
                notified_list.append(att.emp_name)
                att.notified = 1
    except Exception:
        pass
    m.notified = 1
    db.commit()
    return {"notified": len(notified_list), "names": notified_list}


# ==================== Uniform APIs ====================

@app.post("/api/uniform/apply", response_model=UniformApplicationOut)
def create_uniform_application(req: UniformApplicationCreate, db: Session = Depends(get_db)):
    total = req.long_sleeve_qty + req.short_sleeve_qty
    if total > 12:
        raise HTTPException(status_code=400, detail="襯衫/POLO衫總數不可超過12件")
    if req.long_sleeve_qty < 0 or req.short_sleeve_qty < 0:
        raise HTTPException(status_code=400, detail="數量不可為負數")
    now = datetime.now(TW_TZ)
    obj = UniformApplication(
        emp_id=req.emp_id.strip(), emp_name=req.emp_name.strip(),
        dept_code=req.dept_code, location=req.location,
        height=req.height, weight=req.weight,
        chest=req.chest, shoulder=req.shoulder,
        gender=req.gender,
        recommended_size=req.recommended_size, selected_size=req.selected_size,
        uniform_type=req.uniform_type,
        long_sleeve_qty=req.long_sleeve_qty, short_sleeve_qty=req.short_sleeve_qty,
        outerwear_type=req.outerwear_type, outerwear_qty=req.outerwear_qty,
        status='pending', remarks=req.remarks,
        created_at=now.strftime("%Y-%m-%dT%H:%M:%S+08:00"),
    )
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


@app.get("/api/uniform/applications", response_model=list[UniformApplicationOut])
def get_uniform_applications(emp_id: str | None = None, status: str | None = None, db: Session = Depends(get_db)):
    q = db.query(UniformApplication)
    if emp_id:
        q = q.filter(UniformApplication.emp_id == emp_id)
    if status:
        q = q.filter(UniformApplication.status == status)
    return q.order_by(UniformApplication.id.desc()).all()


@app.put("/api/uniform/applications/{app_id}/receive", response_model=UniformApplicationOut)
def receive_uniform(app_id: int, db: Session = Depends(get_db)):
    obj = db.query(UniformApplication).filter(UniformApplication.id == app_id).first()
    if not obj:
        raise HTTPException(status_code=404, detail="申請單不存在")
    if obj.status != 'shipped':
        raise HTTPException(status_code=400, detail="只有已出貨的申請單可以簽收")
    now = datetime.now(TW_TZ)
    obj.status = 'received'
    obj.received_at = now.strftime("%Y-%m-%dT%H:%M:%S+08:00")
    db.commit()
    db.refresh(obj)
    return obj


# --- Roles ---

@app.get("/api/uniform/roles")
def get_uniform_roles(db: Session = Depends(get_db)):
    return [UniformRoleOut.model_validate(r) for r in db.query(UniformRole).order_by(UniformRole.role, UniformRole.emp_name).all()]


@app.get("/api/uniform/my-role")
def get_my_uniform_role(emp_id: str, db: Session = Depends(get_db)):
    r = db.query(UniformRole).filter(UniformRole.emp_id == emp_id).first()
    if r:
        return {"role": r.role, "location": r.location}
    return {"role": "employee", "location": None}


@app.post("/api/uniform/roles", response_model=UniformRoleOut)
def create_uniform_role(req: UniformRoleCreate, db: Session = Depends(get_db)):
    existing = db.query(UniformRole).filter(UniformRole.emp_id == req.emp_id).first()
    now = datetime.now(TW_TZ)
    if existing:
        existing.role = req.role
        existing.location = req.location
        existing.emp_name = req.emp_name
        db.commit()
        db.refresh(existing)
        return existing
    obj = UniformRole(emp_id=req.emp_id, emp_name=req.emp_name, role=req.role, location=req.location, created_at=now.strftime("%Y-%m-%dT%H:%M:%S+08:00"))
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


@app.delete("/api/uniform/roles/{role_id}")
def delete_uniform_role(role_id: int, db: Session = Depends(get_db)):
    obj = db.query(UniformRole).filter(UniformRole.id == role_id).first()
    if not obj:
        raise HTTPException(status_code=404, detail="角色不存在")
    db.delete(obj)
    db.commit()
    return {"ok": True}


@app.get("/api/uniform/sizes", response_model=list[UniformSizeOut])
def get_uniform_sizes(db: Session = Depends(get_db)):
    return db.query(UniformSize).order_by(UniformSize.gender, UniformSize.min_height).all()


@app.put("/api/uniform/applications/{app_id}", response_model=UniformApplicationOut)
def update_uniform_application(app_id: int, data: UniformApplicationUpdate, db: Session = Depends(get_db)):
    obj = db.query(UniformApplication).filter(UniformApplication.id == app_id).first()
    if not obj:
        raise HTTPException(status_code=404, detail="申請單不存在")
    if data.status is not None:
        obj.status = data.status
    if data.approved_by is not None:
        obj.approved_by = data.approved_by
    if data.shipped_by is not None:
        obj.shipped_by = data.shipped_by
    if data.remarks is not None:
        obj.remarks = data.remarks
    db.commit()
    db.refresh(obj)
    return obj


# --- Inventory ---

@app.get("/api/uniform/inventory", response_model=list[UniformInventoryOut])
def get_uniform_inventory(location: str | None = None, db: Session = Depends(get_db)):
    q = db.query(UniformInventory)
    if location:
        q = q.filter(UniformInventory.location == location)
    return q.order_by(UniformInventory.location, UniformInventory.item_type, UniformInventory.size).all()


@app.post("/api/uniform/inventory")
def create_or_update_inventory(location: str, item_type: str, size: str, quantity: int = 0, db: Session = Depends(get_db)):
    obj = db.query(UniformInventory).filter(
        UniformInventory.location == location,
        UniformInventory.item_type == item_type,
        UniformInventory.size == size,
    ).first()
    if obj:
        obj.quantity = quantity
    else:
        obj = UniformInventory(location=location, item_type=item_type, size=size, quantity=quantity)
        db.add(obj)
    db.commit()
    db.refresh(obj)
    return UniformInventoryOut.model_validate(obj)


@app.put("/api/uniform/inventory/{inv_id}")
def update_inventory(inv_id: int, data: UniformInventoryUpdate, db: Session = Depends(get_db)):
    obj = db.query(UniformInventory).filter(UniformInventory.id == inv_id).first()
    if not obj:
        raise HTTPException(status_code=404, detail="庫存紀錄不存在")
    if data.quantity is not None:
        obj.quantity = data.quantity
    if data.safety_months is not None:
        obj.safety_months = data.safety_months
    db.commit()
    db.refresh(obj)
    return UniformInventoryOut.model_validate(obj)


@app.post("/api/uniform/inventory/{inv_id}/adjust")
def adjust_inventory(inv_id: int, change_qty: int, reason: str = "", operator: str = "", db: Session = Depends(get_db)):
    obj = db.query(UniformInventory).filter(UniformInventory.id == inv_id).first()
    if not obj:
        raise HTTPException(status_code=404, detail="庫存紀錄不存在")
    obj.quantity += change_qty
    if obj.quantity < 0:
        obj.quantity = 0
    now = datetime.now(TW_TZ)
    log = UniformInventoryLog(
        inventory_id=inv_id, change_qty=change_qty,
        reason=reason, operator=operator,
        created_at=now.strftime("%Y-%m-%dT%H:%M:%S+08:00"),
    )
    db.add(log)
    db.commit()
    db.refresh(obj)
    return UniformInventoryOut.model_validate(obj)


@app.get("/api/uniform/inventory/logs", response_model=list[UniformInventoryLogOut])
def get_inventory_logs(inventory_id: int | None = None, db: Session = Depends(get_db)):
    q = db.query(UniformInventoryLog)
    if inventory_id:
        q = q.filter(UniformInventoryLog.inventory_id == inventory_id)
    return q.order_by(UniformInventoryLog.id.desc()).limit(200).all()


@app.get("/api/uniform/purchase-alert")
def get_purchase_alert(db: Session = Depends(get_db)):
    """計算各品項月消耗量，回傳低於安全月份的品項"""
    from collections import defaultdict
    # 最近6個月的申請（已核准以上）
    six_months_ago = (datetime.now(TW_TZ) - timedelta(days=180)).strftime("%Y-%m-%d")
    apps = db.query(UniformApplication).filter(
        UniformApplication.created_at >= six_months_ago,
        UniformApplication.status.in_(['approved', 'shipped', 'received']),
    ).all()
    # 統計消耗量
    consumption = defaultdict(int)
    for a in apps:
        if a.uniform_type == 'A_shirt':
            consumption[('shirt_long', a.selected_size)] += a.long_sleeve_qty
            consumption[('shirt_short', a.selected_size)] += a.short_sleeve_qty
            consumption[('blazer', a.selected_size)] += a.outerwear_qty
        else:
            consumption[('polo_long', a.selected_size)] += a.long_sleeve_qty
            consumption[('polo_short', a.selected_size)] += a.short_sleeve_qty
            consumption[('jacket', a.selected_size)] += a.outerwear_qty
    # 月平均
    monthly = {k: round(v / 6, 1) for k, v in consumption.items()}
    # 比對庫存
    alerts = []
    inventories = db.query(UniformInventory).all()
    for inv in inventories:
        key = (inv.item_type, inv.size)
        avg = monthly.get(key, 0)
        months_left = round(inv.quantity / avg, 1) if avg > 0 else 999
        if months_left < inv.safety_months:
            alerts.append({
                "location": inv.location, "item_type": inv.item_type,
                "size": inv.size, "quantity": inv.quantity,
                "monthly_avg": avg, "months_left": months_left,
                "safety_months": inv.safety_months,
                "suggested_qty": max(0, round(avg * inv.safety_months - inv.quantity)),
            })
    return sorted(alerts, key=lambda x: x["months_left"])


# ==================== Static Files ====================

STATIC_DIR = Path(__file__).parent / "static"


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/")
def serve_index():
    content = (STATIC_DIR / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(content, headers={"Cache-Control": "no-cache, no-store, must-revalidate", "Pragma": "no-cache", "Expires": "0"})
