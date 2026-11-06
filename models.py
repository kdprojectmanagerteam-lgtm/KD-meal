from sqlalchemy import Column, Integer, String, Float, DateTime, Date, Text, UniqueConstraint
from database import Base


class MenuItem(Base):
    __tablename__ = "menu_items"

    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(String(20), nullable=False, index=True)       # YYYY-MM-DD
    type = Column(String(20), nullable=False)                    # 早餐/午餐/飲料/晚餐
    vendor = Column(String(100), nullable=False)
    phone = Column(String(50), nullable=True)                     # 店家電話
    name = Column(String(200), nullable=False)
    price = Column(Integer, nullable=False, default=0)
    category = Column(String(20), nullable=True)                  # 飯/麵/特/素
    remarks = Column(String(500), nullable=True)                   # 備註（全素/蛋奶素等）
    sold_out = Column(Integer, nullable=False, default=0)    # 1=售完
    start_time = Column(String(30), nullable=True)               # ISO datetime
    deadline = Column(String(30), nullable=True)                  # ISO datetime


class Order(Base):
    __tablename__ = "orders"

    id = Column(Integer, primary_key=True, autoincrement=True)
    emp_id = Column(String(20), nullable=False, index=True)
    emp_name = Column(String(50), nullable=False)
    date = Column(String(20), nullable=False, index=True)        # YYYY-MM-DD
    type = Column(String(20), nullable=False)                     # 早餐/午餐/飲料/晚餐
    floor = Column(String(10), nullable=True)                      # 3F/5F/15F/16F
    item_name = Column(String(200), nullable=False)
    vendor = Column(String(100), nullable=True)                     # 店家名稱
    qty = Column(Integer, nullable=False, default=1)
    price = Column(Integer, nullable=False, default=0)
    remarks = Column(String(500), nullable=True)                    # 備註
    company_name = Column(String(100), nullable=True)                # 廠商名稱（選填，填寫後自動歸公司負擔）
    company_paid = Column(Integer, nullable=False, default=0)  # 1=公司招待（廠商用餐）
    guest_dept = Column(String(100), nullable=True)                  # 來客/廠商所屬部門
    guest_qty = Column(Integer, nullable=True)                        # 來客/廠商人數
    self_paid = Column(Integer, nullable=False, default=0)     # 1=晚餐自費
    proxy_by = Column(String(50), nullable=True)                   # 代點人姓名
    paid_by = Column(String(20), nullable=True)                    # 主管請客：付款人工號
    paid_by_name = Column(String(50), nullable=True)               # 主管請客：付款人姓名
    treat_reason = Column(String(200), nullable=True)              # 公司請客原因
    created_at = Column(String(50), nullable=True)


class Employee(Base):
    __tablename__ = "employees"

    id = Column(Integer, primary_key=True, autoincrement=True)
    account_id = Column(String(20), nullable=False, unique=True, index=True)  # 工號
    name = Column(String(50), nullable=False)                                  # 姓名
    dept_code = Column(String(20), nullable=True)                              # 部門代碼
    is_active = Column(Integer, nullable=False, default=1)
    password_hash = Column(String(255), nullable=True)                         # bcrypt hash, NULL=預設密碼
    must_change_pwd = Column(Integer, nullable=False, default=1)               # 1=強制改密碼
    is_admin = Column(Integer, nullable=False, default=0)                      # 1=有後台管理權限


class DrinkOption(Base):
    __tablename__ = "drink_options"

    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(String(20), nullable=False, index=True)          # YYYY-MM-DD
    vendor = Column(String(100), nullable=False)                     # 店家
    option_type = Column(String(20), nullable=False)                # topping/sweetness/ice
    name = Column(String(200), nullable=False)
    price = Column(Integer, nullable=False, default=0)


class ExtraFee(Base):
    __tablename__ = "extra_fees"

    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(String(20), nullable=False, index=True)          # YYYY-MM-DD
    type = Column(String(20), nullable=False)                       # 飲料/午餐/晚餐/早餐
    description = Column(String(200), nullable=False)                # 袋子費、運費等
    amount = Column(Integer, nullable=False, default=0)


class Holiday(Base):
    __tablename__ = "holidays"

    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(String(20), nullable=False, unique=True, index=True)  # YYYY-MM-DD
    name = Column(String(50), nullable=False)                            # 端午節、國慶日...


class VendorNote(Base):
    __tablename__ = "vendor_notes"

    id = Column(Integer, primary_key=True, autoincrement=True)
    vendor = Column(String(100), nullable=False, index=True)
    date = Column(String(10), nullable=False, default='')
    note = Column(String(500), nullable=False, default='')
    confirmed_by = Column(String(50), nullable=False, default='')


class DeadlineOverride(Base):
    __tablename__ = "deadline_overrides"

    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(String(20), nullable=False, index=True)          # YYYY-MM-DD
    type = Column(String(20), nullable=False)                       # 午餐/飲料/晚餐/早餐
    deadline_time = Column(String(10), nullable=False)              # HH:MM
    note = Column(String(200), nullable=True)                       # 備註，如：主管會議


class CoffeeItem(Base):
    __tablename__ = "coffee_items"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(200), nullable=False)
    price = Column(Integer, nullable=False, default=0)
    is_active = Column(Integer, nullable=False, default=1)


class CoffeeOrder(Base):
    __tablename__ = "coffee_orders"

    id = Column(Integer, primary_key=True, autoincrement=True)
    emp_id = Column(String(20), nullable=False, index=True)
    emp_name = Column(String(50), nullable=False)
    floor = Column(String(10), nullable=True)
    date = Column(String(20), nullable=False, index=True)
    coffee_name = Column(String(200), nullable=False)
    price = Column(Integer, nullable=False, default=0)
    quantity = Column(Integer, nullable=False, default=1)
    total_amount = Column(Integer, nullable=False, default=0)
    temp = Column(String(10), nullable=False)                    # 冷/熱
    time_pref = Column(String(10), nullable=False)               # 上午/下午/自取
    customizations = Column(String(500), nullable=True)          # JSON string
    note = Column(String(500), nullable=True)
    status = Column(String(20), nullable=False, default='待製作')  # 待製作/已完成
    production_time = Column(String(30), nullable=True)
    company_paid = Column(Integer, nullable=False, default=0)
    department = Column(String(100), nullable=True)
    created_at = Column(String(50), nullable=True)


class TreatNotification(Base):
    __tablename__ = "treat_notifications"

    id = Column(Integer, primary_key=True, autoincrement=True)
    paid_by = Column(String(20), nullable=False, index=True)       # 被請客人工號
    paid_by_name = Column(String(50), nullable=False)               # 被請客人姓名
    orderer_id = Column(String(20), nullable=False)                 # 點餐人工號
    orderer_name = Column(String(50), nullable=False)               # 點餐人姓名
    date = Column(String(20), nullable=True)                        # 點餐日期
    items_summary = Column(String(500), nullable=False)             # 品項摘要
    total = Column(Integer, nullable=False, default=0)              # 總金額
    is_read = Column(Integer, nullable=False, default=0)            # 0=未讀 1=已讀
    created_at = Column(String(50), nullable=True)


class MeetingEvent(Base):
    __tablename__ = "meeting_events"

    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(String(20), nullable=False, index=True)          # YYYY-MM-DD
    title = Column(String(200), nullable=False)                     # 活動名稱
    meal_type = Column(String(20), nullable=True)                   # 午餐/飲料/晚餐/早餐
    deadline_time = Column(String(10), nullable=True)               # HH:MM 截止時間覆蓋
    check_orders = Column(Integer, nullable=False, default=1)       # 1=需檢查訂餐 0=僅調整截止
    notified = Column(Integer, nullable=False, default=0)           # 0=未通知 1=已通知
    created_at = Column(String(50), nullable=True)


class MeetingAttendee(Base):
    __tablename__ = "meeting_attendees"
    __table_args__ = (UniqueConstraint('meeting_id', 'emp_id'),)

    id = Column(Integer, primary_key=True, autoincrement=True)
    meeting_id = Column(Integer, nullable=False, index=True)        # meeting_events.id
    emp_id = Column(String(20), nullable=False)                     # 工號
    emp_name = Column(String(50), nullable=False)                   # 姓名
    notified = Column(Integer, nullable=False, default=0)           # 0=未通知 1=已通知


class OrderReview(Base):
    __tablename__ = "order_reviews"

    id = Column(Integer, primary_key=True, autoincrement=True)
    order_id = Column(Integer, nullable=False, unique=True, index=True)  # 關聯 orders.id
    emp_id = Column(String(20), nullable=False, index=True)
    emp_name = Column(String(50), nullable=False)
    vendor = Column(String(100), nullable=True)
    item_name = Column(String(200), nullable=False)
    order_date = Column(String(20), nullable=True)                        # 訂單日期
    meal_type = Column(String(20), nullable=True)                         # 餐別
    rating = Column(Integer, nullable=False)                              # 1-5 星
    comment = Column(Text, nullable=True)                                 # 文字評論
    created_at = Column(String(50), nullable=True)


# ==================== Uniform ====================

class UniformRole(Base):
    __tablename__ = "uniform_roles"

    id = Column(Integer, primary_key=True, autoincrement=True)
    emp_id = Column(String(20), nullable=False, unique=True, index=True)
    emp_name = Column(String(50), nullable=False)
    role = Column(String(30), nullable=False)         # hr / local_admin / warehouse / super_admin
    location = Column(String(50), nullable=True)       # 當地窗口的據點
    created_at = Column(String(50), nullable=True)


class UniformApplication(Base):
    __tablename__ = "uniform_applications"

    id = Column(Integer, primary_key=True, autoincrement=True)
    emp_id = Column(String(20), nullable=False, index=True)
    emp_name = Column(String(50), nullable=False)
    dept_code = Column(String(20), nullable=True)
    location = Column(String(50), nullable=False)
    height = Column(Float, nullable=False)
    weight = Column(Float, nullable=False)
    chest = Column(Float, nullable=True)
    shoulder = Column(Float, nullable=True)
    gender = Column(String(10), nullable=False)
    recommended_size = Column(String(10), nullable=False)
    selected_size = Column(String(10), nullable=False)
    uniform_type = Column(String(20), nullable=False)       # A_shirt / B_polo
    long_sleeve_qty = Column(Integer, nullable=False, default=0)
    short_sleeve_qty = Column(Integer, nullable=False, default=0)
    outerwear_type = Column(String(20), nullable=False)     # blazer / jacket
    outerwear_qty = Column(Integer, nullable=False, default=0)
    status = Column(String(20), nullable=False, default='pending')
    approved_by = Column(String(50), nullable=True)
    shipped_by = Column(String(50), nullable=True)
    received_at = Column(String(50), nullable=True)
    remarks = Column(String(500), nullable=True)
    created_at = Column(String(50), nullable=True)


class UniformInventory(Base):
    __tablename__ = "uniform_inventory"

    id = Column(Integer, primary_key=True, autoincrement=True)
    location = Column(String(50), nullable=False)         # 據點
    item_type = Column(String(30), nullable=False)        # shirt_long/shirt_short/polo_long/polo_short/blazer/jacket
    size = Column(String(10), nullable=False)              # XS/S/M/L/XL/2XL/3XL
    quantity = Column(Integer, nullable=False, default=0)
    safety_months = Column(Integer, nullable=False, default=6)  # 安全月份數


class UniformInventoryLog(Base):
    __tablename__ = "uniform_inventory_logs"

    id = Column(Integer, primary_key=True, autoincrement=True)
    inventory_id = Column(Integer, nullable=False)
    change_qty = Column(Integer, nullable=False)           # 正=入庫 負=出庫
    reason = Column(String(200), nullable=True)
    operator = Column(String(50), nullable=True)
    created_at = Column(String(50), nullable=True)


class UniformSize(Base):
    __tablename__ = "uniform_sizes"

    id = Column(Integer, primary_key=True, autoincrement=True)
    gender = Column(String(10), nullable=False)
    size_code = Column(String(10), nullable=False)        # 7/9/11/13/15/18 or 37/38/39...
    size_label = Column(String(10), nullable=False)        # S/M/L/XL/2L/3L
    min_height = Column(Float, nullable=False, default=0)
    max_height = Column(Float, nullable=False, default=999)
    min_weight = Column(Float, nullable=False, default=0)
    max_weight = Column(Float, nullable=False, default=999)
    shoulder_width = Column(Float, nullable=True)          # 肩寬 cm
    chest_width = Column(Float, nullable=True)             # 胸寬 cm
    full_length = Column(Float, nullable=True)             # 全長 cm
    sleeve_length = Column(Float, nullable=True)           # 袖長 cm
