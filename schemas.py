from pydantic import BaseModel


# --- MenuItem ---
class MenuItemCreate(BaseModel):
    date: str
    type: str
    vendor: str
    phone: str | None = None
    name: str
    price: int
    category: str | None = None
    start_time: str | None = None
    deadline: str | None = None


class MenuItemBatchImport(BaseModel):
    date: str
    type: str
    start_time: str | None = None
    deadline: str | None = None
    items: list[dict]  # [{"vendor": "...", "name": "...", "price": 100, "category": "..."}, ...]


class MenuItemOut(BaseModel):
    id: int
    date: str
    type: str
    vendor: str
    phone: str | None = None
    name: str
    price: int
    category: str | None = None
    remarks: str | None = None
    sold_out: int = 0
    start_time: str | None = None
    deadline: str | None = None

    model_config = {"from_attributes": True, "serialize_default": True}


# --- Order ---
class OrderCreate(BaseModel):
    emp_id: str
    emp_name: str
    floor: str | None = None
    company_name: str | None = None
    company_paid: bool = False
    guest_dept: str | None = None
    guest_qty: int | None = None
    proxy_by: str | None = None
    paid_by: str | None = None
    paid_by_name: str | None = None
    treat_reason: str | None = None
    items: list[dict]  # [{"date": "...", "type": "...", "item_name": "...", "qty": 1, "price": 100, "remarks": ""}, ...]


class OrderOut(BaseModel):
    id: int
    emp_id: str
    emp_name: str
    date: str
    type: str
    floor: str | None = None
    item_name: str
    vendor: str | None = None
    qty: int
    price: int
    remarks: str | None = None
    company_name: str | None = None
    company_paid: int = 0
    guest_dept: str | None = None
    guest_qty: int | None = None
    self_paid: int = 0
    proxy_by: str | None = None
    paid_by: str | None = None
    paid_by_name: str | None = None
    treat_reason: str | None = None
    created_at: str | None = None

    model_config = {"from_attributes": True, "serialize_default": True}


class OrderUpdate(BaseModel):
    item_name: str | None = None
    qty: int | None = None
    price: int | None = None
    remarks: str | None = None
    self_paid: int | None = None


# --- ExtraFee ---
class ExtraFeeCreate(BaseModel):
    date: str
    type: str
    description: str
    amount: int


class ExtraFeeOut(BaseModel):
    id: int
    date: str
    type: str
    description: str
    amount: int

    model_config = {"from_attributes": True, "serialize_default": True}


# --- Auth ---
class LoginRequest(BaseModel):
    account_id: str
    password: str


class ChangePasswordRequest(BaseModel):
    account_id: str
    old_password: str
    new_password: str


class ForgotPasswordRequest(BaseModel):
    account_id: str
    name: str


# --- Coffee ---
class CoffeeItemCreate(BaseModel):
    name: str
    price: int
    is_active: int = 1


class CoffeeItemOut(BaseModel):
    id: int
    name: str
    price: int
    is_active: int

    model_config = {"from_attributes": True, "serialize_default": True}


class CoffeeOrderCreate(BaseModel):
    emp_id: str
    emp_name: str
    floor: str | None = None
    date: str | None = None
    items: list[dict]
    company_paid: bool = False
    department: str | None = None


class CoffeeOrderOut(BaseModel):
    id: int
    emp_id: str
    emp_name: str
    floor: str | None = None
    date: str
    coffee_name: str
    price: int
    quantity: int
    total_amount: int
    temp: str
    time_pref: str
    customizations: str | None = None
    note: str | None = None
    status: str
    production_time: str | None = None
    company_paid: int = 0
    department: str | None = None
    created_at: str | None = None

    model_config = {"from_attributes": True, "serialize_default": True}


class SetAdminRequest(BaseModel):
    account_id: str
    is_admin: bool


# --- Uniform ---
class UniformRoleCreate(BaseModel):
    emp_id: str
    emp_name: str
    role: str
    location: str | None = None


class UniformRoleOut(BaseModel):
    id: int
    emp_id: str
    emp_name: str
    role: str
    location: str | None = None
    created_at: str | None = None

    model_config = {"from_attributes": True, "serialize_default": True}


class UniformApplicationCreate(BaseModel):
    emp_id: str
    emp_name: str
    dept_code: str | None = None
    location: str
    height: float
    weight: float
    chest: float | None = None
    shoulder: float | None = None
    gender: str
    recommended_size: str
    selected_size: str
    uniform_type: str
    long_sleeve_qty: int
    short_sleeve_qty: int
    outerwear_type: str
    outerwear_qty: int
    remarks: str | None = None


class UniformApplicationOut(BaseModel):
    id: int
    emp_id: str
    emp_name: str
    dept_code: str | None = None
    location: str
    height: float
    weight: float
    chest: float | None = None
    shoulder: float | None = None
    gender: str
    recommended_size: str
    selected_size: str
    uniform_type: str
    long_sleeve_qty: int
    short_sleeve_qty: int
    outerwear_type: str
    outerwear_qty: int
    status: str
    approved_by: str | None = None
    shipped_by: str | None = None
    received_at: str | None = None
    remarks: str | None = None
    created_at: str | None = None

    model_config = {"from_attributes": True, "serialize_default": True}


class UniformApplicationUpdate(BaseModel):
    status: str | None = None
    approved_by: str | None = None
    shipped_by: str | None = None
    remarks: str | None = None


class UniformInventoryOut(BaseModel):
    id: int
    location: str
    item_type: str
    size: str
    quantity: int
    safety_months: int

    model_config = {"from_attributes": True, "serialize_default": True}


class UniformInventoryUpdate(BaseModel):
    quantity: int | None = None
    safety_months: int | None = None


class UniformInventoryLogOut(BaseModel):
    id: int
    inventory_id: int
    change_qty: int
    reason: str | None = None
    operator: str | None = None
    created_at: str | None = None

    model_config = {"from_attributes": True, "serialize_default": True}


class UniformSizeOut(BaseModel):
    id: int
    gender: str
    size_code: str
    size_label: str
    min_height: float
    max_height: float
    min_weight: float
    max_weight: float
    shoulder_width: float | None = None
    chest_width: float | None = None
    full_length: float | None = None
    sleeve_length: float | None = None

    model_config = {"from_attributes": True, "serialize_default": True}
