
from datetime import datetime, date

from sqlalchemy import Column, Integer, String, Float, DateTime, Date, Text, ForeignKey
from sqlalchemy.orm import relationship

from app.database import Base


class ExchangeRate(Base):
    __tablename__ = "exchange_rates"

    id = Column(Integer, primary_key=True, index=True)
    currency = Column(String(10), unique=True, nullable=False)
    rate_to_twd = Column(Float, nullable=False, default=1.0)
    updated_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class PriceSetting(Base):
    __tablename__ = "price_settings"

    id = Column(Integer, primary_key=True, index=True)
    level_name = Column(String(50), unique=True, nullable=False)
    discount_ratio = Column(Float, nullable=False, default=100.0)
    updated_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class Dealer(Base):
    __tablename__ = "dealers"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(255), unique=True, nullable=False)
    level = Column(String(50), nullable=True)
    tax_id = Column(String(50), nullable=True)
    address = Column(String(500), nullable=True)
    phone = Column(String(100), nullable=True)
    email = Column(String(255), nullable=True)
    sales_owner = Column(String(255), nullable=True)
    payment_method = Column(String(100), nullable=True)
    note = Column(Text, nullable=True)
    access_key = Column(String(50), nullable=True)
    can_view_products = Column(Integer, nullable=False, default=1)
    can_export_prices = Column(Integer, nullable=False, default=1)
    can_create_quote = Column(Integer, nullable=False, default=1)
    shipping_note = Column(Text, nullable=True)
    order_note = Column(Text, nullable=True)
    closing_day = Column(String(50), nullable=True)
    payment_day = Column(String(50), nullable=True)
    closing_note = Column(Text, nullable=True)
    signed_month = Column(String(20), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    quotes = relationship("Quote", back_populates="dealer")


class SalesPerson(Base):
    __tablename__ = "sales_people"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(255), unique=True, nullable=False)
    phone = Column(String(100), nullable=True)
    email = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class Product(Base):
    __tablename__ = "products"

    id = Column(Integer, primary_key=True, index=True)
    category = Column(String(100), nullable=True)
    model = Column(String(100), unique=True, nullable=False, index=True)
    name = Column(String(255), nullable=False)
    description = Column(Text, nullable=True)
    note = Column(Text, nullable=True)
    image_url = Column(Text, nullable=True)
    image_path = Column(Text, nullable=True)
    unit = Column(String(50), nullable=True)
    status = Column(String(50), nullable=True)

    source_currency = Column(String(10), nullable=False, default="TWD")
    source_cost = Column(Float, nullable=False, default=0.0)
    shipping_usd = Column(Float, nullable=False, default=0.0)
    duty_rate_pct = Column(Float, nullable=False, default=0.0)
    duty_cost_usd = Column(Float, nullable=False, default=0.0)
    duty_cost_twd = Column(Float, nullable=False, default=0.0)
    outsourced_parts_fee_twd = Column(Float, nullable=False, default=0.0)
    final_cost_twd = Column(Float, nullable=False, default=0.0)
    planning_fee_twd = Column(Float, nullable=False, default=0.0)
    setup_fee_twd = Column(Float, nullable=False, default=0.0)

    special_market_price_ratio = Column(Float, nullable=False, default=1.0)
    special_market_min_ratio = Column(Float, nullable=False, default=1.0)
    special_designer_ratio = Column(Float, nullable=False, default=1.0)
    special_dealer_lv1_ratio = Column(Float, nullable=False, default=1.0)
    special_dealer_lv2_ratio = Column(Float, nullable=False, default=1.0)
    special_branch_ratio = Column(Float, nullable=False, default=1.0)
    special_master_ratio = Column(Float, nullable=False, default=1.0)

    market_min_discount_ratio = Column(Float, nullable=False, default=95.0)
    designer_discount_ratio = Column(Float, nullable=False, default=90.0)
    dealer_lv1_discount_ratio = Column(Float, nullable=False, default=85.0)
    dealer_lv2_discount_ratio = Column(Float, nullable=False, default=80.0)
    branch_discount_ratio = Column(Float, nullable=False, default=75.0)
    master_discount_ratio = Column(Float, nullable=False, default=70.0)

    market_price = Column(Float, nullable=False, default=0.0)
    market_min_price = Column(Float, nullable=False, default=0.0)
    designer_price = Column(Float, nullable=False, default=0.0)
    dealer_lv1_price = Column(Float, nullable=False, default=0.0)
    dealer_lv2_price = Column(Float, nullable=False, default=0.0)
    branch_price = Column(Float, nullable=False, default=0.0)
    master_dealer_price = Column(Float, nullable=False, default=0.0)

    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    quote_items = relationship("QuoteItem", back_populates="product")


class Quote(Base):
    __tablename__ = "quotes"

    id = Column(Integer, primary_key=True, index=True)
    quote_no = Column(String(50), unique=True, nullable=False)
    dealer_id = Column(Integer, ForeignKey("dealers.id"), nullable=True)

    customer_name = Column(String(255), nullable=False)
    contact_name = Column(String(255), nullable=True)
    phone = Column(String(100), nullable=True)
    email = Column(String(255), nullable=True)
    address = Column(String(500), nullable=True)
    price_level = Column(String(50), nullable=False)
    attn = Column(String(255), nullable=True)
    sales_name = Column(String(255), nullable=True)
    sales_phone = Column(String(100), nullable=True)
    sales_email = Column(String(255), nullable=True)
    currency = Column(String(20), nullable=False, default="NTD")

    note = Column(Text, nullable=True)
    product_subtotal = Column(Float, nullable=False, default=0.0)
    planning_fee_total = Column(Float, nullable=False, default=0.0)
    setup_fee_total = Column(Float, nullable=False, default=0.0)
    dispatch_fee = Column(Float, nullable=False, default=0.0)
    planning_multiplier = Column(Float, nullable=False, default=1.0)
    setup_multiplier = Column(Float, nullable=False, default=1.0)
    dispatch_label = Column(String(100), nullable=True)
    lock_install_qty = Column(Integer, nullable=False, default=0)
    lock_install_unit_price = Column(Float, nullable=False, default=3800.0)
    lock_install_fee = Column(Float, nullable=False, default=0.0)
    curtain_install_qty = Column(Float, nullable=False, default=0.0)
    curtain_install_unit = Column(String(50), nullable=True, default='式')
    curtain_install_amount = Column(Float, nullable=False, default=0.0)
    curtain_type = Column(String(50), nullable=True)
    curtain_motor_price = Column(Float, nullable=False, default=0.0)
    curtain_track_unit_price = Column(Float, nullable=False, default=0.0)
    curtain_track_length = Column(Float, nullable=False, default=0.0)
    curtain_fabric_width = Column(Float, nullable=False, default=0.0)
    curtain_fabric_height = Column(Float, nullable=False, default=0.0)
    curtain_fabric_unit_price = Column(Float, nullable=False, default=0.0)
    curtain_note = Column(Text, nullable=True)
    curtain_rows_json = Column(Text, nullable=True)
    weak_current_qty = Column(Float, nullable=False, default=0.0)
    weak_current_unit = Column(String(50), nullable=True, default='式')
    weak_current_amount = Column(Float, nullable=False, default=0.0)
    hardware_qty = Column(Float, nullable=False, default=0.0)
    hardware_unit = Column(String(50), nullable=True, default='式')
    hardware_amount = Column(Float, nullable=False, default=0.0)
    water_elec_qty = Column(Float, nullable=False, default=0.0)
    water_elec_unit = Column(String(50), nullable=True, default='式')
    water_elec_amount = Column(Float, nullable=False, default=0.0)
    custom_fee_json = Column(Text, nullable=True)
    gross_profit_rate = Column(Float, nullable=False, default=0.0)
    negotiated_total = Column(Float, nullable=False, default=0.0)
    negotiated_discount_pct = Column(Float, nullable=False, default=0.0)
    subtotal = Column(Float, nullable=False, default=0.0)
    tax_amount = Column(Float, nullable=False, default=0.0)
    total_amount = Column(Float, nullable=False, default=0.0)
    deposit_1 = Column(Float, nullable=False, default=0.0)
    deposit_2 = Column(Float, nullable=False, default=0.0)
    deposit_3 = Column(Float, nullable=False, default=0.0)
    payment_scheme = Column(String(50), nullable=True)

    quote_date = Column(Date, default=date.today, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    dealer = relationship("Dealer", back_populates="quotes")
    items = relationship("QuoteItem", back_populates="quote", cascade="all, delete-orphan")


class QuoteItem(Base):
    __tablename__ = "quote_items"

    id = Column(Integer, primary_key=True, index=True)
    quote_id = Column(Integer, ForeignKey("quotes.id"), nullable=False)
    product_id = Column(Integer, ForeignKey("products.id"), nullable=False)

    model = Column(String(100), nullable=False)
    product_name = Column(String(255), nullable=False)
    qty = Column(Integer, nullable=False, default=1)
    unit = Column(String(50), nullable=True)
    unit_price_twd = Column(Float, nullable=False, default=0.0)
    line_total_twd = Column(Float, nullable=False, default=0.0)
    cost_total_twd = Column(Float, nullable=False, default=0.0)
    planning_fee_twd = Column(Float, nullable=False, default=0.0)
    setup_fee_twd = Column(Float, nullable=False, default=0.0)
    note = Column(Text, nullable=True)
    image_path = Column(Text, nullable=True)

    quote = relationship("Quote", back_populates="items")
    product = relationship("Product", back_populates="quote_items")




class CategoryOption(Base):
    __tablename__ = "category_options"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), unique=True, nullable=False)
    is_active = Column(Integer, nullable=False, default=1)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class StatusOption(Base):
    __tablename__ = "status_options"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), unique=True, nullable=False)
    is_active = Column(Integer, nullable=False, default=1)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class ProductChangeLog(Base):
    __tablename__ = "product_change_logs"

    id = Column(Integer, primary_key=True, index=True)
    product_id = Column(Integer, ForeignKey("products.id"), nullable=False)
    changed_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    changed_by = Column(String(100), nullable=True, default='internal')
    detail = Column(Text, nullable=True)


class PriceFileArchive(Base):
    __tablename__ = "price_file_archives"

    id = Column(Integer, primary_key=True, index=True)
    version_year = Column(Integer, nullable=False)
    effective_date = Column(Date, nullable=True)
    file_name = Column(String(255), nullable=False)
    note = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class AuditLog(Base):
    __tablename__ = "audit_logs"

    id = Column(Integer, primary_key=True, index=True)
    event_type = Column(String(50), nullable=False)
    target_type = Column(String(50), nullable=True)
    target_id = Column(String(100), nullable=True)
    detail = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class SystemSetting(Base):
    __tablename__ = "system_settings"

    id = Column(Integer, primary_key=True, index=True)
    setting_key = Column(String(100), unique=True, nullable=False)
    setting_value = Column(Text, nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class InternalUser(Base):
    __tablename__ = "internal_users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(100), unique=True, nullable=False)
    password = Column(String(50), nullable=False)
    display_name = Column(String(100), nullable=True)
    role = Column(String(50), nullable=False, default="quote_only")
    is_active = Column(Integer, nullable=False, default=1)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class PasswordResetRequest(Base):
    __tablename__ = "password_reset_requests"

    id = Column(Integer, primary_key=True, index=True)
    account_type = Column(String(20), nullable=False)
    account_identifier = Column(String(100), nullable=False)
    display_name = Column(String(255), nullable=True)
    status = Column(String(20), nullable=False, default="pending")
    admin_note = Column(Text, nullable=True)
    resolved_by = Column(String(100), nullable=True)
    resolved_at = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
