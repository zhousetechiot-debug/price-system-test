from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import List
import re
import shutil

from fastapi import FastAPI, Request, Depends, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import or_, text
from sqlalchemy.orm import Session

from .database import Base, engine, get_db
from . import models
from .services.excel_importer import (
    ensure_default_rates,
    ensure_default_price_settings,
    get_rate_map,
    import_all,
    get_price_setting_map,
    apply_discount_settings_to_products,
)
from .services.pricing import get_price_by_level, get_profit_rate, calculate_final_cost_twd, build_price_levels, PRICE_LEVELS
from .services.quote_exporter import export_quote_to_excel, export_quote_to_pdf, DEFAULT_NOTE

APP_NAME = "智崴物聯價格系統"

app = FastAPI(title=APP_NAME)
Base.metadata.create_all(bind=engine)

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"
UPLOAD_DIR = BASE_DIR.parent / "uploads"
EXPORT_DIR = BASE_DIR.parent / "exports"
IMAGE_DIR = UPLOAD_DIR / "images"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)
IMAGE_DIR.mkdir(parents=True, exist_ok=True)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
app.mount("/uploads", StaticFiles(directory=str(UPLOAD_DIR)), name="uploads")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


def safe_filename(text_value: str) -> str:
    value = (text_value or "").strip()
    value = re.sub(r'[\\/:*?"<>|]+', '_', value)
    return value or "未填"


def quote_output_name(quote) -> str:
    return f"{quote.quote_date.strftime('%Y%m%d')} - {safe_filename(quote.customer_name)} - {safe_filename(quote.contact_name)} -報價單"


def get_image_src(product):
    if product.image_path:
        return product.image_path
    if product.image_url:
        return product.image_url
    return ""


def get_default_price_map(product):
    return {
        "市場報價": round(product.market_price or 0),
        "市場最低價": round(product.market_min_price or 0),
        "設計師價": round(product.designer_price or 0),
        "一級經銷商": round(product.dealer_lv1_price or 0),
        "二級經銷商": round(product.dealer_lv2_price or 0),
        "分公司": round(product.branch_price or 0),
        "總經銷商": round(product.master_dealer_price or 0),
    }


def ensure_sqlite_columns():
    statements = [
        "ALTER TABLE products ADD COLUMN image_path TEXT",
        "ALTER TABLE quotes ADD COLUMN attn VARCHAR(255)",
        "ALTER TABLE quotes ADD COLUMN currency VARCHAR(20) DEFAULT 'NTD'",
    ]
    with engine.begin() as conn:
        for sql in statements:
            try:
                conn.execute(text(sql))
            except Exception:
                pass


@app.on_event("startup")
def startup_seed():
    ensure_sqlite_columns()
    db = next(get_db())
    ensure_default_rates(db)
    ensure_default_price_settings(db)


@app.get("/", response_class=HTMLResponse)
def home(
    request: Request,
    q: str = "",
    category: str = "",
    level: str = "市場報價",
    show_cost: int = 0,
    db: Session = Depends(get_db),
):
    products_query = db.query(models.Product)

    if q:
        keyword = f"%{q}%"
        products_query = products_query.filter(
            or_(
                models.Product.model.like(keyword),
                models.Product.name.like(keyword),
                models.Product.category.like(keyword),
            )
        )

    if category:
        products_query = products_query.filter(models.Product.category == category)

    products = products_query.order_by(models.Product.category, models.Product.model).all()
    categories = [x[0] for x in db.query(models.Product.category).distinct().all() if x[0]]
    discount_map = get_price_setting_map(db)

    return templates.TemplateResponse(
        "home.html",
        {
            "request": request,
            "products": products,
            "categories": categories,
            "selected_category": category,
            "keyword": q,
            "selected_level": level,
            "show_cost": bool(show_cost),
            "discount_map": discount_map,
            "get_price_by_level": get_price_by_level,
            "get_profit_rate": get_profit_rate,
            "get_image_src": get_image_src,
        },
    )


@app.get("/products", response_class=HTMLResponse)
def products_page(request: Request, db: Session = Depends(get_db)):
    products = db.query(models.Product).order_by(models.Product.category, models.Product.model).all()
    return templates.TemplateResponse(
        "products.html",
        {"request": request, "products": products, "get_image_src": get_image_src},
    )


@app.get("/admin/import", response_class=HTMLResponse)
def admin_import_page(request: Request, db: Session = Depends(get_db)):
    rates = get_rate_map(db)
    return templates.TemplateResponse(
        "admin_import.html",
        {"request": request, "rates": rates, "message": None},
    )


@app.post("/admin/import", response_class=HTMLResponse)
async def admin_import_action(
    request: Request,
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
):
    save_path = UPLOAD_DIR / file.filename
    content = await file.read()
    save_path.write_bytes(content)

    result = import_all(db, str(save_path))
    rates = get_rate_map(db)
    message = f"匯入完成：商品 {result['products']} 筆，經銷商 {result['dealers']} 筆。"
    return templates.TemplateResponse(
        "admin_import.html",
        {"request": request, "rates": rates, "message": message},
    )


@app.get("/admin/rates", response_class=HTMLResponse)
def admin_rates_page(request: Request, db: Session = Depends(get_db)):
    rates = get_rate_map(db)
    discount_map = get_price_setting_map(db)
    return templates.TemplateResponse(
        "admin_rates.html",
        {"request": request, "rates": rates, "discount_map": discount_map, "message": None},
    )


@app.post("/admin/rates", response_class=HTMLResponse)
def admin_rates_save(
    request: Request,
    usd_rate: float = Form(...),
    rmb_rate: float = Form(...),
    market_min_discount: float = Form(...),
    designer_discount: float = Form(...),
    dealer_lv1_discount: float = Form(...),
    dealer_lv2_discount: float = Form(...),
    branch_discount: float = Form(...),
    master_discount: float = Form(...),
    apply_to_products: str | None = Form(None),
    db: Session = Depends(get_db),
):
    ensure_default_rates(db, usd_rate=usd_rate, rmb_rate=rmb_rate)

    mapping = {
        "市場最低價": market_min_discount,
        "設計師價": designer_discount,
        "一級經銷商": dealer_lv1_discount,
        "二級經銷商": dealer_lv2_discount,
        "分公司": branch_discount,
        "總經銷商": master_discount,
        "市場報價": 100.0,
    }
    for level, ratio in mapping.items():
        row = db.query(models.PriceSetting).filter(models.PriceSetting.level_name == level).first()
        if row:
            row.discount_ratio = ratio
            row.updated_at = datetime.utcnow()
        else:
            db.add(models.PriceSetting(level_name=level, discount_ratio=ratio))
    db.commit()

    if apply_to_products:
        apply_discount_settings_to_products(db)
        message = "匯率與折扣設定已更新，並已重算全部商品售價。"
    else:
        message = "匯率與折扣設定已更新。"

    return templates.TemplateResponse(
        "admin_rates.html",
        {"request": request, "rates": get_rate_map(db), "discount_map": get_price_setting_map(db), "message": message},
    )


@app.get("/admin/products/new", response_class=HTMLResponse)
def admin_product_new(request: Request, db: Session = Depends(get_db)):
    return templates.TemplateResponse(
        "product_form.html",
        {"request": request, "product": None, "price_levels": PRICE_LEVELS, "action": "/admin/products/new", "title_text": "新增商品"},
    )


@app.post("/admin/products/new")
async def admin_product_create(
    category: str = Form(""),
    model: str = Form(...),
    name: str = Form(...),
    description: str = Form(""),
    note: str = Form(""),
    unit: str = Form("台"),
    source_currency: str = Form("TWD"),
    source_cost: float = Form(0),
    duty_cost_twd: float = Form(0),
    market_price: float = Form(0),
    market_min_price: float = Form(0),
    designer_price: float = Form(0),
    dealer_lv1_price: float = Form(0),
    dealer_lv2_price: float = Form(0),
    branch_price: float = Form(0),
    master_dealer_price: float = Form(0),
    image_url: str = Form(""),
    image_file: UploadFile | None = File(None),
    db: Session = Depends(get_db),
):
    if db.query(models.Product).filter(models.Product.model == model).first():
        raise HTTPException(status_code=400, detail="型號已存在")
    rates = get_rate_map(db)
    final_cost_twd = calculate_final_cost_twd(source_currency, source_cost, rates["USD"], rates["RMB"], duty_cost_twd)
    if market_price <= 0:
        built = build_price_levels(final_cost_twd)
        market_price = built["market_price"]
        market_min_price = market_min_price or built["market_min_price"]
        designer_price = designer_price or built["designer_price"]
        dealer_lv1_price = dealer_lv1_price or built["dealer_lv1_price"]
        dealer_lv2_price = dealer_lv2_price or built["dealer_lv2_price"]
        branch_price = branch_price or built["branch_price"]
        master_dealer_price = master_dealer_price or built["master_dealer_price"]

    product = models.Product(
        category=category,
        model=model,
        name=name,
        description=description,
        note=note,
        unit=unit,
        source_currency=source_currency,
        source_cost=source_cost,
        duty_cost_twd=duty_cost_twd,
        final_cost_twd=final_cost_twd,
        market_price=market_price,
        market_min_price=market_min_price,
        designer_price=designer_price,
        dealer_lv1_price=dealer_lv1_price,
        dealer_lv2_price=dealer_lv2_price,
        branch_price=branch_price,
        master_dealer_price=master_dealer_price,
        image_url=image_url,
        updated_at=datetime.utcnow(),
    )
    db.add(product)
    db.commit()
    db.refresh(product)

    if image_file and image_file.filename:
        ext = Path(image_file.filename).suffix.lower() or ".jpg"
        file_path = IMAGE_DIR / f"product_{product.id}{ext}"
        with file_path.open("wb") as f:
            shutil.copyfileobj(image_file.file, f)
        product.image_path = f"/uploads/images/{file_path.name}"
        db.commit()

    return RedirectResponse(url="/products", status_code=303)


@app.get("/admin/products/{product_id}/edit", response_class=HTMLResponse)
def admin_product_edit(product_id: int, request: Request, db: Session = Depends(get_db)):
    product = db.query(models.Product).filter(models.Product.id == product_id).first()
    if not product:
        raise HTTPException(status_code=404, detail="找不到商品")
    return templates.TemplateResponse(
        "product_form.html",
        {"request": request, "product": product, "price_levels": PRICE_LEVELS, "action": f"/admin/products/{product_id}/edit", "title_text": "修改商品", "get_image_src": get_image_src},
    )


@app.post("/admin/products/{product_id}/edit")
async def admin_product_edit_post(
    product_id: int,
    category: str = Form(""),
    model: str = Form(...),
    name: str = Form(...),
    description: str = Form(""),
    note: str = Form(""),
    unit: str = Form("台"),
    source_currency: str = Form("TWD"),
    source_cost: float = Form(0),
    duty_cost_twd: float = Form(0),
    market_price: float = Form(0),
    market_min_price: float = Form(0),
    designer_price: float = Form(0),
    dealer_lv1_price: float = Form(0),
    dealer_lv2_price: float = Form(0),
    branch_price: float = Form(0),
    master_dealer_price: float = Form(0),
    image_url: str = Form(""),
    image_file: UploadFile | None = File(None),
    db: Session = Depends(get_db),
):
    product = db.query(models.Product).filter(models.Product.id == product_id).first()
    if not product:
        raise HTTPException(status_code=404, detail="找不到商品")
    rates = get_rate_map(db)
    final_cost_twd = calculate_final_cost_twd(source_currency, source_cost, rates["USD"], rates["RMB"], duty_cost_twd)

    product.category = category
    product.model = model
    product.name = name
    product.description = description
    product.note = note
    product.unit = unit
    product.source_currency = source_currency
    product.source_cost = source_cost
    product.duty_cost_twd = duty_cost_twd
    product.final_cost_twd = final_cost_twd
    product.market_price = market_price
    product.market_min_price = market_min_price
    product.designer_price = designer_price
    product.dealer_lv1_price = dealer_lv1_price
    product.dealer_lv2_price = dealer_lv2_price
    product.branch_price = branch_price
    product.master_dealer_price = master_dealer_price
    product.image_url = image_url
    product.updated_at = datetime.utcnow()

    if image_file and image_file.filename:
        ext = Path(image_file.filename).suffix.lower() or ".jpg"
        file_path = IMAGE_DIR / f"product_{product.id}{ext}"
        with file_path.open("wb") as f:
            shutil.copyfileobj(image_file.file, f)
        product.image_path = f"/uploads/images/{file_path.name}"

    db.commit()
    return RedirectResponse(url="/products", status_code=303)


@app.get("/quotes/new", response_class=HTMLResponse)
def quote_new_page(request: Request, db: Session = Depends(get_db)):
    dealers = db.query(models.Dealer).order_by(models.Dealer.name).all()
    products = db.query(models.Product).order_by(models.Product.category, models.Product.model).all()
    return templates.TemplateResponse(
        "quote_form.html",
        {
            "request": request,
            "dealers": dealers,
            "products": products,
            "quote": None,
            "qty_map": {},
            "note_text": DEFAULT_NOTE,
            "price_levels": PRICE_LEVELS,
        },
    )


def create_or_update_quote(
    db: Session,
    quote: models.Quote | None,
    dealer_id: int | None,
    customer_name: str,
    contact_name: str,
    phone: str,
    email: str,
    address: str,
    attn: str,
    quote_date: str,
    quote_no: str,
    note: str,
    product_ids: List[int],
    qtys: List[int],
):
    dealer = db.query(models.Dealer).filter(models.Dealer.id == dealer_id).first() if dealer_id else None
    price_level = dealer.level if dealer and dealer.level else "市場報價"
    parsed_date = datetime.strptime(quote_date, "%Y-%m-%d").date() if quote_date else date.today()

    if not quote:
        if not quote_no:
            quote_no = f"ZQ{parsed_date.strftime('%Y%m%d')}{db.query(models.Quote).count() + 1:03d}"
        quote = models.Quote(
            quote_no=quote_no,
            dealer_id=dealer.id if dealer else None,
            customer_name=customer_name,
            contact_name=contact_name,
            phone=phone,
            email=email,
            address=address,
            price_level=price_level,
            note=note,
            attn=attn,
            quote_date=parsed_date,
            currency="NTD",
        )
        db.add(quote)
        db.flush()
    else:
        quote.dealer_id = dealer.id if dealer else None
        quote.customer_name = customer_name
        quote.contact_name = contact_name
        quote.phone = phone
        quote.email = email
        quote.address = address
        quote.attn = attn
        quote.price_level = price_level
        quote.quote_date = parsed_date
        quote.note = note
        if quote_no:
            quote.quote_no = quote_no
        db.query(models.QuoteItem).filter(models.QuoteItem.quote_id == quote.id).delete()
        db.flush()

    subtotal = 0.0
    discount_map = get_price_setting_map(db)
    for product_id, qty in zip(product_ids, qtys):
        if int(qty) <= 0:
            continue
        product = db.query(models.Product).filter(models.Product.id == product_id).first()
        if not product:
            continue
        unit_price = get_price_by_level(product, price_level, discount_map)
        line_total = unit_price * int(qty)
        subtotal += line_total

        item = models.QuoteItem(
            quote_id=quote.id,
            product_id=product.id,
            model=product.model,
            product_name=product.name,
            qty=int(qty),
            unit=product.unit or "台",
            unit_price_twd=unit_price,
            line_total_twd=line_total,
        )
        db.add(item)

    tax_amount = round(subtotal * 0.05, 0)
    total_amount = round(subtotal + tax_amount, 0)
    quote.subtotal = subtotal
    quote.tax_amount = tax_amount
    quote.total_amount = total_amount
    db.commit()
    db.refresh(quote)
    return quote


@app.post("/quotes/new")
def quote_create(
    dealer_id: int | None = Form(None),
    customer_name: str = Form(...),
    contact_name: str = Form(""),
    phone: str = Form(""),
    email: str = Form(""),
    address: str = Form(""),
    attn: str = Form(""),
    quote_date: str = Form(""),
    quote_no: str = Form(""),
    note: str = Form(DEFAULT_NOTE),
    product_ids: List[int] = Form(...),
    qtys: List[int] = Form(...),
    db: Session = Depends(get_db),
):
    quote = create_or_update_quote(db, None, dealer_id, customer_name, contact_name, phone, email, address, attn, quote_date, quote_no, note, product_ids, qtys)
    return RedirectResponse(url=f"/quotes/{quote.id}", status_code=303)


@app.get("/quotes/{quote_id}", response_class=HTMLResponse)
def quote_detail(quote_id: int, request: Request, db: Session = Depends(get_db)):
    quote = db.query(models.Quote).filter(models.Quote.id == quote_id).first()
    if not quote:
        raise HTTPException(status_code=404, detail="找不到報價單")
    return templates.TemplateResponse(
        "quote_detail.html",
        {"request": request, "quote": quote, "output_name": quote_output_name(quote)},
    )


@app.get("/quotes/{quote_id}/edit", response_class=HTMLResponse)
def quote_edit_page(quote_id: int, request: Request, db: Session = Depends(get_db)):
    quote = db.query(models.Quote).filter(models.Quote.id == quote_id).first()
    if not quote:
        raise HTTPException(status_code=404, detail="找不到報價單")
    dealers = db.query(models.Dealer).order_by(models.Dealer.name).all()
    products = db.query(models.Product).order_by(models.Product.category, models.Product.model).all()
    qty_map = {item.product_id: item.qty for item in quote.items}
    return templates.TemplateResponse(
        "quote_form.html",
        {
            "request": request,
            "dealers": dealers,
            "products": products,
            "quote": quote,
            "qty_map": qty_map,
            "note_text": quote.note or DEFAULT_NOTE,
            "price_levels": PRICE_LEVELS,
        },
    )


@app.post("/quotes/{quote_id}/edit")
def quote_edit(
    quote_id: int,
    dealer_id: int | None = Form(None),
    customer_name: str = Form(...),
    contact_name: str = Form(""),
    phone: str = Form(""),
    email: str = Form(""),
    address: str = Form(""),
    attn: str = Form(""),
    quote_date: str = Form(""),
    quote_no: str = Form(""),
    note: str = Form(DEFAULT_NOTE),
    product_ids: List[int] = Form(...),
    qtys: List[int] = Form(...),
    db: Session = Depends(get_db),
):
    quote = db.query(models.Quote).filter(models.Quote.id == quote_id).first()
    if not quote:
        raise HTTPException(status_code=404, detail="找不到報價單")
    quote = create_or_update_quote(db, quote, dealer_id, customer_name, contact_name, phone, email, address, attn, quote_date, quote_no, note, product_ids, qtys)
    return RedirectResponse(url=f"/quotes/{quote.id}", status_code=303)


@app.get("/quotes/{quote_id}/excel")
def quote_excel_export(quote_id: int, db: Session = Depends(get_db)):
    quote = db.query(models.Quote).filter(models.Quote.id == quote_id).first()
    if not quote:
        raise HTTPException(status_code=404, detail="找不到報價單")

    output_name = quote_output_name(quote)
    output_path = EXPORT_DIR / f"{output_name}.xlsx"
    export_quote_to_excel(quote, str(output_path))
    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_path.name,
    )


@app.get("/quotes/{quote_id}/pdf")
def quote_pdf_export(quote_id: int, db: Session = Depends(get_db)):
    quote = db.query(models.Quote).filter(models.Quote.id == quote_id).first()
    if not quote:
        raise HTTPException(status_code=404, detail="找不到報價單")

    output_name = quote_output_name(quote)
    output_path = EXPORT_DIR / f"{output_name}.pdf"
    export_quote_to_pdf(quote, str(output_path))
    return FileResponse(path=str(output_path), media_type="application/pdf", filename=output_path.name)
