
from typing import Dict

PRICE_LEVELS = [
    "市場報價",
    "市場最低價",
    "設計師價",
    "一級經銷商",
    "二級經銷商",
    "分公司",
    "總經銷商",
]

DEFAULT_DISCOUNT_RATIOS: Dict[str, float] = {
    "市場報價": 100.0,
    "市場最低價": 95.0,
    "設計師價": 90.0,
    "一級經銷商": 85.0,
    "二級經銷商": 80.0,
    "分公司": 75.0,
    "總經銷商": 70.0,
}

PRICE_MULTIPLIERS: Dict[str, float] = {
    "市場銷售價": 2.20,
    "市場報價": 2.20,
}

def safe_number(value) -> float:
    try:
        if value is None or value == "":
            return 0.0
        return float(value)
    except Exception:
        return 0.0

def round_int(value) -> int:
    return int(round(safe_number(value), 0))


def normalize_percentage_value(value) -> float:
    value = safe_number(value)
    if value <= 0:
        return 0.0
    return round(value * 100.0, 4) if 0 < value <= 1 else round(value, 4)

def convert_source_to_twd(source_currency: str, source_cost: float, usd_rate: float, rmb_rate: float) -> float:
    source_currency = (source_currency or "TWD").upper().strip()
    source_cost = safe_number(source_cost)
    if source_currency == "USD":
        return source_cost * safe_number(usd_rate or 0)
    if source_currency == "RMB":
        return source_cost * safe_number(rmb_rate or 0)
    return source_cost

def calculate_duty_cost_usd(source_currency: str, source_cost: float, shipping_usd: float, duty_rate_pct: float, usd_rate: float, rmb_rate: float) -> float:
    source_currency = (source_currency or 'TWD').upper().strip()
    source_cost = safe_number(source_cost)
    shipping_usd = safe_number(shipping_usd)
    duty_rate_pct = safe_number(duty_rate_pct)
    if duty_rate_pct <= 0:
        return 0.0
    if source_currency == 'USD':
        usd_base = source_cost
    elif source_currency == 'RMB':
        usd_base = (source_cost * safe_number(rmb_rate or 0) / safe_number(usd_rate or 1)) if usd_rate else 0
    else:
        usd_base = (source_cost / safe_number(usd_rate or 1)) if usd_rate else 0
    return round((usd_base + shipping_usd) * duty_rate_pct / 100.0, 2)

def calculate_final_cost_twd(source_currency: str, source_cost: float, usd_rate: float, rmb_rate: float, duty_cost_twd: float = 0.0, shipping_usd: float = 0.0, duty_rate_pct: float = 0.0, outsourced_parts_fee_twd: float = 0.0) -> float:
    base_twd = convert_source_to_twd(source_currency, source_cost, usd_rate, rmb_rate)
    shipping_twd = safe_number(shipping_usd) * safe_number(usd_rate or 0)
    if safe_number(duty_rate_pct) > 0:
        duty_cost_twd = round(calculate_duty_cost_usd(source_currency, source_cost, shipping_usd, duty_rate_pct, usd_rate, rmb_rate) * safe_number(usd_rate or 0), 0)
    else:
        duty_cost_twd = safe_number(duty_cost_twd)
    total_cost = base_twd + duty_cost_twd + shipping_twd + safe_number(outsourced_parts_fee_twd)
    return round(total_cost, 0)

def get_discount_price(market_price: float, ratio: float, special_ratio: float = 1.0) -> float:
    market_price = safe_number(market_price)
    ratio = safe_number(ratio)
    special_ratio = safe_number(special_ratio) or 1.0
    if market_price <= 0:
        return 0.0
    return round(market_price * ratio / 100.0 * special_ratio, 0)

def build_price_levels(final_cost_twd: float, specials: dict | None = None, discount_map: dict | None = None) -> dict:
    final_cost_twd = safe_number(final_cost_twd)
    specials = specials or {}
    discount_map = discount_map or DEFAULT_DISCOUNT_RATIOS
    market_price = round(final_cost_twd * PRICE_MULTIPLIERS["市場報價"] * safe_number(specials.get('special_market_price_ratio', 1) or 1), 0)
    return {
        "market_price": market_price,
        "market_min_price": get_discount_price(market_price, discount_map.get("市場最低價", 95), safe_number(specials.get('special_market_min_ratio', 1) or 1)),
        "designer_price": get_discount_price(market_price, discount_map.get("設計師價", 90), safe_number(specials.get('special_designer_ratio', 1) or 1)),
        "dealer_lv1_price": get_discount_price(market_price, discount_map.get("一級經銷商", 85), safe_number(specials.get('special_dealer_lv1_ratio', 1) or 1)),
        "dealer_lv2_price": get_discount_price(market_price, discount_map.get("二級經銷商", 80), safe_number(specials.get('special_dealer_lv2_ratio', 1) or 1)),
        "branch_price": get_discount_price(market_price, discount_map.get("分公司", 75), safe_number(specials.get('special_branch_ratio', 1) or 1)),
        "master_dealer_price": get_discount_price(market_price, discount_map.get("總經銷商", 70), safe_number(specials.get('special_master_ratio', 1) or 1)),
    }

def get_price_by_level(product, level: str, discount_map: dict | None = None) -> float:
    level = (level or "").replace("價", "").strip()
    level_map = {
        "市場報價": safe_number(getattr(product, 'market_price', 0)),
        "市場銷售價": safe_number(getattr(product, 'market_price', 0)),
        "市場最低": safe_number(getattr(product, 'market_min_price', 0)),
        "市場最低價": safe_number(getattr(product, 'market_min_price', 0)),
        "設計師": safe_number(getattr(product, 'designer_price', 0)),
        "設計師價": safe_number(getattr(product, 'designer_price', 0)),
        "一級經銷商": safe_number(getattr(product, 'dealer_lv1_price', 0)),
        "二級經銷商": safe_number(getattr(product, 'dealer_lv2_price', 0)),
        "分公司": safe_number(getattr(product, 'branch_price', 0)),
        "總經銷商": safe_number(getattr(product, 'master_dealer_price', 0)),
    }
    value = level_map.get(level, safe_number(getattr(product, 'market_price', 0)))
    if value > 0:
        return value
    specials = {
        'special_market_price_ratio': getattr(product, 'special_market_price_ratio', 1) or 1,
        'special_market_min_ratio': getattr(product, 'special_market_min_ratio', 1) or 1,
        'special_designer_ratio': getattr(product, 'special_designer_ratio', 1) or 1,
        'special_dealer_lv1_ratio': getattr(product, 'special_dealer_lv1_ratio', 1) or 1,
        'special_dealer_lv2_ratio': getattr(product, 'special_dealer_lv2_ratio', 1) or 1,
        'special_branch_ratio': getattr(product, 'special_branch_ratio', 1) or 1,
        'special_master_ratio': getattr(product, 'special_master_ratio', 1) or 1,
    }
    built = build_price_levels(getattr(product, 'final_cost_twd', 0), specials, discount_map or DEFAULT_DISCOUNT_RATIOS)
    reverse = {
        "市場報價": "market_price",
        "市場銷售價": "market_price",
        "市場最低": "market_min_price",
        "市場最低價": "market_min_price",
        "設計師": "designer_price",
        "設計師價": "designer_price",
        "一級經銷商": "dealer_lv1_price",
        "二級經銷商": "dealer_lv2_price",
        "分公司": "branch_price",
        "總經銷商": "master_dealer_price",
    }
    return safe_number(built.get(reverse.get(level, "market_price"), 0))

def get_profit_rate(cost: float, sale_price: float) -> float:
    cost = safe_number(cost)
    sale_price = safe_number(sale_price)
    if sale_price <= 0:
        return 0.0
    return round((sale_price - cost) / sale_price * 100.0, 1)
