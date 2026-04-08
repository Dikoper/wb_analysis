"""
Бизнес-логика расчётов: группировка товаров, остатки, цены.
"""

import math

import pandas as pd

from bot.config import THRESHOLD_A, THRESHOLD_B, THRESHOLD_C


def assign_group(avg_per_day: float, threshold_a: float = THRESHOLD_A, threshold_b: float = THRESHOLD_B, threshold_c: float = THRESHOLD_C) -> str:
    """
    Определяет группу товара по средним продажам в день (рассчитанным за 14д).

    A: ≥threshold_a шт/день (ходовые)
    B: ≥threshold_b шт/день (средние)
    C: ≥threshold_c шт/день (редкие)
    D: <threshold_c шт/день (почти не продаются, реже 1 шт/неделю)
    """
    if avg_per_day >= threshold_a:
        return 'A'
    elif avg_per_day >= threshold_b:
        return 'B'
    elif avg_per_day >= threshold_c:
        return 'C'
    else:
        return 'D'


def calc_avg_by_group(row) -> float:
    """
    Возвращает среднее продаж в день за 14 дней (единый период для всех групп).
    """
    return row['orders_count_14d'] / 14 if pd.notna(row['orders_count_14d']) else 0


def calc_days_remaining(row) -> float:
    """
    На сколько дней хватит остатка.

    Returns:
        None если нет продаж, иначе кол-во дней
    """
    if row['avg_per_day'] == 0:
        return None
    return row['stock_qty'] / row['avg_per_day']


def get_price_increase(days_remaining: float, days_threshold: int = 7) -> int:
    """
    Возвращает % повышения цены по шкале.

    Шкала:
    - > days_threshold дней: 0%
    - 6-7 дней: 5%
    - 5-6 дней: 10%
    - 4-5 дней: 15%
    - 3-4 дней: 25%
    - 2-3 дней: 30%
    - 1-2 дней: 40%
    - < 1 дня: 50%
    """
    if days_remaining is None or days_remaining > days_threshold:
        return 0
    elif days_remaining >= 6:
        return 5
    elif days_remaining >= 5:
        return 10
    elif days_remaining >= 4:
        return 15
    elif days_remaining >= 3:
        return 25
    elif days_remaining >= 2:
        return 30
    elif days_remaining >= 1:
        return 40
    else:
        return 50


def calc_refill_qty(avg_per_day: float, days: int, reserve_pct: float) -> int:
    """
    Объём пополнения на N дней с процентным запасом.

    refill = ceil(avg_per_day * days * (1 + reserve_pct/100))

    Для avg_per_day <= 0 возвращает 0.
    """
    if not avg_per_day or avg_per_day <= 0:
        return 0
    return math.ceil(avg_per_day * days * (1 + reserve_pct / 100))


# ── Хелперы для блока «Предложение WB» ────────────────────────────────────────

AVAILABILITY_RU = {
    'deficient': 'дефицитный',
    'balanced':  'сбалансированный',
    'actual':    'стабильный',
    'nonActual': 'слабый',
    'nonLiquid': 'неликвид',
}

# Мультипликатор объёма по классификации WB. Все пороги собраны здесь, чтобы
# их можно было крутить в одном месте.
# Главный сигнал — тренд (clamp [-40%..+50%]), затем availability, затем
# простой склада (минимальный вклад: простой может быть от поставщика, а не
# от спроса). Максимальная комбинация (deficient + high miss + strong uptrend)
# — около ~2× базовой; падающий товар (nonActual + trend −40%) → ~0.4× базы.
AVAILABILITY_MULT = {
    'nonLiquid': 0.0,   # не поставлять
    'nonActual': 0.7,
    'balanced':  1.0,
    'actual':    1.05,
    'deficient': 1.25,
}


def availability_ru(val: str) -> str:
    """RU-ярлык классификации WB. Неизвестные/пустые → '—'."""
    return AVAILABILITY_RU.get((val or '').strip(), '—')


def trend_arrow(pct: float) -> str:
    """
    Стрелка тренда по проценту изменения.

    |pct| < 5% → '→ ±X%' (стабильно)
    pct > 0   → '↑ +X%'
    pct < 0   → '↓ -X%'
    """
    if pct is None:
        return '—'
    if abs(pct) < 5:
        return f'→ ±{abs(pct):.0f}%'
    arrow = '↑' if pct > 0 else '↓'
    return f'{arrow} {pct:+.0f}%'


def calc_smart_refill_qty(
    avg: float,
    target_days: int,
    reserve_pct: float,
    availability: str = '',
    miss_days: float = 0.0,
    trend_pct: float = 0.0,
) -> int:
    """
    Умный расчёт объёма поставки с учётом WB-метрик.

    Формула: ceil(base × f_avail × f_miss × f_trend)
        base   = avg * target_days * (1 + reserve_pct/100)
        f_avail — мультипликатор по availability (см. AVAILABILITY_MULT)
        f_miss  — мультипликатор по простою склада (officeMissingTime.days)
        f_trend — линейная корректировка по тренду продаж, clamp [-40%..+50%]
    """
    if not avg or avg <= 0:
        return 0

    f_avail = AVAILABILITY_MULT.get((availability or '').strip(), 1.0)
    if f_avail == 0.0:
        return 0  # неликвид — не поставлять

    if miss_days > 10:
        f_miss = 1.05
    elif miss_days > 5:
        f_miss = 1.03
    else:
        f_miss = 1.0

    t = max(-40.0, min(50.0, trend_pct or 0.0)) / 100.0
    f_trend = 1.0 + t

    base = avg * target_days * (1.0 + reserve_pct / 100.0)
    return int(math.ceil(base * f_avail * f_miss * f_trend))


def aggregate_by_article(rows: list[dict]) -> dict:
    """
    Группировка товаров по supplier_article с агрегацией дублей.

    Returns:
        dict {supplier_article: aggregated_row}
    """
    result = {}
    for r in rows:
        art = r.get('supplier_article')
        if not art:
            continue
        if art not in result:
            result[art] = r.copy()
        else:
            existing = result[art]
            existing['stock_qty'] = existing.get('stock_qty', 0) + r.get('stock_qty', 0)
            existing['stock_qty_clean'] = existing.get('stock_qty_clean', 0) + r.get('stock_qty_clean', 0)
            existing['in_way_from_client'] = existing.get('in_way_from_client', 0) + r.get('in_way_from_client', 0)
            existing['avg_per_day'] = (existing.get('avg_per_day') or 0) + (r.get('avg_per_day') or 0)
            if (r.get('avg_per_day') or 0) > (existing.get('_max_avg') or 0):
                existing['_max_avg'] = r.get('avg_per_day') or 0
                existing['nm_id'] = r.get('nm_id')
                existing['price'] = r.get('price')

    for item in result.values():
        avg = item.get('avg_per_day') or 0
        item['days_remaining'] = (item['stock_qty'] / avg) if avg > 0 else None
        item['product_group'] = assign_group(avg)
        item.pop('_max_avg', None)
    return result


def merge_wh_by_name(wh_list):
    """Объединяет записи складов с одинаковым warehouse_name."""
    merged = {}
    for w in wh_list:
        name = w['warehouse_name']
        if name not in merged:
            merged[name] = {'warehouse_name': name, 'quantity': 0, 'in_way_from_client': 0}
        merged[name]['quantity'] += w['quantity']
        merged[name]['in_way_from_client'] += w['in_way_from_client']
    return list(merged.values())


def wh_compact_str(wh_list: list[dict]) -> str:
    """
    Компактная строка детализации по складам: 'Склад: кол-во | Склад2: кол-во | -N возвр.'.

    Принимает сырой список складов (будет merged внутри).
    Возвращает "—" если нет активных складов.
    """
    if not wh_list:
        return "—"
    wh_list = merge_wh_by_name(wh_list)
    total_iwfc = sum(w.get('in_way_from_client', 0) for w in wh_list)
    wh_active = [w for w in wh_list if w['quantity'] > 0]
    wh_active.sort(key=lambda w: w['quantity'], reverse=True)
    if not wh_active:
        return "—"
    parts = [f"{w['warehouse_name']}: {w['quantity']}" for w in wh_active]
    if total_iwfc > 0:
        parts.append(f"-{total_iwfc} возвр.")
    return " | ".join(parts)


# ── Трансформации данных (перенесено из wb_api.py) ────────────────────────────

def calc_avg_per_day(df: pd.DataFrame, days: int = 7) -> pd.DataFrame:
    """
    Добавляет колонку среднего заказов в день за период.

    Args:
        df: DataFrame с колонкой orders_count_{days}d
        days: период в днях (по умолчанию 7)

    Returns:
        DataFrame с добавленной колонкой avg_per_day_{days}d
    """
    df = df.copy()
    orders_col = f'orders_count_{days}d'
    avg_col = f'avg_per_day_{days}d'
    df[avg_col] = df[orders_col] / days
    return df


def merge_orders_stocks(orders: pd.DataFrame, stocks: pd.DataFrame) -> pd.DataFrame:
    """
    Объединяет остатки и заказы через OUTER JOIN.

    Stocks — основа каталога (содержит метаданные товара).
    Товары без заказов за 14д получают orders_count = 0.
    Товары с заказами, но без остатков — stock_qty = 0.

    Args:
        orders: DataFrame с заказами (nmId, orders_count_7d, orders_count_14d, ...)
        stocks: DataFrame с остатками (nmId, stock_qty, supplierArticle, subject, category, ...)

    Returns:
        DataFrame объединённый, отсортированный по stock_qty (убывание)
    """
    # OUTER JOIN: сохраняем ВСЕ товары из обоих источников
    df = stocks.merge(orders, on='nmId', how='outer', suffixes=('', '_orders'))

    # Заполняем метаданные: приоритет stocks, fallback на orders
    for col in ['supplierArticle', 'subject', 'category']:
        orders_col = f'{col}_orders'
        if orders_col in df.columns:
            df[col] = df[col].fillna(df[orders_col])
            df.drop(columns=[orders_col], inplace=True)

    # Заполняем пропуски числовых полей
    df['stock_qty'] = df['stock_qty'].fillna(0).astype(int)
    df['orders_count_7d'] = df['orders_count_7d'].fillna(0).astype(int)
    df['orders_count_14d'] = df['orders_count_14d'].fillna(0).astype(int)

    df = df.sort_values('stock_qty', ascending=False).reset_index(drop=True)
    return df
