"""
Загрузка данных из WB API и кэширование через SQLite.

Единственная точка входа для получения данных товаров и складов.
"""

import asyncio
import logging

import pandas as pd

from wb_api import get_orders, get_stocks, get_stocks_detailed, get_prices, merge_orders_stocks, calc_avg_per_day
from bot.config import THRESHOLD_A, THRESHOLD_B, DATA_CACHE_TTL
from bot.calculations import assign_group, calc_avg_by_group, calc_days_remaining, get_price_increase
from bot.db import (
    is_data_fresh, get_latest_product_data, save_product_data,
    is_warehouse_data_fresh, get_latest_warehouse_data, save_warehouse_data,
)

logger = logging.getLogger(__name__)


# ── Загрузка данных из API ────────────────────────────────────────────────────

def fetch_store_data(
    token: str = None,
    days_threshold: int = 7,
    threshold_a: float = THRESHOLD_A,
    threshold_b: float = THRESHOLD_B,
) -> list[dict]:
    """
    Загружает данные из WB API, рассчитывает метрики.

    Returns:
        Список словарей — по одному на каждый товар (nm_id).
    """
    logger.info("=== Загрузка данных из WB API ===")

    # === 1. Загрузка данных за 30 дней ===
    try:
        logger.info("Загрузка заказов за 30 дней...")
        orders_30d = get_orders(30, token=token)
        logger.info(f"✓ Заказы 30д: {len(orders_30d)} записей")
    except Exception as e:
        logger.error(f"✗ Ошибка загрузки заказов 30д: {e}")
        raise

    try:
        logger.info("Загрузка остатков...")
        stocks = get_stocks(orders_30d['nmId'].tolist(), token=token)
        logger.info(f"✓ Остатки: {len(stocks)} записей")
    except Exception as e:
        logger.error(f"✗ Ошибка загрузки остатков: {e}")
        raise

    # Объединяем заказы и остатки
    df = merge_orders_stocks(orders_30d, stocks)
    df = calc_avg_per_day(df, days=30)

    # Заполняем поля возвратов если отсутствуют после merge
    for col in ['in_way_from_client', 'stock_qty_clean']:
        if col not in df.columns:
            df[col] = 0
        df[col] = df[col].fillna(0).astype(int)

    # Используем чистый остаток (без товаров в возврате) для расчётов
    df['stock_qty_original'] = df['stock_qty']
    df['stock_qty'] = df['stock_qty_clean']

    logger.info(f"Объединено товаров: {len(df)}")

    # === 2. Группировка товаров (A/B/C) ===
    df['group'] = df['avg_per_day_30d'].apply(lambda x: assign_group(x, threshold_a, threshold_b))

    # === 3. Загрузка данных за 7 и 14 дней ===
    try:
        logger.info("Загрузка заказов за 7 дней...")
        orders_7d = get_orders(7, token=token)
        orders_7d = orders_7d[['nmId', 'orders_count_7d']]
        logger.info(f"✓ Заказы 7д: {len(orders_7d)} записей")
    except Exception as e:
        logger.error(f"✗ Ошибка загрузки заказов 7д: {e}")
        raise

    try:
        logger.info("Загрузка заказов за 14 дней...")
        orders_14d = get_orders(14, token=token)
        orders_14d = orders_14d[['nmId', 'orders_count_14d']]
        logger.info(f"✓ Заказы 14д: {len(orders_14d)} записей")
    except Exception as e:
        logger.error(f"✗ Ошибка загрузки заказов 14д: {e}")
        raise

    # Присоединяем к основной таблице
    df = df.merge(orders_7d, on='nmId', how='left')
    df = df.merge(orders_14d, on='nmId', how='left')

    # === 4. Расчёт среднего по группе ===
    df['avg_per_day'] = df.apply(calc_avg_by_group, axis=1)

    # === 5. Расчёт days_remaining ===
    df['days_remaining'] = df.apply(calc_days_remaining, axis=1)

    # === 6. Расчёт % повышения цены ===
    df['price_increase_pct'] = df['days_remaining'].apply(lambda d: get_price_increase(d, days_threshold))

    # === 7. Загрузка цен ===
    try:
        logger.info("Загрузка цен...")
        prices_map = get_prices(df['nmId'].tolist(), token=token)
        df['price'] = df['nmId'].map(prices_map)
        logger.info(f"✓ Цены загружены: {len(prices_map)} из {len(df)} товаров")
    except Exception as e:
        logger.warning(f"Не удалось загрузить цены: {e}")
        df['price'] = None

    logger.info("=== Данные загружены и рассчитаны ===")

    # Конвертируем в list[dict] для сохранения в БД
    result = []
    for _, row in df.iterrows():
        result.append({
            'nm_id': int(row['nmId']),
            'supplier_article': row.get('supplierArticle'),
            'subject': row.get('subject'),
            'category': row.get('category'),
            'product_group': row['group'],
            'stock_qty': int(row['stock_qty']),
            'in_way_from_client': int(row['in_way_from_client']),
            'stock_qty_clean': int(row['stock_qty_clean']),
            'orders_7d': int(row['orders_count_7d']) if pd.notna(row.get('orders_count_7d')) else None,
            'orders_14d': int(row['orders_count_14d']) if pd.notna(row.get('orders_count_14d')) else None,
            'orders_30d': int(row['orders_count_30d']),
            'avg_per_day': round(row['avg_per_day'], 4),
            'days_remaining': round(row['days_remaining'], 2) if row['days_remaining'] is not None else None,
            'price_increase_pct': int(row['price_increase_pct']),
            'price': float(row['price']) if pd.notna(row.get('price')) else None,
        })

    return result


def fetch_warehouse_data(token: str, nm_ids: list = None) -> list[dict]:
    """
    Загружает детализацию остатков по складам из WB API.

    Returns:
        Список словарей {nm_id, warehouse_name, quantity}
    """
    df = get_stocks_detailed(nm_ids=nm_ids, token=token)
    if df.empty:
        return []
    return [
        {
            'nm_id': int(row['nmId']),
            'warehouse_name': row['warehouseName'],
            'quantity': int(row['quantity']),
            'in_way_from_client': int(row.get('inWayFromClient', 0)),
            'supplier_article': row.get('supplierArticle') or None,
        }
        for _, row in df.iterrows()
    ]


# ── Кэширование ──────────────────────────────────────────────────────────────

async def fetch_or_cache_product(store_id, token, days_threshold, threshold_a, threshold_b):
    """Загружает данные товаров из кэша или API."""
    if await is_data_fresh(store_id, DATA_CACHE_TTL):
        rows, _ = await get_latest_product_data(store_id)
        return rows
    rows = await asyncio.to_thread(
        fetch_store_data, token=token,
        days_threshold=days_threshold,
        threshold_a=threshold_a,
        threshold_b=threshold_b,
    )
    await save_product_data(store_id, rows)
    return rows


async def fetch_or_cache_warehouse(store_id, token):
    """Загружает данные по складам из кэша или API."""
    if await is_warehouse_data_fresh(store_id, DATA_CACHE_TTL):
        wh_data = await get_latest_warehouse_data(store_id)
        if wh_data:
            return wh_data
    wh_data = await asyncio.to_thread(fetch_warehouse_data, token=token)
    await save_warehouse_data(store_id, wh_data)
    return wh_data
