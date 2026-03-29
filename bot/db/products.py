"""
Кэш данных о товарах (product_data).
"""

import os
import asyncio
import aiosqlite

from bot.db.connection import DB_PATH


async def save_product_data(store_id: int, rows: list[dict]):
    """Сохраняет снимок данных по товарам магазина."""
    async with aiosqlite.connect(DB_PATH) as db:
        for row in rows:
            await db.execute(
                '''INSERT INTO product_data
                   (store_id, nm_id, supplier_article, subject, category,
                    product_group, stock_qty, in_way_from_client, stock_qty_clean,
                    orders_7d, orders_14d, orders_30d,
                    avg_per_day, days_remaining, price_increase_pct, price)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                (store_id, row['nm_id'], row.get('supplier_article'),
                 row.get('subject'), row.get('category'),
                 row.get('product_group'), row.get('stock_qty'),
                 row.get('in_way_from_client', 0), row.get('stock_qty_clean'),
                 row.get('orders_7d'), row.get('orders_14d'), row.get('orders_30d'),
                 row.get('avg_per_day'), row.get('days_remaining'),
                 row.get('price_increase_pct'), row.get('price'))
            )
        await db.commit()


async def get_latest_product_data(store_id: int) -> tuple[list[dict], str | None]:
    """
    Возвращает последний снимок данных по магазину.

    Returns:
        (rows, fetched_at) — список товаров и время загрузки, или ([], None)
    """
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            'SELECT fetched_at FROM product_data WHERE store_id = ? '
            'ORDER BY fetched_at DESC LIMIT 1',
            (store_id,)
        )
        ts_row = await cursor.fetchone()
        if not ts_row:
            return [], None

        fetched_at = ts_row[0]

        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            'SELECT * FROM product_data WHERE store_id = ? AND fetched_at = ?',
            (store_id, fetched_at)
        )
        rows = [dict(r) for r in await cursor.fetchall()]
        return rows, fetched_at


async def is_data_fresh(store_id: int, ttl_minutes: int) -> bool:
    """Проверяет, есть ли данные свежее ttl_minutes минут."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            "SELECT 1 FROM product_data WHERE store_id = ? "
            "AND fetched_at > datetime('now', ?)"
            " LIMIT 1",
            (store_id, f'-{ttl_minutes} minutes')
        )
        return await cursor.fetchone() is not None


async def cleanup_old_product_data(days: int) -> int:
    """Удаляет данные о товарах старше days дней. Возвращает кол-во удалённых."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            "DELETE FROM product_data WHERE fetched_at < datetime('now', ?)",
            (f'-{days} days',)
        )
        await db.commit()
        return cursor.rowcount
