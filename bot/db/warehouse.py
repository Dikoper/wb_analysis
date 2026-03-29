"""
Данные по складам (warehouse_stocks).
"""

import aiosqlite

from bot.db.connection import DB_PATH


async def save_warehouse_data(store_id: int, rows: list[dict]):
    """Сохраняет снимок остатков по складам для магазина."""
    async with aiosqlite.connect(DB_PATH) as db:
        for row in rows:
            await db.execute(
                '''INSERT INTO warehouse_stocks
                   (store_id, nm_id, warehouse_name, quantity, in_way_from_client, supplier_article)
                   VALUES (?, ?, ?, ?, ?, ?)''',
                (store_id, row['nm_id'], row.get('warehouse_name', ''),
                 row.get('quantity', 0), row.get('in_way_from_client', 0),
                 row.get('supplier_article'))
            )
        await db.commit()


async def get_latest_warehouse_data(store_id: int) -> list[dict]:
    """Возвращает последний снимок остатков по складам для магазина."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            'SELECT fetched_at FROM warehouse_stocks WHERE store_id = ? '
            'ORDER BY fetched_at DESC LIMIT 1',
            (store_id,)
        )
        ts_row = await cursor.fetchone()
        if not ts_row:
            return []

        fetched_at = ts_row[0]
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            'SELECT nm_id, warehouse_name, quantity, in_way_from_client, supplier_article FROM warehouse_stocks '
            'WHERE store_id = ? AND fetched_at = ?',
            (store_id, fetched_at)
        )
        return [dict(r) for r in await cursor.fetchall()]


async def is_warehouse_data_fresh(store_id: int, ttl_minutes: int) -> bool:
    """Проверяет, есть ли данные по складам свежее ttl_minutes минут."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            "SELECT 1 FROM warehouse_stocks WHERE store_id = ? "
            "AND fetched_at > datetime('now', ?)"
            " LIMIT 1",
            (store_id, f'-{ttl_minutes} minutes')
        )
        return await cursor.fetchone() is not None


async def cleanup_old_warehouse_data(days: int) -> int:
    """Удаляет данные по складам старше days дней."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            "DELETE FROM warehouse_stocks WHERE fetched_at < datetime('now', ?)",
            (f'-{days} days',)
        )
        await db.commit()
        return cursor.rowcount
