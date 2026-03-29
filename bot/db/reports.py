"""
История отчётов (report_history).
"""

import os
import asyncio
import aiosqlite

from bot.db.connection import DB_PATH


async def save_report_history(store_id: int, file_path: str) -> int:
    """Сохраняет запись об отчёте. Возвращает id."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            'INSERT INTO report_history (store_id, file_path) VALUES (?, ?)',
            (store_id, file_path)
        )
        await db.commit()
        return cursor.lastrowid


async def get_last_report(store_id: int) -> dict | None:
    """Возвращает последний отчёт для магазина."""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            'SELECT id, store_id, file_path, created_at FROM report_history '
            'WHERE store_id = ? ORDER BY created_at DESC LIMIT 1',
            (store_id,)
        )
        row = await cursor.fetchone()
        return dict(row) if row else None


async def cleanup_old_reports(days: int) -> int:
    """Удаляет отчёты старше days дней из БД и с диска. Возвращает кол-во удалённых."""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            "SELECT id, file_path FROM report_history "
            "WHERE created_at < datetime('now', ?)",
            (f'-{days} days',)
        )
        rows = await cursor.fetchall()

        deleted = 0
        for row in rows:
            fp = row['file_path']
            if fp and await asyncio.to_thread(os.path.exists, fp):
                await asyncio.to_thread(os.remove, fp)
            await db.execute('DELETE FROM report_history WHERE id = ?', (row['id'],))
            deleted += 1

        await db.commit()
    return deleted
