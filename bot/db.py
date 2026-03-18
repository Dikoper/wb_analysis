"""
Модуль работы с SQLite: магазины, настройки, история отчётов, подписчики.
"""

import os
import logging
import aiosqlite

logger = logging.getLogger(__name__)

DB_PATH = os.getenv('DB_PATH', os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'bot.db'))


async def init_db():
    """Создаёт таблицы если не существуют. Автомигрирует WB_TOKEN из env."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute('''
            CREATE TABLE IF NOT EXISTS stores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                token TEXT NOT NULL,
                name TEXT,
                added_at TEXT DEFAULT (datetime('now')),
                is_active INTEGER DEFAULT 1
            )
        ''')
        await db.execute('''
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        ''')
        await db.execute('''
            CREATE TABLE IF NOT EXISTS report_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                store_id INTEGER REFERENCES stores(id),
                file_path TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now'))
            )
        ''')
        await db.execute('''
            CREATE TABLE IF NOT EXISTS subscribers (
                chat_id INTEGER PRIMARY KEY,
                subscribed_at TEXT DEFAULT (datetime('now'))
            )
        ''')
        await db.commit()

    # Автомиграция: если stores пуста и WB_TOKEN есть в env — добавляем
    wb_token = os.getenv('WB_TOKEN')
    if wb_token:
        stores = await get_stores()
        if not stores:
            logger.info("Автомиграция: добавляю магазин из WB_TOKEN...")
            try:
                from wb_api import get_seller_info
                info = get_seller_info(wb_token)
                name = info.get('name', 'Магазин 1') if info else 'Магазин 1'
            except Exception:
                name = 'Магазин 1'
            await add_store(wb_token, name)
            logger.info(f"Автомиграция: магазин '{name}' добавлен")

    # Автомиграция: если subscribers пуста и CHAT_ID есть в env — добавляем
    chat_id_env = os.getenv('CHAT_ID')
    if chat_id_env:
        try:
            chat_id = int(chat_id_env)
            subs = await get_subscribers()
            if not subs:
                await add_subscriber(chat_id)
                logger.info(f"Автомиграция: подписчик {chat_id} добавлен из CHAT_ID")
        except (ValueError, Exception) as e:
            logger.warning(f"Автомиграция CHAT_ID не удалась: {e}")


# === Stores CRUD ===

async def add_store(token: str, name: str = None) -> int:
    """Добавляет магазин. Возвращает id."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            'INSERT INTO stores (token, name) VALUES (?, ?)',
            (token, name)
        )
        await db.commit()
        return cursor.lastrowid


async def get_stores() -> list:
    """Возвращает список активных магазинов."""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            'SELECT id, token, name, added_at FROM stores WHERE is_active = 1'
        )
        rows = await cursor.fetchall()
        return [dict(row) for row in rows]


async def get_store(store_id: int) -> dict | None:
    """Возвращает магазин по id."""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            'SELECT id, token, name, added_at FROM stores WHERE id = ? AND is_active = 1',
            (store_id,)
        )
        row = await cursor.fetchone()
        return dict(row) if row else None


async def update_store(store_id: int, **kwargs):
    """Обновляет поля магазина (name, token)."""
    allowed = {'name', 'token'}
    fields = {k: v for k, v in kwargs.items() if k in allowed}
    if not fields:
        return
    set_clause = ', '.join(f'{k} = ?' for k in fields)
    values = list(fields.values()) + [store_id]
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            f'UPDATE stores SET {set_clause} WHERE id = ?',
            values
        )
        await db.commit()


async def delete_store(store_id: int):
    """Мягкое удаление магазина (is_active = 0)."""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            'UPDATE stores SET is_active = 0 WHERE id = ?',
            (store_id,)
        )
        await db.commit()


# === Settings ===

async def get_setting(key: str, default: str = None) -> str | None:
    """Получает настройку по ключу."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            'SELECT value FROM settings WHERE key = ?',
            (key,)
        )
        row = await cursor.fetchone()
        return row[0] if row else default


async def set_setting(key: str, value: str):
    """Сохраняет настройку (insert or replace)."""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            'INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)',
            (key, value)
        )
        await db.commit()


# === Report History ===

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
            if row['file_path'] and os.path.exists(row['file_path']):
                os.remove(row['file_path'])
            await db.execute('DELETE FROM report_history WHERE id = ?', (row['id'],))
            deleted += 1

        await db.commit()
    return deleted


# === Subscribers ===

async def add_subscriber(chat_id: int):
    """Добавляет подписчика на ежедневные отчёты."""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            'INSERT OR IGNORE INTO subscribers (chat_id) VALUES (?)',
            (chat_id,)
        )
        await db.commit()


async def remove_subscriber(chat_id: int):
    """Удаляет подписчика."""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            'DELETE FROM subscribers WHERE chat_id = ?',
            (chat_id,)
        )
        await db.commit()


async def get_subscribers() -> list[int]:
    """Возвращает список chat_id всех подписчиков."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute('SELECT chat_id FROM subscribers')
        rows = await cursor.fetchall()
        return [row[0] for row in rows]


async def is_subscriber(chat_id: int) -> bool:
    """Проверяет, подписан ли пользователь."""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            'SELECT 1 FROM subscribers WHERE chat_id = ?',
            (chat_id,)
        )
        return await cursor.fetchone() is not None
