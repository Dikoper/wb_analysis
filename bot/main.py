"""
Точка входа телеграм-бота WB Analiz.

Запускает бота с интерактивным меню и планировщиком отчётов.
"""

import asyncio
import logging
import os
import sys

from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage

# Добавляем родительскую директорию для импорта
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bot.config import TELEGRAM_TOKEN, LOGS_DIR
from bot.handlers import register_routers
from bot.db import init_db
from bot.scheduler import setup_scheduler

# Создаём директорию для логов
os.makedirs(LOGS_DIR, exist_ok=True)

# Логирование в консоль и файл
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(os.path.join(LOGS_DIR, 'bot.log'), encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


async def main():
    """Основная функция запуска бота."""
    if not TELEGRAM_TOKEN:
        logger.error("TELEGRAM_TOKEN не задан!")
        sys.exit(1)

    # Инициализация БД (создание таблиц + автомиграция токена из env)
    await init_db()

    # Инициализация бота и диспетчера
    bot = Bot(token=TELEGRAM_TOKEN)
    dp = Dispatcher(storage=MemoryStorage())

    # Подключение роутеров
    register_routers(dp)

    # Планировщик
    scheduler = await setup_scheduler(bot)

    # Запуск polling
    logger.info("Бот запущен!")
    try:
        await dp.start_polling(bot)
    finally:
        scheduler.shutdown()
        await bot.session.close()


if __name__ == '__main__':
    asyncio.run(main())
