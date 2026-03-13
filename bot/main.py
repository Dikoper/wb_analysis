"""
Точка входа телеграм-бота WB Analiz.

Запускает бота с планировщиком отправки отчётов в 8:00 МСК.
"""

import asyncio
import logging
import os
import sys

from aiogram import Bot, Dispatcher
from aiogram.types import FSInputFile
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz

# Добавляем родительскую директорию для импорта
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bot.config import TELEGRAM_TOKEN, CHAT_ID, REPORT_TIME, TIMEZONE, LOGS_DIR
from bot.handlers import router
from bot.report import generate_report

# Создаём директорию для логов
os.makedirs(LOGS_DIR, exist_ok=True)

# Логирование в консоль и файл
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),  # консоль
        logging.FileHandler(os.path.join(LOGS_DIR, 'bot.log'), encoding='utf-8')  # файл
    ]
)
logger = logging.getLogger(__name__)


async def send_daily_report(bot: Bot):
    """
    Отправляет ежедневный отчёт в указанный чат.

    Вызывается планировщиком в 8:00 МСК.
    """
    logger.info("Генерация ежедневного отчёта...")

    try:
        report_path = await asyncio.to_thread(generate_report)
        document = FSInputFile(report_path, filename=os.path.basename(report_path))

        await bot.send_document(
            chat_id=CHAT_ID,
            document=document,
            caption="📊 Ежедневный отчёт по ценам"
        )
        logger.info(f"Отчёт отправлен в чат {CHAT_ID}")

    except Exception as e:
        logger.error(f"Ошибка отправки отчёта: {e}")
        # Уведомляем об ошибке
        await bot.send_message(
            chat_id=CHAT_ID,
            text=f"❌ Ошибка генерации отчёта: {e}"
        )


async def main():
    """Основная функция запуска бота."""
    # Проверяем наличие токенов
    if not TELEGRAM_TOKEN:
        logger.error("TELEGRAM_TOKEN не задан!")
        sys.exit(1)

    if not CHAT_ID:
        logger.error("CHAT_ID не задан!")
        sys.exit(1)

    # Инициализация бота и диспетчера
    bot = Bot(token=TELEGRAM_TOKEN)
    dp = Dispatcher()
    dp.include_router(router)

    # Настройка планировщика
    scheduler = AsyncIOScheduler(timezone=pytz.timezone(TIMEZONE))

    # Парсим время отправки
    hour, minute = map(int, REPORT_TIME.split(':'))

    # Добавляем задачу на каждый день в 8:00 МСК
    scheduler.add_job(
        send_daily_report,
        CronTrigger(hour=hour, minute=minute),
        args=[bot],
        id='daily_report',
        replace_existing=True
    )

    scheduler.start()
    logger.info(f"Планировщик запущен. Отчёт будет отправляться в {REPORT_TIME} {TIMEZONE}")

    # Запуск polling
    logger.info("Бот запущен!")
    try:
        await dp.start_polling(bot)
    finally:
        scheduler.shutdown()
        await bot.session.close()


if __name__ == '__main__':
    asyncio.run(main())
