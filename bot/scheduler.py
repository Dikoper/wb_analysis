"""
Планировщик ежедневных отчётов для всех магазинов.
"""

import asyncio
import os
import logging

from aiogram import Bot
from aiogram.types import FSInputFile
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz

from bot.config import CHAT_ID, TIMEZONE
from bot.db import get_stores, get_setting, save_report_history
from bot.report import generate_report
from wb_api import WBTokenError

logger = logging.getLogger(__name__)

DEFAULT_REPORT_TIME = "09:00"


async def send_daily_reports(bot: Bot):
    """
    Генерирует и отправляет отчёты по всем активным магазинам.
    Вызывается планировщиком.
    """
    stores = await get_stores()
    if not stores:
        logger.warning("Нет активных магазинов для отчёта")
        return

    logger.info(f"Генерация отчётов для {len(stores)} магазинов...")

    for store in stores:
        name = store.get('name') or f"Магазин #{store['id']}"
        try:
            report_path = await asyncio.to_thread(
                generate_report, token=store['token'], store_name=name
            )
            await save_report_history(store['id'], report_path)

            document = FSInputFile(report_path, filename=os.path.basename(report_path))
            await bot.send_document(
                chat_id=CHAT_ID,
                document=document,
                caption=f"📊 Ежедневный отчёт: {name}"
            )
            logger.info(f"Отчёт для {name} отправлен")

        except WBTokenError as e:
            await bot.send_message(
                chat_id=CHAT_ID,
                text=(
                    f"🔑 <b>Ошибка токена: {name}</b>\n\n"
                    f"{e}\n\n"
                    "Обновите токен: /menu → ⚙️ Настройки → Магазины"
                ),
                parse_mode="HTML"
            )
            logger.error(f"Ошибка токена для {name}: {e}")

        except Exception as e:
            await bot.send_message(
                chat_id=CHAT_ID,
                text=f"❌ Ошибка отчёта для {name}: {e}"
            )
            logger.error(f"Ошибка отчёта для {name}: {e}", exc_info=True)


async def setup_scheduler(bot: Bot) -> AsyncIOScheduler:
    """Настраивает и запускает планировщик."""
    scheduler = AsyncIOScheduler(timezone=pytz.timezone(TIMEZONE))

    report_time = await get_setting('report_time', DEFAULT_REPORT_TIME)
    hour, minute = map(int, report_time.split(':'))

    scheduler.add_job(
        send_daily_reports,
        CronTrigger(hour=hour, minute=minute),
        args=[bot],
        id='daily_reports',
        replace_existing=True
    )

    scheduler.start()
    logger.info(f"Планировщик запущен. Отчёты в {report_time} {TIMEZONE}")
    return scheduler
