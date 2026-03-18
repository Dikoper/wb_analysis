"""
Планировщик ежедневных отчётов для всех магазинов.
"""

import asyncio
import os
import logging
from datetime import datetime

from aiogram import Bot
from aiogram.types import FSInputFile
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz

from bot.config import TIMEZONE
from bot.db import get_stores, get_setting, save_report_history, get_subscribers
from bot.keyboards import store_display_name
from bot.report import generate_report
from wb_api import WBTokenError

logger = logging.getLogger(__name__)

DEFAULT_REPORT_TIME = "09:00"

_scheduler: "AsyncIOScheduler | None" = None
_bot: "Bot | None" = None


async def send_daily_reports(bot: Bot):
    """
    Генерирует и отправляет отчёты по всем активным магазинам всем подписчикам.
    Вызывается планировщиком.
    """
    subscribers = await get_subscribers()
    if not subscribers:
        logger.warning("Нет подписчиков для рассылки отчётов")
        return

    stores = await get_stores()
    if not stores:
        logger.warning("Нет активных магазинов для отчёта")
        return

    logger.info(f"Генерация отчётов для {len(stores)} магазинов, подписчиков: {len(subscribers)}")

    # Шапка рассылки
    tz = pytz.timezone(TIMEZONE)
    now = datetime.now(tz)
    report_time = await get_setting('report_time', DEFAULT_REPORT_TIME)
    store_names = ", ".join(store_display_name(s) for s in stores)

    months_ru = {
        1: "января", 2: "февраля", 3: "марта", 4: "апреля",
        5: "мая", 6: "июня", 7: "июля", 8: "августа",
        9: "сентября", 10: "октября", 11: "ноября", 12: "декабря",
    }
    date_str = f"{now.day} {months_ru[now.month]} {now.year}"

    header = (
        f"📊 <b>Ежедневный отчёт WB</b>\n"
        f"{date_str} · {report_time} МСК\n\n"
        f"🏪 {store_names}"
    )
    for chat_id in subscribers:
        await bot.send_message(chat_id, header, parse_mode="HTML")

    # Отчёт по каждому магазину
    for store in stores:
        name = store_display_name(store)
        try:
            report_path = await asyncio.to_thread(
                generate_report, token=store['token'], store_name=name
            )
            await save_report_history(store['id'], report_path)

            document = FSInputFile(report_path, filename=os.path.basename(report_path))
            for chat_id in subscribers:
                await bot.send_document(
                    chat_id=chat_id,
                    document=document,
                    caption=f"📄 {name}"
                )
            logger.info(f"Отчёт для {name} отправлен {len(subscribers)} подписчикам")

        except WBTokenError as e:
            for chat_id in subscribers:
                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        f"🔑 <b>Ошибка токена: {name}</b>\n\n"
                        f"{e}\n\n"
                        "Обновите токен: /menu → ⚙️ Настройки → Магазины"
                    ),
                    parse_mode="HTML"
                )
            logger.error(f"Ошибка токена для {name}: {e}")

        except Exception as e:
            for chat_id in subscribers:
                await bot.send_message(
                    chat_id=chat_id,
                    text=f"❌ Ошибка отчёта для {name}: {e}"
                )
            logger.error(f"Ошибка отчёта для {name}: {e}", exc_info=True)


def reschedule_daily_reports(time_str: str):
    """Перепланирует задачу без перезапуска бота."""
    if _scheduler is None:
        logger.warning("Планировщик не инициализирован, перепланирование невозможно")
        return
    hour, minute = map(int, time_str.split(':'))
    _scheduler.reschedule_job(
        'daily_reports',
        trigger=CronTrigger(hour=hour, minute=minute, timezone=pytz.timezone(TIMEZONE))
    )
    logger.info(f"Планировщик перепланирован: отчёты в {time_str} {TIMEZONE}")


async def setup_scheduler(bot: Bot) -> AsyncIOScheduler:
    """Настраивает и запускает планировщик."""
    global _scheduler, _bot
    _bot = bot
    _scheduler = AsyncIOScheduler(timezone=pytz.timezone(TIMEZONE))

    report_time = await get_setting('report_time', DEFAULT_REPORT_TIME)
    hour, minute = map(int, report_time.split(':'))

    _scheduler.add_job(
        send_daily_reports,
        CronTrigger(hour=hour, minute=minute),
        args=[bot],
        id='daily_reports',
        replace_existing=True
    )

    _scheduler.start()
    logger.info(f"Планировщик запущен. Отчёты в {report_time} {TIMEZONE}")
    return _scheduler
