"""
Главное меню: /start, /menu, навигация.
"""

import logging
from datetime import datetime, timezone, timedelta

from aiogram import Router, F
from aiogram.types import Message, CallbackQuery
from aiogram.filters import Command

from bot.keyboards import MenuCB, NavCB, main_menu_kb, store_display_name
from bot.db import get_stores, get_last_report, get_setting, is_subscriber

_MSK = timezone(timedelta(hours=3))

logger = logging.getLogger(__name__)

router = Router()

DEFAULT_REPORT_TIME = "09:00"


async def _send_main_menu(target, stores: list = None):
    """Отправляет или редактирует главное меню с информацией."""
    if stores is None:
        stores = await get_stores()

    if isinstance(target, Message):
        chat_id = target.chat.id
    else:
        chat_id = target.message.chat.id

    subscribed = await is_subscriber(chat_id)
    report_time = await get_setting('report_time', DEFAULT_REPORT_TIME)

    text = "📊 <b>WB Analiz Bot</b>\n\n"

    # Магазины с датой последнего отчёта
    text += f"🏪 Магазинов: {len(stores)}\n"
    if stores:
        for s in stores:
            name = store_display_name(s)
            last = await get_last_report(s['id'])
            if last:
                # created_at хранится в UTC (SQLite datetime('now')), конвертируем в МСК
                try:
                    raw = last['created_at'][:16].replace('T', ' ')
                    utc_dt = datetime.strptime(raw, '%Y-%m-%d %H:%M').replace(tzinfo=timezone.utc)
                    msk_dt = utc_dt.astimezone(_MSK)
                    dt_short = msk_dt.strftime('%d.%m %H:%M')
                except Exception:
                    dt_short = last['created_at'][:16]
                text += f"  • {name} — {dt_short}\n"
            else:
                text += f"  • {name} — нет отчётов\n"
    else:
        text += "  ⚠️ Нет подключённых магазинов\n"

    # Статус рассылки
    sub_icon = "🔔" if subscribed else "🔕"
    sub_label = "активна" if subscribed else "отключена"
    text += f"\n⏰ Рассылка: {report_time} МСК · {sub_icon} {sub_label}\n"
    text += "\nВыберите действие:"

    if isinstance(target, Message):
        await target.answer(text, reply_markup=main_menu_kb(), parse_mode="HTML")
    elif isinstance(target, CallbackQuery):
        await target.message.edit_text(text, reply_markup=main_menu_kb(), parse_mode="HTML")
        await target.answer()


@router.message(Command('start'))
async def cmd_start(message: Message):
    """Обработчик /start — показывает главное меню."""
    await _send_main_menu(message)


@router.message(Command('menu'))
async def cmd_menu(message: Message):
    """Обработчик /menu — показывает главное меню."""
    await _send_main_menu(message)


@router.callback_query(NavCB.filter(F.target == "main"))
async def nav_main(callback: CallbackQuery):
    """Возврат в главное меню."""
    await _send_main_menu(callback)
