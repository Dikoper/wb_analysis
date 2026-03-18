"""
Главное меню: /start, /menu, навигация.
"""

import logging

from aiogram import Router, F
from aiogram.types import Message, CallbackQuery
from aiogram.filters import Command

from bot.keyboards import MenuCB, NavCB, main_menu_kb
from bot.db import get_stores, is_subscriber

logger = logging.getLogger(__name__)

router = Router()


async def _send_main_menu(target, stores: list = None):
    """Отправляет или редактирует главное меню с информацией."""
    if stores is None:
        stores = await get_stores()

    if isinstance(target, Message):
        chat_id = target.chat.id
    else:
        chat_id = target.message.chat.id

    subscribed = await is_subscriber(chat_id)
    sub_status = "🔔 Рассылка подключена" if subscribed else "🔕 Рассылка отключена"

    text = (
        "📊 <b>WB Analiz Bot</b>\n\n"
        f"🏪 Магазинов подключено: {len(stores)}\n"
    )
    for s in stores:
        name = s.get('name') or f"Магазин #{s['id']}"
        text += f"  • {name}\n"

    if not stores:
        text += "  ⚠️ Нет подключённых магазинов\n"

    text += f"\n{sub_status}\n\nВыберите действие:"

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
