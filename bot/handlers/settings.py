"""
Настройки: время отчётов.
"""

import re
import logging

from aiogram import Router, F
from aiogram.types import CallbackQuery, Message
from aiogram.filters import StateFilter
from aiogram.fsm.context import FSMContext

from bot.keyboards import MenuCB, SettingsCB, NavCB, settings_kb, cancel_kb
from bot.states import MenuStates
from bot.db import get_setting, set_setting

logger = logging.getLogger(__name__)

router = Router()

DEFAULT_REPORT_TIME = "09:00"


@router.callback_query(MenuCB.filter(F.action == "settings"))
async def menu_settings(callback: CallbackQuery):
    """Показать меню настроек."""
    current_time = await get_setting('report_time', DEFAULT_REPORT_TIME)
    await callback.message.edit_text(
        "⚙️ <b>Настройки</b>",
        reply_markup=settings_kb(current_time),
        parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(NavCB.filter(F.target == "settings"))
async def nav_settings(callback: CallbackQuery):
    """Возврат в настройки."""
    current_time = await get_setting('report_time', DEFAULT_REPORT_TIME)
    await callback.message.edit_text(
        "⚙️ <b>Настройки</b>",
        reply_markup=settings_kb(current_time),
        parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(SettingsCB.filter(F.action == "time"))
async def ask_report_time(callback: CallbackQuery, state: FSMContext):
    """Запрос нового времени отчёта."""
    await callback.message.edit_text(
        "🕐 Введите новое время отправки отчёта в формате <b>ЧЧ:ММ</b>\n"
        "(например: 09:00, 18:30)",
        reply_markup=cancel_kb(),
        parse_mode="HTML"
    )
    await state.set_state(MenuStates.set_report_time)
    await callback.answer()


@router.message(MenuStates.set_report_time, F.text)
async def set_report_time(message: Message, state: FSMContext):
    """Обработка ввода нового времени."""
    text = message.text.strip()

    # Валидация формата ЧЧ:ММ
    if not re.match(r'^([01]\d|2[0-3]):([0-5]\d)$', text):
        await message.answer(
            "❌ Неверный формат. Введите время как <b>ЧЧ:ММ</b> (например: 09:00)",
            reply_markup=cancel_kb(),
            parse_mode="HTML"
        )
        return

    await set_setting('report_time', text)
    await state.clear()

    await message.answer(
        f"✅ Время отчёта изменено на <b>{text}</b> МСК\n\n"
        "Изменения вступят в силу при следующем запуске бота.",
        parse_mode="HTML"
    )
    logger.info(f"Время отчёта изменено на {text}")
