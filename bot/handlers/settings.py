"""
Настройки: время отчётов.
"""

import re
import logging

from aiogram import Router, F, Bot
from aiogram.types import CallbackQuery, Message
from aiogram.fsm.context import FSMContext

from bot.keyboards import MenuCB, SettingsCB, NavCB, SubscribeCB, settings_kb, cancel_kb
from bot.states import MenuStates
from bot.db import get_setting, set_setting, is_subscriber, add_subscriber, remove_subscriber
from bot.scheduler import reschedule_daily_reports

logger = logging.getLogger(__name__)

router = Router()

DEFAULT_REPORT_TIME = "09:00"


async def _send_settings(callback: CallbackQuery):
    """Показывает меню настроек с актуальным состоянием."""
    current_time = await get_setting('report_time', DEFAULT_REPORT_TIME)
    subscribed = await is_subscriber(callback.message.chat.id)
    await callback.message.edit_text(
        "⚙️ <b>Настройки</b>",
        reply_markup=settings_kb(current_time, subscribed),
        parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(MenuCB.filter(F.action == "settings"))
async def menu_settings(callback: CallbackQuery):
    """Показать меню настроек."""
    await _send_settings(callback)


@router.callback_query(NavCB.filter(F.target == "settings"))
async def nav_settings(callback: CallbackQuery):
    """Возврат в настройки."""
    await _send_settings(callback)


@router.callback_query(SubscribeCB.filter(F.action == "toggle"))
async def toggle_subscription(callback: CallbackQuery):
    """Подписка/отписка от ежедневных отчётов."""
    chat_id = callback.message.chat.id
    subscribed = await is_subscriber(chat_id)

    if subscribed:
        await remove_subscriber(chat_id)
        await callback.answer("Вы отписались от рассылки", show_alert=True)
        logger.info(f"Пользователь {chat_id} отписался от отчётов")
    else:
        await add_subscriber(chat_id)
        await callback.answer("Вы подписались на ежедневные отчёты!", show_alert=True)
        logger.info(f"Пользователь {chat_id} подписался на отчёты")

    await _send_settings(callback)


@router.callback_query(SettingsCB.filter(F.action == "time"))
async def ask_report_time(callback: CallbackQuery, state: FSMContext):
    """Запрос нового времени отчёта."""
    await callback.message.edit_text(
        "🕐 Введите новое время отправки отчёта в формате <b>ЧЧ:ММ</b>\n"
        "(например: 09:00, 18:30)",
        reply_markup=cancel_kb(),
        parse_mode="HTML"
    )
    await state.update_data(bot_msg_id=callback.message.message_id)
    await state.set_state(MenuStates.set_report_time)
    await callback.answer()


@router.message(MenuStates.set_report_time, F.text)
async def set_report_time(message: Message, state: FSMContext, bot: Bot):
    """Обработка ввода нового времени."""
    text = message.text.strip()

    # Удалить сообщение пользователя
    try:
        await message.delete()
    except Exception:
        pass

    data = await state.get_data()
    bot_msg_id = data.get('bot_msg_id')
    chat_id = message.chat.id

    current_time = await get_setting('report_time', DEFAULT_REPORT_TIME)
    subscribed = await is_subscriber(chat_id)

    async def edit_bot_msg(msg_text: str, reply_markup=None):
        if bot_msg_id:
            try:
                await bot.edit_message_text(
                    msg_text, chat_id=chat_id, message_id=bot_msg_id,
                    reply_markup=reply_markup, parse_mode="HTML"
                )
                return
            except Exception:
                pass
        await bot.send_message(chat_id, msg_text, reply_markup=reply_markup, parse_mode="HTML")

    if not re.match(r'^([01]\d|2[0-3]):([0-5]\d)$', text):
        await edit_bot_msg(
            "❌ Неверный формат. Введите время как <b>ЧЧ:ММ</b> (например: 09:00)\n\n"
            "🕐 Введите новое время отправки отчёта в формате <b>ЧЧ:ММ</b>\n"
            "(например: 09:00, 18:30)",
            reply_markup=cancel_kb()
        )
        return

    await set_setting('report_time', text)
    reschedule_daily_reports(text)
    await state.clear()

    await edit_bot_msg(
        f"✅ Время отчёта изменено на <b>{text}</b> МСК\n\n"
        "⚙️ <b>Настройки</b>",
        reply_markup=settings_kb(text, subscribed)
    )
    logger.info(f"Время отчёта изменено на {text}")
