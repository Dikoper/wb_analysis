"""
Управление магазинами: добавление, редактирование, удаление.
"""

import asyncio
import logging

from aiogram import Router, F
from aiogram.types import CallbackQuery, Message
from aiogram.fsm.context import FSMContext

from bot.keyboards import (
    SettingsCB, StoreCB, NavCB,
    store_management_kb, confirm_delete_kb, cancel_kb,
)
from bot.states import MenuStates
from bot.db import get_stores, get_store, add_store, update_store, delete_store
from wb_api import get_seller_info, WBTokenError

logger = logging.getLogger(__name__)

router = Router()


@router.callback_query(SettingsCB.filter(F.action == "stores"))
async def manage_stores(callback: CallbackQuery):
    """Показать список магазинов для управления."""
    stores = await get_stores()
    await callback.message.edit_text(
        "🏪 <b>Управление магазинами</b>",
        reply_markup=store_management_kb(stores),
        parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(NavCB.filter(F.target == "store_mgmt"))
async def nav_store_mgmt(callback: CallbackQuery, state: FSMContext):
    """Возврат к управлению магазинами."""
    await state.clear()
    stores = await get_stores()
    await callback.message.edit_text(
        "🏪 <b>Управление магазинами</b>",
        reply_markup=store_management_kb(stores),
        parse_mode="HTML"
    )
    await callback.answer()


# === Добавление магазина ===

@router.callback_query(StoreCB.filter(F.action == "add"))
async def ask_store_token(callback: CallbackQuery, state: FSMContext):
    """Запрос токена нового магазина."""
    await callback.message.edit_text(
        "🔑 Отправьте <b>токен WB API</b> нового магазина.\n\n"
        "Токен можно получить в личном кабинете WB → Настройки → Доступ к API",
        reply_markup=cancel_kb(),
        parse_mode="HTML"
    )
    await state.set_state(MenuStates.add_store_token)
    await callback.answer()


@router.message(MenuStates.add_store_token, F.text)
async def process_store_token(message: Message, state: FSMContext):
    """Проверка токена и добавление магазина."""
    token = message.text.strip()

    if len(token) < 10:
        await message.answer(
            "❌ Токен слишком короткий. Проверьте правильность.",
            reply_markup=cancel_kb()
        )
        return

    await message.answer("⏳ Проверяю токен...")

    try:
        info = await asyncio.to_thread(get_seller_info, token)
        name = info.get('name', 'Новый магазин') if info else 'Новый магазин'
    except WBTokenError:
        await message.answer(
            "❌ Токен невалиден или истёк.\n"
            "Проверьте правильность токена и попробуйте снова.",
            reply_markup=cancel_kb()
        )
        return
    except Exception as e:
        logger.warning(f"Не удалось получить имя магазина: {e}")
        name = 'Новый магазин'

    store_id = await add_store(token, name)
    await state.clear()

    await message.answer(
        f"✅ Магазин <b>{name}</b> добавлен (ID: {store_id})",
        parse_mode="HTML"
    )
    logger.info(f"Добавлен магазин: {name} (ID: {store_id})")


# === Редактирование магазина ===

@router.callback_query(StoreCB.filter(F.action == "edit"))
async def ask_edit_store(callback: CallbackQuery, callback_data: StoreCB, state: FSMContext):
    """Запрос нового имени магазина."""
    store = await get_store(callback_data.store_id)
    if not store:
        await callback.answer("Магазин не найден", show_alert=True)
        return

    name = store.get('name') or f"Магазин #{store['id']}"
    await state.update_data(edit_store_id=callback_data.store_id)
    await state.set_state(MenuStates.edit_store_name)

    await callback.message.edit_text(
        f"✏️ Текущее имя: <b>{name}</b>\n\n"
        "Введите новое имя для магазина:",
        reply_markup=cancel_kb(),
        parse_mode="HTML"
    )
    await callback.answer()


@router.message(MenuStates.edit_store_name, F.text)
async def process_edit_store(message: Message, state: FSMContext):
    """Сохранение нового имени магазина."""
    data = await state.get_data()
    store_id = data.get('edit_store_id')
    if not store_id:
        await state.clear()
        return

    new_name = message.text.strip()
    await update_store(store_id, name=new_name)
    await state.clear()

    await message.answer(
        f"✅ Магазин переименован в <b>{new_name}</b>",
        parse_mode="HTML"
    )
    logger.info(f"Магазин #{store_id} переименован в {new_name}")


# === Удаление магазина ===

@router.callback_query(StoreCB.filter(F.action == "delete"))
async def ask_delete_store(callback: CallbackQuery, callback_data: StoreCB):
    """Подтверждение удаления магазина."""
    store = await get_store(callback_data.store_id)
    if not store:
        await callback.answer("Магазин не найден", show_alert=True)
        return

    name = store.get('name') or f"Магазин #{store['id']}"
    await callback.message.edit_text(
        f"❓ Удалить магазин <b>{name}</b>?",
        reply_markup=confirm_delete_kb(callback_data.store_id),
        parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(StoreCB.filter(F.action == "confirm_delete"))
async def confirm_delete_store(callback: CallbackQuery, callback_data: StoreCB):
    """Удаление магазина после подтверждения."""
    store = await get_store(callback_data.store_id)
    name = store.get('name', '?') if store else '?'

    await delete_store(callback_data.store_id)

    stores = await get_stores()
    await callback.message.edit_text(
        f"🗑 Магазин <b>{name}</b> удалён.\n\n"
        "🏪 <b>Управление магазинами</b>",
        reply_markup=store_management_kb(stores),
        parse_mode="HTML"
    )
    await callback.answer()
    logger.info(f"Магазин #{callback_data.store_id} ({name}) удалён")
