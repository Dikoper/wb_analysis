"""
Анализ остатков: выбор магазина, генерация/просмотр отчётов.
"""

import os
import asyncio
import logging

from aiogram import Router, F
from aiogram.types import CallbackQuery, FSInputFile

from bot.keyboards import (
    MenuCB, StoreCB, NavCB,
    stores_list_kb, store_actions_kb,
)
from bot.db import get_stores, get_store, get_last_report, save_report_history
from bot.report import generate_report
from wb_api import WBTokenError

logger = logging.getLogger(__name__)

router = Router()


@router.callback_query(MenuCB.filter(F.action == "analysis"))
async def menu_analysis(callback: CallbackQuery):
    """Показывает список магазинов для анализа."""
    stores = await get_stores()
    await callback.message.edit_text(
        "📊 <b>Анализ остатков</b>\n\nВыберите магазин:",
        reply_markup=stores_list_kb(stores, action="select"),
        parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(NavCB.filter(F.target == "analysis"))
async def nav_analysis(callback: CallbackQuery):
    """Возврат к списку магазинов."""
    stores = await get_stores()
    await callback.message.edit_text(
        "📊 <b>Анализ остатков</b>\n\nВыберите магазин:",
        reply_markup=stores_list_kb(stores, action="select"),
        parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(StoreCB.filter(F.action == "select"))
async def store_selected(callback: CallbackQuery, callback_data: StoreCB):
    """Магазин выбран — показать действия."""
    store = await get_store(callback_data.store_id)
    if not store:
        await callback.answer("Магазин не найден", show_alert=True)
        return

    last_report = await get_last_report(callback_data.store_id)
    last_time = None
    if last_report:
        last_time = last_report['created_at'][:16].replace('T', ' ')

    name = store.get('name') or f"Магазин #{store['id']}"
    await callback.message.edit_text(
        f"🏪 <b>{name}</b>\n\nВыберите действие:",
        reply_markup=store_actions_kb(callback_data.store_id, last_time),
        parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(StoreCB.filter(F.action == "last"))
async def send_last_report(callback: CallbackQuery, callback_data: StoreCB):
    """Отправить последний сохранённый отчёт."""
    last_report = await get_last_report(callback_data.store_id)
    if not last_report or not os.path.exists(last_report['file_path']):
        await callback.answer("Отчёт не найден. Сгенерируйте новый.", show_alert=True)
        return

    document = FSInputFile(
        last_report['file_path'],
        filename=os.path.basename(last_report['file_path'])
    )
    await callback.message.answer_document(
        document,
        caption=f"📄 Отчёт от {last_report['created_at'][:16]}"
    )
    await callback.answer()


@router.callback_query(StoreCB.filter(F.action == "new"))
async def generate_new_report(callback: CallbackQuery, callback_data: StoreCB):
    """Генерация нового отчёта для магазина."""
    store = await get_store(callback_data.store_id)
    if not store:
        await callback.answer("Магазин не найден", show_alert=True)
        return

    name = store.get('name') or f"Магазин #{store['id']}"
    await callback.message.edit_text(
        f"⏳ Генерирую отчёт для <b>{name}</b>...\n"
        "Загрузка данных из WB API.",
        parse_mode="HTML"
    )
    await callback.answer()

    try:
        report_path = await asyncio.to_thread(
            generate_report, token=store['token'], store_name=name
        )
        await save_report_history(callback_data.store_id, report_path)

        document = FSInputFile(report_path, filename=os.path.basename(report_path))
        await callback.message.answer_document(document, caption=f"📊 Отчёт для {name} готов!")

    except WBTokenError as e:
        await callback.message.edit_text(
            f"🔑 <b>Ошибка токена для {name}</b>\n\n"
            f"{e}\n\n"
            "Обновите токен: ⚙️ Настройки → Управление магазинами",
            parse_mode="HTML"
        )
    except Exception as e:
        logger.error(f"Ошибка генерации отчёта для {name}: {e}", exc_info=True)
        await callback.message.edit_text(
            f"❌ Ошибка генерации отчёта для {name}:\n{e}"
        )
