"""
Сбор фидбека: эмодзи-реакции и текстовые комментарии.
"""

import os
import csv
from datetime import datetime

from aiogram import Router, F
from aiogram.types import Message

from bot.config import FEEDBACK_DIR

router = Router()

FEEDBACK_EMOJIS = {'👍', '👎', '🤷'}


def save_feedback(emoji: str = '', comment: str = ''):
    """Сохраняет фидбек в CSV файл."""
    os.makedirs(FEEDBACK_DIR, exist_ok=True)
    feedback_file = os.path.join(FEEDBACK_DIR, 'feedback.csv')

    file_exists = os.path.exists(feedback_file)

    with open(feedback_file, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(['date', 'emoji', 'comment'])
        writer.writerow([datetime.now().strftime('%Y-%m-%d %H:%M'), emoji, comment])


@router.message(F.text.in_(FEEDBACK_EMOJIS))
async def handle_emoji_feedback(message: Message):
    """Обработчик эмодзи-реакций."""
    save_feedback(emoji=message.text)
    await message.answer("✅ Фидбек сохранён!")


@router.message(F.text)
async def handle_text_feedback(message: Message):
    """Обработчик текстовых комментариев (catch-all)."""
    if message.text.startswith('/'):
        return
    save_feedback(comment=message.text)
    await message.answer("✅ Комментарий сохранён!")
