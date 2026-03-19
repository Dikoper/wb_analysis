"""
Middleware авторизации: проверяет доступ пользователя по паролю.
"""

import logging

from aiogram import BaseMiddleware

from bot.config import BOT_PASSWORD
from bot.db import is_authorized
from bot.states import MenuStates

logger = logging.getLogger(__name__)


class AuthMiddleware(BaseMiddleware):
    """Блокирует неавторизованных пользователей."""

    async def __call__(self, handler, event, data):
        user = data.get("event_from_user")
        if not user:
            return await handler(event, data)

        # Если пароль не задан — пропускаем всех (режим разработки)
        if not BOT_PASSWORD:
            return await handler(event, data)

        # Если авторизован — пропускаем
        if await is_authorized(user.id):
            return await handler(event, data)

        # Если в состоянии ввода пароля — пропускаем к хендлеру
        state = data.get("state")
        if state:
            current = await state.get_state()
            if current == MenuStates.waiting_password.state:
                return await handler(event, data)

        # Если это команда /start — пропускаем к хендлеру
        if hasattr(event, 'text') and event.text and event.text.startswith('/start'):
            return await handler(event, data)

        # Всё остальное — блокируем
        logger.debug(f"Blocked request from unauthorized user {user.id}")
        if hasattr(event, 'answer'):
            await event.answer("🔒 Доступ ограничен. Введите /start для авторизации.")
        return
