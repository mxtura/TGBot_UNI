from aiogram.dispatcher.middlewares import BaseMiddleware
from aiogram.types import Message
from aiogram.dispatcher.handler import SkipHandler
import pytz

timezone = pytz.timezone('Europe/Moscow')
        
class SkipHandlerMiddleware(BaseMiddleware):
    """Пропускает сообщения отправленные боту, когда он был выключен."""
    def __init__(self, bot_start_time):
        self.bot_start_time = bot_start_time
        super().__init__()

    async def on_pre_process_message(self, message: Message, data: dict):
        bot_start_time = self.bot_start_time
        message_date = message.date.astimezone(timezone).strftime('%Y-%m-%d %H:%M:%S')
        if bot_start_time > message_date:
            raise SkipHandler()