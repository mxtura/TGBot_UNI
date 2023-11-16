from aiogram import Dispatcher
from aiogram.types import Message
from aiogram.dispatcher import FSMContext
import tgbot.keyboards.reply as rkb
from tgbot.models.database_instance import db


async def logout(message: Message, state: FSMContext):
    """Удаляет пользователя из базы данных."""
    await db.delete_user_id(message.from_user.id)
    await message.answer("Вы больше не авторизованы.\n"
                         "<b>Чтобы авторизоваться введите команду /start</b>", reply_markup=rkb.empty_keyboard)
    await state.finish()


def register_logout(dp: Dispatcher):
    dp.register_message_handler(logout, commands=['logout'], state="*")
