from aiogram.types import Message
import tgbot.keyboards.reply as rkb
from tgbot.models.database_instance import db

async def open_personal_data_form(message: Message):
    """Отправляет сообщение с данными пользователя."""
    user_id = message.from_user.id
    name = await db.get_fio(user_id)
    user_type = await db.get_user_type(user_id)
    login = await db.get_login(user_id)
    personal_data = f"Личная информация\n<b>ФИО</b>: {name}\n<b>Логин</b>: {login}\n"
    if user_type == "student":
        group = await db.get_user_group_name(user_id)
        personal_data += f"<b>Роль</b>: студент\n<b>Группа</b>: {group}"
    elif user_type == "teacher":
        personal_data += "<b>Роль</b>: преподаватель"
    else:
        personal_data += "<b>Роль</b>: менеджер"
    await message.answer(text=personal_data, reply_markup=rkb.personal_data_editing_keyboard)