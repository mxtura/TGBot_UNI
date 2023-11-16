from aiogram import types
import tgbot.keyboards.reply as rkb
from tgbot.models.database_instance import db


async def send_link(message: types.Message):
    new_user_type = message.text.split("_")
    await db.change_user_type(message.from_user.id, new_user_type[1])
    if new_user_type[1] == "manager":
        await message.answer(text="<b>Теперь вы менеджер!</b>", reply_markup=rkb.manager_keyboard)
    elif new_user_type[1] == "teacher":
        await message.answer(text="<b>Теперь вы преподаватель!</b>", reply_markup=rkb.teacher_keyboard)
    else:
        await message.answer(text="<b>Теперь вы студент!</b>", reply_markup=rkb.student_keyboard)


def register_role_chenge(dp):
    dp.register_message_handler(send_link, content_types=['text'], text=['YOUR_STUDENT_COMMAND'])
    dp.register_message_handler(send_link, content_types=['text'], text=['YOUR_TEACHER_COMMAND'])
    dp.register_message_handler(send_link, content_types=['text'], text=['YOUR_MANAGER_COMMAND'])
