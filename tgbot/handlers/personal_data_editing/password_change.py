from aiogram.types import Message, CallbackQuery
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.inline as ikb
import tgbot.keyboards.reply as rkb
from tgbot.misc.states import PasswordChangeStates
from tgbot.filters.user_type import UserTypeFilter
from tgbot.handlers.cancel import cq_cancel_all
from tgbot.models.database_instance import db


password_request_text = "Пожалуйста введите ваш текущий пароль:"
password_confirmed_text = "Пароль подтвержден"
new_password_request_text = "Введите новый пароль:"
wrong_password_text = "*Пароль введен неверно!*"
password_change_confirmation_text = "Вы уверены, что хотите сменить пароль на:"
password_changed_text = "Пароль изменен"


async def send_password_request(message: Message, state: FSMContext):
    """Запрашивает пароль пользователя."""
    del_msg = await message.answer(text=password_request_text,
                                   reply_markup=rkb.password_input_cancel_keyboard)
    await state.update_data(del_msg=del_msg)
    await PasswordChangeStates.checking_password.set()


async def check_password(message: Message, state: FSMContext):
    """Проверяет пароль пользователя."""
    data = await state.get_data()
    del_msg = data.get("del_msg")
    await del_msg.delete()
    await message.delete()

    password = message.text
    real_password = await db.get_password(message.from_user.id)
    if password == real_password:
        del_msg = await message.answer(text=f"{password_confirmed_text}\n\n{new_password_request_text}")
        await state.update_data(del_msg=del_msg)
        await PasswordChangeStates.getting_new_password.set()
    else:
        del_msg = await message.answer(text=f"{wrong_password_text}\n\n{password_request_text}",
                                       parse_mode="MARKDOWN")
        await state.update_data(del_msg=del_msg)


async def get_new_password(message: Message, state: FSMContext):
    """Получает пароль и запрашивает подтверждение смены данных."""
    data = await state.get_data()
    del_msg = data.get("del_msg")
    await del_msg.delete()
    await message.delete()
    new_password = message.text
    del_msg = await message.answer(text=f"{password_change_confirmation_text} {new_password}?",
                                   parse_mode="MARKDOWN",
                                   reply_markup=ikb.confirm_name_kb)
    await state.update_data(new_password=new_password, del_msg=del_msg)
    await PasswordChangeStates.confirming_new_password.set()


async def confirm_password_change(callback_query: CallbackQuery, state: FSMContext):
    """В случае подтверждения смены пароля вызывает функцию смены пароля или отменяет смену пароля."""
    data = await state.get_data()
    del_msg = data.get("del_msg")
    await del_msg.delete()

    if callback_query.data == "yes":
        await change_password(callback_query, state)
    await cq_cancel_all(callback_query, state)


async def change_password(callback_query, state):
    """Меняет пароль в БД."""
    data = await state.get_data()
    new_password = data.get("new_password")
    await db.change_password(callback_query.from_user.id, new_password)
    await callback_query.message.answer(password_changed_text)




def register_password_change(dp):
    dp.register_message_handler(
        send_password_request, ~UserTypeFilter(None),
        content_types=['text'],
        text=['Сменить пароль']
        )
    dp.register_message_handler(
        check_password, 
        state=PasswordChangeStates.checking_password
        )
    dp.register_message_handler(
        get_new_password, 
        state=PasswordChangeStates.getting_new_password
        )
    dp.register_callback_query_handler(
        confirm_password_change, 
        state=PasswordChangeStates.confirming_new_password
        )
