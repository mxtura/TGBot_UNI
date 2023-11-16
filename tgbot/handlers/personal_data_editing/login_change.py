from aiogram.types import Message, CallbackQuery
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.inline as ikb
import tgbot.keyboards.reply as rkb
from tgbot.misc.states import LoginChangeStates
from tgbot.filters.user_type import UserTypeFilter
from tgbot.handlers.cancel import cq_cancel_all
from tgbot.models.database_instance import db

login_request_text = "Пожалуйста введите новый логин:"
login_repeated_text = "*Логин занят!*"
login_change_confirmation_text = "Вы уверены, что хотите сменить логин на:"
login_changed_text = "Логин изменен"

async def send_login_request(message: Message, state: FSMContext):
    """Просит ввести новый логин."""
    del_msg = await message.answer(text=login_request_text,
                                   reply_markup=rkb.login_input_cancel_keyboard)
    await state.update_data(del_msg=del_msg)
    await LoginChangeStates.getting_new_login.set()


async def get_new_login(message: Message, state: FSMContext):
    """Получает логин и проверяет его наличие в БД."""
    data = await state.get_data()
    del_msg = data.get("del_msg")
    await del_msg.delete()
    await message.delete()

    new_login = message.text
    login_list = await db.get_logins()

    is_added = 0
    for login in login_list:
        if new_login == login[0]:
            is_added = 1
            break

    if is_added:
        del_msg = await message.answer(text=f"{login_repeated_text}\n\n{login_request_text}",
                                       parse_mode="MARKDOWN")
        await state.update_data(del_msg=del_msg)
    else:
        del_msg = await message.answer(text=f"{login_change_confirmation_text} {new_login}?",
                                       reply_markup=ikb.confirm_name_kb)
        await state.update_data(new_login=new_login, del_msg=del_msg)
        await LoginChangeStates.confirming_login_change.set()


async def confirm_login_change(callback_query: CallbackQuery, state: FSMContext):
    """Подтверждает смену логина."""
    data = await state.get_data()
    del_msg = data.get("del_msg")
    await del_msg.delete()

    if callback_query.data == "yes":
        await change_login(callback_query, state)
    await cq_cancel_all(callback_query, state)


async def change_login(callback_query, state):
    """Меняет логин в БД."""
    data = await state.get_data()
    new_login = data.get("new_login")
    await db.change_login(callback_query.from_user.id, new_login)
    await callback_query.message.answer(login_changed_text)


def register_login_change(dp):
    dp.register_message_handler(send_login_request, ~UserTypeFilter(None), content_types=['text'],
                                text=['Сменить логин'])
    dp.register_message_handler(get_new_login, state=LoginChangeStates.getting_new_login)
    dp.register_callback_query_handler(confirm_login_change, state=LoginChangeStates.confirming_login_change)
