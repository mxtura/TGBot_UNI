from aiogram.types import Message, CallbackQuery
from aiogram.dispatcher import FSMContext
import random
import asyncio

import tgbot.keyboards.reply as rkb
from tgbot.misc.states import EmployeeAddStates
import tgbot.keyboards.inline as ikb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.models.database_instance import db
from tgbot.misc.login_generator import generate_login


async def start_adding(message: Message, state: FSMContext):
    msgs_to_del = [await message.answer(text="Введите ФИО пользователя", reply_markup=rkb.name_input_cancel_keyboard)]
    await state.update_data(msgs_to_del=msgs_to_del)
    await EmployeeAddStates.name_waiting_state.set()


async def check_user_name(message: Message, state: FSMContext):
    words = message.text.split(" ")
    new_name = message.text.strip()
    is_added = 0
    await clear_chat(message, state, 1)
    if len(words) < 2:
        msgs_to_del = [await message.answer(text="<b>ФИО не может состоять из одного слова</b>"),
                       await message.answer(text="Введите ФИО пользователя",
                                            reply_markup=rkb.name_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
    elif len(words) > 3:
        msgs_to_del = [await message.answer(text="<b>ФИО не должно содержать более трех слов</b>"),
                       await message.answer(text="Введите ФИО пользователя",
                                            reply_markup=rkb.name_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
    else:
        names = await db.get_user_names()
        for name in names:
            if f"{name[0]} {name[1]} {name[2]}".strip() == new_name:
                msgs_to_del = [await message.answer(text="<b>Пользователь с таким именем уже существует!</b>\n\n"),
                               await message.answer(text="Введите ФИО пользователя\n\n",
                                                    reply_markup=rkb.name_input_cancel_keyboard)]
                await state.update_data(msgs_to_del=msgs_to_del)
                is_added = 1
        if not is_added:
            await state.update_data(new_user_name=new_name)
            msgs_to_del = [await message.answer(text=f"ФИО пользователя: <b>{new_name}</b>"),
                           await message.answer(text="Выберите тип пользователя", reply_markup=ikb.user_type_keyboard)]
            await state.update_data(msgs_to_del=msgs_to_del)
            await EmployeeAddStates.user_type_selection_state.set()


async def send_add_confirmation(callback_query: CallbackQuery, state: FSMContext):
    new_user_type = callback_query.data
    await callback_query.answer()
    await clear_chat(callback_query.message, state, 0)
    data = await state.get_data()
    new_user_name = data.get("new_user_name")
    await state.update_data(new_user_type=new_user_type)
    if new_user_type == "teacher":
        msgs_to_del = [await callback_query.message.answer(
            f"Вы уверены, что хотите добавить нового преподавателя:\n<b>{new_user_name}</b>?",
            reply_markup=ikb.confirmation_kb)]
    else:
        msgs_to_del = [await callback_query.message.answer(
            f"Вы уверены, что хотите добавить нового менеджера:\n<b>{new_user_name}</b>?",
            reply_markup=ikb.confirmation_kb)]
    await state.update_data(msgs_to_del=msgs_to_del)
    await EmployeeAddStates.new_user_confirmation_state.set()


async def add_user(callback_query: CallbackQuery, state: FSMContext):
    await callback_query.answer()
    await clear_chat(callback_query.message, state, 0)
    if callback_query.data == "yes":
        data = await state.get_data()
        new_user_name = data.get("new_user_name")
        new_user_type = data.get("new_user_type")
        fio = new_user_name.split()
        user = []
        for name in fio:
            user.append(name)
        if len(user) < 3:
            user.append(None)
        login = await generate_login(fio, new_user_type)
        password = random.randint(100000, 999999)
        await db.add_user(list(fio), login, password, new_user_type)
        user_data = f"<b>ФИО:</b> {new_user_name}\n<b>Логин:</b> {login}\n<b>Пароль:</b> {password}\n"
        await callback_query.message.answer(text="Пользователь добавлен")
        personal_data = await callback_query.message.answer(text=f"Данные пользователя:\n\n{user_data}\n\n"
                                                                 f"<b>Это сообщение будет удалено через 10 минут, "
                                                                 f"поэтому рекомендуется записать эти данные</b>")
        asyncio.create_task(del_msg(personal_data, 600))
        await callback_query.message.answer(text="Меню менеджера", reply_markup=rkb.manager_keyboard)
        await state.finish()
    else:
        msgs_to_del = [await callback_query.message.answer(text="Введите ФИО пользователя",
                                                           reply_markup=rkb.name_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
        await EmployeeAddStates.name_waiting_state.set()


async def del_msg(message, time):
    await asyncio.sleep(time)
    await message.delete()


async def clear_chat(message: Message, state: FSMContext, del_user_msg):
    data = await state.get_data()
    msgs_to_del = data.get("msgs_to_del")
    for msg in msgs_to_del:
        await msg.delete()
    if del_user_msg:
        await message.delete()


def register_employee_add(dp):
    dp.register_message_handler(start_adding, UserTypeFilter("manager"), content_types=['text'],
                                text='Добавить сотрудника')
    dp.register_message_handler(check_user_name, state=EmployeeAddStates.name_waiting_state)
    dp.register_callback_query_handler(send_add_confirmation, state=EmployeeAddStates.user_type_selection_state)
    dp.register_callback_query_handler(add_user, state=EmployeeAddStates.new_user_confirmation_state)
