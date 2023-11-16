import re

from aiogram.types import Message, CallbackQuery
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.models.db import Database
from tgbot.misc.states import EmployeesDelStates
import tgbot.keyboards.inline as ikb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.models.database_instance import db

user_name_request = "Введите ФИО пользователя:"

user_duplicate_alert = " пользователю "

no_user_selected = "пользователи не выбраны"


# Убрать возможность удалить себя
# Сделать покрасивее, поделить на функции
async def get_employee_names(message: Message, state: FSMContext):
    users = await db.get_employees()
    if len(users) == 1:
        await message.answer(text="Сотрудники еще не добавлены", reply_markup=rkb.manager_keyboard)
    else:
        await state.update_data(user_list=users)
        users_msg = ""
        for user in users:
            if user[2] is not None:
                users_msg += f"◦ `{user[0]} {user[1]} {user[2]}` : {user[3]};\n"
            else:
                users_msg += f"◦ `{user[0]} {user[1]}` : {user[3]};\n"
        msgs_to_del = [await message.answer(text=f"*Сотрудники:*\n{users_msg}\n"
                                                 f"Введите имена сотрудников, которых хотите удалить, через запятую\n\n",
                                            reply_markup=rkb.employees_del_cancel_keyboard, parse_mode="MARKDOWN")]
        await state.update_data(msgs_to_del=msgs_to_del)
        await EmployeesDelStates.confirm_deleting_state.set()


async def confirm_deleting(message: Message, state: FSMContext):
    selected_users_msg = re.sub(r",\s+", ",", message.text)
    selected_users_msg = selected_users_msg.split(",")
    selected_users = []
    await clear_chat(message, state, 1)
    user_list = await db.get_employees()
    users_msg = ""
    for user in selected_users_msg:
        for usr in user_list:
            if usr[2] is not None:
                if user == f"{usr[0]} {usr[1]} {usr[2]}":
                    users_msg += f"◦ `{usr[0]} {usr[1]} {usr[2]}` : {usr[3]};\n"
                    selected_users.append([usr[0], usr[1], usr[2]])
            else:
                if user == f"{usr[0]} {usr[1]}":
                    users_msg += f"◦ `{usr[0]} {usr[1]}` : {usr[3]};\n"
                    selected_users.append([usr[0], usr[1]])
    if len(selected_users) == 0:
        for user in user_list:
            if user[2] is not None:
                users_msg += f"◦ `{user[0]} {user[1]} {user[2]}` : {user[3]};\n"
            else:
                users_msg += f"◦ `{user[0]} {user[1]}` : {user[3]};\n"
        msgs_to_del = [await message.answer(text=f"*Вы не выбрали ни одного сотрудника!*\n\n"
                                                 f"*Все сотрудники:*>\n{users_msg}\n"
                                                 f"Введите имена сотрудников, которых хотите удалить, через запятую\n\n",
                                            parse_mode="MARKDOWN")]
        await state.update_data(msgs_to_del=msgs_to_del)
    else:
        await state.update_data(selected_users=selected_users)
        msgs_to_del = [await message.answer(text=f"*Выбранные сотрудники:* \n{users_msg}\n"
                                                 f"*Вы уверены, что хотите удалить выбранных сотрудников?*",
                                            reply_markup=ikb.confirmation_kb, parse_mode="MARKDOWN")]
        await state.update_data(msgs_to_del=msgs_to_del)
        await EmployeesDelStates.next()


async def delete_employees(callback_query: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    await clear_chat(callback_query.message, state, 0)
    if callback_query.data == "no":
        user_list = await db.get_employees()
        users_msg = ""
        for user in user_list:
            if user[2] is not None:
                users_msg += f"◦ `{user[0]} {user[1]} {user[2]}` : {user[3]};\n"
            else:
                users_msg += f"◦ `{user[0]} {user[1]}` : {user[3]};\n"
        msgs_to_del = [await callback_query.message.answer(text=f"*Все сотрудники:*\n{users_msg}\n"
                                                                f"Введите имена сотрудников, которых хотите удалить, "
                                                                f"через запятую", parse_mode="MARKDOWN")]
        await state.update_data(msgs_to_del=msgs_to_del)
        await EmployeesDelStates.first()
    else:
        users = data.get("selected_users")
        msg = ""
        for user in users:
            if len(user) == 3:
                msg += f"◦ {user[0]} {user[1]} {user[2]};\n"
                await db.delete_user(user)
            else:
                msg += f"◦ {user[0]} {user[1]};\n"
                await db.delete_user_without_mn(user)
        msgs_to_del = [await callback_query.message.answer(f"<b>Следующие сотрудники были удалены:</b>\n{msg}",
                                                           reply_markup=rkb.manager_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
        await state.finish()


async def clear_chat(message: Message, state: FSMContext, del_user_msg):
    data = await state.get_data()
    msgs_to_del = data.get("msgs_to_del")
    for msg in msgs_to_del:
        await msg.delete()
    if del_user_msg:
        await message.delete()


def register_employees_del(dp):
    dp.register_message_handler(get_employee_names, UserTypeFilter("manager"), content_types=['text'],
                                text='Удалить сотрудников')
    dp.register_message_handler(confirm_deleting, state=EmployeesDelStates.confirm_deleting_state)
    dp.register_callback_query_handler(delete_employees, state=EmployeesDelStates.deleting_employees_state)
