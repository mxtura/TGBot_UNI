import asyncio

from aiogram.types import Message
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.misc.states import ChangeUserPasswordStates
from tgbot.filters.user_type import UserTypeFilter
from tgbot.models.database_instance import db


# Разобраться в последнюю очередь

async def start_change(message: Message, state: FSMContext):
    bot = message.bot
    chat_id = message.from_user.id
    if message.text == 'Сменить пароль сотрудника':
        manager_list = await db.get_managers()
        teacher_list = await db.get_teachers()
        manager_list_msg = ""
        teacher_list_msg = ""
        for manager in manager_list:
            if manager[2] is not None:
                manager_list_msg += f"◦ `{manager[0]} {manager[1]} {manager[2]}` : {manager[3]};\n"
            else:
                manager_list_msg += f"◦ `{manager[0]} {manager[1]}` : {manager[3]};\n"
        for teacher in teacher_list:
            if teacher[2] is not None:
                teacher_list_msg += f"◦ `{teacher[0]} {teacher[1]} {teacher[2]}` : {teacher[3]};\n"
            else:
                teacher_list_msg += f"◦ `{teacher[0]} {teacher[1]}` : {teacher[3]};\n"
        msgs_to_del = [await message.answer(f"*Смена пароля сотрудника*\n\n"
                                            f"*Менеджеры:*\n{manager_list_msg}\n"
                                            f"*Преподаватели:*\n{teacher_list_msg}",
                                            reply_markup=rkb.name_input_cancel_keyboard, parse_mode="MARKDOWN"),
                       await bot.send_message(text="<b>Пожалуйста введите имя сотрудника, которому хотите сменить "
                                                   "пароль</b>", chat_id=chat_id)]
        await state.update_data(msgs_to_del=msgs_to_del)
        await ChangeUserPasswordStates.select_user_to_change_password_state.set()
    elif message.text == 'Сменить пароль студента':
        group_list = await db.get_group_names()
        group_msg = ""
        for gr in group_list:
            group_msg += f"◦ `{gr[0]}`;\n"
        msgs_to_del = [await bot.send_message(text=f"*Смена пароля студента*\n\n"
                                                   f"*Список групп:*\n{group_msg}",
                                              chat_id=chat_id, reply_markup=rkb.group_name_cancel_keyboard,
                                              parse_mode="MARKDOWN"),
                       await bot.send_message(text="<b>Пожалуйста введите название группы студента, "
                                                   "которому хотите сменить пароль</b>", chat_id=chat_id)]
        await state.update_data(msgs_to_del=msgs_to_del)
        await ChangeUserPasswordStates.select_group_to_change_password_state.set()


async def select_user(message: Message, state: FSMContext):  # Сделать проверку на корректность ввода имени
    selected_user = message.text
    await clear_chat(message, state, 1)
    await state.update_data(selected_user=selected_user)
    msgs_to_del = [await message.answer(f"<b>Введите новый пароль для пользователя</b>\n\n{selected_user}",
                                        reply_markup=rkb.password_input_cancel_keyboard)]
    await state.update_data(msgs_to_del=msgs_to_del)
    await ChangeUserPasswordStates.get_new_user_password_state.set()


# меняет пароли только для зарегистрированных пользователей
async def change_password(message: Message, state: FSMContext):
    await clear_chat(message, state, 1)
    data = await state.get_data()
    selected_user = data.get("selected_user")
    user_id = await db.get_user_id_by_fio(selected_user.split())
    await db.change_password(user_id[0], message.text)
    msg = await message.answer(text=f"<b>Пароль пользователя</b>\n{selected_user}\n<b>успешно изменен</b>\n\n"
                                    f"<b>Новый пароль:</b> {message.text}", reply_markup=rkb.manager_keyboard)
    asyncio.create_task(del_msg(msg, 600))
    await state.finish()


async def del_msg(message, time):
    await asyncio.sleep(time)
    await message.delete()


async def select_group(message: Message, state: FSMContext):  # Проверка бла бла бла...
    bot = message.bot
    chat_id = message.from_user.id
    selected_group = message.text
    await state.update_data(selected_group=selected_group)
    await clear_chat(message, state, 1)
    students = await db.get_students_by_group_name(selected_group)
    students_msg = ""
    for student in students:  # Переделать
        students_msg += f"◦ {student[0]} {student[1]}"
        if student[2] is not None:
            students_msg += " " + student[2]
        students_msg += f" : {student[3]};\n"
    msgs_to_del = [await message.answer(f"<b>Студенты группы</b> {selected_group}:\n{students_msg}"),
                   await bot.send_message(text="<b>Пожалуйста введите имя студента из списка, которому хотите сменить "
                                               "пароль</b>", chat_id=chat_id,
                                          reply_markup=rkb.name_input_cancel_keyboard)]
    await state.update_data(msgs_to_del=msgs_to_del)
    await ChangeUserPasswordStates.select_user_to_change_password_state.set()


async def clear_chat(message: Message, state: FSMContext, del_user_msg):
    data = await state.get_data()
    msgs_to_del = data.get("msgs_to_del")
    for msg in msgs_to_del:
        await msg.delete()
    if del_user_msg:
        await message.delete()


def register_change_user_password(dp):
    dp.register_message_handler(start_change, UserTypeFilter("manager"), content_types=['text'],
                                text='Сменить пароль сотрудника')
    dp.register_message_handler(start_change, UserTypeFilter("manager"), content_types=['text'],
                                text='Сменить пароль студента')
    dp.register_message_handler(select_user, state=ChangeUserPasswordStates.select_user_to_change_password_state)
    dp.register_message_handler(select_group, state=ChangeUserPasswordStates.select_group_to_change_password_state)
    dp.register_message_handler(change_password, state=ChangeUserPasswordStates.get_new_user_password_state)
