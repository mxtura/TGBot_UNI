import re
from aiogram.dispatcher import FSMContext
from tgbot.misc.states import MailingsStates
from aiogram.types import Message
from tgbot.misc.decorators.log_decorator import log_function_call

from tgbot.models.database_instance import db
from tgbot.handlers.mailing.mailing_text_data import std_groups, std_managers, std_teachers
from tgbot.handlers.mailing.mailing_data_requests import request_group_selection, request_manager_selection, request_teacher_selection
from tgbot.handlers.mailing.manager_mailing_form import create_manager_mailing_form
from tgbot.handlers.mailing.teacher_mailing_form import create_teacher_mailing_form


async def get_mailing_text(message: Message, state: FSMContext):
    await message.delete()
    state_data = await state.get_data()
    user_type = state_data['user_type']
    del_msgs = state_data["del_msgs"]
    for msg in del_msgs:
        await msg.delete()
    await state.update_data(mailing_text=message.text, del_msgs=[])
    if user_type == "manager":
        await create_manager_mailing_form(message, state)
    else:
        await create_teacher_mailing_form(message, state)


# Сюда внести правки, для большей свободы ввода даты и времени
# "Введите 'Подтвердить', чтобы отправить сообщение мгновенно"
async def get_mailing_date(message: Message, state: FSMContext):
    await message.delete()
    state_data = await state.get_data()
    del_msgs = state_data["del_msgs"]
    for msg in del_msgs:
        await msg.delete()
    date_length = len(message.text)
    if date_length != 10 or message.text[2] != "-" or message.text[5] != "-" or (message.text[:2]).isdigit() is False or (
        message.text[3:5]).isdigit() is False or (message.text[7:]).isdigit() is False:
        del_msgs = [await message.answer(text="<b>Дата введена некорректно!</b>\n\n"
                                         "Пожалуйста введите дату отправки в формате: <b>ДД-ММ-ГГГГ</b>")]
        await state.update_data(del_msgs=del_msgs)
    else:
        await state.update_data(date=message.text, del_msgs=[])
        user_type = state_data['user_type']
        if user_type == "manager":
            await create_manager_mailing_form(message, state)
        else:
            await create_teacher_mailing_form(message, state)


async def get_mailing_time(message: Message, state: FSMContext):
    await message.delete()
    state_data = await state.get_data()
    del_msgs = state_data["del_msgs"]
    for msg in del_msgs:
        await msg.delete()
    time_len = len(message.text)
    if time_len != 5 or message.text[2] != ":" or (message.text[:2]).isdigit() is False or (
            message.text[3:5]).isdigit() is False:
        del_msgs = [await message.answer(text="<b>Время введено некорректно!</b>\n\n"
                                             "Пожалуйста введите время отправки в формате: <b>ЧЧ:ММ</b>")]
        await state.update_data(del_msgs=del_msgs)
    else:
        await state.update_data(time=message.text, del_msgs=[])
        user_type = state_data['user_type']
        if user_type == "manager":
            await create_manager_mailing_form(message, state)
        else:
            await create_teacher_mailing_form(message, state)


# Добавить возможность очищать группы
async def get_mailing_groups(message: Message, state: FSMContext):
    await message.delete()
    state_data = await state.get_data()
    user_type = state_data['user_type']
    del_msgs = state_data["del_msgs"]
    for msg in del_msgs:
        await msg.delete()
    if message.text == 'Подтвердить':
        await state.update_data(del_msgs=[])
        if user_type == "manager":
            await create_manager_mailing_form(message, state)
        elif user_type == "teacher":
            await create_teacher_mailing_form(message, state)
    else:
        new_groups = re.sub(r",\s+", ",", message.text)
        new_groups = new_groups.split(",")
        selected_groups = state_data.get("groups")
        # Возможно стоит записать all_groups в состояние
        all_groups = await db.get_group_names()

        all_group_list = []
        for group in all_groups:
            all_group_list.append(group[0])

        for new_group in new_groups:
            if new_group in selected_groups:
                selected_groups.remove(new_group)
            else:
                if new_group in all_group_list:
                    if std_groups == selected_groups:
                        selected_groups.remove(std_groups[0])
                    selected_groups.append(new_group)

        if len(selected_groups) == 0:
            selected_groups.append(std_groups[0])

        await state.update_data(groups=selected_groups)
        await request_group_selection(message, state)


async def get_mailing_managers(message: Message, state: FSMContext):
    await message.delete()
    state_data = await state.get_data()
    user_type = state_data.get("user_type")
    del_msgs = state_data["del_msgs"]
    for msg in del_msgs:
        await msg.delete()
    if message.text == "Подтвердить":
        await state.update_data(del_msgs=[])
        if user_type == "manager":
            await create_manager_mailing_form(message, state)
        elif user_type == "teacher":
            await create_teacher_mailing_form(message, state)
    else:
        new_managers = re.sub(r",\s+", ",", message.text)
        new_managers = new_managers.split(",")
        selected_managers = state_data.get("managers")

        all_managers = await db.get_users_by_type("manager")
        all_manager_list = []
        for manager in all_managers:
            if manager[2] is not None:
                all_manager_list.append(f"{manager[0]} {manager[1]} {manager[2]}")
            else:
                all_manager_list.append(f"{manager[0]} {manager[1]}")
        for new_manager in new_managers:
            if new_manager in selected_managers:
                selected_managers.remove(new_manager)
            else:
                if new_manager in all_manager_list:
                    if std_managers == selected_managers:
                        selected_managers.remove(std_managers[0])
                    selected_managers.append(new_manager)

        if len(selected_managers) == 0:
            selected_managers.append(std_managers[0])

        await state.update_data(managers=selected_managers)
        await request_manager_selection(message, state)


async def get_mailing_teachers(message: Message, state: FSMContext):
    await message.delete()
    state_data = await state.get_data()
    user_type = state_data.get("user_type")
    del_msgs = state_data["del_msgs"]
    for msg in del_msgs:
        await msg.delete()
    if message.text == "Подтвердить":
        await state.update_data(del_msgs=[])
        if user_type == "manager":
            await create_teacher_mailing_form(message, state)
        elif user_type == "teacher":
            await create_teacher_mailing_form(message, state)
    else:
        new_teachers = re.sub(r",\s+", ",", message.text)
        new_teachers = new_teachers.split(",")
        selected_teachers = state_data.get("teachers")

        all_teachers = await db.get_users_by_type("teacher")
        all_teacher_list = []
        for teacher in all_teachers:
            if teacher[2] is not None:
                all_teacher_list.append(f"{teacher[0]} {teacher[1]} {teacher[2]}")
            else:
                all_teacher_list.append(f"{teacher[0]} {teacher[1]}")

        for new_teacher in new_teachers:
            if new_teacher in selected_teachers:
                selected_teachers.remove(new_teacher)
            else:
                if new_teacher in all_teacher_list:
                    if std_teachers == selected_teachers:
                        selected_teachers.remove(std_teachers[0])
                    selected_teachers.append(new_teacher)

        if len(selected_teachers) == 0:
            selected_teachers.append(std_teachers[0])

        await state.update_data(teachers=selected_teachers)
        await request_teacher_selection(message, state)


def register_mailing_data_processing(dp):
    dp.register_message_handler(get_mailing_text, state=MailingsStates.mailing_text_input_state)
    dp.register_message_handler(get_mailing_date, state=MailingsStates.date_input_state)
    dp.register_message_handler(get_mailing_time, state=MailingsStates.time_input_state)
    dp.register_message_handler(get_mailing_groups, state=MailingsStates.group_selection_state)
    dp.register_message_handler(get_mailing_managers, state=MailingsStates.manager_selection_state)
    dp.register_message_handler(get_mailing_teachers, state=MailingsStates.teacher_selection_state)