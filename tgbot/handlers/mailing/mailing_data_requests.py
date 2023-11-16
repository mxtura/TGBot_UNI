from aiogram.dispatcher import FSMContext
from tgbot.misc.states import MailingsStates
from aiogram.types import Message
from tgbot.models.database_instance import db
from tgbot.handlers.mailing.mailing_text_data import create_group_list, create_manager_list, create_teacher_list


async def request_mailing_text(message: Message, state: FSMContext):
    del_msgs = [await message.answer(text="Введите текст рассылки:")]
    await state.update_data(del_msgs=del_msgs)
    await MailingsStates.mailing_text_input_state.set()


async def request_date(message: Message, state: FSMContext):
    del_msgs = [await message.answer(text="<b>Пожалуйста введите дату отправки в формате: ДД-ММ-ГГГГ</b>")]
    await state.update_data(del_msgs=del_msgs)
    await MailingsStates.date_input_state.set()


async def request_time(message: Message, state: FSMContext):
    del_msgs = [await message.answer(text="<b>Пожалуйста введите время отправки в формате: ЧЧ:ММ</b>")]
    await state.update_data(del_msgs=del_msgs)
    await MailingsStates.time_input_state.set()


async def request_group_selection(message: Message, state: FSMContext):
    state_data = await state.get_data()
    selected_group_list = state_data['groups']
    selected_groups = await create_group_list(selected_group_list)

    all_groups = await db.get_group_names()
    for group in all_groups:
        if group[0] in selected_group_list:
            all_groups.remove(group)

    available_groups = ""
    for group in all_groups:
        available_groups += f"◦ `{group[0]}`;\n"

    # Вынести фразы в отдельные переменные, чтобы улучшить читаемость кода
    del_msgs = [await message.answer(f"*Доступные группы:*\n{available_groups}\n"
                                     f"*Выбранные группы:*\n{selected_groups}\n"
                                     "Отправьте названия групп, которые хотите добавить или удалить через запятую\n\n"
                                     "Чтобы вернуться к форме рассылки отправьте: `Подтвердить`\n", parse_mode="MARKDOWN")]
    await state.update_data(del_msgs=del_msgs)
    await MailingsStates.group_selection_state.set()


async def request_manager_selection(message: Message, state: FSMContext):
    state_data = await state.get_data()
    selected_manager_list = state_data['managers']
    selected_managers = await create_manager_list(selected_manager_list)

    all_managers = await db.get_users_by_type("manager")
    all_manager_list = []
    for manager in all_managers:
        if manager[2] is not None:
            all_manager_list.append(f"{manager[0]} {manager[1]} {manager[2]}")
        else:
            all_manager_list.append(f"{manager[0]} {manager[1]}")

    for manager in all_manager_list:
        if manager in selected_manager_list:
            all_manager_list.remove(manager)

    available_managers = ""
    for manager in all_manager_list:
        available_managers += f"◦ `{manager}`;\n"

    del_msgs = [await message.answer(f"*Менеджеры:*\n{available_managers}\n"
                                     f"*Выбранные менеджеры:*\n{selected_managers}\n"
                                     "Отправьте имена менеджеров, которых хотите добавить или удалить через запятую\n\n"
                                     "Чтобы вернуться к форме рассылки отправьте: `Подтвердить`\n", parse_mode="MARKDOWN")]
    await state.update_data(del_msgs=del_msgs)
    await MailingsStates.manager_selection_state.set()


async def request_teacher_selection(message: Message, state: FSMContext):
    state_data = await state.get_data()
    selected_teacher_list = state_data['teachers']
    selected_teachers = await create_teacher_list(selected_teacher_list)

    all_teachers = await db.get_users_by_type("teacher")
    all_teacher_list = []
    for teacher in all_teachers:
        if teacher[2] is not None:
            all_teacher_list.append(f"{teacher[0]} {teacher[1]} {teacher[2]}")
        else:
            all_teacher_list.append(f"{teacher[0]} {teacher[1]}")
            
    for teacher in all_teacher_list:
        if teacher in selected_teacher_list:
            all_teacher_list.remove(teacher)

    available_teachers = ""
    for teacher in all_teacher_list:
        available_teachers += f"◦ `{teacher}`;\n"

    del_msgs = [await message.answer(f"*Преподаватели:*\n{available_teachers}\n"
                                     f"*Выбранные преподаватели:*\n{selected_teachers}\n"
                                     "Отправьте имена преподавателей, которых хотите добавить или удалить через запятую\n\n"
                                     "Чтобы вернуться к форме рассылки отправьте: `Подтвердить`\n", parse_mode="MARKDOWN")]
    await state.update_data(del_msgs=del_msgs)
    await MailingsStates.teacher_selection_state.set()