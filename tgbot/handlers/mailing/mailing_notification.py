from aiogram.dispatcher import FSMContext
from aiogram.types import Message
from tgbot.handlers.mailing.mailing_text_data import std_groups, std_managers, std_teachers
from tgbot.handlers.mailing.mailing_text_data import create_date_and_time_variable
from tgbot.handlers.mailing.mailing_text_data import create_group_list, create_manager_list, create_teacher_list


async def send_mailing_notification(message: Message, state: FSMContext):
    state_data = await state.get_data()
    date_and_time = await create_date_and_time_variable(state)
    text = state_data['mailing_text']
    groups = state_data['groups']
    managers = state_data['managers']
    teachers = state_data['teachers']
    notification_text = ""
    if groups == std_groups and managers == std_managers and teachers == std_teachers:
        notification_text = f"*Получатели не выбраны!*\nСоздание рассылки отменено"
    else:
        notification_text += f"*Текст рассылки:*\n{text}\n\n"
        if groups != std_groups:
            notification_text += f"*Группы получатели:*\n{await create_group_list(groups)}\n"
        if managers != std_managers:
            notification_text += f"*Менеджеры получатели:*\n{await create_manager_list(managers)}\n"
        if teachers != std_teachers:
            notification_text += f"*Преподаватели получатели:*\n{await create_teacher_list(teachers)}\n"
        notification_text += f"*{date_and_time}*"
    await message.answer(notification_text, parse_mode="MARKDOWN")