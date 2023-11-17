"""Отправляет все ссылки, запрашиваемые в одно действие."""
from aiogram.types import Message
from tgbot.filters.user_type import UserTypeFilter
from tgbot.models.database_instance import db


async def send_schedule(message: Message):
    """Отправляет расписание."""
    schedule_url = await db.get_schedule_url()
    if not schedule_url:
        await message.answer(text="<b>Ссылка на расписание еще не загружена менеджером!</b>")
    else:
        await message.answer(text=schedule_url)


async def send_work_schedule(message: Message):
    """Отправляет график работы."""
    work_schedule_url = await db.get_work_schedule_url()
    if not work_schedule_url:
        await message.answer(text="<b>Ссылка на график работы еще не загружена менеджером!</b>")
    else:
        await message.answer(text=work_schedule_url)


async def send_learning_schedule(message: Message):
    """Отправляет график учебы."""
    learning_schedule_url = await db.get_learning_schedule_url()
    if not learning_schedule_url:
        await message.answer(text="<b>Ссылка на график учебы еще не загружена менеджером!</b>")
    else:
        await message.answer(text=learning_schedule_url)


async def send_mailings_table(message: Message):
    """Отправляет таблицу рассылок"""
    mailings_table_url = await db.get_mailings_url()
    if not mailings_table_url:
        await message.answer(text="<b>Ссылка на таблицу рассылок еще не загружена менеджером!</b>")
    else:
        await message.answer(text=mailings_table_url)


async def send_reports(message: Message):
    """Отправляет ссылку на ведомости."""
    reports_url = await db.get_report_cards_url()
    if not reports_url:
        await message.answer(text="<b>Ссылка на ведомости еще не загружена менеджером!</b>")
    else:
        await message.answer(text=reports_url[0])


async def send_retakes(message: Message):
    """Отправляет ссылку на ведомости пересдач."""
    retakes_url = await db.get_retake_cards_url()
    if not retakes_url:
        await message.answer(text="<b>Ссылка на ведомости пересдач еще не загружена менеджером!</b>")
    else:
        await message.answer(text=retakes_url[0])


async def send_problem_reporting_email(message: Message):
    """Отправляет почту для жалоб."""
    problem_reporting_email = await db.get_problem_reporting_email()
    if not problem_reporting_email:
        await message.answer(text="<b>Почта поддержки еще не загружена менеджером!</b>")
    else:
        await message.answer(f"Почта поддержки: {problem_reporting_email[0]}\n\n"
                             "Пожалуйста пишите сюда, в случае обнаружения проблем в работе бота.")


def register_requested_data_sending(dp):
    dp.register_message_handler(send_schedule, ~UserTypeFilter(None), 
                                content_types=['text'], text='Расписание')
    dp.register_message_handler(send_work_schedule, UserTypeFilter("teacher"), 
                                content_types=['text'], text='График работы')
    dp.register_message_handler(send_work_schedule, UserTypeFilter("manager"), 
                                content_types=['text'], text='График работы')
    dp.register_message_handler(send_learning_schedule, UserTypeFilter("student"), 
                                content_types=['text'], text='График учебы')
    dp.register_message_handler(send_learning_schedule, UserTypeFilter("manager"), 
                                content_types=['text'], text='График учебы')
    dp.register_message_handler(send_mailings_table, UserTypeFilter("manager"), 
                                content_types=['text'], text=['Таблица рассылок'])
    dp.register_message_handler(send_mailings_table, UserTypeFilter("teacher"), 
                                content_types=['text'], text=['Таблица рассылок'])
    dp.register_message_handler(send_reports, UserTypeFilter("manager"), 
                                content_types=['text'], text=['Получить ведомости'])
    dp.register_message_handler(send_reports, UserTypeFilter("teacher"), 
                                content_types=['text'], text=['Получить ведомости'])
    dp.register_message_handler(send_retakes, UserTypeFilter("manager"), 
                                content_types=['text'], text='Получить пересдачи')
    dp.register_message_handler(send_retakes, UserTypeFilter("teacher"), 
                                content_types=['text'], text='Получить ссылку на таблицу')
    dp.register_message_handler(send_problem_reporting_email, ~UserTypeFilter(None), 
                                content_types=['text'], text='Сообщить о проблеме')
