import gspread
import datetime
from aiogram.dispatcher import FSMContext
from aiogram.types import Message, CallbackQuery
from tgbot.handlers.mailing.mailing_text_data import std_date, std_time, std_groups, std_managers, std_teachers
from tgbot.models.database_instance import db

google_credentials = gspread.service_account(filename='creds.json')
sheet = google_credentials.open("Бот")
worksheet = sheet.worksheet("Рассылка")


async def record_mailing(callback_query: CallbackQuery, state: FSMContext):
    state_data = await state.get_data()
    text = state_data['mailing_text']
    date = state_data['date']
    time = state_data['time']
    groups = state_data['groups']
    managers = state_data['managers']
    teachers = state_data['teachers']

    if date == std_date:
        today = datetime.date.today()
        date = today.strftime('%d-%m-%Y')

    if time == std_time:
        time = "00:00"

    last_name = await db.get_user_ln(callback_query.from_user.id)
    first_name = await db.get_user_fn(callback_query.from_user.id)
    middle_name = await db.get_user_mn(callback_query.from_user.id)
    if middle_name[0] is not None:
        sender = f"{last_name[0]} {first_name[0]} {middle_name[0]}"
    else:
        sender = f"{last_name[0]} {first_name[0]}"

    values = worksheet.col_values(1)
    non_empty_values = list(filter(None, values))
    first_empty_row = len(non_empty_values) + 1
    worksheet.update_acell("A{}".format(first_empty_row), f"'{date}")
    worksheet.update_acell("B{}".format(first_empty_row), f"'{time}")
    worksheet.update_acell("C{}".format(first_empty_row), text)
    worksheet.update_acell("D{}".format(first_empty_row), sender)
    if groups != std_groups:
        worksheet.update_acell("E{}".format(first_empty_row), ", ".join(groups))
    if managers != std_managers:
        worksheet.update_acell("F{}".format(first_empty_row), ", ".join(managers))
    if teachers != std_teachers:
        worksheet.update_acell("F{}".format(first_empty_row), ", ".join(teachers))