import gspread
import datetime
from tgbot.models.database_instance import db

google_credentials = gspread.service_account(filename='creds.json')
sheet = google_credentials.open("Бот").worksheet("Рассылка")
offset = datetime.timezone(datetime.timedelta(hours=3))


# Добавить возможность вводить дату и время более свободно
async def check_msgs_list(bot):
    date_col_num = sheet.find("Дата отправки").col
    dates = sheet.col_values(date_col_num)

    now = datetime.datetime.now(offset)
    current_date = now.strftime('%d-%m-%Y')
    current_time = now.strftime('%H:%M')

    for i, date_str in enumerate(dates):
        if i == 0:
            continue
        if date_str == current_date:
            time_str = sheet.cell(i+1, date_col_num + 1).value
            if time_str == current_time:
                message = sheet.cell(i+1, date_col_num + 2).value
                sender = sheet.cell(i+1, date_col_num + 3).value
                groups = sheet.cell(i+1, date_col_num + 4).value
                managers = sheet.cell(i+1, date_col_num + 5).value
                teachers = sheet.cell(i+1, date_col_num + 6).value
                await send_mailing(bot, message, sender, groups, managers, teachers)


# Что делать если пользователь удалил телеграм аккаунт? Подумать в последнюю очередь!
# Проверить работоспособность
async def send_mailing(bot, message, sender, groups, managers, teachers):
    if groups is not None:
        groups = groups.split(", ")
        for group in groups:
            student_ids = await db.get_user_ids_by_group(group)
            for student_id in student_ids:
                try:
                    await bot.send_message(student_id[0], f"{message}\n\n<b>отправитель:</b> {sender}")
                except:
                    print("Упс")
    if managers is not None:
        managers = managers.split(", ")
        for manager in managers:
            manager_id = await db.get_user_id_by_name(manager.split(" "))
            if manager_id is not None:
                try:
                    await bot.send_message(manager_id[0], f"{message}\n\n<b>отправитель:</b> {sender}")
                except:
                    print("Упс")
    if teachers is not None:
        teachers = teachers.split(", ")
        for teacher in teachers:
            teacher_id = await db.get_user_id_by_name(teacher.split(" "))
            if teacher_id is not None:
                try:
                    await bot.send_message(teacher_id[0], f"{message}\n\n<b>отправитель:</b> {sender}")
                except:
                    print("Упс")