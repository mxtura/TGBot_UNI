from aiogram.dispatcher import FSMContext


std_text = "Текст не введен"
std_groups = ["Группы не выбраны"]
std_managers = ["Менеджеры не выбраны"]
std_teachers = ["Преподаватели не выбраны"]
std_date = "Дата не выбрана"
std_time = "Время не выбрано"

mailing_recorded_msg = "Рассылка записана в таблицу рассылок и будет отправлена в указанное время"
mailing_send_msg = "Рассылка отправлена"


# Три одинаковые функции (нужно объединить в одну)
async def create_group_list(groups):
    group_list = ""
    if groups != std_groups:
        for group in groups:
            group_list += f"◦ `{group}`;\n"
    else:
        group_list = f"{groups[0]}\n"
    return group_list


async def create_manager_list(managers):
    manager_list = ""
    if managers != std_managers:
        for manager in managers:
            manager_list += f"◦ `{manager}`;\n"
    else:
        manager_list += f"{managers[0]}\n"
    return manager_list


async def create_teacher_list(teachers):
    teacher_list = ""
    if teachers != std_teachers:
        for teacher in teachers:
            teacher_list += f"◦ `{teacher}`;\n"
    else:
        teacher_list += f"{teachers[0]}\n"
    return teacher_list


async def create_date_and_time_variable(state: FSMContext):
    state_data = await state.get_data()
    date = state_data['date']
    time = state_data['time']
    if date == std_date and time == std_time:
        mailing_date_and_time = "Сообщение будет отправленно мгновенно"
    elif date == std_date:
        mailing_date_and_time = f"Сообщение будет отправлено сегодня в {time}"
    elif time == std_time:
        mailing_date_and_time = f"Сообщение будет отправлено {date} в 00:00"
    else:
        mailing_date_and_time = f"Сообщение будет отправлено {date} в {time}"
    return mailing_date_and_time