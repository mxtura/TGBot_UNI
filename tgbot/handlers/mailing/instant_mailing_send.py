from aiogram.dispatcher import FSMContext
from aiogram.types import CallbackQuery
from tgbot.models.database_instance import db
from tgbot.handlers.mailing.mailing_text_data import std_groups, std_managers, std_teachers


async def send_mailing(callback_query: CallbackQuery, state: FSMContext):
    bot = callback_query.bot
    state_data = await state.get_data()
    mailing_text = groups = state_data["mailing_text"]
    groups = state_data["groups"]
    managers = state_data["managers"]
    teachers = state_data["teachers"]
    last_name = await db.get_user_ln(callback_query.from_user.id)
    first_name = await db.get_user_fn(callback_query.from_user.id)
    middle_name = await db.get_user_mn(callback_query.from_user.id)
    if middle_name[0] is not None:
        sender = f"{last_name[0]} {first_name[0]} {middle_name[0]}"
    else:
        sender = f"{last_name[0]} {first_name[0]}"

    if groups != std_groups:
        for group in groups:
            student_ids = await db.get_user_ids_by_group(group)
            for student_id in student_ids:
                try:
                    await bot.send_message(student_id[0], f"{mailing_text}\n\n<b>отправитель:</b> {sender}")
                except:
                    print("Упс")
    if managers != std_managers:
        for manager in managers:
            manager_id = await db.get_user_id_by_name(manager.split(" "))
            if manager_id is not None:
                try:
                    await bot.send_message(manager_id[0], f"{mailing_text}\n\n<b>отправитель:</b> {sender}")
                except:
                    print("Упс")
    if teachers != std_teachers:
        for teacher in teachers:
            teacher_id = await db.get_user_id_by_name(teacher.split(" "))
            if teacher_id is not None:
                try:
                    await bot.send_message(teacher_id[0], f"{mailing_text}\n\n<b>отправитель:</b> {sender}")
                except:
                    print("Упс")