"""Содержит функции вызывающие разные меню при вводе определенного текста или при нажатии на кнопку."""
from aiogram.types import Message
from tgbot.filters.user_type import UserTypeFilter
import tgbot.keyboards.reply as rkb
from .personal_data_editing.personal_data_form import open_personal_data_form


async def open_links_menu(message: Message):
    await message.answer(text="Меню редактирования ссылок", reply_markup=rkb.links_menu_keyboard)


async def open_users_editing_menu(message: Message):
    await message.answer(text="Меню редактирования пользователей", reply_markup=rkb.users_editing_keyboard)


async def open_check_links_menu(message: Message):
    await message.answer(text="Меню проверки ссылок", reply_markup=rkb.check_links_menu_keyboard)


async def open_mailing_menu(message: Message):
    await message.answer(text="Меню рассылки", reply_markup=rkb.mailing_keyboard)


async def open_employees_editing_menu(message: Message):
    await message.answer("Меню редактирования сотрудников", reply_markup=rkb.employees_editing_keyboard)


async def open_students_editing_menu(message: Message):
    await message.answer("<b>Меню редактирования студентов</b>", reply_markup=rkb.students_editing_keyboard)

async def open_students_debts_menu(message: Message):
    await message.answer("<b>Меню долгов</b>", reply_markup=rkb.student_debts_keyboard)
    


def register_menu(dp):
    dp.register_message_handler(open_links_menu, UserTypeFilter("manager"),
                                content_types=['text'], text='Редактирование ссылок')
    dp.register_message_handler(open_users_editing_menu, UserTypeFilter("manager"),
                                content_types=['text'], text='Редактирование пользователей')
    dp.register_message_handler(open_mailing_menu, UserTypeFilter("manager"),
                                content_types=['text'], text=['Рассылка'])
    dp.register_message_handler(open_mailing_menu, UserTypeFilter("teacher"),
                                content_types=['text'], text=['Рассылка'])
    dp.register_message_handler(open_check_links_menu, UserTypeFilter("manager"),
                                content_types=['text'], text='Проверка ссылок')
    dp.register_message_handler(open_personal_data_form, ~UserTypeFilter(None),
                                content_types=['text'], text=['Личная информация'])
    dp.register_message_handler(open_employees_editing_menu, UserTypeFilter("manager"),
                                content_types=['text'], text=['Сотрудники'])
    dp.register_message_handler(open_students_editing_menu, UserTypeFilter("manager"),
                                content_types=['text'], text=['Студенты'])
    dp.register_message_handler(open_students_debts_menu, UserTypeFilter("student"),
                                content_types=['text'], text=['Долги'])