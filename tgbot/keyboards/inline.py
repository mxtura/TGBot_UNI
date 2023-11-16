"""Содержит все Inline клавиатуры."""
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
from tgbot.models.database_instance import db

confirmation_kb = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(text='Да', callback_data='yes'),
            InlineKeyboardButton(text='Нет', callback_data='no')
        ]
    ]
)


user_type_keyboard = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(text='Менеджер', callback_data='manager'),
            InlineKeyboardButton(text='Преподаватель', callback_data='teacher')
        ]
    ]
)

confirm_name_kb = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(text='Да', callback_data='yes'),
        ],
        [
            InlineKeyboardButton(text='Нет', callback_data='no')
        ]
    ]
)


teacher_mailing_keyboard = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(text='Ввести сообщение', callback_data='mailing_text_input'),
            InlineKeyboardButton(text='Выбрать группы', callback_data='groups_selection'),
        ],
        [
            InlineKeyboardButton(text='Ввести дату', callback_data='date_input'),
            InlineKeyboardButton(text='Ввести время', callback_data='time_input')
        ],
        [
            InlineKeyboardButton(text='Подтвердить и отправить', callback_data='confirmation_of_mailing'),
            InlineKeyboardButton(text='Отмена', callback_data='cancellation_of_mailing_form')
        ]
    ]
)

manager_mailing_keyboard = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(text='Ввести сообщение', callback_data='mailing_text_input'),
            InlineKeyboardButton(text='Выбрать группы', callback_data='groups_selection'),
        ],
        [
            InlineKeyboardButton(text='Ввести дату', callback_data='date_input'),
            InlineKeyboardButton(text='Выбрать менеджеров', callback_data='managers_selection'),
        ],
        [
            InlineKeyboardButton(text='Ввести время', callback_data='time_input'),
            InlineKeyboardButton(text='Выбрать преподавателей', callback_data='teachers_selection'),
        ],
        [
            InlineKeyboardButton(text='Подтвердить и отправить', callback_data='confirmation_of_mailing'),
            InlineKeyboardButton(text='Отмена', callback_data='cancellation_of_mailing_form')
        ]
    ]
)


group_edit_keyboard = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(text="Сменить название группы", callback_data='change_group_name'),
            InlineKeyboardButton(text="Сменить таблицу успеваемости", callback_data='change_performance_list')
        ],
        [
            InlineKeyboardButton(text="Список студентов", callback_data='display_student_list'),
            InlineKeyboardButton(text="Обновить студентов", callback_data='update_student_list')
        ],
        [
            InlineKeyboardButton(text="Отмена", callback_data='cancel_group_editing')
        ]
    ]
)

student_list_keyboard = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(text="Назад", callback_data='back'),
            InlineKeyboardButton(text="Отмена", callback_data='cancel'),
        ]
    ]
)