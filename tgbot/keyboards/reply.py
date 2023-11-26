"""Содержит все Reply клавиатуры."""
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove

action_request = "Выберите что хотите сделать..."

student_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Успеваемость"),
            KeyboardButton(text="Долги")
        ],
        [
            KeyboardButton(text="Расписание"),
            KeyboardButton(text="График учебы")
        ],
        [
            KeyboardButton(text="Личная информация"),
            KeyboardButton(text="Сообщить о проблеме")
        ]
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

teacher_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Расписание"),
            KeyboardButton(text="График работы")
        ],
        [
            KeyboardButton(text="Получить ведомости"),
            KeyboardButton(text="Пересдачи")
        ],
        [
            KeyboardButton(text="Рассылка"),
            KeyboardButton(text="Личная информация")
        ],
        [
            KeyboardButton(text="Сообщить о проблеме")
        ]
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

manager_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Редактирование пользователей"),
            KeyboardButton(text="Редактирование ссылок")
        ],
        [
            KeyboardButton(text="Рассылка"),
            KeyboardButton(text="Проверка ссылок")
        ],
        [
            KeyboardButton(text="Личная информация"),
            KeyboardButton(text="Сообщить о проблеме")
        ]
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

users_editing_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Сотрудники"),
            KeyboardButton(text="Студенты")
        ],
        [
            KeyboardButton(text="Отмена")
        ]
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

links_menu_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Заменить расписание"),
            KeyboardButton(text="Заменить график работы"),
            KeyboardButton(text="Заменить график учебы")
        ],
        [
            KeyboardButton(text="Заменить ведомости"),
            KeyboardButton(text="Заменить пересдачи"),
            KeyboardButton(text="Заменить таблицу рассылок")
        ],
        [
            KeyboardButton(text="Заменить почту поддержки"),
            KeyboardButton(text="Отмена")
        ]
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

check_links_menu_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Получить ведомости"),
            KeyboardButton(text="Получить пересдачи"),
        ],
        [
            KeyboardButton(text="График работы"),
            KeyboardButton(text="График учебы")
        ],
        [
            KeyboardButton(text="Расписание"),
            KeyboardButton(text="Отмена")
        ]
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

mailing_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Создать рассылку"),
            KeyboardButton(text="Таблица рассылок")
        ],
        [
            KeyboardButton(text="Отмена")
        ],
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

personal_data_editing_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Сменить логин"),
            KeyboardButton(text="Сменить пароль")
        ],
        [
            KeyboardButton(text="Отмена"),
        ],
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

employees_editing_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Добавить сотрудника"),
            KeyboardButton(text="Удалить сотрудников")
        ],
        [
            KeyboardButton(text="Сменить пароль сотрудника"),
            KeyboardButton(text="Отмена")
        ]
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

students_editing_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Добавить группу"),
            KeyboardButton(text="Удалить группы"),
            KeyboardButton(text="Редактировать группу")
        ],
        [
            KeyboardButton(text="Сменить пароль студента"),
            KeyboardButton(text="Отмена")
        ]
    ],
    resize_keyboard=True
)

back_save_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Сохранить "),
            KeyboardButton(text="Назад")
        ]
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder="группа1, группа2..."
)

student_debts_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Посмотреть долги"),
            KeyboardButton(text="Записаться на пересдачу")
        ],
        [
            KeyboardButton(text="Отмена"),
        ],
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

teacher_retakes_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Получить ссылку на таблицу"),
            KeyboardButton(text="Получить список пересдач")
        ],
        [
            KeyboardButton(text="Отмена"),
        ],
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

confirmation_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text="Да"),
            KeyboardButton(text="Нет")
        ],
    ],
    resize_keyboard=True, one_time_keyboard=False, input_field_placeholder=action_request
)

# group_change_keyboard = ReplyKeyboardMarkup(
#     keyboard=[
#         [
#             KeyboardButton(text="Сменить название группы"),
#             KeyboardButton(text="Сменить таблицу успеваемости")
#         ],
#         [
#             KeyboardButton(text="Отмена")
#         ]
#     ],
#     resize_keyboard=True
# )

# Переименовать их более корректно

login_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="Логин"
)

password_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="Пароль"
)

name_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="Иванов Иван Иванович"
)


url_change_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="https://..."
)

email_change_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="...@gmail.com"
)


group_name_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True,
    input_field_placeholder="Название группы..."
)


employees_del_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True,
    input_field_placeholder="Иванов Иван Иванович, Петров Петр Петрович..."
)

faculty_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="Название факультета"
)

direction_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="Название направления"
)

course_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="Курс"
)

recod_book_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="Номер зачетки"
)

subject_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="Название предмета"
)

date_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="01.01.2024"
)

control_input_cancel_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Отмена")]],
    resize_keyboard=True, one_time_keyboard=True, input_field_placeholder="экзамен"
)

empty_keyboard = ReplyKeyboardRemove()
