import asyncio
import random
import itertools

from aiogram import Dispatcher
from aiogram.types import Message, CallbackQuery
from tgbot.filters.user_type import UserTypeFilter
from tgbot.misc.decorators.log_decorator import log_function_call
import tgbot.keyboards.reply as rkb
import tgbot.keyboards.inline as ikb
from tgbot.misc.states import AuthorizationStates, FirstManagerAuthorizationStates
from aiogram.dispatcher import FSMContext
from transliterate import translit
from tgbot.models.database_instance import db
from tgbot.misc.login_generator import generate_login


@log_function_call
async def authorized_notification(message: Message, state: FSMContext):
    """Объявляет, что пользователь уже авторизован и отправляет меню."""
    await message.answer("<b>Вы уже авторизованы!</b>")
    user_type = await db.get_user_type(message.from_user.id)
    if user_type == "student":
        await message.answer("Меню студента", reply_markup=rkb.student_keyboard)
    elif user_type == "teacher":
        await message.answer("Меню преподавателя", reply_markup=rkb.teacher_keyboard)
    elif user_type == "manager":
        await message.answer("Меню менеджера", reply_markup=rkb.manager_keyboard)
    await state.finish()

# !Можно настроить отмену действия и удаление лишних сообщений спустя время, если пользователь ничего не вводит
# !Регистрацию менеджера и авторизацию других пользователей можно распределить в разные файлы
@log_function_call
async def start_authorization(message: Message, state: FSMContext):
    """Определяет является ли пользователь первым пользователем бота, если да запускает регистрацию первого менеджера, иначе запрашивает логин."""
    await message.answer("<b>Добро пожаловать в бота для заочников!</b>")
    if await db.users_exist():
        msgs_to_del = [await message.answer(text="Введите логин", reply_markup=rkb.login_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
        await AuthorizationStates.getting_login.set()
    else:
        msgs_to_del = [await message.answer(text="Вы будете зарегистрированы как менеджер"),
                       await message.answer(text="Пожалуйста введите ваше ФИО кириллицей через пробел",
                                            reply_markup=rkb.name_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
        await FirstManagerAuthorizationStates.getting_name.set()


# !Сделать проверку на непонятные символы
@log_function_call
async def get_login(message: Message, state: FSMContext):
    """Получает логин и проверяет корректность введенных данных."""
    msg_text = message.text.split()
    await clear_chat(message, state, 1)
    if len(msg_text) != 1:
        msgs_to_del = [await message.answer(text="<b>Логин введен некорректно</b>"),
                       await message.answer(text="Введите логин", reply_markup=rkb.name_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
    else:
        await state.update_data(login=msg_text[0])
        msgs_to_del = [await message.answer(text="Введите пароль", reply_markup=rkb.password_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
        await AuthorizationStates.getting_password.set()


@log_function_call
async def get_password(message: Message, state: FSMContext):
    """Получает пароль"""
    await clear_chat(message, state, 1)
    await state.update_data(password=message.text)
    await AuthorizationStates.checking_data.set()
    await check_data(message, state)


@log_function_call
async def check_data(message: Message, state: FSMContext):
    """Проверяет совпадают ли значения логина и пароля с каким-либо зарегистрированным пользователем."""
    data = await state.get_data()
    login = data.get("login")
    password = data.get("password")
    if await db.check_auth_data(login, password) == "correct":
        await authorize_user(message, state, login)
    else:
        msgs_to_del = [await message.answer(text="<b>Данные неверны</b>"),
                       await message.answer(text="Введите логин", reply_markup=rkb.login_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
        await AuthorizationStates.getting_login.set()


@log_function_call
async def authorize_user(message: Message, state: FSMContext, login):
    """Авторизирует пользователя"""
    await db.connect_telegram_id(message.from_user.id, login)
    user_type = await db.get_user_type(message.from_user.id)
    if user_type == "student":
        await message.answer("Меню студента", reply_markup=rkb.student_keyboard)
    elif user_type == "teacher":
        await message.answer("Меню преподавателя", reply_markup=rkb.teacher_keyboard)
    elif user_type == "manager":
        await message.answer("Меню менеджера", reply_markup=rkb.manager_keyboard)
    await state.finish()


@log_function_call
async def check_name(message: Message, state: FSMContext):
    """Проверяет корректность введенного имени. (для первого менеджера)"""
    msg_text = message.text.split(" ")
    await clear_chat(message, state, 1)
    if len(msg_text) < 2:
        msgs_to_del = [await message.answer(text="<b>ФИО не может состоять из одного слова</b>"),
                       await message.answer(text="Пожалуйста введите ваше ФИО кириллицей через пробел",
                                            reply_markup=rkb.name_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
    elif len(msg_text) > 3:
        msgs_to_del = [await message.answer(text="<b>ФИО не должно содержать более трех слов</b>"),
                       await message.answer(text="Пожалуйста введите ваше ФИО кириллицей через пробел",
                                            reply_markup=rkb.name_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
    else:
        msgs_to_del = [await message.answer(text=f"Ваше ФИО: <b>{message.text}</b> \nвведено корректно?",
                                            reply_markup=ikb.confirmation_kb)]
        await state.update_data(name=message.text, msgs_to_del=msgs_to_del)
        await FirstManagerAuthorizationStates.authorizing.set()


# Можно поделить эту функцию на несколько
@log_function_call
async def register_first_manager(callback_query: CallbackQuery, state: FSMContext):
    """Регистрирует первого менеджера."""
    await callback_query.answer()
    await clear_chat(callback_query.message, state, 0)
    data = await state.get_data()
    if callback_query.data == "no":
        msgs_to_del = [await callback_query.message.answer("Пожалуйста введите ваше ФИО кириллицей через пробел\n\n",
                                                           reply_markup=rkb.name_input_cancel_keyboard)]
        await state.update_data(msgs_to_del=msgs_to_del)
        await FirstManagerAuthorizationStates.getting_name.set()
    else:
        fio = data.get("name").split()
        if len(fio) < 3:
            fio.append(None)
        user_type = "manager"
        login = await generate_login(fio, user_type)
        password = random.randint(100000, 999999)
        await db.add_user(list(fio), login, password, user_type)
        await db.connect_telegram_id(callback_query.from_user.id, login)
        personal_data = await callback_query.message.answer(
            text="Вы успешно зарегистрированы!\n\n"
                 f"<b>Ваш логин:</b> {login}\n"
                 f"<b>Ваш пароль:</b> {password}\n\n"
                 f"<b>Это сообщение будет удалено через 10 минут, поэтому рекомендуется записать эти данные</b>")
        asyncio.create_task(del_msg(personal_data, 600))
        await callback_query.message.answer(text="Меню менеджера", reply_markup=rkb.manager_keyboard)
        await state.finish()


# Вынести в отдельный файл
async def del_msg(message, time):
    """Удаляет сообщение через заданное время."""
    await asyncio.sleep(time)
    await message.delete()


async def clear_chat(message: Message, state: FSMContext, del_user_msg):
    """Удаляет все сообщения, отправленные в процессе регистрации/авторизации."""
    data = await state.get_data()
    msgs_to_del = data.get("msgs_to_del")
    for msg in msgs_to_del:
        await msg.delete()
    if del_user_msg:
        await message.delete()


def register_authorization(dp: Dispatcher):
    dp.register_message_handler(start_authorization, UserTypeFilter(None), commands=['start'], state="*")
    dp.register_message_handler(authorized_notification, UserTypeFilter("manager"), commands=['start'], state="*")
    dp.register_message_handler(authorized_notification, UserTypeFilter("teacher"), commands=['start'], state="*")
    dp.register_message_handler(authorized_notification, UserTypeFilter("student"), commands=['start'], state="*")
    dp.register_message_handler(get_login, content_types=['text'], state=AuthorizationStates.getting_login)
    dp.register_message_handler(get_password, content_types=['text'], state=AuthorizationStates.getting_password)
    dp.register_message_handler(check_name, content_types=['text'],
                                state=FirstManagerAuthorizationStates.getting_name)
    dp.register_callback_query_handler(register_first_manager,
                                       state=FirstManagerAuthorizationStates.authorizing)
