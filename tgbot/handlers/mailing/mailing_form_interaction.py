import tgbot.keyboards.reply as rkb
from aiogram.dispatcher import FSMContext
from tgbot.misc.states import MailingsStates
from aiogram.types import Message, CallbackQuery
from tgbot.filters.user_type import UserTypeFilter
from tgbot.misc.decorators.log_decorator import log_function_call
from tgbot.handlers.mailing.mailing_text_data import std_groups, std_managers, std_teachers, std_date, std_time
from tgbot.handlers.mailing.mailing_text_data import mailing_recorded_msg, mailing_send_msg
from tgbot.handlers.mailing.mailing_data_requests import request_mailing_text, request_date, request_time
from tgbot.handlers.mailing.mailing_data_requests import request_group_selection, request_manager_selection, request_teacher_selection
from tgbot.handlers.mailing.mailing_notification import send_mailing_notification
from tgbot.handlers.mailing.mailing_recording import record_mailing
from tgbot.handlers.mailing.instant_mailing_send import send_mailing


# настроить уведомления через текстовые переменные
@log_function_call
async def handle_mailing_form_interaction(callback_query: CallbackQuery, state: FSMContext):
    cq_data = callback_query.data
    state_data = await state.get_data()
    del_msgs = state_data["del_msgs"]
    for msg in del_msgs:
        await msg.delete()
    await state.update_data(del_msgs=[])
    if cq_data == 'mailing_text_input':
        await request_mailing_text(callback_query.message, state)
    elif cq_data == 'date_input':
        await request_date(callback_query.message, state)
    elif cq_data == 'time_input':
        await request_time(callback_query.message, state)
    elif cq_data == 'groups_selection':
        await request_group_selection(callback_query.message, state)
    elif cq_data == 'managers_selection':
        await request_manager_selection(callback_query.message, state)
    elif cq_data == 'teachers_selection':
        await request_teacher_selection(callback_query.message, state)
    elif cq_data == 'confirmation_of_mailing':
        await send_mailing_notification(callback_query.message, state)
        user_type = state_data['user_type']
        if (state_data['groups'] == std_groups and 
            state_data['managers'] == std_managers and 
            state_data['teachers'] == std_teachers):
            await close_mailing_form(callback_query.message, state)
        elif state_data['date'] == std_date and state_data['time'] == std_time:
            await send_mailing(callback_query, state)
            if user_type == "manager":
                await callback_query.message.answer(mailing_send_msg, reply_markup=rkb.manager_keyboard)
            elif user_type == "teacher":
                await callback_query.message.answer(mailing_send_msg, reply_markup=rkb.teacher_keyboard)
        else:
            msg = await callback_query.message.answer("Рассылка записывается...")
            await record_mailing(callback_query, state)
            await msg.delete()
            if user_type == "manager":
                await callback_query.message.answer(mailing_recorded_msg, reply_markup=rkb.manager_keyboard)
            elif user_type == "teacher":
                await callback_query.message.answer(mailing_recorded_msg, reply_markup=rkb.teacher_keyboard)
        await state.finish()
    elif cq_data == 'cancellation_of_mailing_form':
        await close_mailing_form(callback_query.message, state)
        await state.finish()


async def close_mailing_form(message: Message, state: FSMContext):
    state_data = await state.get_data()
    user_type = state_data["user_type"]
    if user_type == "student":
        await message.answer("Меню студента", reply_markup=rkb.student_keyboard)
    elif user_type == "teacher":
        await message.answer("Меню преподавателя", reply_markup=rkb.teacher_keyboard)
    elif user_type == "manager":
        await message.answer("Меню менеджера", reply_markup=rkb.manager_keyboard)
    else:
        await message.answer("Пожалуйста авторизируйтесь\n"
                             "Для этого воспользуйтесь командой /start", reply_markup=rkb.empty_keyboard)


def register_mailing_form_interaction(dp):
    dp.register_callback_query_handler(handle_mailing_form_interaction, UserTypeFilter("manager"),
                                       state=MailingsStates.mailing_form_interaction_state)
    dp.register_callback_query_handler(handle_mailing_form_interaction, UserTypeFilter("teacher"),
                                       state=MailingsStates.mailing_form_interaction_state)