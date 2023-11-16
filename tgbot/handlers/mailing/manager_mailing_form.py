import tgbot.keyboards.inline as ikb
import tgbot.keyboards.reply as rkb

from aiogram.dispatcher import FSMContext
from tgbot.misc.states import MailingsStates
from aiogram.types import Message
from tgbot.filters.user_type import UserTypeFilter
from tgbot.misc.decorators.log_decorator import log_function_call

from tgbot.handlers.mailing.mailing_text_data import std_text, std_groups, std_managers, std_teachers, std_date, std_time
from tgbot.handlers.mailing.mailing_text_data import create_group_list, create_manager_list, create_teacher_list
from tgbot.handlers.mailing.mailing_text_data import create_date_and_time_variable


@log_function_call
async def prepare_manager_initial_data(message: Message, state: FSMContext):
    await state.update_data(mailing_text=std_text, groups=std_groups, managers=std_managers, teachers=std_teachers,
                            date=std_date, time=std_time, user_type="manager", del_msgs=[])
    await create_manager_mailing_form(message, state)


@log_function_call
async def create_manager_mailing_form(message: Message, state: FSMContext):
    state_data = await state.get_data()
    del_msgs = state_data["del_msgs"]
    del_msgs.append(await message.answer(text="<b>Форма создания рассылки</b>", reply_markup=rkb.empty_keyboard))
    # Можно переписать по красивше
    mailing_text, group_list, manager_list, teacher_list, mailing_date_and_time = await prepare_manager_mailing_form_data(state)
    del_msgs.append(await message.answer(text=f"*Текст рассылки:*\n{mailing_text}\n\n"
                                         f"*Группы:*\n{group_list}\n"
                                         f"*Менеджеры:*\n{manager_list}\n"
                                         f"*Преподаватели:*\n{teacher_list}\n"
                                         f"*{mailing_date_and_time}*",
                                         reply_markup=ikb.manager_mailing_keyboard,
                                         parse_mode="MARKDOWN"))
    await state.update_data(del_msgs=del_msgs)
    await MailingsStates.mailing_form_interaction_state.set()


async def prepare_manager_mailing_form_data(state: FSMContext):
    state_data = await state.get_data()
    mailing_text = state_data['mailing_text']
    groups = state_data['groups']
    managers = state_data['managers']
    teachers = state_data['teachers']
    group_list = await create_group_list(groups)
    manager_list = await create_manager_list(managers)
    teacher_list = await create_teacher_list(teachers)
    mailing_date_and_time = await create_date_and_time_variable(state)
    return mailing_text, group_list, manager_list, teacher_list, mailing_date_and_time


def register_manager_mailing_form(dp):
    dp.register_message_handler(prepare_manager_initial_data, UserTypeFilter("manager"),
                                content_types=['text'], text='Создать рассылку')