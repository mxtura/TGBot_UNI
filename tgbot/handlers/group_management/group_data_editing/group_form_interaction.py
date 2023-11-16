from aiogram.types import CallbackQuery
from aiogram.dispatcher import FSMContext

from tgbot.filters.user_type import UserTypeFilter
from .group.group_data_requests import request_new_group_name, request_new_performance_list_url
from .students.student_list_update import check_year_of_enrollment
from .students.student_list_form import send_student_list_form
from tgbot.handlers.cancel import cq_cancel_all
from tgbot.misc.states import GroupEditingStates


async def handle_group_form_interaction(callback_query: CallbackQuery, state: FSMContext):
    cq_data = callback_query.data
    state_data = await state.get_data()
    del_msg = state_data["del_msg"]
    if del_msg is not None:
        await del_msg.delete()
        await state.update_data(del_msg=None)

    if cq_data == 'change_group_name':
        await request_new_group_name(callback_query.message, state)
    elif cq_data == 'change_performance_list':
        await request_new_performance_list_url(callback_query.message, state)
    elif cq_data == 'update_student_list':
        await check_year_of_enrollment(callback_query.message, state)
    elif cq_data == 'display_student_list':
        await send_student_list_form(callback_query.message, state)
    elif cq_data == 'cancel_group_editing':
        await cq_cancel_all(callback_query, state)


def register_group_form_interaction(dp):
    dp.register_callback_query_handler(
        handle_group_form_interaction,
        UserTypeFilter("manager"),
        state=GroupEditingStates.group_form_interaction
        )