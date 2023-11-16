from aiogram.types import CallbackQuery
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.handlers.group_management.group_data_editing.group_edit_form import send_group_edit_form
from tgbot.handlers.personal_data_editing.change_user_password import start_change
from tgbot.handlers.cancel import cq_cancel_all
from tgbot.misc.states import GroupEditingStates


async def handle_student_list_interaction(callback_query: CallbackQuery, state: FSMContext):
    cq_data = callback_query.data
    state_data = await state.get_data()
    del_msg = state_data["del_msg"]
    if del_msg is not None:
        await del_msg.delete()
        await state.update_data(del_msg=None)

    # if cq_data == "change_student_password":
    #     await start_change()

    if cq_data == 'back':
        await send_group_edit_form(callback_query.message, state)
    elif cq_data == 'cancel':
        await cq_cancel_all(callback_query, state)


def register_student_list_interaction(dp):
    dp.register_callback_query_handler(
        handle_student_list_interaction,
        UserTypeFilter("manager"),
        state=GroupEditingStates.student_list_interaction
        )