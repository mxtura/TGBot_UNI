from aiogram.types import Message
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.misc.states import GroupEditingStates
from .group_edit_form import send_group_edit_form
from .group.group_data_requests import request_group_name
from tgbot.models.database_instance import db


async def get_group_name(message: Message, state: FSMContext):
    groups = await db.get_group_names()
    await state.update_data(del_msg=None)
    if len(groups) == 0:
        await message.answer(text="Группы еще не добавлены", reply_markup=rkb.manager_keyboard)
    else:
        await request_group_name(message, state, groups)
        await GroupEditingStates.getting_group_name.set()


async def check_group_name(message: Message, state: FSMContext):
    await message.delete()
    state_data = await state.get_data()
    del_msg = state_data["del_msg"]
    await del_msg.delete()

    group_name = message.text
    groups = await db.get_group_names()
    group_exist = False

    for group in groups:
        if group[0] == group_name:
            group_exist = True
            break
    if group_exist:
        await state.update_data(group_name=group_name)
        await send_group_edit_form(message, state)
    else:
        del_msg = await message.answer(text="<b>Такой группы не существует!</b>")
        await state.update_data(del_msg=del_msg)
        await request_group_name(message, state, groups)
    

def register_group_selection(dp):
    dp.register_message_handler(
        get_group_name, 
        UserTypeFilter("manager"), 
        content_types=['text'],
        text='Редактировать группу'
        ) 
    
    dp.register_message_handler(
        check_group_name, 
        state=GroupEditingStates.getting_group_name)