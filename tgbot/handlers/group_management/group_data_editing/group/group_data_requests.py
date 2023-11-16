from aiogram.types import Message
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.misc.states import GroupNameChangeStates, GroupURLChangeStates, GroupEditingStates


async def request_group_name(message: Message, state: FSMContext, groups):
    state_data = await state.get_data()
    del_msg = state_data["del_msg"]
    if del_msg is not None:
        await del_msg.delete()

    group_msg = ""
    for group in groups:
        group_msg += f"◦ `{group[0]}`;\n"
    del_msg = await message.answer(f"*Доступные группы:*\n{group_msg}\n"
                                   f"*Введите название группы, которую хотите выбрать*",
                                   reply_markup=rkb.group_name_cancel_keyboard,
                                   parse_mode="MARKDOWN")
    await state.update_data(del_msg=del_msg)


async def request_new_group_name(message: Message, state: FSMContext):
    del_msg = await message.answer(text="Введите новое название группы:",
                                   reply_markup=rkb.group_name_cancel_keyboard)
    await state.update_data(del_msg=del_msg)
    await GroupNameChangeStates.getting_new_group_name.set()


async def request_new_performance_list_url(message: Message, state: FSMContext):
    del_msg = await message.answer(text="Введите новую ссылку на таблицу успеваемости:",
                                   reply_markup=rkb.url_change_cancel_keyboard)
    await state.update_data(del_msg=del_msg)
    await GroupURLChangeStates.getting_new_performance_list_url.set()