from aiogram.types import Message
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.misc.states import GroupURLChangeStates
from tgbot.models.database_instance import db


async def check_new_performance_list_url(message: Message, state: FSMContext):
    state_data = await state.get_data()
    del_msg = state_data["del_msg"]
    await del_msg.delete()
    await message.delete()

    new_performance_list_url = message.text
    performance_url_list = await db.get_performance_urls()
    is_duplicate = 0
    for url in performance_url_list:
        if url[0] == new_performance_list_url:
            is_duplicate = 1
            break

    if is_duplicate:
        connected_group_name = await db.get_group_name_by_url(new_performance_list_url)
        del_msg = await message.answer(f"f*Такая ссылка уже привязана к группе* {connected_group_name}\n\n"
                                       "*Введите другую ссылку на таблицу успеваемости*",
                                       parse_mode="MARKDOWN")
        await state.update_data(del_msg=del_msg)
    else:
        await state.update_data(new_performance_list_url=new_performance_list_url)
        await change_performance_list(message, state)


async def change_performance_list(message: Message, state = FSMContext):
    state_data = await state.get_data()
    group_name = state_data["group_name"]
    new_performance_list_url = state_data["new_performance_list_url"]
    await db.change_performance_list_url(group_name, new_performance_list_url)
    await message.answer(f"*Таблица успеваемости группы* {group_name} *успешно изменена*",
                         parse_mode="MARKDOWN",
                         reply_markup=rkb.manager_keyboard)
    await state.finish()
    

def register_performance_list_change(dp):
    dp.register_message_handler(
        check_new_performance_list_url,
        content_types=['text'],
        state=GroupURLChangeStates.getting_new_performance_list_url
        )