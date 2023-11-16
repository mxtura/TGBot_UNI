from aiogram.types import Message
import tgbot.keyboards.inline as ikb
from tgbot.misc.states import GroupEditingStates
from aiogram.dispatcher import FSMContext
from tgbot.models.database_instance import db

async def send_group_edit_form(message: Message, state: FSMContext):
    state_data = await state.get_data()
    current_group_name = state_data["group_name"]
    current_performance_list = await db.get_performance_list_by_group_name(current_group_name)
    current_performance_list = current_performance_list.replace("_", r"\_").replace("*", r"\*")
    del_msg = await message.answer(f"*Название группы:* {current_group_name}\n\n"
                                   f"*Таблица успеваемости:* {current_performance_list}",
                                   reply_markup=ikb.group_edit_keyboard,
                                   parse_mode="MARKDOWN")
    await state.update_data(del_msg=del_msg)
    await GroupEditingStates.group_form_interaction.set()

