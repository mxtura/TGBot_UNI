from aiogram.types import Message
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.inline as ikb
from tgbot.misc.states import GroupEditingStates
from tgbot.models.database_instance import db

async def send_student_list_form(message: Message, state: FSMContext):
    state_data = await state.get_data()
    group_name = state_data['group_name']
    student_names = await db.get_students_by_group_name(group_name)
    student_list = [f"*Список студентов группы* {group_name}*:*\n"]
    n = 1
    for name in student_names:
        if name[2] is not None:
            student_list.append(f"◦ {n}. `{name[0]} {name[1]} {name[2]}`;\n")
        else:
            student_list.append(f"◦ {n}. `{name[0]} {name[1]}`;\n")
        n += 1
    del_msg = await message.answer(text="".join(student_list),
                                   reply_markup=ikb.student_list_keyboard,
                                   parse_mode="MARKDOWN")
    await state.update_data(del_msg=del_msg)
    await GroupEditingStates.student_list_interaction.set()
        