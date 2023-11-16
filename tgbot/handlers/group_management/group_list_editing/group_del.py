import re

from aiogram.types import Message, CallbackQuery
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.misc.states import GroupsDelStates
import tgbot.keyboards.inline as ikb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.models.database_instance import db

no_groups_selected = "группы не выбраны"


async def get_group_names(message: Message, state: FSMContext):
    groups = await db.get_group_names()
    if len(groups) == 0:
        await message.answer(text="Группы еще не добавлены", reply_markup=rkb.manager_keyboard)
    else:
        await state.update_data(group_list=groups)
        group_msg = ""
        for group in groups:
            group_msg += f"◦ `{group[0]}`;\n"
        # group_msg = sorted(group_msg)
        await message.answer(text=f"*Доступные группы:*\n{group_msg}\n"
                                  f"*Введите названия групп, которые хотите удалить, через запятую*",
                             reply_markup=rkb.group_name_cancel_keyboard, parse_mode="MARKDOWN")
        await GroupsDelStates.first()


async def confirm_deleting(message: Message, state: FSMContext):
    selected_groups_msg = re.sub(r",\s+", ",", message.text)
    selected_groups_msg = selected_groups_msg.split(",")
    selected_groups = ""
    data = await state.get_data()
    group_list = data.get("group_list")
    for group in selected_groups_msg:
        for gr in group_list:
            if group in gr:
                selected_groups += f"◦ `{group}`;\n"
    if selected_groups == "":
        groups_msg = ""
        for group in group_list:
            groups_msg += f"◦ `{group[0]}`;\n"
        await message.answer(text=f"*Вы не выбрали ни одной группы!*\n\n"
                                  f"Доступные группы:\n{groups_msg}\n"
                                  f"Введите названия групп, которые хотите удалить, через запятую\n\n",
                             parse_mode="MARKDOWN")
    else:
        await state.update_data(selected_groups=selected_groups_msg)
        await message.answer(text=f"*Выбранные группы:* \n{selected_groups}\n"
                                  f"*Вы уверены, что хотите удалить выбранные группы?*",
                             reply_markup=ikb.confirmation_kb, parse_mode="MARKDOWN")
        await GroupsDelStates.next()


async def delete_groups(callback_query: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    if callback_query.data == "no":
        group_list = data.get("selected_groups")
        groups_msg = ""
        for group in group_list:
            groups_msg += f"◦ `{group[0]}`;\n"
        await callback_query.message.answer(text=f"*Доступные группы:*\n{groups_msg}\n"
                                                 f"Введите названия групп, которые хотите удалить, через запятую\n\n",
                                            parse_mode="MARKDOWN")
        await GroupsDelStates.first()
    else:
        groups = data.get("selected_groups")
        msg = ""
        for group in groups:
            msg += f"◦ `{group}`;\n"
            await db.delete_group(group)
        await callback_query.message.answer(f"*Следующие группы были удалены:* \n{msg}",
                                            reply_markup=rkb.manager_keyboard, parse_mode="MARKDOWN")
        await state.finish()


def register_group_del(dp):
    dp.register_message_handler(get_group_names, UserTypeFilter("manager"),
                                content_types=['text'], text='Удалить группы')
    dp.register_message_handler(confirm_deleting, state=GroupsDelStates.confirm_groups_deleting_state)
    dp.register_callback_query_handler(delete_groups, state=GroupsDelStates.deleting_groups_state)
