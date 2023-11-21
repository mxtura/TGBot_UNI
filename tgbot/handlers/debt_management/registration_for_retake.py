from aiogram.types import Message, CallbackQuery
from aiogram.dispatcher import FSMContext

import logging
import tgbot.keyboards.inline as ikb
import tgbot.keyboards.reply as rkb
from tgbot.misc.states import LoginChangeStates
from tgbot.filters.user_type import UserTypeFilter
from tgbot.models.database_instance import db

reg_retake_request_text = "Пожалуйста, напишите название предмета, который хотите пересдать:"

async def send_reg_retake_request(message: Message, state: FSMContext):
    """Просит ввести название предмета"""
    del_msg = await message.answer(text=reg_retake_request_text,
                                   reply_markup=rkb.retake_input_cancel_keyboard)
    await state.update_data(del_msg=del_msg)


def register_reg_retake(dp):
    dp.register_message_handler(send_reg_retake_request, UserTypeFilter("student"), content_types=['text'],
                                text=['Записаться на пересдачу'])
