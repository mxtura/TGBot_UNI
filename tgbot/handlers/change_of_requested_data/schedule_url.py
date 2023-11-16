from aiogram.types import Message
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.misc.states import ScheduleURLChangeStates
from tgbot.misc.URL_validator import is_string_an_url
from tgbot.models.database_instance import db


async def request_schedule_url(message: Message):
    await message.answer(text="Пожалуйста введите ссылку на расписание",
                         reply_markup=rkb.url_change_cancel_keyboard)
    await ScheduleURLChangeStates.first()


async def change_schedule_url(message: Message, state: FSMContext):
    if is_string_an_url(message.text):
        await db.change_schedule_url(message.text)
        await message.answer(text="Расписание изменено", reply_markup=rkb.manager_keyboard)
    else:
        await message.answer(text="Ссылка введена некорректно", reply_markup=rkb.manager_keyboard)
    await state.finish()


def register_schedule_change(dp):
    dp.register_message_handler(request_schedule_url, UserTypeFilter("manager"),
                                content_types=['text'], text='Заменить расписание')
    dp.register_message_handler(change_schedule_url, state=ScheduleURLChangeStates.change_schedule_url)
