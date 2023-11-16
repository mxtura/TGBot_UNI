from aiogram.types import Message
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.misc.states import RetakesURLChangeStates
from tgbot.misc.URL_validator import is_string_an_url
from tgbot.models.database_instance import db


# Добавить подтверждение
async def request_retakes_url(message: Message):
    await message.answer(text="Пожалуйста введите ссылку на каталог ведомостей пересдач",
                         reply_markup=rkb.url_change_cancel_keyboard)
    await RetakesURLChangeStates.first()


async def change_retakes_url(message: Message, state: FSMContext):
    if is_string_an_url(message.text):
        await db.change_retake_cards_url(message.text)
        await message.answer(text="Ссылка на каталок ведомостей изменена", reply_markup=rkb.manager_keyboard)
    else:
        await message.answer(text="Ссылка введена некорректно", reply_markup=rkb.manager_keyboard)
    await state.finish()


def register_retakes_change(dp):
    dp.register_message_handler(request_retakes_url, UserTypeFilter("manager"),
                                content_types=['text'], text='Заменить пересдачи')
    dp.register_message_handler(change_retakes_url, state=RetakesURLChangeStates.change_retakes_url_state)
