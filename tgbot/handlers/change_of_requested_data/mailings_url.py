from aiogram.types import Message
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.misc.states import MailingsURLChangeStates
from tgbot.misc.URL_validator import is_string_an_url
from tgbot.models.database_instance import db


async def request_mailings_url(message: Message):
    await message.answer(text="Введите ссылку на таблицу сообщений",
                         reply_markup=rkb.url_change_cancel_keyboard)
    await MailingsURLChangeStates.first()


async def change_mailings_url(message: Message, state: FSMContext):
    if is_string_an_url(message.text):
        await db.change_mailings_url(message.text)
        await message.answer(text="Таблица рассылок изменена", reply_markup=rkb.manager_keyboard)
    else:
        await message.answer(text="Ссылка введена некорректно", reply_markup=rkb.manager_keyboard)
    await state.finish()


def register_mailings_link_change(dp):
    dp.register_message_handler(request_mailings_url, UserTypeFilter("manager"),
                                content_types=['text'], text='Заменить таблицу рассылок')
    dp.register_message_handler(change_mailings_url, state=MailingsURLChangeStates.change_mailings_url_state)
