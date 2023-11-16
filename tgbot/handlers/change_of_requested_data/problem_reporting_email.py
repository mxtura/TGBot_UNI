from aiogram.types import Message
from aiogram.dispatcher import FSMContext

import tgbot.keyboards.reply as rkb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.misc.states import ProblemReportingEmailChangeStates
from tgbot.models.database_instance import db


async def request_problem_reporting_email(message: Message):
    await message.answer(text="Пожалуйста введите адрес электронной почты, на которую пользователи смогут сообщать о проблемах в работе бота",
                         reply_markup=rkb.email_change_cancel_keyboard)
    await ProblemReportingEmailChangeStates.first()


async def change_problem_reporting_email(message: Message, state: FSMContext):
    await db.change_problem_reporting_email(message.text)
    await message.answer(text="Почта поддержки изменена", reply_markup=rkb.manager_keyboard)
    await state.finish()


def register_problem_reporting_email_change(dp):
    dp.register_message_handler(request_problem_reporting_email, UserTypeFilter("manager"),
                                content_types=['text'], text='Заменить почту поддержки')
    dp.register_message_handler(change_problem_reporting_email, state=ProblemReportingEmailChangeStates.change_problem_reporting_email)