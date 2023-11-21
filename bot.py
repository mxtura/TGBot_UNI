import asyncio
import logging
from datetime import datetime
import pytz


from aiogram import Bot, Dispatcher
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from tgbot.filters.user_type import UserTypeFilter
from tgbot.middlewares.skip_handlers import SkipHandlerMiddleware
from tgbot.config import load_config
from tgbot.handlers.mailing.delayed_mailing_send import check_msgs_list
from tgbot.misc.commands import set_commands

from tgbot.handlers.cancel import register_cancel
from tgbot.handlers.keyboard_menu import register_menu
from tgbot.handlers.authorization import register_authorization
from tgbot.handlers.logout import register_logout
from tgbot.handlers.change_of_requested_data.schedule_url import register_schedule_change
from tgbot.handlers.change_of_requested_data.work_schedule_url import register_work_schedule_change
from tgbot.handlers.change_of_requested_data.learning_schedule_url import register_learning_schedule_change
from tgbot.handlers.change_of_requested_data.mailings_url import register_mailings_link_change
from tgbot.handlers.change_of_requested_data.reports_url import register_reports_change
from tgbot.handlers.change_of_requested_data.retakes_url import register_retakes_change
from tgbot.handlers.change_of_requested_data.problem_reporting_email import register_problem_reporting_email_change
from tgbot.handlers.performance_debts import register_performance_debts_check
from tgbot.handlers.group_management.group_list_editing.group_add import register_group_add
from tgbot.handlers.group_management.group_list_editing.group_del import register_group_del
from tgbot.handlers.group_management.group_data_editing.group_selection import register_group_selection
from tgbot.handlers.group_management.group_data_editing.group_form_interaction import register_group_form_interaction
from tgbot.handlers.group_management.group_data_editing.group.group_name_change import register_group_name_change
from tgbot.handlers.group_management.group_data_editing.group.performance_list_change import register_performance_list_change
from tgbot.handlers.group_management.group_data_editing.students.student_list_form_interaction import register_student_list_interaction
from tgbot.handlers.employee_management.employee_add import register_employee_add
from tgbot.handlers.employee_management.employees_del import register_employees_del
from tgbot.handlers.personal_data_editing.change_user_password import register_change_user_password
from tgbot.handlers.personal_data_editing.login_change import register_login_change
from tgbot.handlers.personal_data_editing.password_change import register_password_change
from tgbot.handlers.mailing.mailing_data_processing import register_mailing_data_processing
from tgbot.handlers.mailing.mailing_form_interaction import register_mailing_form_interaction
from tgbot.handlers.mailing.manager_mailing_form import register_manager_mailing_form
from tgbot.handlers.mailing.teacher_mailing_form import register_teacher_mailing_form
from tgbot.handlers.requested_data_sending import register_requested_data_sending
from tgbot.handlers.role_change import register_role_chenge
from tgbot.handlers.debt_management.registration_for_retake import register_reg_retake

from tgbot.models.database_instance import db

logger = logging.getLogger(__name__)
timezone = pytz.timezone('Europe/Moscow')

def register_all_middlewares(dp, bot_start_time):
    """Регистрирует все мидлвари."""
    # Нет антиспам мидлвари. Нужно сделать
    dp.setup_middleware(SkipHandlerMiddleware(bot_start_time))


def register_all_filters(dp):
    """Регистрирует все фильтры."""
    dp.filters_factory.bind(UserTypeFilter)


def register_all_handlers(dp):
    """Регистрирует все хэндлеры."""
    register_cancel(dp)
    register_authorization(dp)
    register_logout(dp)
    register_menu(dp)
    register_schedule_change(dp)
    register_work_schedule_change(dp)
    register_learning_schedule_change(dp)
    register_mailings_link_change(dp)
    register_reports_change(dp)
    register_retakes_change(dp)
    register_problem_reporting_email_change(dp)
    register_performance_debts_check(dp)
    register_group_add(dp)
    register_employee_add(dp)
    register_change_user_password(dp)
    register_group_del(dp)
    register_employees_del(dp)
    register_login_change(dp)
    register_password_change(dp)
    register_mailing_data_processing(dp)
    register_mailing_form_interaction(dp)
    register_manager_mailing_form(dp)
    register_teacher_mailing_form(dp)
    register_requested_data_sending(dp)
    register_role_chenge(dp)
    register_group_selection(dp)
    register_group_form_interaction(dp)
    register_group_name_change(dp)
    register_performance_list_change(dp)
    register_student_list_interaction(dp)
    register_reg_retake(dp)


def set_scheduled_jobs(scheduler, bot):
    """Устанавливает запланированные события."""
    scheduler.add_job(check_msgs_list, "interval", seconds=60, args=(bot,))


async def main():
    """Задает все необходимые переменные для работы бота и запускает его."""
    logging.basicConfig(
        level=logging.INFO,
        format=u'%(filename)s:%(lineno)d #%(levelname)-8s [%(asctime)s] - %(name)s - %(message)s',
    )

    bot_start_time = datetime.now(timezone).strftime('%Y-%m-%d %H:%M:%S')
    logger.info("Starting bot")
    config = load_config(".env")
    storage = MemoryStorage()

    bot = Bot(token=config.tg_bot.token, parse_mode='HTML')
    dp = Dispatcher(bot, storage=storage)
    bot['config'] = config
    scheduler = AsyncIOScheduler()

    set_scheduled_jobs(scheduler, bot)
    register_all_middlewares(dp, bot_start_time)
    register_all_filters(dp)
    register_all_handlers(dp)

    await db.create_tables()
    await set_commands(bot)

    try:
        scheduler.start()
        await dp.start_polling()
    finally:
        await dp.storage.close()
        await dp.storage.wait_closed()
        await bot.session.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except(KeyboardInterrupt, SystemExit):
        logger.error("Bot stopped!")
