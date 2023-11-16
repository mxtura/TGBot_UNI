from aiogram.dispatcher.filters import BoundFilter
from tgbot.models.database_instance import db


class UserTypeFilter(BoundFilter):
    def __init__(self, user_type):
        self.user_type = user_type

    async def check(self, message):
        user = await db.get_user_type(message.from_user.id)
        return bool(user == self.user_type)
