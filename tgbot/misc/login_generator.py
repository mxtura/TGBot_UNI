from transliterate import translit
import itertools
from tgbot.models.database_instance import db


async def generate_login(fio, year):
    """Генерирует логин."""
    first_name_en = translit(fio[0], 'ru', reversed=True).lower()
    last_name_en = translit(fio[1], 'ru', reversed=True).lower()
    if fio[2] is None:
        initials = first_name_en[0] + last_name_en[0]
    else:
        middle_name_en = translit(fio[2], 'ru', reversed=True).lower()
        initials = first_name_en[0] + last_name_en[0] + middle_name_en[0]
    login_base = initials + "." + year

    logins = await db.get_logins()
    for i in itertools.count(start=1):
        login = f"{login_base}.{str(i)}"
        if login not in [log[0] for log in logins]:
            return login

        logins.append((login,))