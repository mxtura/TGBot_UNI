"""Осуществляет все запросы к базе данных."""
import aiosqlite


class Database:
    def __init__(self, db_file):
        self.db_name = db_file

    async def create_tables(self):
        """Создает все необходимые таблицы."""
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute('''CREATE TABLE IF NOT EXISTS users 
                (login TEXT PRIMARY KEY,
                telegram_id INTEGER,
                last_name TEXT,
                first_name TEXT,
                middle_name TEXT,
                password TEXT,
                user_type TEXT,
                group_name TEXT,
                FOREIGN KEY (group_name) REFERENCES groups(group_name))''')

                await cursor.execute('''CREATE TABLE IF NOT EXISTS groups 
                (group_name TEXT PRIMARY KEY,
                performance_list_url TEXT)''')

                await cursor.execute('''CREATE TABLE IF NOT EXISTS report_cards 
                (report_cards_url TEXT PRIMARY KEY)''')

                await cursor.execute('''CREATE TABLE IF NOT EXISTS retake_cards 
                (retake_cards_url TEXT PRIMARY KEY)''')

                await cursor.execute('''CREATE TABLE IF NOT EXISTS schedule 
                (schedule_url TEXT PRIMARY KEY)''')

                await cursor.execute('''CREATE TABLE IF NOT EXISTS work_schedule 
                (work_schedule_url TEXT PRIMARY KEY)''')

                await cursor.execute('''CREATE TABLE IF NOT EXISTS learning_schedule 
                                (learning_schedule_url TEXT PRIMARY KEY)''')

                await cursor.execute('''CREATE TABLE IF NOT EXISTS mailings 
                            (mailings_url TEXT PRIMARY KEY)''')
                
                await cursor.execute('''CREATE TABLE IF NOT EXISTS problem_reporting_email 
                            (email TEXT PRIMARY KEY)''')

                await db.commit()

    async def get_user_type(self, telegram_id):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT user_type FROM users WHERE telegram_id = ?", (telegram_id,))
                result = await cursor.fetchone()
                if result is not None:
                    return result[0]
                else:
                    return None

    async def users_exist(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM users")
                result = await cursor.fetchone()
                if result[0] == 0:
                    return False
                else:
                    return True

    async def add_user(self, fio, login, password, user_type, group_name=None):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                res = await cursor.execute(
                    "INSERT INTO 'users' (last_name, first_name, middle_name, login, password, group_name, "
                    "user_type)"
                    "VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (fio[0], fio[1], fio[2], login, password, group_name, user_type))
                await db.commit()
                return res

    async def connect_telegram_id(self, telegram_id, login):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                result = await cursor.execute("UPDATE 'users' SET telegram_id = ? WHERE login = ?",
                                              (telegram_id, login))
                await db.commit()
                return result

    async def delete_user_id(self, telegram_id):
        new_id = None
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                result = await cursor.execute("UPDATE 'users' SET telegram_id = ? WHERE telegram_id = ?",
                                              (new_id, telegram_id))
                await db.commit()
                return result

    async def check_auth_data(self, login, password):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT password FROM 'users' WHERE login = ?", (login,))
                check = await cursor.fetchone()
                if check is not None and check[0] == password:
                    result = "correct"
                else:
                    result = "incorrect"
                return result

    # Сделать норм return
    async def get_fio(self, telegram_id):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT last_name, first_name, middle_name FROM 'users' WHERE telegram_id = ?",
                                     (telegram_id,))
                fio = await cursor.fetchone()
                if fio[2] is None:
                    return fio[0] + " " + fio[1]
                else:
                    return fio[0] + " " + fio[1] + " " + fio[2]

    async def get_user_fn(self, telegram_id):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT first_name FROM 'users' WHERE telegram_id = ?", (telegram_id,))
                result = await cursor.fetchone()
                return result

    async def get_user_mn(self, telegram_id):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT middle_name FROM 'users' WHERE telegram_id = ?", (telegram_id,))
                result = await cursor.fetchone()
                return result

    async def get_user_ln(self, telegram_id):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT last_name FROM 'users' WHERE telegram_id = ?", (telegram_id,))
                result = await cursor.fetchone()
                return result

    async def get_login(self, telegram_id):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT login FROM 'users' WHERE telegram_id = ?", (telegram_id,))
                result = await cursor.fetchone()
                return result[0]

    async def get_report_cards_url(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT * FROM 'report_cards'")
                result = await cursor.fetchone()
                return result

    async def get_retake_cards_url(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT * FROM 'retake_cards'")
                result = await cursor.fetchone()
                return result
            
    async def get_problem_reporting_email(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT * FROM 'problem_reporting_email'")
                result = await cursor.fetchone()
                return result

    async def change_schedule_url(self, new_url):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'schedule'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    await cursor.execute("INSERT INTO 'schedule' ('schedule_url') VALUES (?)", (new_url,))
                else:
                    await cursor.execute(f"UPDATE 'schedule' SET 'schedule_url' = ?", (new_url,))
                await db.commit()

    async def change_work_schedule_url(self, new_url):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'work_schedule'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    await cursor.execute("INSERT INTO 'work_schedule' ('work_schedule_url') VALUES (?)", (new_url,))
                else:
                    await cursor.execute(f"UPDATE 'work_schedule' SET 'work_schedule_url' = ?", (new_url,))
                await db.commit()

    async def change_learning_schedule_url(self, new_url):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'learning_schedule'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    await cursor.execute("INSERT INTO 'learning_schedule' ('learning_schedule_url') VALUES (?)", (new_url,))
                else:
                    await cursor.execute(f"UPDATE 'learning_schedule' SET 'learning_schedule_url' = ?", (new_url,))
                await db.commit()

    async def change_report_cards_url(self, new_url):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'report_cards'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    await cursor.execute("INSERT INTO 'report_cards' ('report_cards_url') VALUES (?)", (new_url,))
                else:
                    await cursor.execute(f"UPDATE 'report_cards' SET 'report_cards_url' = ?", (new_url,))
                await db.commit()

    async def change_retake_cards_url(self, new_url):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'retake_cards'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    await cursor.execute("INSERT INTO 'retake_cards' ('retake_cards_url') VALUES (?)", (new_url,))
                else:
                    await cursor.execute(f"UPDATE 'retake_cards' SET 'retake_cards_url' = ?", (new_url,))
                await db.commit()

    async def change_mailings_url(self, new_url):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'mailings'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    await cursor.execute("INSERT INTO 'mailings' ('mailings_url') VALUES (?)", (new_url,))
                else:
                    await cursor.execute(f"UPDATE 'mailings' SET 'mailings_url' = ?", (new_url,))
                await db.commit()

    async def change_problem_reporting_email(self, new_email):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'problem_reporting_email'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    await cursor.execute("INSERT INTO 'problem_reporting_email' ('email') VALUES (?)", (new_email,))
                else:
                    await cursor.execute(f"UPDATE 'problem_reporting_email' SET 'email' = ?", (new_email,))
                await db.commit()

    async def get_user_group_name(self, telegram_id):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT group_name FROM 'users' WHERE telegram_id = ?", (telegram_id,))
                result = await cursor.fetchone()
                return result[0]

    async def get_performance_list_by_group_name(self, group_name):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT performance_list_url FROM 'groups' WHERE group_name = ?", (group_name,))
                result = await cursor.fetchone()
                return result[0]

    async def get_schedule_url(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'schedule'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    return 0
                else:
                    await cursor.execute("SELECT * FROM 'schedule'")
                    res = await cursor.fetchone()
                    return res[0]

    async def get_work_schedule_url(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'work_schedule'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    return 0
                else:
                    await cursor.execute("SELECT * FROM 'work_schedule'")
                    res = await cursor.fetchone()
                    return res[0]

    async def get_learning_schedule_url(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'learning_schedule'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    return 0
                else:
                    await cursor.execute("SELECT * FROM 'learning_schedule'")
                    res = await cursor.fetchone()
                    return res[0]

    async def get_mailings_url(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT COUNT(*) FROM 'mailings'")
                count = await cursor.fetchone()
                if count[0] == 0:
                    return 0
                else:
                    await cursor.execute("SELECT * FROM 'mailings'")
                    res = await cursor.fetchone()
                    return res[0]

    async def get_group_names(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT group_name FROM 'groups'")
                result = await cursor.fetchall()
                return result

    async def get_performance_urls(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT performance_list_url FROM 'groups'")
                result = await cursor.fetchall()
                return result

    async def get_group_name_by_url(self, performance_url):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT group_name FROM 'groups' WHERE performance_list_url = ?",
                                     (performance_url,))
                result = await cursor.fetchone()
                return result

    async def get_group_url(self, group_name):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT performance_list_url FROM 'groups' WHERE group_name = ?",
                                     (group_name,))
                result = await cursor.fetchone()
                return result

    async def add_group(self, group_name, performance_list_url):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("INSERT INTO 'groups' (group_name, performance_list_url) VALUES (?, ?)",
                                     (group_name, performance_list_url,))
                await db.commit()

    async def change_group_name(self, group_name, new_group_name):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                res = await cursor.execute("UPDATE 'groups' SET group_name = ? WHERE group_name = ?",
                                           (new_group_name, group_name))
                await db.commit()
                return res
            
    async def change_group_name_for_students(self, group_name, new_group_name):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                res = await cursor.execute("UPDATE 'users' SET group_name = ? WHERE group_name = ?",
                                           (new_group_name, group_name))
                await db.commit()
                return res
            
    async def change_performance_list_url(self, group_name, new_performance_list_url):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                res = await cursor.execute("UPDATE 'groups' SET performance_list_url = ? WHERE group_name = ?",
                                           (new_performance_list_url, group_name))
                await db.commit()
                return res

    async def get_user_names(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT last_name, first_name, middle_name FROM 'users'")
                result = await cursor.fetchall()
                return result

    async def get_managers(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT last_name, first_name, middle_name, login FROM 'users' WHERE "
                                     "user_type == 'manager' ORDER BY user_type")
                result = await cursor.fetchall()
                return result

    async def get_teachers(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT last_name, first_name, middle_name, login FROM 'users' WHERE "
                                     "user_type == 'teacher' ORDER BY user_type")
                result = await cursor.fetchall()
                return result

    # Переписать менее жижно
    async def get_user_id_by_name(self, user_name):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                if len(user_name) < 3:
                    await cursor.execute("SELECT telegram_id FROM 'users' WHERE last_name = ? AND first_name = ? AND "
                                         "middle_name IS NULL AND telegram_id IS NOT NULL", (user_name[0], user_name[1],))
                elif len(user_name) == 3:
                    if user_name[1][0] == "(":
                        await cursor.execute("SELECT telegram_id FROM 'users' WHERE last_name = ? AND first_name = ? AND "
                                            "middle_name IS NULL AND telegram_id IS NOT NULL", 
                                            (f"{user_name[0]} {user_name[1]}", user_name[2],))
                    else:
                        await cursor.execute("SELECT telegram_id FROM 'users' WHERE last_name = ? AND first_name = ? AND "
                                            "middle_name = ? AND telegram_id IS NOT NULL", 
                                            (user_name[0], user_name[1], user_name[2],))
                elif len(user_name) > 3:
                    if user_name[1][0] == "(":
                        last_name = f"{user_name[0]} {user_name[1]}"
                        first_name = ""
                        for i in range(2, len(user_name) - 1):
                            first_name += f"{user_name[i]} "
                        first_name = first_name[:-1]     
                        middle_name = user_name[-1]                                                             
                        await cursor.execute("SELECT telegram_id FROM 'users' WHERE last_name = ? AND first_name = ? AND "
                                            "middle_name = ? AND telegram_id IS NOT NULL", 
                                            (last_name, first_name, middle_name,))
                    else:
                        last_name = user_name[0]
                        first_name = ""
                        for i in range(1, len(user_name) - 1):
                            first_name += f"{user_name[i]} "
                        first_name = first_name[:-1]     
                        middle_name = user_name[-1]
                        await cursor.execute("SELECT telegram_id FROM 'users' WHERE last_name = ? AND first_name = ? AND "
                                            "middle_name = ? AND telegram_id IS NOT NULL", 
                                            (last_name, first_name, middle_name,))
                else:
                    return None
                result = await cursor.fetchone()
                return result

    async def change_password(self, telegram_id, new_password):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                res = await cursor.execute("UPDATE 'users' SET password = ? WHERE telegram_id = ?",
                                           (new_password, telegram_id,))
                await db.commit()
                return res

    async def get_students_by_group_name(self, group_name):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute(
                    "SELECT last_name, first_name, middle_name, login FROM 'users' WHERE group_name = ?",
                    (group_name,))
                result = await cursor.fetchall()
                return result

    async def get_student_names_by_group_name(self, group_name):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute(
                    "SELECT last_name, first_name, middle_name FROM 'users' WHERE group_name = ?",
                    (group_name,))
                result = await cursor.fetchall()
                return result

    async def get_logins(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT login FROM 'users'")
                result = await cursor.fetchall()
                return result

    async def change_login(self, telegram_id, new_login):  # Возможно не нужно ничего возвращать в этих функциях?
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                res = await cursor.execute("UPDATE 'users' SET login = ? WHERE telegram_id = ?",
                                           (new_login, telegram_id))
                await db.commit()
                return res

    async def change_user_type(self, telegram_id, new_user_type):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                res = await cursor.execute("UPDATE 'users' SET user_type = ? WHERE telegram_id = ?",
                                           (new_user_type, telegram_id))
                await db.commit()
                return res

    async def get_password(self, telegram_id):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT password FROM 'users' WHERE telegram_id = ?", (telegram_id,))
                result = await cursor.fetchone()
                return result[0]

    async def delete_group(self, group_name):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("DELETE FROM groups WHERE group_name=?", (group_name,))
                await cursor.execute("DELETE FROM users WHERE group_name=?", (group_name,))
                await db.commit()

    async def get_employees(self):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT last_name, first_name, middle_name, user_type FROM 'users' WHERE "
                                     "user_type != 'student'  ORDER BY user_type")
                result = await cursor.fetchall()
                return result

    async def delete_user(self, user):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("DELETE FROM users WHERE last_name=? AND first_name=? AND middle_name=?",
                                     (user[0], user[1], user[2],))
                await db.commit()
        
    async def delete_user_by_login(self, login):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("DELETE FROM users WHERE login=?",
                                     (login,))
                await db.commit()

    async def delete_user_without_mn(self, user):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("DELETE FROM users WHERE last_name=? AND first_name=?", (user[0], user[1],))
                await db.commit()

    async def get_users_by_type(self, role):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT last_name, first_name, middle_name FROM 'users' WHERE user_type = ?",
                                     (role,))
                result = await cursor.fetchall()
                return result

    async def get_user_ids_by_group(self, group_name):
        async with aiosqlite.connect(self.db_name) as db:
            async with db.cursor() as cursor:
                await cursor.execute("SELECT telegram_id FROM 'users' WHERE group_name = ? AND "
                                     "telegram_id IS NOT NULL", (group_name,))
                result = await cursor.fetchall()
                return result
