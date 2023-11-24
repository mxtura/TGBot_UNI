from aiogram.types import Message, CallbackQuery
from aiogram.dispatcher import FSMContext

import logging
import openpyxl
import io
from urllib.request import urlopen
import tgbot.keyboards.reply as rkb
from tgbot.filters.user_type import UserTypeFilter
from tgbot.models.database_instance import db
from tgbot.misc.states import RegRetakeFSM
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from oauth2client.service_account import ServiceAccountCredentials
import httplib2



faculty_request_text = "Пожалуйста, напишите название факультета:"
direction_request_text = "Пожалуйста, напишите название направления:"
group_request_text = "Пожалуйста, напишите номер группы:"
record_book_request_text = "Пожалуйста, напишите номер зачетки:"
subject_request_text = "Пожалуйста, напишите название предмета:"
_request_text = "Пожалуйста, напишите тип контроля:"
teacher_request_text = "Пожалуйста, напишите ФИО преподавателя:"
date_request_text = "Пожалуйста, напишите дату сдачи:"
time_request_text = "Пожалуйста, напишите время сдачи:"

logger = logging.getLogger(__name__)
retakes_data = []

async def getting_faculty(message: Message, state: FSMContext):
    """Просит ввести название факультета"""
    retakes_data.clear()
    await message.answer(text=faculty_request_text,
                                   reply_markup=rkb.faculty_input_cancel_keyboard)
    
    await RegRetakeFSM.faculty.set()
    

async def getting_direction(message: Message, state: FSMContext):
    
    await state.update_data(faculty=message.text)
    
    del_msg = await message.answer(text=direction_request_text,
                                   reply_markup=rkb.direction_input_cancel_keyboard)
    
    await RegRetakeFSM.direction.set()
    

async def getting_group(message: Message, state: FSMContext):
    
    await state.update_data(direction=message.text)  
    
    del_msg = await message.answer(text=group_request_text,
                                   reply_markup=rkb.group_input_cancel_keyboard)
    
    await RegRetakeFSM.group.set()

async def getting_record_book(message: Message, state: FSMContext):
    
    await state.update_data(group=message.text)  
            
    del_msg = await message.answer(text=record_book_request_text,
                                   reply_markup=rkb.recod_book_input_cancel_keyboard)

    await RegRetakeFSM.record_book.set()
    
    
async def getting_subject(message: Message, state: FSMContext):
    
    await state.update_data(record_book=message.text)  
    

    del_msg = await message.answer(text=subject_request_text,
                                   reply_markup=rkb.subject_input_cancel_keyboard)
    
    await RegRetakeFSM.subject.set()
    
    

async def getting_teacher(message: Message, state: FSMContext):

    await state.update_data(subject=message.text)  
    
    
    del_msg = await message.answer(text=teacher_request_text,
                                   reply_markup=rkb.name_input_cancel_keyboard)

    await RegRetakeFSM.teacher.set()
    
    
async def getting_date(message: Message, state: FSMContext):

    # fio = message.text.split()

    # # Проверка количества элементов в ФИО
    # if len(fio) >= 3:
    #     teacher_fn, teacher_mn, teacher_ln = fio[:3]
        
    # elif len(fio) == 2:
    #     teacher_fn, teacher_ln = fio
    #     teacher_mn = ""
    # else:
    #     # Отправить сообщение об ошибке пользователю
    #     await message.answer("Пожалуйста, введите ФИО полностью.")
    #     return
    
    
    await state.update_data(teacher=message.text)  
    

    del_msg = await message.answer(text=date_request_text,
                                   reply_markup=rkb.date_input_cancel_keyboard)
    
    await RegRetakeFSM.date.set()
    

async def getting_time(message: Message, state: FSMContext):

    await state.update_data(date=message.text)  
    

    del_msg = await message.answer(text=time_request_text,
                                   reply_markup=rkb.time_input_cancel_keyboard)

    await RegRetakeFSM.time.set()
    

async def agreement(message: Message, state: FSMContext):
    await state.update_data(time=message.text)  

    teacher_data = await state.get_data()

    await message.answer(teacher_data["faculty"])
    await message.answer(teacher_data["direction"])
    await message.answer(teacher_data["group"])
    await message.answer(teacher_data["record_book"])
    await message.answer(teacher_data["subject"])
    await message.answer(teacher_data["teacher"])
    await message.answer(teacher_data["date"])
    await message.answer(teacher_data["time"])

    await RegRetakeFSM.agreement.set()


async def write_retakes_list(message: Message, state: FSMContext):
    """Записывает данные о пересдачах в excel-файл на google drive"""
    
    student_fn = str(await db.get_user_fn(message.from_user.id)).strip("(,')")
    student_mn = str(await db.get_user_mn(message.from_user.id)).strip("(,')")
    student_ln = str(await db.get_user_ln(message.from_user.id)).strip("(,')")

    retakes_url = await db.get_retake_cards_url()
    file_id = retakes_url[0].split('/')[-2]
    file_url = f'https://drive.google.com/u/0/uc?id={file_id}&export=download'
    fn = io.BytesIO(urlopen(file_url).read())
    wb = openpyxl.load_workbook(fn)
    
    ws = wb["Общие"]
    logger.info(ws.title)

    teacher_data = await state.get_data()
    
    
    
    ws.append([teacher_data["faculty"], teacher_data["direction"], teacher_data["group"][0], teacher_data["group"], student_fn, student_mn, student_ln, teacher_data["record_book"], teacher_data["subject"], teacher_data["teacher"].split()[0], teacher_data["teacher"].split()[1], teacher_data["teacher"].split()[2], teacher_data["date"], teacher_data["time"]])
    
    # Сохранение файла локально
    local_filename = "temp.xlsx"
    wb.save(local_filename)
    wb.close()

    # Загрузка файла обратно на Google Drive
    credentials = ServiceAccountCredentials.from_json_keyfile_name('path/to/your/credentials.json', ['https://www.googleapis.com/auth/drive'])
    http = credentials.authorize(httplib2.Http())
    drive_service = build('drive', 'v3', http=http)

    file_metadata = {'name': 'updated_file.xlsx'}
    media = MediaFileUpload(local_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    drive_service.files().update(fileId=file_id, body=file_metadata, media_body=media).execute()

    await message.answer("Данные успешно записаны")



def register_reg_retake(dp):
    dp.register_message_handler(getting_faculty, UserTypeFilter("student"), content_types=['text'],
                                text=['Записаться на пересдачу'])
    dp.register_message_handler(getting_direction, state=RegRetakeFSM.faculty)
    dp.register_message_handler(getting_group, state=RegRetakeFSM.direction)
    dp.register_message_handler(getting_record_book, state=RegRetakeFSM.group)
    dp.register_message_handler(getting_subject, state=RegRetakeFSM.record_book)
    dp.register_message_handler(getting_teacher, state=RegRetakeFSM.subject)
    dp.register_message_handler(getting_date, state=RegRetakeFSM.teacher)
    dp.register_message_handler(getting_time, state=RegRetakeFSM.date)
    dp.register_message_handler(agreement, state=RegRetakeFSM.time)
    dp.register_message_handler(write_retakes_list, state=RegRetakeFSM.agreement)
    
