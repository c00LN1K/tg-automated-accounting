import os
import re
import shutil
import traceback
import datetime
from http.client import RemoteDisconnected
from time import sleep
import requests
import telebot
from telebot import types
from telebot.types import InputFile
import config
from db import Task
from work_with_excel import *

TOKEN = config.TOKEN

bot = telebot.TeleBot(TOKEN)
number_starts = 0


def restricted(func):
    def wrapped(message):
        user_id = message.from_user.id
        if str(user_id) in config.ALLOWED_USERS:
            func(message)
        else:
            bot.send_message(message.chat.id, "Извините, у вас нет доступа к этому боту.")

    return wrapped


@bot.message_handler(commands=['start'])
@restricted
def menu_handler(message):
    keyboard = types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
    add_task_button = types.KeyboardButton(text="Добавить запись💸")
    check_tasks_button = types.KeyboardButton(text="Просмотреть активные👀")
    download_button = types.KeyboardButton(text="Скачать excel📝")
    keyboard.add(add_task_button, check_tasks_button, download_button)
    bot.send_message(message.chat.id, 'Добро пожаловать в главное меню. Выберите желаемое действие',
                     reply_markup=keyboard)


@bot.message_handler(func=lambda message: message.text.lower() in ('добавить запись', 'добавить запись💸'))
@restricted
def add_task_handler(message):
    task = []
    keyboard = types.ReplyKeyboardMarkup(row_width=3, resize_keyboard=True)
    today_button = types.KeyboardButton(text='Сегодня')
    tomorrow_button = types.KeyboardButton(text='Завтра')
    another_button = types.KeyboardButton(text='Другое')
    menu_button = types.KeyboardButton(text='Главное меню')
    keyboard.add(today_button, tomorrow_button, another_button, menu_button)
    bot.send_message(message.chat.id, "Укажите дату планируемой стирки:", reply_markup=keyboard)
    bot.register_next_step_handler(message, get_date, task)


def get_date(message, task: list):
    keyboard = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    button = types.KeyboardButton('Главное меню')
    keyboard.add(button)
    date = ''
    if message.text == 'Сегодня':
        date = datetime.date.today()
    elif message.text == 'Завтра':
        date = datetime.date.today() + datetime.timedelta(days=1)
    elif message.text == 'Другое':
        bot.send_message(message.chat.id, 'Укажите, пожалуйста, дату в формате: DD.MM.YYYY. Например, 30.01.2024',
                         reply_markup=keyboard)
        bot.send_message(message.chat.id,
                         'Вы также можете указать только день, тогда будет взять текущий месяц и год. Или указать день и месяц, тогда - только текущий год.\n'
                         'Но будьте внимательны при указании дня. В разных месяцах может содержаться разное число дней')
        bot.register_next_step_handler(message, get_date, task)
    else:
        if message.text.lower() == 'главное меню':
            menu_handler(message)
            return
        try:
            if message.text.count('.') == 2:
                date = datetime.datetime.strptime(message.text, '%d.%m.%Y').date()
            elif message.text.count('.') == 1:
                d = datetime.datetime.strptime(message.text, '%d.%m').date()
                date = d.replace(year=datetime.datetime.now().year)
            else:
                d = datetime.datetime.strptime(message.text, '%d').date()
                date = d.replace(year=datetime.datetime.now().year, month=datetime.datetime.now().month)

        except ValueError as _ex:
            bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
            bot.register_next_step_handler(message, get_date, task)

    if date:
        task.append(date)
        bot.send_message(message.chat.id, 'Введите имя заказчика', reply_markup=keyboard)
        bot.register_next_step_handler(message, get_name, task)


def get_name(message, task: list):
    if message.text == 'Главное меню':
        menu_handler(message)
    else:
        task.append(message.text)
        bot.send_message(message.chat.id, 'Введите время начала стирки в формате HH:MM. Например, 12:00')
        bot.register_next_step_handler(message, get_time, task)


def get_time(message, task: list):
    if message.text == 'Главное меню':
        menu_handler(message)
    elif re.fullmatch(r'[0-2][0-9]:[0-9][0-9]', message.text):
        hour, minute = map(int, message.text.split(":"))
        task.append(datetime.time(hour=hour, minute=minute))
        keyboard = types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
        button30 = types.KeyboardButton(text='30')
        button60 = types.KeyboardButton(text='60')
        menu_button = types.KeyboardButton(text='Главное меню')
        keyboard.add(button30, button60, menu_button)
        bot.send_message(message.chat.id, 'Выберете или введите длительность стирки в минутах', reply_markup=keyboard)
        bot.register_next_step_handler(message, get_duration, task)
    else:
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_time, task)


def get_duration(message, task: list):
    if message.text == 'Главное меню':
        menu_handler(message)
    elif message.text.isdigit():
        task.append(int(message.text))
        keyboard = types.ReplyKeyboardMarkup(row_width=3, resize_keyboard=True)
        button50 = types.KeyboardButton(text='50')
        button85 = types.KeyboardButton(text='85')
        button100 = types.KeyboardButton(text='100')
        menu_button = types.KeyboardButton(text='Главное меню')
        keyboard.add(button50, button85, button100, menu_button)
        bot.send_message(message.chat.id, 'Выберете или введите свою стоимость стирки', reply_markup=keyboard)
        bot.register_next_step_handler(message, get_price, task)
    else:
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_duration, task)


def get_price(message, task: list):
    if message.text == 'Главное меню':
        menu_handler(message)
    elif message.text.isdigit():
        task.append(int(message.text))
        add_task(message, task)
    else:
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_price, task)


def add_task(message, task: list):
    print(task)
    res = Task.add_task(task)
    out = f'Ура! Запись на {task[0]} в {task[2].strftime("%H:%M")} для {task[1]} со стоимостью {task[4]} руб. успешно создана' if res \
        else 'Ой, кажется произошла ошибка при создании записи. Возможно, заказ на данное время и дату уже существует'
    keyboard = types.ReplyKeyboardRemove()
    bot.send_message(message.chat.id, out, reply_markup=keyboard)
    menu_handler(message)


@bot.message_handler(func=lambda message: message.text.lower() in ('просмотреть активные', 'просмотреть активные👀'))
@restricted
def check_active(message):
    active_tasks = Task.get_unfinished_tasks()
    if len(active_tasks) == 0:
        bot.send_message(message.chat.id,
                         'Ой, кажется, у вас нет активных заказов. Хорошая или плохая эта новость, решать вам)')
        menu_handler(message)
        return

    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    menu_button = types.KeyboardButton(text='Главное меню')
    keyboard.add(menu_button)
    bot.send_message(message.chat.id, 'Вот список активных заказов:🙊', reply_markup=keyboard)

    for data in active_tasks:
        index_task = data.id
        keyboard = types.InlineKeyboardMarkup()
        delete_button = types.InlineKeyboardButton(text='Удалить', callback_data=f'del_{index_task}')
        finish_button = types.InlineKeyboardButton(text='Завершить', callback_data=f'fin_{index_task}')
        keyboard.add(finish_button, delete_button)
        timee = (datetime.datetime.combine(datetime.datetime.today(), data.time) + datetime.timedelta(
            minutes=data.duration)).time()
        data_out = (
            f"Дата: {data.day}.{data.month}.{data.year} ({WEEKDAY[data.month]})\n"
            f"Имя: {data.name}\n"
            f"Длительность: {data.duration} мин.\n"
            f"Время: {data.time.strftime('%H:%M')}-{timee.strftime('%H:%M')}\n"
            f"Стоимость: {data.price}")
        bot.send_message(message.chat.id, data_out, reply_markup=keyboard)


@bot.callback_query_handler(func=lambda call: True)
def active_btn(call):
    message = call.message
    action, index = call.data.split('_')

    if action == 'del':
        Task.delete_task(int(index))
        bot.edit_message_text(chat_id=message.chat.id, message_id=message.id, text=f'Заказ успешно удален')
    else:
        Task.finish_task(int(index))
        bot.edit_message_text(chat_id=message.chat.id, message_id=message.id, text=f'Заказ успешно завершен')


@bot.message_handler(func=lambda message: message.text.lower() in ('скачать excel', 'скачать excel📝'))
@restricted
def send_excel(message):
    keyboard = types.ReplyKeyboardMarkup(row_width=3, resize_keyboard=True)
    day_button = types.KeyboardButton(text='День')
    week_button = types.KeyboardButton(text='Месяц')
    year_button = types.KeyboardButton(text='Год')
    all_button = types.KeyboardButton(text='Все')
    menu_button = types.KeyboardButton(text='Главное меню')
    keyboard.add(day_button, week_button, year_button, all_button, menu_button)
    bot.send_message(message.chat.id, 'За какой промежуток вы хотите получить excel?\nВыберите подходящий вариант:',
                     reply_markup=keyboard)
    bot.register_next_step_handler(message, get_period)


def get_period(message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    menu_button = types.KeyboardButton(text='Главное меню')
    keyboard.add(menu_button)
    if message.text == 'День':
        bot.send_message(message.chat.id, 'Укажите, пожалуйста, дату в формате: DD.MM.YYYY. Например, 30.01.2024',
                         reply_markup=keyboard)
        bot.send_message(message.chat.id,
                         'Вы можете указать только день, тогда будет взять текущий месяц и год. Или указать только день и месяц, тогда - только текущий год')
        bot.register_next_step_handler(message, get_day)
    elif message.text == 'Месяц':
        bot.send_message(message.chat.id, 'Напишите название или номер месяца, который вас интересует :)', reply_markup=keyboard)
        bot.register_next_step_handler(message, get_month)
    elif message.text == 'Год':
        bot.send_message(message.chat.id, 'Укажите интересующий вас год -_-', reply_markup=keyboard)
        bot.register_next_step_handler(message, get_year)
    elif message.text == 'Все':
        filename = create_all_excel_file()
        check_and_send_file(message, filename)
    elif message.text.lower() == 'главное меню':
        menu_handler(message)
    else:
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_period)


def get_day(message):
    if message.text == 'Главное меню':
        menu_handler(message)
        return
    try:
        if message.text.count('.') == 2:
            date = datetime.datetime.strptime(message.text, '%d.%m.%Y').date()
        elif message.text.count('.') == 1:
            date = datetime.datetime.strptime(message.text, '%d.%m').date()
            date = date.replace(year=datetime.datetime.now().year)
        else:
            date = datetime.datetime.strptime(message.text, '%d').date()
            date = date.replace(year=datetime.datetime.now().year, month=datetime.datetime.now().month)

        filename = create_one_day_excel_file(date)
        check_and_send_file(message, filename)

    except ValueError as _ex:
        traceback.print_tb(_ex.__traceback__)
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_day)


def get_month(message):
    if message.text.lower() == 'главное меню':
        menu_handler(message)
        return
    elif message.text.lower() in MONTH or (message.text.isdigit() and 1 <= int(message.text.isdigit()) <= 12):
        if message.text.isdigit():
            month = int(message.text)
        else:
            month = MONTH.index(message.text.lower())
        filename = create_month_excel_file(datetime.date.today().replace(month=month))
        check_and_send_file(message, filename)

    else:
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_month)


def get_year(message):
    if message.text.lower() == 'главное меню':
        menu_handler(message)
        return
    elif message.text.isdigit():
        year = int(message.text.lower())
        filename = create_year_excel_file(datetime.date.today().replace(year=year))
        check_and_send_file(message, filename)
    else:
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_year)


@bot.message_handler(content_types=['text'])
@restricted
def bot_message(message):
    if message.text.lower() in ('главное меню', 'меню', 'menu'):
        menu_handler(message)
    else:
        bot.send_message(message.chat.id, "Ой, я вас немного недопонял👉👈. Убедитесь, пожалуйста, в корректности ввода")


def check_and_send_file(message, filename):
    if filename is None:
        bot.send_message(message.chat.id, 'Упс. Кажется, нет данных за этот период')
    else:
        send_doc(message.chat.id, filename=filename)
        os.remove(filename)
    menu_handler(message)


def send_doc(chat_id, filename, attempts=3):
    try:
        bot.send_message(chat_id, 'Пытаюсь получить файлы...')
        if attempts:
            bot.send_document(chat_id, InputFile(filename))
        else:
            bot.send_message(chat_id,
                             'Ой, похоже какая проблема с Интернетом. Попробуйте запросить excel немного попозже')
    except Exception as _ex:
        if attempts == 0:
            bot.send_message(chat_id, 'Упс, похоже какая проблема с Интернетом. Подождите 30 сек и попробуйте ещё раз.')
            raise
        else:
            send_doc(chat_id, filename, attempts - 1)


def write_about_error(error):
    # check_root()
    with open('error.txt', 'a') as file:
        file.write(
            f'{datetime.datetime.now().isoformat()} --> {error}\n\n')


# Запускаем бота
while True:
    try:
        # write_about_error('Bot is running!')
        # get_active_tasks()
        # check_excel_dir()
        print('Бот запущен!')
        bot.polling(none_stop=True, interval=0)
    except (requests.exceptions.ConnectionError, RemoteDisconnected) as ex:
        traceback.print_tb(ex.__traceback__)
        # write_about_error(traceback.format_exc())
        bot.send_message(config.FEEDBACK_ID, f'Ошибка с соединением')
        sleep(30)

    except Exception as _ex:
        traceback.print_tb(_ex.__traceback__)
        # write_about_error(traceback.format_exc())
        # save_active_tasks()
        # make_zip()
        # bot.send_document(config.FEEDBACK_ID, open('error.txt', 'rb'))
        # bot.send_document(config.FEEDBACK_ID, open('excel.zip', 'rb'))
        sleep(5)
