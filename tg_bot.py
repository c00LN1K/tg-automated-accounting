import re
import shutil
import traceback
from time import sleep

import telebot
from telebot import types

from work_with_excel import *

TOKEN = '6054870053:AAFlMbXn4Our2zsWm_M9Loahu5DH7Zt0xUg'

bot = telebot.TeleBot(TOKEN)
number_starts = 0


@bot.message_handler(commands=['start'])
def menu_handler(message):
    keyboard = types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
    add_task_button = types.KeyboardButton(text="Добавить запись")
    check_tasks_button = types.KeyboardButton(text="Просмотреть активные")
    download_button = types.KeyboardButton(text="Скачать excel")
    keyboard.add(add_task_button, check_tasks_button, download_button)
    bot.send_message(message.chat.id, 'Добро пожаловать в главное меню. Выберите желаемое действие',
                     reply_markup=keyboard)


@bot.message_handler(func=lambda message: message.text.lower() == 'добавить запись')
def add_task_handler(message):
    task = []
    keyboard = types.ReplyKeyboardMarkup(row_width=3)
    today_button = types.KeyboardButton(text='Сегодня')
    tomorrow_button = types.KeyboardButton(text='Завтра')
    another_button = types.KeyboardButton(text='Другое')
    menu_button = types.KeyboardButton(text='Главное меню')
    keyboard.add(today_button, tomorrow_button, another_button, menu_button)
    bot.send_message(message.chat.id, "Укажите дату планируемой стирки:", reply_markup=keyboard)
    bot.register_next_step_handler(message, get_date, task)


def get_date(message, task: list):
    keyboard = types.ReplyKeyboardMarkup(row_width=1)
    button = types.KeyboardButton('Главное меню')
    keyboard.add(button)
    date = ''
    if message.text == 'Сегодня':
        date = datetime.date.today()
    elif message.text == 'Завтра':
        date = datetime.date.today() + datetime.timedelta(days=1)
    elif message.text == 'Другое':
        # TODO: доработать логику (указан год или нет и тд)
        bot.send_message(message.chat.id, 'Укажите, пожалуйста, дату в формате: DD.MM.YYYY. Например, 30.01.2024',
                         reply_markup=keyboard)
        bot.send_message(message.chat.id,
                         'Вы также можете указать только день, тогда будет взять текущий месяц и год. Или указать только день и месяц, тогда - только текущий год')
        bot.register_next_step_handler(message, get_date, task)
    else:
        if message.text == 'Главное меню':
            menu_handler(message)
            return
        try:
            if message.text.count('.') == 2:
                date = datetime.datetime.strptime(message.text, '%d.%m.%Y')
            elif message.text.count('.') == 1:
                date = datetime.datetime.strptime(message.text, '%d.%m')
                date = date.replace(year=datetime.datetime.now().year)
            else:
                date = datetime.datetime.strptime(message.text, '%d')
                date = date.replace(year=datetime.datetime.now().year, month=datetime.datetime.now().month)

        except ValueError:
            bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
            bot.register_next_step_handler(message, get_date)

    if date:
        task.append(date.strftime('%Y-%m-%d'))
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
        task.append(message.text)
        keyboard = types.ReplyKeyboardMarkup(row_width=3, resize_keyboard=True)
        button50 = types.KeyboardButton(text='50')
        button85 = types.KeyboardButton(text='85')
        button100 = types.KeyboardButton(text='100')
        menu_button = types.KeyboardButton(text='Главное меню')
        keyboard.add(button50, button85, button100, menu_button)
        bot.send_message(message.chat.id, 'Выберете или введите свою стоимость стирки', reply_markup=keyboard)
        bot.register_next_step_handler(message, get_price, task)

        # bot.send_message(message.chat.id, 'Введите длительность стирки в минутах')
        # bot.register_next_step_handler(message, get_duration, task)
    else:
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_time, task)


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
    task.append(False)
    ans = write_task(task)
    out = f'Ура! Запись на {task[0]} в {task[2]} для {task[1]} со стоимостью {task[3]} руб. успешно создана' if ans \
        else 'Ой, кажется произошла ошибка при создании записи. Возможно, данный такой заказ уже существует'
    keyboard = types.ReplyKeyboardRemove()
    bot.send_message(message.chat.id, out, reply_markup=keyboard)
    menu_handler(message)


@bot.message_handler(func=lambda message: message.text.lower() == 'просмотреть активные')
def check_active(message):
    if len(active_tasks) == 0:
        bot.send_message(message.chat.id,
                         'Ой, кажется, у вас нет активных заказов. Хорошая или плохая эта новость, решать вам)')
        menu_handler(message)
        return
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    menu_button = types.KeyboardButton(text='Главное меню')
    keyboard.add(menu_button)
    bot.send_message(message.chat.id, 'Вот список активных заказов:', reply_markup=keyboard)

    for index_task, data in active_tasks.items():
        keyboard = types.InlineKeyboardMarkup()
        delete_button = types.InlineKeyboardButton(text='Удалить', callback_data=f'del_{index_task}')
        finish_button = types.InlineKeyboardButton(text='Завершить', callback_data=f'fin_{index_task}')
        keyboard.add(finish_button, delete_button)
        data_out = f'Дата: {data[0]}\nИмя: {data[1]}\nВремя: {data[2]}\nСтоимость: {data[3]}'
        bot.send_message(message.chat.id, data_out, reply_markup=keyboard)


@bot.callback_query_handler(func=lambda call: True)
def active_btn(call):
    message = call.message
    action, index = call.data.split('_')
    task = active_tasks.pop(int(index))
    change_status(action, task)
    res = 'удален' if action == 'del' else 'завершен'

    bot.edit_message_text(chat_id=message.chat.id, message_id=message.id, text=f'Заказ успешно {res}')


@bot.message_handler(func=lambda message: message.text.lower() == 'скачать excel')
def send_excel(message):
    # TODO: продумать возврат (всю папку, конкретный месяц или файл)
    if 'excel' not in os.listdir(ROOT):
        bot.send_message(message.chat.id,
                         'Ой, похоже мне нечего вам прислать. Возможно, вы ещё не создали ни одной записи. Так чего же вы ждёте)')
        menu_handler(message)
        return

    keyboard = types.ReplyKeyboardMarkup(row_width=3, resize_keyboard=True)
    day_button = types.KeyboardButton(text='День')
    week_button = types.KeyboardButton(text='Месяц')
    year_button = types.KeyboardButton(text='Год')
    all_button = types.KeyboardButton(text='Все')
    menu_button = types.KeyboardButton(text='Главное меню')
    keyboard.add(day_button, week_button, year_button, all_button, menu_button)
    bot.send_message(message.chat.id, 'За какой промежуток вы хотите получить excel? Выберите подходящий вариант:',
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
        bot.send_message(message.chat.id, 'Напишите какой месяц вас интересует :)')
        bot.register_next_step_handler(message, get_month)
    elif message.text == 'Год':
        bot.send_message(message.chat.id, 'Укажите интересующий вас год -_-')
        bot.register_next_step_handler(message, get_year)
    elif message.text == 'Все':
        make_zip()
        bot.send_document(message.chat.id, open('excel.zip', 'rb'))
    elif message.text != 'Главное меню':
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_period)


def get_day(message):
    if message.text == 'Главное меню':
        menu_handler(message)
        return
    try:
        if message.text.count('.') == 2:
            date = datetime.datetime.strptime(message.text, '%d.%m.%Y')
        elif message.text.count('.') == 1:
            date = datetime.datetime.strptime(message.text, '%d.%m')
            date = date.replace(year=datetime.datetime.now().year)
        else:
            date = datetime.datetime.strptime(message.text, '%d')
            date = date.replace(year=datetime.datetime.now().year, month=datetime.datetime.now().month)

        filename = get_filename(date)
        if filename:
            bot.send_document(message.chat.id, filename)
        else:
            bot.send_message(message.chat.id, 'Похоже, что такого файла за такой период нет')

    except ValueError:
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_day)


def get_month(message):
    if message.text == 'Главное меню':
        menu_handler(message)
        return
    try:
        if message.text.count('.') == 2:
            date = datetime.datetime.strptime(message.text, '%d.%m.%Y')
        elif message.text.count('.') == 1:
            date = datetime.datetime.strptime(message.text, '%d.%m')
            date = date.replace(year=datetime.datetime.now().year)
        else:
            date = datetime.datetime.strptime(message.text, '%d')
            date = date.replace(year=datetime.datetime.now().year, month=datetime.datetime.now().month)
        return date

    except ValueError:
        bot.send_message(message.chat.id, 'Похоже возникла ошибка. Проверьте корректность вводимых данных')
        bot.register_next_step_handler(message, get_day)


@bot.message_handler(content_types=['text'])
def bot_message(message):
    if message.text.lower() in ('главное меню', 'меню', 'menu'):
        menu_handler(message)
    else:
        bot.send_message(message.chat.id, "Ой, я вас немного недопонял👉👈. Убедитесь, пожалуйста, в корректности ввода")


def check_error_txt():
    f = open('error.txt', 'a')
    f.write(f'{datetime.datetime.now().isoformat()} --> Bot is running!\n\n')
    f.close()


def check_excel_dir():
    if 'excel' not in os.listdir():
        os.mkdir('excel')
    else:
        if 'achieve' not in os.listdir():
            os.mkdir('achieve')
        try:
            shutil.move('excel', 'achieve')
        except:
            shutil.rmtree('achieve\\excel')
            shutil.move('excel', 'achieve')
        finally:
            os.mkdir('excel')


def make_zip():
    check_root()
    if 'excel.zip' in os.listdir():
        shutil.rmtree('excel.zip')
    shutil.make_archive('excel', 'zip', 'excel')


# Запускаем бота
while True:
    try:
        check_root()
        check_error_txt()
        check_excel_dir()
        print('Бот запущен!')
        bot.polling(none_stop=True, interval=0)
    except Exception as _ex:
        traceback.print_tb(_ex.__traceback__)
        make_zip()
        with open('error.txt', 'a') as file:
            file.write(
                f'{datetime.datetime.now().isoformat()} --> {traceback.format_exc()}\n\n')
        bot.send_message('1081588278', f'Ошибка')
        bot.send_document('1081588278', open('error.txt', 'rb'))
        bot.send_document('1081588278', open('excel.zip', 'rb'))
        shutil.rmtree('excel', ignore_errors=True)
        sleep(60)
