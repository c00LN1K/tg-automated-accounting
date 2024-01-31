import datetime
import openpyxl
import os

import numpy as np
import pandas as pd
from openpyxl.styles import Font

MONTH = ['', 'январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь',
         'декабрь']

ROOT = os.getcwd()
active_tasks = {}
last_index = 0


def check_root():
    if os.getcwd() != ROOT:
        os.chdir(ROOT)


def get_filename(date, is_create=True):
    check_root()
    dir_year = f'{date.year}'
    # Дисбаланс из-за того, что мы берем месяц от переданного дня, а начало-конец недели может быть в разных месяцах
    # Из-за этого получается дублирование. И одном и другом месяце будут содержаться те же недели
    # Решение, брать месяц из первого дня недели
    first_day_week = date - datetime.timedelta(days=date.weekday())
    dir_month = f'{MONTH[first_day_week.month]}'

    last_day_week = (date + datetime.timedelta(days=6 - date.weekday())).strftime("%d.%m")
    excel_week = f'{first_day_week.strftime("%d.%m")}-{last_day_week}' + '.xlsx'
    sheet = f'{date.day:02}'

    if 'excel' not in os.listdir():
        raise Exception('No such directory excel!')
    os.chdir('excel')

    if dir_year not in os.listdir():
        if not is_create:
            return
        os.mkdir(dir_year)
    os.chdir(dir_year)

    if dir_month not in os.listdir():
        if not is_create:
            return
        os.mkdir(dir_month)
    os.chdir(dir_month)

    if excel_week not in os.listdir():
        if not is_create:
            return
        workbook = openpyxl.Workbook()
        ws = workbook.active
        ws.title = sheet
        workbook.save(excel_week)
        df = pd.DataFrame({'Дата': [], 'Имя': [], 'Время': [], 'Стоимость': [], 'Оплачено': []})
        with pd.ExcelWriter(os.getcwd() + '\\' + excel_week, mode='a', engine='openpyxl',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

    wb = openpyxl.load_workbook(excel_week)
    if sheet not in wb.sheetnames:
        if not is_create:
            return
        wb.save(excel_week)
        df = pd.DataFrame({'Дата': [], 'Имя': [], 'Время': [], 'Стоимость': [], 'Оплачено': []})
        with pd.ExcelWriter(os.getcwd() + '\\' + excel_week, mode='a', engine='openpyxl',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

    return os.getcwd() + '\\' + excel_week


def write_task(task: list):
    """
    Функция, которая добавляет новую запись
    """
    date = task[0]
    print(task)
    file_name = get_filename(datetime.datetime.fromisoformat(date))
    sheet = date[-2:]

    # Чтение данных из листа
    df = pd.read_excel(file_name, sheet_name=sheet, usecols="A:E")
    summ = np.sum(df[df['Оплачено'] == True]['Стоимость'])
    # Проверка наличия задачи в DataFrame
    mask = (df == task).all(axis=1)
    if mask.any():
        print("Такая задача уже существует.")
        return False
    else:
        # Добавление новой строки в DataFrame
        df.loc[len(df)] = task
        df.sort_values(by='Время', inplace=True)

        # Записываем DataFrame в тот же лист
        with pd.ExcelWriter(file_name, mode='a', engine='openpyxl',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

    wb = openpyxl.load_workbook(filename=file_name)
    ws = wb.get_sheet_by_name(sheet)
    ws['H1'] = 'Сумма'
    ws['H1'].font = Font(bold=True)
    ws['H2'] = summ
    wb.save(file_name)

    global last_index
    active_tasks[last_index] = task
    last_index += 1
    print('Обновление индекса:', last_index)
    return True


def change_status(status, task: list):
    """
    Функция, которая будет изменять статус записи (удалить или завершить)
    """

    date = task[0]
    file_name = get_filename(datetime.datetime.fromisoformat(date))
    sheet = date[-2:]
    summ = 0
    with pd.ExcelWriter(file_name, engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as writer:
        df = pd.read_excel(file_name, sheet_name=sheet, usecols="A:E")
        mask = (df == task).all(axis=1)
        if mask.any():
            location = df.index[mask].tolist()
            print(df.loc[location], location)
            if status == 'del':
                df.drop(location, inplace=True)
            else:
                df.loc[location, 'Оплачено'] = True

            summ = np.sum(df[df['Оплачено'] == True]['Стоимость'])
            df.to_excel(writer, sheet_name=sheet, index=False)
        else:
            print("Строка не найдена в DataFrame.")

    if summ:
        wb = openpyxl.load_workbook(filename=file_name)
        ws = wb.get_sheet_by_name(sheet)
        ws['H1'] = 'Сумма'
        ws['H1'].font = Font(bold=True)
        ws['H2'] = summ
        wb.save(file_name)
