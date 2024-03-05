import os
import shutil
import sqlite3
import pythoncom
import win32com.client as win32
from datetime import datetime as dt

from PyQt6.QtCore import pyqtSignal, QObject
import pandas as pd

from siteparser import parse


# Создание БД
def create_db():
    con = sqlite3.connect('Database.db')
    cur = con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS update_info ( \
                    table_name VARCHAR(20) PRIMARY KEY, \
                    last_update DATETIME)")
    cur.execute("INSERT INTO update_info (table_name, last_update) \
                    VALUES ('certificates_fstek', CURRENT_TIMESTAMP), \
                    ('certificates_fsb', CURRENT_TIMESTAMP) \
                ")
    con.commit()
    con.close()


# Получаем дату последнего обновления указанной таблицы
def get_update_date(table):
    update_date = None
    con = sqlite3.connect('Database.db')
    cur = con.cursor()
    exists = cur.execute("SELECT name FROM sqlite_master \
                                WHERE type='table' \
                                AND name='update_info' \
                                ").fetchall()
    # Существует ли указанная таблица в БД
    if not exists:
        create_db()

    result = cur.execute(f"SELECT last_update FROM update_info \
                        WHERE table_name = '{table}'").fetchall()
    con.close()
    update_date = dt.strptime(result[0][0], '%Y-%m-%d %H:%M:%S')
    return update_date


# Обновление БД
def upload_db(df, table):
    con = sqlite3.connect('Database.db')
    cur = con.cursor()
    cur.execute(f"DROP TABLE IF EXISTS {table}")
    df.to_sql(table, con, index=False)
    cur.execute(f"UPDATE update_info SET \
                last_update = CURRENT_TIMESTAMP \
                WHERE table_name = '{table}'")
    con.commit()
    con.close()


# Замена двойных пробелов на одинарные
def strip_spaces(a_str_with_spaces):
    return a_str_with_spaces.replace('  ', ' ')


# ФСТЭК
class WorkerFstek(QObject):
    message = pyqtSignal(str)
    database_date = pyqtSignal(dt)
    reestr_date = pyqtSignal(dt)
    update_commited = pyqtSignal()
    error = pyqtSignal()
    finished = pyqtSignal()

    def __init__(self, parent=None):
        super(WorkerFstek, self).__init__(parent)
        self.db_path = os.getcwd() + "\Database.db"
        self.file_name = os.getcwd() + r"\fstek_reg.csv"
        self.df = None
        self.headers = [
            "id", "date_start", "date_end", "name",
            "docs", "scheme", "lab", "certification",
            "applicant", "requisites", "support"
            ]


    def run(self):
        if os.path.exists(self.db_path):
            local_update = get_update_date("certificates_fstek")
            self.database_date.emit(local_update)
            fstek_update = parse("FSTEK", True)
            if fstek_update:
                self.reestr_date.emit(fstek_update)
                if isinstance(fstek_update, dt):
                    if local_update < fstek_update:
                        self.update_commited.emit()
                        self.get_reestr()
                    else:
                        self.message.emit("Обновление БД не требуется.")
                else:
                    self.message.emit("Невозможно сравнить даты.")
            else:
                self.message.emit("Нет соединения с сайтом ФСТЭК России.")
                self.error.emit()
                return
        else:
            self.message.emit("Создаём БД...")
            create_db()
            self.update_commited.emit()
            self.get_reestr()
        self.finished.emit()

    # Получение данных из реестра
    def get_reestr(self):
        self.message.emit("Получаем данные с сайта ФСТЭК России...")
        parse("FSTEK")
        df = pd.read_csv(self.file_name)
        if df is not None and df is not False:
            fstek_update = parse("FSTEK", True)
            self.reestr_date.emit(fstek_update)
            self.message.emit("Загружаем данные в БД...")
            df.columns = self.headers
            self.convert_date(df, 10)
            upload_db(df, "certificates_fstek")
            local_update = get_update_date("certificates_fstek")
            self.database_date.emit(local_update)
            os.remove(self.file_name)
        else:
            self.message.emit("Нет соединения с сайтом ФСТЭК России.")
            self.error.emit()
            return

    # Конвертирование даты в dataframe
    def convert_date(self, df, col_num):
        column = self.headers[col_num]
        for i in range(len(df.values)):
            value = df.iloc[i][column]
            try:
                new_value = dt.date(dt.strptime(value, '%d.%m.%Y'))
                df.iloc[i][column] = new_value
            except:
                if value is None:
                    df.iloc[i][column] = None


# ФСБ
class WorkerFsb(QObject):
    message = pyqtSignal(str)
    database_date = pyqtSignal(dt)
    error = pyqtSignal()
    finished = pyqtSignal()

    def __init__(self, parent=None):
        super(WorkerFsb, self).__init__(parent)
        self.db_path = os.getcwd() + "\Database.db"
        self.fileNameDoc = os.getcwd() + "\\fsb_reg.doc"
        self.fileNameHtml = os.getcwd() + "\\fsb_reg.html"
        self.DirHtml = os.getcwd() + "\\fsb_reg.files"
        self.df = None
        self.headers = [
            "id", "date_end", "name", "function"
            ]

    def run(self):
        if os.path.exists(self.db_path):
            local_update = get_update_date("certificates_fsb")
            self.database_date.emit(local_update)
            if parse("FSB") is False:
                self.message.emit("Нет соединения с сайтом ФСБ России.")
                self.error.emit()
                return
            self.get_reestr()
        else:
            self.message.emit("Создаём БД...")
            create_db()
            self.get_reestr()
        self.finished.emit()

    # Получение данных из реестра
    def get_reestr(self):
            self.message.emit("Получаем данные с сайта ФСБ России...")
            self.message.emit("Считываем данные...")
            self.convert_word_to_html()
            self.message.emit("Обрабатываем данные...")
            df = self.clean_df()
            self.message.emit("Загружаем данные в БД...")
            upload_db(df, "certificates_fsb")
            local_update = get_update_date("certificates_fsb")
            self.database_date.emit(local_update)

    # На конвертирование в html уходит 4 сек
    def convert_word_to_html(self):
        doc = win32.GetObject(self.fileNameDoc, pythoncom.CoInitialize())
        doc.SaveAs(FileName=self.fileNameHtml, FileFormat=8)
        doc.Close()
        os.remove(self.fileNameDoc)

    # Считывание данных из html и их редактирование
    def clean_df(self):
        convert_func = {
            'Условное  наименование (индекс)': strip_spaces,
            'Выполняемая  функция': strip_spaces,
            'Изготовитель': strip_spaces
            }
        date_start, date_end = [], []
        df = pd.read_html(self.fileNameHtml, converters=convert_func)[0]

        # После прочтения файла удаляем его и каталог
        os.remove(self.fileNameHtml)
        shutil.rmtree(self.DirHtml)

        # Убираем дублирующиеся колонки
        df = df.drop(df.columns[[0, 2, 4, 6, 8]], axis=1)
        df = df[:-1]
        df.columns = self.headers
        date_column = df['date_end']

        # Конвертируем даты и вносим их в списки
        for row in date_column:
            date_tuple = row.split()
            start = dt.date(dt.strptime(date_tuple[0], '%d.%m.%Y'))
            end = dt.date(dt.strptime(date_tuple[1], '%d.%m.%Y'))
            date_start.append(start)
            date_end.append(end)

        df.insert(1, 'date_start', date_start)
        df['date_end'] = date_end
        return df
