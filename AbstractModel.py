# -*- coding: utf-8 -*-
# https://stackoverflow.com/questions/17697352/pyqt-implement-a-qabstracttablemodel-for-display-in-qtableview
# https://www.pythonguis.com/tutorials/qtableview-modelviews-numpy-pandas/
# https://doc.qt.io/qtforpython-5/overviews/model-view-programming.html?highlight=layoutabouttobechanged

import os
import sqlite3 as db
from calendar import monthrange
from datetime import datetime, date

from PyQt6.QtGui import QColor
from PyQt6.QtCore import QAbstractTableModel, Qt, QModelIndex


def convert_date(data):
    return datetime.date(datetime.strptime(str(data), "%Y-%m-%d"))

# Получаем дату следующего полугодия
def half_year():
    cur_date = datetime.now()
    month = cur_date.month + 5
    year = cur_date.year + month // 12
    month = month % 12 + 1
    day = min(cur_date.day, monthrange(year, month)[1])
    return date(year, month, day)


# Cоздание модели данных на базе QAbstractTableModel.
class FstekTableModel(QAbstractTableModel):

    def __init__(self):
        super(FstekTableModel, self).__init__()
        self.headers = [
            "№\nсертификата", "Дата\nвнесения\nв реестр", "Срок\nдействия\nсертификата",
            "Наименование\nсредства (шифр)",
            "Наименования документов,\nтребованиям которых\nсоответствует средство",
            "Схема\nсертификации", "Испытательная\nлаборатория", "Орган по\nсертификации",
            "Заявитель", "Реквизиты заявителя\n(индекс, адрес, телефон)",
            "Информация об\nокончании срока\nтехнической\nподдержки,\nполученная от\nзаявителя"
            ]
        self._data = None
        self.conn = None
        self.cur = None
        self.update()

    def update(self):
        # Существует ли БД
        if os.path.exists(os.getcwd() + "\Database.db"):
            self.conn = db.connect('Database.db')
            self.cur = self.conn.cursor()
            table = self.cur.execute("SELECT name FROM sqlite_master \
                                     WHERE type='table' \
                                     AND name='certificates_fstek' \
                                     ").fetchall()
            # Существует ли указанная таблица в БД
            if table:
                self.layoutAboutToBeChanged.emit()
                result = self.cur.execute("SELECT * FROM certificates_fstek").fetchall()
                self._data = result
                self.layoutChanged.emit()

    def exec(self, query):
        if self.cur:
            self.layoutAboutToBeChanged.emit()
            result = self.cur.execute(f"SELECT * FROM certificates_fstek WHERE {query}").fetchall()
            self._data = result
            self.layoutChanged.emit()

    def rowCount(self, index = QModelIndex()):
        if self._data:
            return len(self._data)
        else:
            return 0

    # Принимает первый вложенный список и возвращает длину.
    # (только если все строки имеют одинаковую длину)
    def columnCount(self, index = QModelIndex()):
        return 11

    # Параметр role описывает, какого рода информацию
    # метод должен возвращать при этом вызове.
    def data(self, index, role):
        now_date = datetime.date(datetime.today())
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.BackgroundRole:
            
            # Колонка с датами окончания сертификата
            date_end = self._data[index.row()][2]

            # Колонка с датами окончания поддержки
            sup = self._data[index.row()][10]

            # Если и сертификат, и подержка не действительны
            try:
                if ('прекращено' in date_end or \
                    'приостановлено' in date_end or \
                    convert_date(date_end) < now_date) and \
                    convert_date(sup) < now_date:
                    return QColor('#ff7f7f')  # красный
            except ValueError:
                pass

            # Если поддержка не действительна
            try:
                if convert_date(sup) < now_date:
                    return QColor('#ff9fc3')  # розовый
            except ValueError:
                pass

            # Если сертификат не действителен
            try:
                if convert_date(date_end) < now_date:
                    return QColor('#ffecb7')  # желтый

                # Если сертификат истечет через полгода
                elif convert_date(date_end) <= half_year():
                    return QColor('#e8eaed')  # светло-серый
            except ValueError:
                if "прекращено" in date_end or \
                    "приостановлено" in date_end:
                    return QColor('#ffecb7')  # желтый

        # DisplayRole фактически принимает только строковые значения.
        # В иных случаях необходимо форматировать данные в строку.
        elif role == Qt.ItemDataRole.DisplayRole:
            return self._data[index.row()][index.column()]

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and \
            orientation == Qt.Orientation.Horizontal:
            return self.headers[section]
        return None
        

# Cоздание модели данных на базе QAbstractTableModel.
class FsbTableModel(QAbstractTableModel):

    def __init__(self):
        super(FsbTableModel, self).__init__()
        self.headers = [
            "Рег. номер\nсертификата\nсоответствия",
            'Дата внесения\nв реестр',
            "Срок действия\nсертификата\nсоответствия",
            "Условное наименование (индекс)",
            "Выполняемая функция", "Изготовитель"
            ]
        self._data = None
        self.conn = None
        self.cur = None

    def get_update_date(self):
        result = self.cur.execute(f"SELECT last_update FROM update_info \
                                    WHERE table_name = 'certificates_fsb'\
                                  ").fetchall()
        update_date = datetime.strptime(result[0][0], '%Y-%m-%d %H:%M:%S')
        return update_date

    def update(self):
        # Существует ли БД
        if os.path.exists(os.getcwd() + "\Database.db"):
            self.conn = db.connect('Database.db')
            self.cur = self.conn.cursor()
            exists = self.cur.execute("SELECT name FROM sqlite_master \
                                        WHERE type='table' \
                                        AND name='certificates_fsb' \
                                        ").fetchall()
            # Существует ли указанная таблица в БД
            if exists:
                self.layoutAboutToBeChanged.emit()
                result = self.cur.execute(f"SELECT * FROM certificates_fsb").fetchall()
                self._data = result
                self.layoutChanged.emit()
                return self.get_update_date()

    def exec(self, query):
        if self.cur:
            self.layoutAboutToBeChanged.emit()
            result = self.cur.execute(f"SELECT * FROM certificates_fsb WHERE {query}").fetchall()
            self._data = result
            self.layoutChanged.emit()

    def rowCount(self, index = QModelIndex()):
        if self._data:
            return len(self._data)
        else:
            return 0

    # Принимает первый вложенный список и возвращает длину.
    # (только если все строки имеют одинаковую длину)
    def columnCount(self, index = QModelIndex()):
        return 6

    # Параметр role описывает, какого рода информацию
    # метод должен возвращать при этом вызове.
    def data(self, index, role):
        now_date = datetime.date(datetime.today())
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.BackgroundRole:
            
            # Колонка с датами окончания сертификата
            date_end = self._data[index.row()][2]

            # Если сертификат не действителен
            try:
                if convert_date(date_end) < now_date:
                    return QColor('#ffecb7')  # желтый

                # Если сертификат истечет через полгода
                elif convert_date(date_end) <= half_year():
                    return QColor('#e8eaed')  # светло-серый
            except ValueError:
                pass

        # DisplayRole фактически принимает только строковые значения.
        # В иных случаях необходимо форматировать данные в строку.
        elif role == Qt.ItemDataRole.DisplayRole:
            return self._data[index.row()][index.column()]

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and \
            orientation == Qt.Orientation.Horizontal:
            return self.headers[section]
        return None

