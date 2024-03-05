# -*- coding: utf-8 -*-
# https://stackoverflow.com/questions/64184256/qtableview-placeholder-text-before-table-appears
# https://stackoverflow.com/questions/17604243/pyside-get-list-of-all-visible-rows-in-a-table
# https://stackoverflow.com/questions/53353450/how-to-highlight-a-words-in-qtablewidget-from-a-searchlist
# https://stackoverflow.com/questions/60353152/qtablewidget-resizerowstocontents-very-slow


# TO DO
# Менять ширину колонок динамически, в зависимости от разрешения экрана?

import os
import re
import sys
from ctypes import windll
from datetime import datetime

from PyQt6.QtCore import (QRect, QThread, QPoint,
                          Qt, QTimer, pyqtSlot)
from PyQt6.QtGui import QIcon, QStatusTipEvent
from PyQt6.QtWidgets import (QApplication, QMainWindow, QButtonGroup,
                             QMessageBox, QFileDialog, QLabel, QFrame)


from myWindow import Ui_MainWindow
from AbstractModel import FstekTableModel, FsbTableModel
from HeaderView import FilterHeader
from ProxyModel import MyProxyModel
from Delegate import MyDelegate

from threadworkers import WorkerFstek, WorkerFsb
from creatingfiles import save_excel_file, save_word_file

# Для отображения значка на панели.
myappid = "'ООО ЦБИ'. ДДАС. Бондаренко М.А."
windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def export_table_data(table, model, rows):
    # Цикл по строкам proxy модели
    for row in range(rows):
        tbl_row = []

        for column in range(model.columnCount()):

            # Считываем данные каждой ячейки таблицы,
            # если они имеются.
            tbl_row.append("{}".format(
                            model.index(row, column).data() or ""))
        table.append(tbl_row)
    return table


# Поиск и конвертирование формата даты
def convert_date_format(text):
    new_reg = text
    d = "".join(re.findall('[.]*\d*', text))
    if "." in d:
        lst = d.split(".")
        # Если чисел больше одного
        if len(lst) >= 2:
            # меняем местами первый и последний
            lst[0], lst[-1] = lst[-1], lst[0]
            new_date = "-".join(lst)
        # Если число одно
        elif len(lst) == 1:
            if d[0] == "." and d[-1] == ".":
                pass
            elif d[0] == ".":
                date_line = d.replace(".","")
                date_line += "."
            elif d[-1] == ".":
                date_line = d.replace(".", "")
                date_line = "." + date_line
            new_date = "".join(date_line).replace(".","-")
        new_date = new_date.replace(".","-")
        new_reg = text.replace(d, new_date)
    return new_reg


# Вертикальная линия для статусбара
class VLine(QFrame):
    # a simple VLine, like the one you get from designer
    def __init__(self):
        super(VLine, self).__init__()
        self.setFrameShape(self.Shape.VLine)
        self.setFrameShadow(self.Shadow.Sunken)


class Table(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        # Инициализация нашего дизайна
        self.setupUi(self)
        self.setWindowTitle("Сертификаты")
        self.showMaximized()
        self.menu.menuAction().setStatusTip("Создание файла")
        icon_path = resource_path("icon.ico")
        self.setWindowIcon(QIcon(icon_path))
        self.statusbar = self.statusBar()
        self.status = QLabel()
        self.statusbar.addPermanentWidget(VLine())
        self.statusbar.addPermanentWidget(self.status)
        self.statusbar.addPermanentWidget(VLine())
        self.conn = None
        self.update_flag = False
        self.keywords = 'date_end LIKE "%прекращено%" or \
                        date_end LIKE "%приостановлено%"'
        self.keywords_not = 'date_end NOT LIKE "%прекращено%" and \
                            date_end NOT LIKE "%приостановлено%"'

        column_width = [
            80, 70, 110, 300, 200, 150,
            150, 210, 300, 300, 103
            ]

        self.thread_worker = QThread()
        self.init_worker = WorkerFstek()

        # Модель
        self.model = FstekTableModel()
        self.start_update_database()

        # Прокси модель (поиск и сортировка по возрастанию\убыванию)
        self.proxy = MyProxyModel("Fstek")
        self.proxy.setSourceModel(self.model)

        # Поиск по всей таблице (все колонки)
        self.proxy.setFilterKeyColumn(-1)
        self.proxy.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

        # Делегат
        self._delegate = MyDelegate()
        self.tableView.setItemDelegate(self._delegate)
        self.tableView.setModel(self.model)

        # Представление
        header = FilterHeader(self.tableView)
        self.tableView.setHorizontalHeader(header)
        self.tableView.setModel(self.proxy)
        self.header = self.tableView.horizontalHeader()
        self.header.setFilterBoxes(12)
        self.header.filterActivated.connect(self.handleFilterActivated)

        # Активируем возможность сортировки
        # по заголовкам в представлении.
        self.tableView.setSortingEnabled(True)

        # Приведение заголовков и первых видимых
        # строк таблицы к желаемому виду.
        for col, w in enumerate(column_width):
            self.tableView.horizontalHeader().resizeSection(col, w)

        # Выполнение функции с задержкой
        QTimer.singleShot(100, self.resize_visible_rows_fstek)
        self.init_fsb_tab()
        
        # Создаем группу фильтров
        self.filters = QButtonGroup(self)
        self.filters.addButton(self.radioButton_red)
        self.filters.addButton(self.radioButton_pink)
        self.filters.addButton(self.radioButton_yellow)
        self.filters.addButton(self.radioButton_gray)
        self.filters.addButton(self.radioButton_white)

        # # Соединяем виджеты с функциями
        self.action_Excel.triggered.connect(self.save_file)
        self.action_Word.triggered.connect(self.save_file)
        self.refreshButton.clicked.connect(self.start_update_database)
        self.radioButton_red.clicked.connect(self.red)
        self.radioButton_pink.clicked.connect(self.pink)
        self.radioButton_yellow.clicked.connect(self.yellow)
        self.radioButton_gray.clicked.connect(self.gray)
        self.radioButton_white.clicked.connect(self.white)
        self.checkBox_sup.stateChanged.connect(self.white)
        self.resetButton.clicked.connect(self.reset)
        self.clearSearch.clicked.connect(self.clear)
        self.proxy.layoutChanged.connect(self.resize_visible_rows_fstek)
        self.tableView.verticalScrollBar().valueChanged.connect(self.resize_visible_rows_fstek)
        self.tabWidget.currentChanged.connect(self.resize_visible_rows_fsb)
        self.tabWidget.currentChanged.connect(self.row_count_status)

    # Инициализация таблицы ФСБ
    def init_fsb_tab(self):
        column_width = [
            90, 90, 90,
            300, 900, 300
            ]

        self.thread_worker_fsb = QThread()
        self.init_worker_fsb = WorkerFsb()

        # Модель
        self.model_fsb = FsbTableModel()

        # Обновляем БД и получаем её актуальную дату
        self.set_db_date(self.model_fsb.update())

        # Прокси модель (поиск и сортировка по возрастанию\убыванию)
        self.proxy_fsb = MyProxyModel("Fsb")
        self.proxy_fsb.setSourceModel(self.model_fsb)
        self.proxy_fsb.setFilterKeyColumn(-1)
        self.proxy_fsb.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

        # Делегат
        self._delegate_fsb = MyDelegate()
        self.tableViewFsb.setItemDelegate(self._delegate_fsb)
        self.tableViewFsb.setModel(self.model_fsb)

        # Представление
        header = FilterHeader(self.tableViewFsb, 30)
        self.tableViewFsb.setHorizontalHeader(header)
        self.tableViewFsb.setModel(self.proxy_fsb)
        self.tableViewFsb.setSortingEnabled(True)
        self.header_fsb = self.tableViewFsb.horizontalHeader()
        self.header_fsb.filterActivated.connect(self.handleFilterActivated)
        self.header_fsb.setFilterBoxes(6)

        # Приведение заголовков и первых видимых
        # строк таблицы к желаемому виду.
        for col, width in enumerate(column_width):
            self.tableViewFsb.horizontalHeader().resizeSection(col, width)

        self.refreshButtonFsb.clicked.connect(self.start_update_fsb)
        self.proxy_fsb.layoutChanged.connect(self.resize_visible_rows_fsb)
        self.tableViewFsb.verticalScrollBar().valueChanged.connect(self.resize_visible_rows_fsb)
        self.radioButton_yellow_fsb.clicked.connect(self.yellow_fsb)
        self.radioButton_gray_fsb.clicked.connect(self.gray_fsb)
        self.radioButton_white_fsb.clicked.connect(self.white_fsb)
        self.resetButtonFsb.clicked.connect(self.reset_fsb)
        self.clearSearchFsb.clicked.connect(self.clear)
        
    # Обновление БД ФСБ
    def start_update_fsb(self):
        self.refreshButtonFsb.setEnabled(False)
        self.statusbar.setStyleSheet("background-color: #FFFF89")
        self.change_status("Загрузка...")

        self.init_worker_fsb.moveToThread(self.thread_worker_fsb)
        self.thread_worker_fsb.started.connect(self.init_worker_fsb.run)
        self.init_worker_fsb.database_date.connect(self.set_db_date)
        self.init_worker_fsb.message.connect(self.change_status)
        self.init_worker_fsb.error.connect(self.cancel_update_database)
        self.init_worker_fsb.finished.connect(self.finish_update_database)
        self.thread_worker_fsb.start()

    def handleFilterActivated(self):
        tab = self.tabWidget.currentIndex()
        header = self.header
        view = self.tableView
        if tab == 1:
            header = self.header_fsb
            view = self.tableViewFsb
        for index in range(header.count()):
            text = header.filterText(index)
            self.search(text, index)
            view.viewport().update()

    @pyqtSlot(str)
    def change_status(self, text):
        self.status.setText(text)
        self.statusbar.repaint()

    @pyqtSlot()
    def update_status(self):
        self.update_flag = True

    # Отображение кол-ва строк
    def row_count_status(self):
        n = self.proxy.rowCount()
        if self.tabWidget.currentIndex() == 1:
            n = self.proxy_fsb.rowCount()
        if n != 0:
            text = f'Сертификатов: {n}.'
        else:
            text = 'По данному запросу не найдено.'
        self.statusbar.setStyleSheet("background-color : #D8D8D8")
        self.statusbar.showMessage(text)

    # При наведении курсора на меню "Файл"
    # строка состояния становилась пустой.
    def event(self, e):
        # Событие изменения строки состояния. (в нижнем левом углу)
        if e.type() == 112 and e.tip() == '':
            proxy = self.proxy
            if self.tabWidget.currentIndex() == 1:
                proxy = self.proxy_fsb
            e = QStatusTipEvent(
                f'Сертификатов: {proxy.rowCount()}.')
        return super().event(e)

    # Изменение высоты только видимых строк
    def resize_visible_rows_fstek(self):
        viewport_rect = QRect(QPoint(0, 0),
                                self.tableView.viewport().size())

        for row in range(0, self.proxy.rowCount() + 1):
            # Выбираем любую видимую колонку. (костыль!)
            rect = self.tableView.visualRect(self.proxy.index(row, 4))
            if viewport_rect.intersects(rect):  # если видимые строки
                for _ in range(0, 20):
                    self.tableView.resizeRowToContents(row)
                    row += 1
                break

    # Изменение высоты только видимых строк
    def resize_visible_rows_fsb(self):
        viewport_rect = QRect(QPoint(0, 0),
                                self.tableViewFsb.viewport().size())

        for row in range(0, self.proxy_fsb.rowCount() + 1):
            # Выбираем любую видимую колонку. (костыль!)
            rect = self.tableViewFsb.visualRect(self.proxy_fsb.index(row, 4))
            if viewport_rect.intersects(rect):  # если видимые строки
                for _ in range(0, 20):
                    self.tableViewFsb.resizeRowToContents(row)
                    row += 1
                break

    # Отображение полученных дат в приложении
    def set_db_date(self, actual_date):
        sender = self.sender()
        if actual_date:
            actual_date_str = datetime.strftime(actual_date, "%d.%m.%Y г. %H:%M")
        else:
            actual_date_str = "Нет данных."
        actual = f"Актуальность текущей базы: {actual_date_str}"
        if isinstance(sender, WorkerFstek):
            self.last_update_date.setText(actual)
        else:
            self.last_update_fsb.setText(actual)

    # Изменение даты обновления БД
    def set_reestr_date(self, actual_date):
        actual_date_str = datetime.strftime(actual_date, "%d.%m.%Y г. %H:%M")
        actual = f"Актуальность базы ФСТЭК: {actual_date_str}"
        self.actual_date.setText(actual)
    
    # Инициализация и запуск класса WorkerThread
    def start_update_database(self):
        self.refreshButton.setEnabled(False)
        self.refreshButtonFsb.setEnabled(False)
        self.statusbar.setStyleSheet("background-color: #FFFF89")
        self.change_status("Проверка актуальности данных...")

        self.init_worker.moveToThread(self.thread_worker)
        self.thread_worker.started.connect(self.init_worker.run)
        self.init_worker.database_date.connect(self.set_db_date)
        self.init_worker.reestr_date.connect(self.set_reestr_date)
        self.init_worker.message.connect(self.change_status)
        self.init_worker.update_commited.connect(self.update_status)
        self.init_worker.error.connect(self.cancel_update_database)
        self.init_worker.finished.connect(self.finish_update_database)
        self.thread_worker.start()
    
    # При отсутствии соединения с сайтом ФСТЭК/ФСБ
    def cancel_update_database(self):
        sender = self.sender()
        if isinstance(sender, WorkerFstek):
            name = "ФСТЭК"
            actual = f"Актуальность базы ФСТЭК: нет соединения."
            self.actual_date.setText(actual)
            self.thread_worker.quit()
        else:
            name = "ФСБ"
            self.thread_worker_fsb.quit()
        # бледно-красный
        self.statusbar.setStyleSheet("background-color: #FF9090")
        self.change_status(f"Нет соединения с сайтом {name} " +
                                "России. Проверьте интернет-" +
                                "соединение или повторите " +
                                "попытку позже.") 
        self.refreshButton.setEnabled(True)
        self.refreshButtonFsb.setEnabled(True)

    # Завершение процесса обновления БД
    def finish_update_database(self):
        sender = self.sender()
        self.statusbar.setStyleSheet("background-color: #FFFF89")
        self.change_status('Загрузка таблицы...')
        if isinstance(sender, WorkerFstek):
            self.model.update()
            n = self.proxy.rowCount()
            if self.update_flag:
                self.update_flag = False
                self.statusbar.setStyleSheet("background-color: #ADFF94")
                self.change_status("База данных ФСТЭК успешно обновлена.")
                self.statusbar.showMessage(f"Сертификатов: {n}")
            else:
                self.statusbar.setStyleSheet("background-color: #D8D8D8")
                self.change_status(f"База данных ФСТЭК уже последней версии.")
                self.statusbar.showMessage(f"Сертификатов: {n}")
            self.thread_worker.quit()
        else:
            self.model_fsb.update()
            n = self.proxy_fsb.rowCount()
            self.statusbar.setStyleSheet("background-color: #ADFF94")
            self.change_status("База данных ФСБ успешно обновлена.")
            self.statusbar.showMessage(f"Сертификатов: {n}")
            self.thread_worker_fsb.quit()
        self.refreshButton.setEnabled(True)
        self.refreshButtonFsb.setEnabled(True)

    # Сохранение таблицы в файл
    def save_file(self):
        # Берём за основу Proxy-модель для экспорта
        # таблицы с учётом применённых фильтров
        tab = self.tabWidget.currentIndex()
        model = self.proxy
        name = "ФСТЭК"
        if tab == 1:
            model = self.proxy_fsb
            name = "ФСБ"
        rows = model.rowCount()
        err = False
        r = "(*.docx)"
        table = []
        now = datetime.date(datetime.today())
        date_time = now.strftime("%d.%m.%Y")
        a = self.sender()

        # Если строк 0, то отменяем сохранение.
        if rows == 0: 
            QMessageBox.information(self, "Сохранение файла",
                                    "Нечего сохранять.\n" +
                                    f"Строк {rows} шт.")
            return

        if a.text() == 'Экспортировать в Excel-файл':
            r = "(*.xlsx)"

        table = export_table_data(table, model, rows)

        fileName, _ = QFileDialog.getSaveFileName(
                                self, "Сохранить файл",
                                f"./Сертификаты {name} {rows} " +
                                f"шт. {date_time}", f"All Files{r}")

        if not fileName:    # кнопка "отмена"
            return 

        self.statusbar.setStyleSheet("background-color: #FFFF89")
        if r == "(*.xlsx)":
            self.change_status('Загрузка таблицы в Excel-файл...')
            self.statusbar.repaint()
            err = save_excel_file(self, fileName, table, tab)
        else:
            self.change_status('Загрузка таблицы в Word-файл...')
            err = save_word_file(self, fileName, table, tab)
        
        if not err:
            QMessageBox.information(
                            self, "Сохранение файла",
                            f"Данные сохранены в файле: \n{fileName}")

            self.change_status('Загрузка завершена.')
        else:
            self.change_status('Ошибка загрузки.')
        self.statusbar.setStyleSheet("background-color: #D8D8D8") # серый


    # Поиск по таблице
    def search(self, text, index):
        tab = self.tabWidget.currentIndex()
        proxy = self.proxy_fsb
        if tab == 0:
            proxy = self.proxy
            if index in [1, 2, 10]:
                text = convert_date_format(text)

        self.statusbar.setStyleSheet("background-color: #FFFF89")
        self.statusbar.showMessage('Поиск...')
        proxy.setRegex(text, index)

        proxy.layoutChanged.emit()
        n = proxy.rowCount()
        if n != 0:
            # Серый цвет
            self.statusbar.setStyleSheet("background-color: #D8D8D8")
            self.statusbar.showMessage(f'Сертификатов: {n}.')
        else:
            self.statusbar.setStyleSheet("background-color: #D8D8D8")
            self.statusbar.showMessage('По данному запросу не найдено.')

    # Очищение полей поиска
    def clear(self):
        tab = self.tabWidget.currentIndex()
        proxy = self.proxy
        header = self.header
        if tab == 1:
            proxy = self.proxy_fsb
            header = self.header_fsb
        proxy.clearFilters()
        proxy.layoutChanged.emit()
        header.clearFilters()
        self.row_count_status()

    # Фильтры:
    # 1. Сертификат и поддержка не действительны
    def red(self):
        red_filter = f'(date_end <= CURRENT_DATE or \
                        {self.keywords}) and \
                        support <= CURRENT_DATE and\
                        support NOT LIKE "%бессрочно%"'
        if self.radioButton_red.isChecked():
            self.model.exec(red_filter)
            self.row_count_status()
            self.reset_checkBoxes()

    # 2. Поддержка не действительна
    def pink(self):
        pink_filter = f'(date_end > CURRENT_DATE and \
                        {self.keywords_not}) and \
                        support IS NOT NULL and \
                        support <= CURRENT_DATE'
        if self.radioButton_pink.isChecked():
            self.model.exec(pink_filter)
            self.row_count_status()
            self.reset_checkBoxes()

    # 3. Сертификат не действителен
    def yellow(self):
        yellow_filter = f'(support > CURRENT_DATE or \
                        support IS NULL) and \
                        (date_end <= CURRENT_DATE or \
                        {self.keywords})'
        if self.radioButton_yellow.isChecked():
            self.model.exec(yellow_filter)
            self.row_count_status()
            self.reset_checkBoxes()
            
    def yellow_fsb(self):
        yellow_filter = f'date_end <= CURRENT_DATE'
        if self.radioButton_yellow_fsb.isChecked():
            self.model_fsb.exec(yellow_filter)
            self.row_count_status()

    # 4. Сертификат истечёт менее, чем через полгода
    # (почему-то пропадает последний столбец)
    def gray(self):
        gray_filter = 'date_end BETWEEN CURRENT_DATE AND \
                        DATE(CURRENT_DATE,"+6 months")'
        if self.radioButton_gray.isChecked():
            self.model.exec(gray_filter)
            self.row_count_status()
            self.reset_checkBoxes()

    def gray_fsb(self):
        gray_filter = 'date_end BETWEEN CURRENT_DATE AND \
                        DATE(CURRENT_DATE,"+6 months")'
        if self.radioButton_gray_fsb.isChecked():
            self.model_fsb.exec(gray_filter)
            self.row_count_status()

    # 5. Сертификат и поддержка действительны
    def white(self):
        white_filter = f'date_end > CURRENT_DATE and \
                        {self.keywords_not} and \
                        support > CURRENT_DATE'
        if self.radioButton_white.isChecked():
            self.checkBox_sup.setEnabled(True)
            if self.checkBox_sup.isChecked():
                white_filter = f'date_end > CURRENT_DATE and \
                                {self.keywords_not} and \
                                (support > CURRENT_DATE or \
                                support IS NULL)'
            self.model.exec(white_filter)
            self.row_count_status()

    def white_fsb(self):
        white_filter = f'date_end > CURRENT_DATE'
        if self.radioButton_white_fsb.isChecked():
            self.model_fsb.exec(white_filter)
            self.row_count_status()

    # Сброс флагов
    def reset_checkBoxes(self):
        # Если хотя бы один флаг включён
        if self.checkBox_sup.isEnabled:
            self.checkBox_sup.setEnabled(False)

    # Сброс фильтров
    def reset(self):
        self.reset_checkBoxes()
        
        # Делаем переключатели уникальными
        # и отдельно выключаем каждый.
        self.filters.setExclusive(False)
        self.radioButton_red.setChecked(False)
        self.radioButton_pink.setChecked(False)
        self.radioButton_yellow.setChecked(False)
        self.radioButton_gray.setChecked(False)
        self.radioButton_white.setChecked(False)
        self.filters.setExclusive(True)
        self.model.update()
        self.row_count_status()

    def reset_fsb(self):
        self.radioButton_yellow_fsb.setChecked(False)
        self.radioButton_gray_fsb.setChecked(False)
        self.radioButton_white_fsb.setChecked(False)
        self.model_fsb.update()
        self.row_count_status()

    # Событие закрытия приложения
    def closeEvent(self, event):
        if self.thread_worker.isRunning():
            print("Пожалуйста, дождитесь заверешния работы программы.")
            event.ignore()
        else:
            self.statusbar.setStyleSheet("background-color: #FFFF89")
            self.change_status('Закрываем...')
            event.accept()


app = QApplication(sys.argv)
# screen_size = app.primaryScreen().size()
# width = screen_size.width()
# height = screen_size.height()
win = Table()
win.show()
sys.exit(app.exec())
