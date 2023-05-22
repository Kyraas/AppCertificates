from datetime import datetime

from PyQt6.QtWidgets import QStyledItemDelegate
from PyQt6.QtCore import Qt

# Подсвечивает искомый текст, изменение цвета выделенных строк.
class MyDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super(MyDelegate, self).__init__(parent)

    def paint(self, painter, option, index):
        option.displayAlignment = Qt.AlignmentFlag.AlignCenter
        QStyledItemDelegate.paint(self, painter, option, index)

    # Изменение формата даты из ГГГГ-ММ-ДД в ДД.ММ.ГГГГ, не препятствуя сортировке
    def displayText(self, value, locale):
        try:
            value = datetime.strptime(value, "%Y-%m-%d").date()
            value = value.strftime("%d.%m.%Y")
        except ValueError:
            pass
        return value


