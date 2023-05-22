from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QTableView


class MyTableView(QTableView):
    def paintEvent(self, event):
        super().paintEvent(event)
        if self.model() is not None and self.model().rowCount() > 0:
            return
        painter = QtGui.QPainter(self.viewport())
        painter.save()
        col = self.palette().placeholderText().color()
        painter.setPen(col)
        fm = self.fontMetrics()
        elided_text = fm.elidedText("Нет данных", QtCore.Qt.TextElideMode.ElideRight, self.viewport().width())
        painter.drawText(self.viewport().rect(), QtCore.Qt.AlignmentFlag.AlignCenter, elided_text)
        painter.restore()