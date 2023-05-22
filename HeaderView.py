# https://forum.qt.io/topic/123087/adding-custom-qheaderview-upon-request/2
# https://stackoverflow.com/questions/44343738/how-to-inject-widgets-between-qheaderview-and-qtableview

from PyQt6.QtCore import pyqtSignal, Qt
from PyQt6.QtWidgets import  QHeaderView, QWidget, QLineEdit


class FilterHeader(QHeaderView):
    filterActivated = pyqtSignal()

    def __init__(self, parent, y=80, x=5):
        super().__init__(Qt.Orientation.Horizontal, parent)
        self._editors = []
        self._padding = 4
        self.compensate_y = y
        self.compensate_x = x
        self.setStretchLastSection(True)
        self.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        self.setSortIndicatorShown(True)
        self.setSectionsClickable(True)
        self.setHighlightSections(True)
        self.show()
        self.sectionResized.connect(self.adjustPositions)
        parent.horizontalScrollBar().valueChanged.connect(self.adjustPositions)

    def setFilterBoxes(self, count):
        while self._editors:
            editor = self._editors.pop()
            editor.deleteLater()
        for index in range(count):
            editor = self.create_editor(self.parent(), index)
            self._editors.append(editor)
            editor.show()
        self.adjustPositions()

    def create_editor(self, parent, index):
        editor = QLineEdit(parent)
        editor.setPlaceholderText("Поиск")
        editor.setAlignment(Qt.AlignmentFlag.AlignCenter)
        editor.returnPressed.connect(self.filterActivated)
        return editor   

    def sizeHint(self):
        size = super().sizeHint()
        if self._editors:
            height = self._editors[0].sizeHint().height()
            size.setHeight(size.height() + height + self._padding)
        return size

    def updateGeometries(self):
        if self._editors:
            height = self._editors[0].sizeHint().height()
            self.setViewportMargins(0, 0, 0, height + self._padding)
        else:
            self.setViewportMargins(0, 0, 0, 0)
        super().updateGeometries()
        self.adjustPositions()

    def adjustPositions(self):
        for index, editor in enumerate(self._editors):
            if not isinstance(editor, QWidget):
                continue
            height = editor.sizeHint().height()
            editor.move(
                self.sectionPosition(index) - self.offset() + self.compensate_x,
                height + (self._padding // 2) + 3 + self.compensate_y,
            )
            editor.resize(self.sectionSize(index), height)

    def filterText(self, index):
        if 0 <= index < len(self._editors):
            return self._editors[index].text()
        return ''

    def clearFilters(self):
        for editor in self._editors:
            editor.clear()