# https://stackoverflow.com/questions/47201539/how-to-filter-multiple-column-in-qtableview

import re

from PyQt6.QtCore import QSortFilterProxyModel, Qt


class MyProxyModel(QSortFilterProxyModel):
    def __init__(self, source, parent=None):
        super().__init__(parent)
        self.source = source
        self.search = dict()

    def lessThan(self, left, right):
        leftData=self.sourceModel().data(left, Qt.ItemDataRole.DisplayRole)
        rightData=self.sourceModel().data(right, Qt.ItemDataRole.DisplayRole)
        if leftData is None:
            return True
        elif rightData is None:
            return False
        elif self.source == "Fstek" and left.column() == right.column() == 0:
            if "/" in leftData:
                leftData = leftData.split("/")[0]
            if "/" in rightData:
                rightData = rightData.split("/")[0]
            return int(leftData) > int(rightData)
        else:
            return leftData > rightData
    
    def setRegex(self, regex, column):
        if isinstance(regex, str):
            regex = re.compile(regex, re.IGNORECASE)
        self.search[column] = regex
        self.invalidateFilter()

    def clearFilter(self, column):
        del self.search[column]
        self.invalidateFilter()

    def clearFilters(self):
        self.search = {}
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        if not self.search:
            return True
        results = []

        if self.search:
            for key, regex in self.search.items():
                text = ''
                index = self.sourceModel().index(source_row, key, source_parent)
                if index.isValid():
                    text = self.sourceModel().data(index, Qt.ItemDataRole.DisplayRole)
                    if text is None:
                        text = ''
                results.append(regex.search(text))
            return all(results)
