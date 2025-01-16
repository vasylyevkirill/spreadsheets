from PyQt5.QtCore import QDate, QPoint, Qt
from PyQt5.QtGui import QColor, QPainter, QPixmap, QIcon
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QWidget, QCompleter, QDateTimeEdit, QItemDelegate, QLineEdit, QColorDialog, QFontDialog

from dataframes.models import Spreadsheet


class SpreadSheetDelegate(QItemDelegate):
    def __init__(self, parent = None):
        super(SpreadSheetDelegate, self).__init__(parent)

    def createEditor(self, parent, styleOption, index):
        if index.column() == 1:
            editor = QDateTimeEdit(parent)
            editor.setDisplayFormat(self.parent().currentDateFormat)
            editor.setCalendarPopup(True)
            return editor

        editor = QLineEdit(parent)
        # create a completer with the strings in the column as model
        allStrings = []
        for i in range(1, index.model().rowCount()):
            strItem = index.model().data(index.sibling(i, index.column()), Qt.EditRole)
            if strItem not in allStrings:
                allStrings.append(strItem)

        autoComplete = QCompleter(allStrings)
        editor.setCompleter(autoComplete)
        editor.editingFinished.connect(self.commitAndCloseEditor)
        return editor

    def commitAndCloseEditor(self):
        editor = self.sender()
        self.commitData.emit(editor)
        self.closeEditor.emit(editor, QItemDelegate.NoHint)

    def setEditorData(self, editor, index):
        if isinstance(editor, QLineEdit):
            editor.setText(index.model().data(index, Qt.EditRole))
        elif isinstance(editor, QDateTimeEdit):
            editor.setDate(QDate.fromString(
                index.model().data(index, Qt.EditRole), self.parent().currentDateFormat))

    def setModelData(self, editor, model, index):
        if isinstance(editor, QLineEdit):
            model.setData(index, editor.text())
        elif isinstance(editor, QDateTimeEdit):
            model.setData(index, editor.date().toString(self.parent().currentDateFormat))


class TableWidget(QTableWidget):
    def __init__(self, data = [], rows_count: int = -1, columns_count: int = -1, parent: QWidget | None = None):
        super(TableWidget, self).__init__(rows_count, columns_count, parent)
        self.data = data
        self.resize(rows_count, columns_count)
        self.setItemPrototype(self.item(rows_count - 1, columns_count - 1))
        self.setItemDelegate(SpreadSheetDelegate(self))
        self.currentItemChanged.connect(self.updateItemColor)
        
    def resize(self, rows_count: int, columns_count: int) -> QWidget:
        for column_index in range(columns_count):
            character = chr(ord('A') + column_index)
            if self.data:
                character = self.data.ALLOWED_Y_CHARACTER[column_index]

            self.setHorizontalHeaderItem(column_index, QTableWidgetItem(character))

    def reset_data(self, data: Spreadsheet):
        self.resize(data.shape[0], data.shape[1])
        
        return

    def updateItemColor(self, item):
        pixmap = QPixmap(16, 16)
        color = QColor()
        if item:
            color = item.background().color()
        if not color.isValid():
            color = self.palette().base().color()
        painter = QPainter(pixmap)
        painter.fillRect(0, 0, 16, 16, color)
        lighter = color.lighter()
        painter.setPen(lighter)
        painter.drawPolyline(QPoint(0, 15), QPoint(0, 0), QPoint(15, 0))
        painter.setPen(color.darker())
        painter.drawPolyline(QPoint(1, 15), QPoint(15, 15), QPoint(15, 1))
        painter.end()
        self.parent().colorAction.setIcon(QIcon(pixmap))

    def selectColor(self):
        item = self.currentItem()
        color = item and QColor(item.background()) or self.palette().base().color()
        color = QColorDialog.getColor(color, self.parent())
        if not color.isValid():
            return
        selected = self.selectedItems()
        if not selected:
            return
        for i in selected:
            i and i.setBackground(color)
        self.updateItemColor(self.currentItem())

    def selectFont(self):
        selected = self.selectedItems()
        if not selected:
            return
        font, ok = QFontDialog.getFont(self.font(), self)
        if not ok:
            return
        for i in selected:
            i and i.setFont(font)

    def set_table_item(self, pos_x: int, pos_y: int, data: str):
        self.table.setItem(pos_x, pos_y, SpreadSheetItem(data))


