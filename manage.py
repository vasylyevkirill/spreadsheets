#!/usr/bin/env python

from pandas import DataFrame, read_excel, read_csv, read_json

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QColor, QIcon, QKeySequence, QPixmap
from PyQt5.QtWidgets import (
    QAction, QActionGroup, QApplication, QFileDialog,
    QLabel, QLineEdit, QMainWindow, QToolBar, QMessageBox
)
from PyQt5.QtPrintSupport import QPrinter, QPrintPreviewDialog

from components.TableWidget import TableWidget
from components.InputDialog import InputDialog
from components.AboutWindow import show_about_window
from visuals.spreadsheetitem import SpreadSheetItem
from visuals.printview import PrintView
from util import decode_pos, encode_pos


class SpreadSheet(QMainWindow):

    dateFormats = ['dd/M/yyyy', 'yyyy/M/dd', 'dd.MM.yyyy']

    currentDateFormat = dateFormats[0]

    def __init__(self, rows, cols, parent=None):
        super(SpreadSheet, self).__init__(parent)

        self.toolBar = QToolBar()
        self.addToolBar(self.toolBar)
        self.formulaInput = QLineEdit()
        self.cellLabel = QLabel(self.toolBar)
        self.cellLabel.setMinimumSize(80, 0)
        self.toolBar.addWidget(self.cellLabel)
        self.toolBar.addWidget(self.formulaInput)
        self.table = TableWidget(rows_count=rows, columns_count=cols, parent=self)
        self.table.itemChanged.connect(self.updateStatus)
        self.table.currentItemChanged.connect(self.updateLineEdit)

        self.createActions()
        self.table.updateItemColor(0)
        self.setupMenuBar()
        self.setupContents()
        self.setupContextMenu()
        self.setCentralWidget(self.table)
        self.statusBar()
        self.formulaInput.returnPressed.connect(self.returnPressed)
        self.setWindowTitle('ЭксЭксЭль')

    def setupContextMenu(self):
        self.addAction(self.cell_addAction)
        self.addAction(self.cell_subAction)
        self.addAction(self.cell_mulAction)
        self.addAction(self.cell_divAction)
        self.addAction(self.cell_sumAction)
        self.addAction(self.firstSeparator)
        self.addAction(self.colorAction)
        self.addAction(self.fontAction)
        self.addAction(self.secondSeparator)
        self.addAction(self.clearAction)
        self.setContextMenuPolicy(Qt.ActionsContextMenu)

    def createActions(self):
        self.cell_sumAction = QAction('Сумма', self)
        self.cell_sumAction.triggered.connect(self.actionSum)

        self.cell_addAction = QAction('&Добавить', self)
        self.cell_addAction.setShortcut(Qt.CTRL | Qt.Key_Plus)
        self.cell_addAction.triggered.connect(self.actionAdd)

        self.cell_subAction = QAction('&Отнять', self)
        self.cell_subAction.setShortcut(Qt.CTRL | Qt.Key_Minus)
        self.cell_subAction.triggered.connect(self.actionSubtract)

        self.cell_mulAction = QAction('&Умножить', self)
        self.cell_mulAction.setShortcut(Qt.CTRL | Qt.Key_multiply)
        self.cell_mulAction.triggered.connect(self.actionMultiply)

        self.cell_divAction = QAction('&Поделить', self)
        self.cell_divAction.setShortcut(Qt.CTRL | Qt.Key_division)
        self.cell_divAction.triggered.connect(self.actionDivide)

        self.fontAction = QAction('Шрифт...', self)
        self.fontAction.setShortcut(Qt.CTRL | Qt.Key_F)
        self.fontAction.triggered.connect(self.table.selectFont)

        self.colorAction = QAction(QIcon(QPixmap(16, 16)), 'Цвет &фона...', self)
        self.colorAction.triggered.connect(self.table.selectColor)

        self.clearAction = QAction('Очистить', self)
        self.clearAction.setShortcut(Qt.Key_Delete)
        self.clearAction.triggered.connect(self.clear)

        self.aboutSpreadSheet = QAction('О приложении', self)
        self.aboutSpreadSheet.triggered.connect(self.show_about)

        self.exitAction = QAction('В&ыйти', self)
        self.exitAction.setShortcut(QKeySequence.Quit)
        self.exitAction.triggered.connect(QApplication.instance().quit)

        self.printAction = QAction('&Печать', self)
        self.printAction.setShortcut(QKeySequence.Print)
        self.printAction.triggered.connect(self.print_)

        self.importAction = QAction('&Импортировать', self)
        self.importAction.triggered.connect(self.runImportDialog)

        self.firstSeparator = QAction(self)
        self.firstSeparator.setSeparator(True)

        self.secondSeparator = QAction(self)
        self.secondSeparator.setSeparator(True)

    def setupMenuBar(self):
        self.fileMenu = self.menuBar().addMenu('&Файл')
        self.dateFormatMenu = self.fileMenu.addMenu('&Формат даты')
        self.dateFormatGroup = QActionGroup(self)
        for f in self.dateFormats:
            action = QAction(f, self, checkable=True, triggered=self.changeDateFormat)
            self.dateFormatGroup.addAction(action)
            self.dateFormatMenu.addAction(action)
            if f == self.currentDateFormat:
                action.setChecked(True)

        self.fileMenu.addAction(self.importAction)
        self.fileMenu.addAction(self.printAction)
        self.fileMenu.addAction(self.exitAction)
        self.cellMenu = self.menuBar().addMenu('&Клетка')
        self.cellMenu.addAction(self.cell_addAction)
        self.cellMenu.addAction(self.cell_subAction)
        self.cellMenu.addAction(self.cell_mulAction)
        self.cellMenu.addAction(self.cell_divAction)
        self.cellMenu.addAction(self.cell_sumAction)
        self.cellMenu.addSeparator()
        self.cellMenu.addAction(self.colorAction)
        self.cellMenu.addAction(self.fontAction)
        self.menuBar().addSeparator()
        self.aboutMenu = self.menuBar().addMenu('&Помощь')
        self.aboutMenu.addAction(self.aboutSpreadSheet)

    def changeDateFormat(self):
        action = self.sender()
        oldFormat = self.currentDateFormat
        newFormat = self.currentDateFormat = action.text()
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 1)
            date = QDate.fromString(item.text(), oldFormat)
            item.setText(date.toString(newFormat))

    def updateStatus(self, item):
        if item and item == self.table.currentItem():
            self.statusBar().showMessage(item.data(Qt.StatusTipRole), 1000)
            self.cellLabel.setText(
                'Cell: (%s)' % encode_pos(
                    self.table.row(item),
                    self.table.column(item)
                )
            )

    def updateLineEdit(self, item):
        if item != self.table.currentItem():
            return
        if item:
            self.formulaInput.setText(item.data(Qt.EditRole))
        else:
            self.formulaInput.clear()

    def returnPressed(self):
        text = self.formulaInput.text()
        row = self.table.currentRow()
        col = self.table.currentColumn()
        item = self.table.item(row, col)
        if not item:
            self.table.setItem(row, col, SpreadSheetItem(text))
        else:
            item.setData(Qt.EditRole, text)
        self.table.viewport().update()

    def actionSum(self):
        row_first = 0
        row_last = 0
        row_cur = 0
        col_first = 0
        col_last = 0
        col_cur = 0
        selected = self.table.selectedItems()
        if selected:
            first = selected[0]
            last = selected[-1]
            row_first = self.table.row(first)
            row_last = self.table.row(last)
            col_first = self.table.column(first)
            col_last = self.table.column(last)

        current = self.table.currentItem()
        if current:
            row_cur = self.table.row(current)
            col_cur = self.table.column(current)

        cell1 = encode_pos(row_first, col_first)
        cell2 = encode_pos(row_last, col_last)
        out = encode_pos(row_cur, col_cur)
        ok, cell1, cell2, out = self.runInputDialog(
            'Sum cells', 'First cell:',
            'Last cell:', u'\N{GREEK CAPITAL LETTER SIGMA}', 'Output to:',
            cell1, cell2, out
        )
        if ok:
            row, col = decode_pos(out)
            self.table.item(row, col).setText('sum %s %s' % (cell1, cell2))

    def actionMath_helper(self, title, op):
        cell1 = 'C1'
        cell2 = 'C2'
        out = 'C3'
        current = self.table.currentItem()
        if current:
            out = encode_pos(self.table.currentRow(), self.table.currentColumn())
        ok, cell1, cell2, out = self.runInputDialog(
            title, 'Cell 1', 'Cell 2',
            op, 'Output to:', cell1, cell2, out
        )
        if ok:
            row, col = decode_pos(out)
            self.table.item(row, col).setText('%s %s %s' % (op, cell1, cell2))

    def actionAdd(self):
        self.actionMath_helper('Addition', '+')

    def actionSubtract(self):
        self.actionMath_helper('Subtraction', '-')

    def actionMultiply(self):
        self.actionMath_helper('Multiplication', '*')

    def actionDivide(self):
        self.actionMath_helper('Division', '/')

    def clear(self):
        for i in self.table.selectedItems():
            i.setText('')

    def runInputDialog(self, *args, **kwargs):
        addDialog = InputDialog(*args, table=self.table, **kwargs)
        if addDialog.exec_():
            cell1 = addDialog.cell1ColInput.currentText() + addDialog.cell1RowInput.currentText()
            cell2 = addDialog.cell2ColInput.currentText() + addDialog.cell2RowInput.currentText()
            outCell = addDialog.outColInput.currentText() + addDialog.outRowInput.currentText()
            return True, cell1, cell2, outCell
        return False, None, None, None

    def runImportDialog(self, *args, **kwargs) -> str:
        filename, _ = QFileDialog.getOpenFileName(self, 'Открыть файл', '', 'Табличные файлы (*.xlsx *.csv *.json)')
        if filename:
            try:
                self.read_table_file(filename)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при чтении файла: {e}") 

    def read_table_file(self, filename):
        with open(filename) as file: 
            if filename.endswith('.xlsx'):
                return read_excel(file)
            elif filename.endswith('.json'):
                return read_json(filename)
            elif filename.endswith('.csv'):
                return read_csv(filename)
            else:
                QMessageBox.warning(self, 'Ошибка', 'Неподдерживаемый формат файла.')

            QMessageBox.critical(self, 'Ошибка', f'Ошибка при чтении файла: {e}')

    def setupContents(self):
        titleBackground = QColor(Qt.lightGray)
        titleFont = self.table.font()
        titleFont.setBold(True)
        # column 0
        self.table.setItem(0, 0, SpreadSheetItem('Item'))
        self.table.item(0, 0).setBackground(titleBackground)
        self.table.item(0, 0).setToolTip('This column shows the purchased item/service')
        self.table.item(0, 0).setFont(titleFont)
        self.table.setItem(1, 0, SpreadSheetItem('AirportBus'))
        self.table.setItem(2, 0, SpreadSheetItem('Flight (Munich)'))
        self.table.setItem(3, 0, SpreadSheetItem('Lunch'))
        self.table.setItem(4, 0, SpreadSheetItem('Flight (LA)'))
        self.table.setItem(5, 0, SpreadSheetItem('Taxi'))
        self.table.setItem(6, 0, SpreadSheetItem('Dinner'))
        self.table.setItem(7, 0, SpreadSheetItem('Hotel'))
        self.table.setItem(8, 0, SpreadSheetItem('Flight (Oslo)'))
        self.table.setItem(9, 0, SpreadSheetItem('Total:'))
        self.table.item(9, 0).setFont(titleFont)
        self.table.item(9, 0).setBackground(Qt.lightGray)
        # column 1
        self.table.setItem(0, 1, SpreadSheetItem('Date'))
        self.table.item(0, 1).setBackground(titleBackground)
        self.table.item(0, 1).setToolTip('This column shows the purchase date, double click to change')
        self.table.item(0, 1).setFont(titleFont)
        self.table.setItem(1, 1, SpreadSheetItem('15/6/2006'))
        self.table.setItem(2, 1, SpreadSheetItem('15/6/2006'))
        self.table.setItem(3, 1, SpreadSheetItem('15/6/2006'))
        self.table.setItem(4, 1, SpreadSheetItem('21/5/2006'))
        self.table.setItem(5, 1, SpreadSheetItem('16/6/2006'))
        self.table.setItem(6, 1, SpreadSheetItem('16/6/2006'))
        self.table.setItem(7, 1, SpreadSheetItem('16/6/2006'))
        self.table.setItem(8, 1, SpreadSheetItem('18/6/2006'))
        self.table.setItem(9, 1, SpreadSheetItem())
        self.table.item(9, 1).setBackground(Qt.lightGray)
        # column 2
        self.table.setItem(0, 2, SpreadSheetItem('Price'))
        self.table.item(0, 2).setBackground(titleBackground)
        self.table.item(0, 2).setToolTip('This column shows the price of the purchase')
        self.table.item(0, 2).setFont(titleFont)
        self.table.setItem(1, 2, SpreadSheetItem('150'))
        self.table.setItem(2, 2, SpreadSheetItem('2350'))
        self.table.setItem(3, 2, SpreadSheetItem('-14'))
        self.table.setItem(4, 2, SpreadSheetItem('980'))
        self.table.setItem(5, 2, SpreadSheetItem('5'))
        self.table.setItem(6, 2, SpreadSheetItem('120'))
        self.table.setItem(7, 2, SpreadSheetItem('300'))
        self.table.setItem(8, 2, SpreadSheetItem('1240'))
        self.table.setItem(9, 2, SpreadSheetItem())
        self.table.item(9, 2).setBackground(Qt.lightGray)
        # column 3
        self.table.setItem(0, 3, SpreadSheetItem('Currency'))
        self.table.item(0, 3).setBackground(titleBackground)
        self.table.item(0, 3).setToolTip('This column shows the currency')
        self.table.item(0, 3).setFont(titleFont)
        self.table.setItem(1, 3, SpreadSheetItem('NOK'))
        self.table.setItem(2, 3, SpreadSheetItem('NOK'))
        self.table.setItem(3, 3, SpreadSheetItem('EUR'))
        self.table.setItem(4, 3, SpreadSheetItem('EUR'))
        self.table.setItem(5, 3, SpreadSheetItem('USD'))
        self.table.setItem(6, 3, SpreadSheetItem('USD'))
        self.table.setItem(7, 3, SpreadSheetItem('USD'))
        self.table.setItem(8, 3, SpreadSheetItem('USD'))
        self.table.setItem(9, 3, SpreadSheetItem())
        self.table.item(9, 3).setBackground(Qt.lightGray)
        # column 4
        self.table.setItem(0, 4, SpreadSheetItem('Ex. Rate'))
        self.table.item(0, 4).setBackground(titleBackground)
        self.table.item(0, 4).setToolTip('This column shows the exchange rate to NOK')
        self.table.item(0, 4).setFont(titleFont)
        self.table.setItem(1, 4, SpreadSheetItem('1'))
        self.table.setItem(2, 4, SpreadSheetItem('1'))
        self.table.setItem(3, 4, SpreadSheetItem('8'))
        self.table.setItem(4, 4, SpreadSheetItem('8'))
        self.table.setItem(5, 4, SpreadSheetItem('7'))
        self.table.setItem(6, 4, SpreadSheetItem('7'))
        self.table.setItem(7, 4, SpreadSheetItem('7'))
        self.table.setItem(8, 4, SpreadSheetItem('7'))
        self.table.setItem(9, 4, SpreadSheetItem())
        self.table.item(9, 4).setBackground(Qt.lightGray)
        # column 5
        self.table.setItem(0, 5, SpreadSheetItem('NOK'))
        self.table.item(0, 5).setBackground(titleBackground)
        self.table.item(0, 5).setToolTip('This column shows the expenses in NOK')
        self.table.item(0, 5).setFont(titleFont)
        self.table.setItem(1, 5, SpreadSheetItem('* C2 E2'))
        self.table.setItem(2, 5, SpreadSheetItem('* C3 E3'))
        self.table.setItem(3, 5, SpreadSheetItem('* C4 E4'))
        self.table.setItem(4, 5, SpreadSheetItem('* C5 E5'))
        self.table.setItem(5, 5, SpreadSheetItem('* C6 E6'))
        self.table.setItem(6, 5, SpreadSheetItem('* C7 E7'))
        self.table.setItem(7, 5, SpreadSheetItem('* C8 E8'))
        self.table.setItem(8, 5, SpreadSheetItem('* C9 E9'))
        self.table.setItem(9, 5, SpreadSheetItem('sum F2 F9'))
        self.table.item(9, 5).setBackground(Qt.lightGray)

    def show_about(self):
        return show_about_window(self)

    def print_(self):
        printer = QPrinter(QPrinter.ScreenResolution)
        dlg = QPrintPreviewDialog(printer)
        view = PrintView()
        view.setModel(self.table.model())
        dlg.paintRequested.connect(view.print_)
        dlg.exec_()


if __name__ == '__main__':

    import sys

    app = QApplication(sys.argv)
    sheet = SpreadSheet(10, 6)
    sheet.setWindowIcon(QIcon(QPixmap(':/images/interview.png')))
    sheet.show()
    sys.exit(app.exec_())
