from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QLabel, QComboBox, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QGroupBox
from dataframes.models import Coordinates
from visuals.spreadsheetitem import SpreadSheetItem
from components.TableWidget import TableWidget


class InputDialog(QDialog):
    def __init__(self, title: str, c1Text: str, c2Text: str, opText: str,
        outText: str, cell1: SpreadSheetItem, cell2: SpreadSheetItem,
                 outCell: SpreadSheetItem, table: TableWidget):
        super(InputDialog, self).__init__()
        rows = []
        cols = []
        for r in range(table.rowCount()):
            rows.append(str(r + 1))
        for c in range(table.columnCount()):
            cols.append(chr(ord('A') + c))
        self.setWindowTitle(title)
        group = QGroupBox(title, self)
        group.setMinimumSize(250, 100)
        # First cell configuration
        cell1Label = QLabel(c1Text, group)
        self.cell1RowInput = QComboBox(group)
        cell1Coordinates = Coordinates(cell1)
        self.cell1RowInput.addItems(rows)
        self.cell1RowInput.setCurrentIndex(cell1Coordinates.x)
        self.cell1ColInput = QComboBox(group)
        self.cell1ColInput.addItems(cols)
        self.cell1ColInput.setCurrentIndex(cell1Coordinates.y)

        operatorLabel = QLabel(opText, group)
        operatorLabel.setAlignment(Qt.AlignHCenter)

        cell2Label = QLabel(c2Text, group)
        self.cell2RowInput = QComboBox(group)
        cell2Coordinates = Coordinates(cell2)
        self.cell2RowInput.addItems(rows)
        self.cell2RowInput.setCurrentIndex(cell2Coordinates.x)
        self.cell2ColInput = QComboBox(group)
        self.cell2ColInput.addItems(cols)
        self.cell2ColInput.setCurrentIndex(cell2Coordinates.y)

        equalsLabel = QLabel('=', group)
        equalsLabel.setAlignment(Qt.AlignHCenter)

        outLabel = QLabel(outText, group)
        self.outRowInput = QComboBox(group)
        outCoordinates = Coordinates(outCell)
        self.outRowInput.addItems(rows)
        self.outRowInput.setCurrentIndex(outCoordinates.x)
        self.outColInput = QComboBox(group)
        self.outColInput.addItems(cols)
        self.outColInput.setCurrentIndex(outCoordinates.y)

        cancelButton = QPushButton('Cancel', self)
        cancelButton.clicked.connect(self.reject)
        okButton = QPushButton('OK', self)
        okButton.setDefault(True)
        okButton.clicked.connect(self.accept)
        buttonsLayout = QHBoxLayout()
        buttonsLayout.addStretch(1)
        buttonsLayout.addWidget(okButton)
        buttonsLayout.addSpacing(10)
        buttonsLayout.addWidget(cancelButton)

        dialogLayout = QVBoxLayout(self)
        dialogLayout.addWidget(group)
        dialogLayout.addStretch(1)
        dialogLayout.addItem(buttonsLayout)

        cell1Layout = QHBoxLayout()
        cell1Layout.addWidget(cell1Label)
        cell1Layout.addSpacing(10)
        cell1Layout.addWidget(self.cell1ColInput)
        cell1Layout.addSpacing(10)
        cell1Layout.addWidget(self.cell1RowInput)

        cell2Layout = QHBoxLayout()
        cell2Layout.addWidget(cell2Label)
        cell2Layout.addSpacing(10)
        cell2Layout.addWidget(self.cell2ColInput)
        cell2Layout.addSpacing(10)
        cell2Layout.addWidget(self.cell2RowInput)

        outLayout = QHBoxLayout()
        outLayout.addWidget(outLabel)
        outLayout.addSpacing(10)
        outLayout.addWidget(self.outColInput)
        outLayout.addSpacing(10)
        outLayout.addWidget(self.outRowInput)

        vLayout = QVBoxLayout(group)
        vLayout.addItem(cell1Layout)
        vLayout.addWidget(operatorLabel)
        vLayout.addItem(cell2Layout)
        vLayout.addWidget(equalsLabel)
        vLayout.addStretch(1)
        vLayout.addItem(outLayout)
