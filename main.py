import sys
import os

from PyQt5.QtGui import QIcon, QColor, QStandardItemModel, QFont, QStandardItem
from PyQt5.QtCore import QObject, QAbstractTableModel, pyqtSignal, Qt, QThread, QDir, QVariant, QModelIndex
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog, QStyledItemDelegate, QFileDialog, QMessageBox, QCheckBox, QTableView

from pymodbus.client.sync import ModbusSerialClient
from PyQt5.QtSerialPort import QSerialPortInfo

import csv
import time
import struct

from MainWindow import MainWin
from DialogWindow import SettingsWin
from AboutWindow import AboutWin


class TableModel(QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data
        self.header = ['Регистр', 'Тип данных', 'ID', 'Статус', 'Наименование', 'Значение']

    def data(self, index, role):
        if role == Qt.DisplayRole:
            # See below for the nested-list data structure.
            # .row() indexes into the outer list,
            # .column() indexes into the sub-list
            return self._data[index.row()][index.column()]

        elif role == Qt.TextAlignmentRole:
            if index.column() == 5:
                return Qt.AlignVCenter | Qt.AlignLeft
            return Qt.AlignVCenter | Qt.AlignHCenter

        elif role == Qt.ForegroundRole:
            if index.column() == 4:
                return QColor('grey')

    def rowCount(self, index):
        # The length of the outer list.
        return len(self._data)

    def columnCount(self, index):
        # The following takes the first sub-list, and returns
        # the length (only works if all rows are an equal length)
        return len(self._data[0])

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return QVariant(self.header[col])
        return QVariant()

    def insertRows(self, rows, parent=QModelIndex()):
        self.beginInsertRows(parent, 0, -1)
        for i, row in enumerate(rows):
            self._data[i][5] = row
        self.endInsertRows()
        return True

    def flags(self, index):
        if index.isValid():
            if index.column() == 0:
                return Qt.ItemIsSelectable | Qt.ItemIsEditable



class Ui(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        super(Ui, self).__init__()

        self.main_window = MainWin()
        self.main_window.setupUi(self)

        self.table = self.main_window.tableView
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setShowGrid(True)
        self.table.setGeometry(10, 50, 780, 645)
        self.table.setSelectionMode(0)

        # Buttons connections
        self.main_window.open_action.triggered.connect(self.loadCsv)
        self.main_window.export_action.triggered.connect(self.testAdd)

        self.test = []


    def testAdd(self):
        self.model.insertRows(self.test)
        self.test = [str(i + 21) for i in range(len(self.test))]

    def loadCsv(self, fileName):
        fileName, _ = QFileDialog.getOpenFileName(self, "Выбрать карту Modbus",
                                                  (QDir.homePath() + "/Modbus RTU map.txt"),
                                                  "Modbus RTU map(*.txt)")

        if fileName:
            self.loadCsvOnOpen(fileName)

    def loadCsvOnOpen(self, fileName):
        self.list_of_items = []
        if fileName:
            print(fileName + " loaded")
            font = QFont()
            font.setBold(True)
            f = open(fileName, 'r')
            with f:
                self.fileName = fileName
                self.fname = os.path.splitext(str(fileName))[0].split("///")[-1]
                self.setWindowTitle('MODBUSSER   ' + self.fname)
                regs = []
                self.ints = []
                self.rows = []
                name_label = []
                for row in csv.reader(f, delimiter="|"):
                    if len(row) == 1:
                        name_label.append(row[0])
                        continue
                    if row == []:
                        continue
                    if row[1] == ' INT ':
                        self.ints.append(int(row[0].strip(' ').strip()))
                    if row[4].strip()[0] == '=':
                        self.ints.append(row[4].strip(' ='))  # .split(':')[1].strip())
                    if row[0] == '     ':
                        continue

                    row[4] = row[4].strip(' ')
                    row.append('')
                    self.items = self.rows.append([r for r in row])

                    #self.model.insertRows(self.items)
                    #cort = (int(row[0]), row[1].strip())
                    #regs.append(cort)
                    #row.insert(0, QCheckBox(row[0]))

            self.test = [i[0] for i in self.rows]
            print(self.test)
            self.model = TableModel(self.rows)
            self.table.setModel(self.model)
            self.table.resizeColumnsToContents()
            self.dict = dict()
            current_digit = None
            for el in self.ints:
                if isinstance(el, int):
                    current_digit = el
                    self.dict[el] = []
                else:
                    self.dict[current_digit].append(el)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Ui()
    window.show()
    app.exec_()