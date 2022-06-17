from ctypes import resize
import sys
import os
import winreg

from PyQt5 import QtGui, QtCore, QtWidgets

from PyQt5.QtGui import QIcon, QColor, QFont
from PyQt5.QtCore import QObject, QAbstractTableModel, pyqtSignal, Qt, QThread, QDir, QVariant, QModelIndex, \
    QPersistentModelIndex, pyqtSlot, QTimer
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog, QFileDialog, QMessageBox, QCheckBox, QStyledItemDelegate

from pymodbus.client.sync import ModbusSerialClient
from PyQt5.QtSerialPort import QSerialPortInfo

import csv
import time
import struct

from MainWindow import MainWin
from DialogWindow import SettingsWin
from AboutWindow import AboutWin



class CustomDelegate(QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        value = index.data(QtCore.Qt.CheckStateRole)
        model = index.model()
        model.setData(index, QtCore.Qt.Unchecked, QtCore.Qt.CheckStateRole)
        super().initStyleOption(option, index)
        option.direction = QtCore.Qt.LeftToRight
        option.displayAlignment = QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter


class RowColorDelegateSilver(QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        if option.text.strip():
            option.backgroundBrush = QtGui.QColor("silver")


class RowColorDelegateWhite(QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        if option.text.strip():
            option.backgroundBrush = QtGui.QColor("white")


class TableModel(QtCore.QAbstractTableModel):
    def __init__(self):
        super(TableModel, self).__init__()
        self._data = []
        self.checkStates = {}
        self.header = ['Регистр', 'Тип данных', 'ID', 'Статус', 'Наименование', 'Значение']
        self.flag = False

    def data(self, index, role):
        if role == QtCore.Qt.DisplayRole:
            return self._data[index.row()][index.column()]
        elif role == QtCore.Qt.TextAlignmentRole:
            if index.column() == 5:
                return QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft
            return QtCore.Qt.AlignVCenter | QtCore.Qt.AlignHCenter
        elif role == QtCore.Qt.ForegroundRole:
            if index.column() == 4:
                return QtGui.QColor('grey')
        elif role == QtCore.Qt.CheckStateRole:
            value = self.checkStates.get(QtCore.QPersistentModelIndex(index))
            if value is not None:
                return value

    def setData(self, index, value, role=QtCore.Qt.EditRole):
        row = index.row()
        column = index.column()
        if role == QtCore.Qt.EditRole:
            if not value:
                self._data[row][column] = self._data[index.row()][index.column()]
            else:
                self._data[row][column] = value
            self.dataChanged.emit(index, index, (role,))
            return True
        if role == QtCore.Qt.CheckStateRole:
            self.checkStates[QtCore.QPersistentModelIndex(index)] = value
            self.dataChanged.emit(index, index, (role,))
            if self.flag:
                upd.row_brush_signal.emit(index.row(), value)
            return True
        return False

    def rowCount(self, index):
        return len(self._data)

    def columnCount(self, index):
        return 6

    def headerData(self, col, orientation, role):
        if orientation == QtCore.Qt.Horizontal and role == QtCore.Qt.DisplayRole:
            return QtCore.QVariant(self.header[col])
        return QtCore.QVariant()

    def insertRows(self, rows, parent=QtCore.QModelIndex()):
        self.beginInsertRows(parent, 0, len(rows)-1)
        """for i, row in enumerate(rows):
            self._data[i] = row"""
        self._data.extend(rows)
        self.endInsertRows()
        return True

    def removeRows(self, position, rows, parent=QtCore.QModelIndex()):
        start, end = position, rows 
        self.beginRemoveRows(parent, start, end)
        self._data.clear()
        self.endRemoveRows()
        return True

    def updateTable(self):
        index_1 = self.index(0, 0)
        index_2 = self.index(0, -1)
        self.dataChanged.emit(index_1, index_2, [Qt.DisplayRole])

    def flags(self, index):
        if index.isValid():
            return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable | QtCore.Qt.ItemIsUserCheckable


class Ui(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        super(Ui, self).__init__()

        self.main_window = MainWin()
        self.main_window.setupUi(self)

        self.table = self.main_window.tableView
        self.table.horizontalHeader().setStretchLastSection(True)
        #self.table.setShowGrid(True)
        #self.table.setGeometry(10, 50, 780, 645)
        #self.table.setSelectionMode(0)

        """# Table model object
        self.rows = [['' for i in range(6)]]
        self.rowCount = 0
        self.model = TableModel(self.rows)
        self.table.setModel(self.model)"""

        # Row color delegates
        self.silver = RowColorDelegateSilver()
        self.white = RowColorDelegateWhite()

        # Buttons connections
        self.main_window.open_action.triggered.connect(self.loadCsv)
        self.main_window.export_action.triggered.connect(self.writeCsv)
        self.main_window.about_action.triggered.connect(lambda x: about.exec_())
        self.main_window.status_box.stateChanged.connect(self.hide_minuses)
        self.main_window.status_box.setChecked(True)
        self.main_window.resetButton.clicked.connect(self.resetRowBrush)

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\QtProject\\OrganizationDefaults\\FileDialog")
        except FileNotFoundError:
            print('Not Found')

        self.last_path = winreg.QueryValueEx(key, "lastVisited")[0].split('///')[1].replace("%60", "`")

        self.tableResizeAccept = False
        upd.row_brush_signal.connect(lambda row, state: self.rowColor(row, state))

    def resetRowBrush(self):
        for row in range(self.model.rowCount(0)):
            self.rowColor(row, 0)
        if self.model.checkStates:
            for key in self.model.checkStates:
                self.model.checkStates[key] = 0

    def rowColor(self, row, state):
        if state:
            self.table.setItemDelegateForRow(row, self.silver)
        else:
            self.table.setItemDelegateForRow(row, self.white)

    def loadCsv(self, fileName):
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Выбрать карту Modbus",
                                                  (self.last_path + "/Modbus RTU map.txt"),
                                                  "*.txt")
        if fileName:
            self.loadCsvOnOpen(fileName)
            self.model = TableModel()
            self.table.setModel(self.model)
            self.model.insertRows(self.rows)
            if not self.tableResizeAccept:
                self.delegate = CustomDelegate()
                self.table.setItemDelegateForColumn(0, self.delegate)
                self.table.resizeColumnsToContents()
                self.tableResizeAccept = True
            self.model.flag = True
            self.hide_minuses()
            self.resetRowBrush()

    def loadCsvOnOpen(self, fileName):
        if child.client:
            child.stop()
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
                self.regs = []
                self.ints = []
                self.rows = []
                self.max_len = 0
                self.name = ()
                for row in csv.reader(f, delimiter="|"):
                    #self.rowCount += 1
                    if len(row) == 1:
                        self.name += row[0],
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
                    if len(row[4]) > self.max_len:
                        self.max_len = len(row[4])

                    row.append('')
                    self.rows.append(row)
                    cort = (int(row[0]), row[1].strip())
                    self.regs.append(cort)

            self.rowCount = 0
            self.main_window.name_label.setText(self.name[0])
            self.dict = {}
            self.minus_rows = [i for i, e in enumerate(self.rows) if ' - ' in e]

            current_digit = None
            for el in self.ints:
                if isinstance(el, int):
                    current_digit = el
                    self.dict[el] = []
                else:
                    self.dict[current_digit].append(el)

    def writeCsv(self, fileName):
        if child.client:
            child.stop()

        checkedStates = [v for v in self.model.checkStates.values()]
        checkedRows = [i for i, row in enumerate(checkedStates) if row == 2]
        rows = [range(self.model.rowCount(0)), checkedRows][any(checkedRows)]
        jsonName = self.name[0].split(':')[1].strip(' .json')

        if checkedRows:
            jsonName += ' selected rows'

        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Экспорт данных",
                                                  (f'{QtCore.QDir.homePath()}/{self.fname} {jsonName}'),
                                                  "(*.txt)")
        if fileName:
            print(fileName)
            f = open(fileName, 'w', newline='')
            with f:
                writer = csv.writer(f, delimiter='|')
                f.write(jsonName)
                writer.writerow("")
                writer.writerow("")

                for row in rows:
                    row_data = []
                    for col in range(self.model.columnCount(0)):
                        row_data.append(
                            self.model.index(row, col, QtCore.QModelIndex()).data(QtCore.Qt.DisplayRole))
                        if row_data == '':
                            continue
                    row_data[4] = row_data[4].ljust(self.max_len + 1, ' ')
                    writer.writerow(row_data)

    def hide_minuses(self):
        try:
            for idx in self.minus_rows:
                if self.main_window.status_box.checkState():
                    self.main_window.tableView.setRowHidden(idx, True)
                else:
                    self.main_window.tableView.setRowHidden(idx, False)
        except:
            print("Failed hide minuses")


class ChildWindow(QDialog, Ui):
    def __init__(self):
        super(Ui, self).__init__()
        QtWidgets.QDialog.__init__(self, None, QtCore.Qt.WindowFlags(QtCore.Qt.WA_DeleteOnClose))
        self.settings = SettingsWin()
        self.settings.setupUi(self)

        self.cb_data = []
        self.com_stat_list = []
        self.client = None

        win.main_window.slave_id_sb.clear()
        win.main_window.slave_id_sb.setValue(1)

        win.main_window.scan_rate_sb.clear()
        win.main_window.scan_rate_sb.setValue(500)

        self.settings.buttonBox.accepted.connect(self.accept_data)
        self.settings.buttonBox.rejected.connect(self.reject_data)
        win.main_window.start_btn.clicked.connect(self.first_start)

        speeds = ['1200', '2400', '4800', '9600', '19200', '38400', '57600', '115200']
        self.settings.baud_set.clear()
        self.settings.baud_set.addItems(speeds)
        self.settings.baud_set.setCurrentIndex(5)

        parity = ['N', 'E']
        self.settings.parity_set.clear()
        self.settings.parity_set.addItems(parity)

        stop_bits = ['1', '2']
        self.settings.sb_set.clear()
        self.settings.sb_set.addItems(stop_bits)

        self.accept_data()
        self.update_ports()

    def update_ports(self):
        self.settings.com_set.clear()
        info_list = QSerialPortInfo()
        serial_list = info_list.availablePorts()
        self.serial_ports = [(f'{port.portName()} {port.description()}') for port in serial_list]
        self.settings.com_set.addItems(self.serial_ports)
        self.cb_data[0] = self.settings.com_set.currentText()
        self.set_text()

    def set_text(self):
        self.com_port_label_text = (f'{self.cb_data[0]} ({self.cb_data[1]}-{self.cb_data[2]}-{self.cb_data[3]})')
        self.com_stat_list.append(self.com_port_label_text)
        win.main_window.com_port_label.setText(self.com_port_label_text)

    def accept_data(self):
        self.cb_data.clear()
        self.cb_data.append(self.settings.com_set.currentText())
        self.cb_data.append(self.settings.baud_set.currentText())
        self.cb_data.append(self.settings.parity_set.currentText())
        self.cb_data.append(self.settings.sb_set.currentText())
        self.set_text()
        self.close()
        del self.com_stat_list[0]

    def thread_start(self):
        self.threadd = QtCore.QThread(parent=self)
        worker.moveToThread(self.threadd)
        self.threadd.started.connect(worker.run)

    def reject_data(self):
        try:
            self.settings.com_set.setCurrentText(self.cb_data[0])
            self.settings.baud_set.setCurrentText(self.cb_data[1])
            self.settings.parity_set.setCurrentText(self.cb_data[2])
            self.settings.sb_set.setCurrentText(self.cb_data[3])
        except:
            pass
        self.close()

    def first_start(self):
        self.thread_start()
        self.start()

    def start(self):
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(win.model.updateTable)
        self.timer.start(400)
        try:
            self.connect()
        except:
            win.main_ui.status_label.setText('Error')
        self._isRunning = True
        self.threadd.start()
        win.main_window.start_btn.setText('Stop')
        win.main_window.start_btn.clicked.disconnect()
        win.main_window.start_btn.clicked.connect(self.stop)
        win.main_window.status_label.setText('Connection...')

    def connect(self):
        self.client = ModbusSerialClient(
            method='rtu',
            port=self.settings.com_set.currentText().split(' ')[0].strip(),
            baudrate=int(self.settings.baud_set.currentText()),
            timeout=1,
            parity=self.settings.parity_set.currentText(),
            stopbits=int(self.settings.sb_set.currentText()),
            bytesize=8)
        self.client.connect()

    def connected(self):
        win.main_window.status_label.setText('Connected!')
        win.main_window.status_label.setStyleSheet('color: rgb(0, 128, 0);')

    def stop(self):
        self.threadd.terminate()
        self.timer.stop()
        self._isRunning = False
        try:
            self.client.socket.close()
        except:
            win.main_window.status_label.setText('Может хватит? ;)')
        win.main_window.status_label.setStyleSheet('color: rgb(230, 0, 0);')
        win.main_window.status_label.setText('Disconnected')
        win.main_window.start_btn.setText('Connect')
        win.main_window.start_btn.clicked.disconnect()
        win.main_window.start_btn.clicked.connect(self.start)
        self.client = None

    def error(self):
        self.stop()
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText("Ошибка подключения")
        msg.setInformativeText("Проверьте настройки, затем подключите кабель к любому USB порту и нажмите 'Ок'")
        msg.setWindowTitle("MODBUSSER")
        msg.setWindowIcon(QIcon('icon.ico'))
        msg.exec_()
        self.update_ports()


class AboutWindow(QtWidgets.QDialog):
    def __init__(self):
        QtWidgets.QDialog.__init__(self, None, QtCore.Qt.WindowFlags(QtCore.Qt.WA_DeleteOnClose))
        self.about = AboutWin()
        self.about.setupUi(self)
        self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint)
        self.about.ok_btn.clicked.connect(self.close)


class Upd(QtCore.QObject):
    update_signal = QtCore.pyqtSignal()
    stop_signal = QtCore.pyqtSignal()
    err_signal = QtCore.pyqtSignal()
    row_brush_signal = QtCore.pyqtSignal(int, int)


class ReadHoldingRegisters(QtCore.QObject):
    def __init__(self):
        super(ReadHoldingRegisters, self).__init__()
        self.upd = Upd()
        self.upd.update_signal.connect(child.connected)
        self.upd.stop_signal.connect(child.stop)
        self.upd.err_signal.connect(child.error)

    def new_it(self, response):
        for i in range(len(win.rows)):
            win.rows[i][-1] = response[i]
        if child._isRunning:
            self.upd.update_signal.emit()

    def run(self):
        mb_data = []  # Список считанных данных
        response = []  # Сконвертированный отформатированный список данных
        count = 125  # Максимальное кол-во регистров, для
        # одновременного считывания (ограничение протокола модбас)

        slaveID = int(win.main_window.slave_id_sb.value())  # Адрес slave устройства

        #try:
        while child._isRunning:
            time.sleep((int(win.main_window.scan_rate_sb.value())) / 1000)  # Scan Rate
            reg_blocks = win.regs[-1][0] // count  # Количесвто блоков по 125 регистров
            last_block = win.regs[-1][0] - reg_blocks * count  # Последний блок с
            # остаточным количесвтом регистров
            adr = 0

            for i in range(reg_blocks):
                self.res = child.client.read_holding_registers(address=adr,
                                                                count=count,
                                                                unit=slaveID)
                adr = adr + count
                mb_data.extend(self.res.registers)

            self.res = child.client.read_holding_registers(address=adr,
                                                            count=last_block,
                                                            unit=slaveID)
            mb_data.extend(self.res.registers)

            for i in range(len(win.regs)):
                self.register, data_type = win.regs[i][0], win.regs[i][1]

                if data_type == 'MEA':
                    self.reg = struct.pack('>HH', mb_data[self.register], mb_data[self.register - 1])
                    self.reg_float = [struct.unpack('>f', self.reg)]
                    response.append(str(round(self.reg_float[0][0], 3)))

                elif data_type == 'BIN':
                    self.reg = mb_data[self.register - 1]

                    ones = []
                    for x in range(16):
                        if (self.reg >> x) & 0x1:
                            ones.append(' ' + (str(x)))
                    ones.reverse()

                    self.reg = f"{mb_data[self.register - 1]:>016b}"
                    for j in range(len(ones)):
                        self.reg = str(self.reg) + ' ' + ones[j]
                    if self.reg[15] == '1':
                        self.reg = 'ДА   ' + self.reg
                    elif self.reg[15] == '0':
                        self.reg = 'НЕТ  ' + self.reg
                    if self.reg[14] == '1' and self.reg[20] == '1' and int(self.reg[5:20], 2) < 65:
                        self.reg = self.reg + ' (Тревога)'
                    elif self.reg[13] == '1' and self.reg[20] == '1' and int(self.reg[5:20], 2) < 129:
                        self.reg = self.reg + ' (Предупреждение)'
                    response.append(self.reg)

                elif data_type == 'INT':
                    self.reg = mb_data[self.register - 1]
                    try:
                        if int(self.reg) > int(len(win.dict[self.register][-1])):
                            continue
                        else:
                            self.reg = str(win.dict[self.register][self.reg]).replace(': ', ' - ')
                    except:
                        pass
                    response.append(self.reg)

            self.new_it(response)
            win.model.updateTable()
            mb_data.clear()
            response.clear()

        """except Exception:
            if child._isRunning:
                self.upd.err_signal.emit()
"""

def child_exec():
    if child.client:
        child.stop()
    else:
        child.update_ports()
    child.show()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QIcon('icon.ico'))
    upd = Upd()

    win = Ui()
    child = ChildWindow()
    about = AboutWindow()
    worker = ReadHoldingRegisters()
    win.show()
    btn = win.main_window.set_btn
    btn.clicked.connect(child_exec)
    sys.exit(app.exec_())

