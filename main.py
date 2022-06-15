import sys
import os
import winreg

from PyQt5.QtGui import QIcon, QColor, QFont
from PyQt5.QtCore import QObject, QAbstractTableModel, pyqtSignal, Qt, QThread, QDir, QVariant, QModelIndex, QSettings
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog, QFileDialog, QMessageBox, QTableView

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

    """def flags(self, index):
        if index.isValid():
            if index.column() == 0:
                return Qt.ItemIsSelectable | Qt.ItemIsEditable"""


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
        self.main_window.export_action.triggered.connect(self.writeCsv)
        self.main_window.about_action.triggered.connect(lambda x: about.exec_())

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\QtProject\\OrganizationDefaults\\FileDialog")
        except FileNotFoundError:
            print('Not Found')

        self.last_path = winreg.QueryValueEx(key, "lastVisited")[0].split('///')[1]


    def loadCsv(self, fileName):
        fileName, _ = QFileDialog.getOpenFileName(self, "Выбрать карту Modbus",
                                                  (self.last_path + "/Modbus RTU map.txt"),
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
                self.regs = []
                self.ints = []
                self.rows = []
                self.max_len = 0
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

                    if len(row[4]) > self.max_len:
                        self.max_len = len(row[4])

                    row.append('')
                    self.items = self.rows.append([r for r in row])
                    cort = (int(row[0]), row[1].strip())
                    self.regs.append(cort)

            self.name = name_label[0].split(':')[1].strip().split('.')[0]
            self.model = TableModel(self.rows)
            self.table.setModel(self.model)
            self.table.resizeColumnsToContents()
            self.dict = {}
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
        fileName, _ = QFileDialog.getSaveFileName(self, "Экспорт данных",
                        (QDir.homePath() + "/" + self.fname + ' ' + self.name),
                        "(*.txt)")
        if fileName:
            print(fileName)
            f = open(fileName, 'w', newline='')
            with f:
                writer = csv.writer(f, delimiter='|')
                f.write(self.name)
                writer.writerow('')

                for row in range(self.model.rowCount(0)):
                    row_data = []
                    for col in range(self.model.columnCount(0)):
                        row_data.append(self.model.index(row, col, QModelIndex()).data(Qt.DisplayRole))
                        if row_data == '':
                            continue
                    row_data[4] = row_data[4].ljust(self.max_len + 1, ' ')
                    writer.writerow(row_data)


class ChildWindow(QDialog, Ui):
    def __init__(self):
        super(Ui, self).__init__()
        QDialog.__init__(self, None, Qt.WindowFlags(Qt.WA_DeleteOnClose))
        self.settings = SettingsWin()
        self.settings.setupUi(self)

        self.cb_data = []
        self.com_stat_list = []
        self.client = None

        win.main_window.slave_id_sb.clear()
        win.main_window.slave_id_sb.setValue(1)

        win.main_window.scan_rate_sb.clear()
        win.main_window.scan_rate_sb.setValue(1000)

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
        try:
            if self.com_stat_list[0] != self.com_stat_list[1]:
                self.stop()
        except:
            pass
        del self.com_stat_list[0]

    def thread_start(self):
        self.threadd = QThread(parent=self)
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
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText("Ошибка подключения")
        msg.setInformativeText("Проверьте настройки, затем подключите кабель к любому USB порту и нажмите 'Ок'")
        msg.setWindowTitle("MODBUSSER")
        msg.setWindowIcon(QIcon('icon.ico'))
        msg.exec_()
        self.update_ports()


class AboutWindow(QDialog):
    def __init__(self):
        QDialog.__init__(self, None, Qt.WindowFlags(Qt.WA_DeleteOnClose))
        self.about = AboutWin()
        self.about.setupUi(self)
        self.setWindowFlag(Qt.WindowCloseButtonHint)
        self.about.ok_btn.clicked.connect(self.close)


class Upd(QObject):
    update_signal = pyqtSignal()
    stop_signal = pyqtSignal()
    err_signal = pyqtSignal()


class ReadHoldingRegisters(QObject):
    def __init__(self):
        super(ReadHoldingRegisters, self).__init__()
        self.upd = Upd()
        self.upd.update_signal.connect(child.connected)
        self.upd.stop_signal.connect(child.stop)
        self.upd.err_signal.connect(child.error)

    def new_it(self, response):
        win.model.insertRows(response)
        if child._isRunning:
            self.upd.update_signal.emit()

    def run(self):
        mb_data = []  # Список считанных данных
        response = []  # Сконвертированный отформатированный список данных

        count = 125  # Максимальное кол-во регистров, для
        # одновременного считывания (ограничение протокола модбас)

        slaveID = int(win.main_window.slave_id_sb.value())  # Адрес slave устройства

        while child._isRunning:
            time.sleep((int(win.main_window.scan_rate_sb.value())) / 1000)  # Scan Rate
            reg_blocks = win.regs[-1][0] // count  # Количесвто блоков по 125 регистров
            last_block = win.regs[-1][0] - reg_blocks * count  # Последний блок с
                                                               # остаточным количесвтом регистров
            adr = 0
            delta = 1

            for i in range(reg_blocks):
                delta = 0
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
                    self.reg = struct.pack('>HH', mb_data[self.register - delta], mb_data[self.register - 1 - delta])
                    self.reg_float = [struct.unpack('>f', self.reg)]
                    response.append(round(self.reg_float[0][0], 5))

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
            mb_data.clear()
            response.clear()

        #except:
         #   if child._isRunning:
          #      self.upd.err_signal.emit()


def child_exec():
    if child.client:
        child.stop()
    else:
        child.update_ports()
    child.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('icon.ico'))
    win = Ui()
    child = ChildWindow()
    about = AboutWindow()
    worker = ReadHoldingRegisters()
    win.show()
    btn = win.main_window.set_btn
    btn.clicked.connect(child_exec)
    sys.exit(app.exec_())

