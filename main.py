import sys
import os
import winreg as reg

import pymodbus
from PyQt5 import QtGui, QtCore, QtWidgets
from pymodbus.client.sync import ModbusSerialClient
from PyQt5.QtSerialPort import QSerialPortInfo
from PyQt5.Qt import Qt

import csv
import time
import struct

from MainWindow import MainWin
from DialogWindow import SettingsWin
from AboutWindow import AboutWin


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
    """
    Custom model for our table
    """
    def __init__(self):
        super(TableModel, self).__init__()
        self.searchProxy = QtCore.QSortFilterProxyModel(self)
        self._data = []
        # check states for 0(first) column
        self.checkStates = {}
        self.header = 'Register', 'Data type', 'ID', 'Status', 'Name', 'Value',
        self.flag = False

    def data(self, index, role):
        col = index.column()
        row = index.row()
        if index.isValid():
            if role == QtCore.Qt.DisplayRole:
                return self._data[row][col]

            if role == QtCore.Qt.TextAlignmentRole:
                if col == 5:
                    return QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft
                return QtCore.Qt.AlignVCenter | QtCore.Qt.AlignHCenter

            if role == QtCore.Qt.ForegroundRole:
                if col == 4:
                    return QtGui.QColor('brown')

            if role == QtCore.Qt.FontRole:
                font = QtGui.QFont()
                if col == 5:
                    font.setWeight(50)
                    return font

            if role == QtCore.Qt.CheckStateRole:
                if index.column() == 0:
                    return QtCore.Qt.Checked if self.checkStates[self._data[row][0]] == 1 else QtCore.Qt.Unchecked

    def setData(self, index, value, role=QtCore.Qt.EditRole):
        row = index.row()
        col = index.column()

        if role == QtCore.Qt.EditRole:
            if value:
                self._data[row][col] = value
            self.dataChanged.emit(index, index)
            return True
        if role == Qt.CheckStateRole and col == 0:
            self.checkStates[self._data[row][0]] = 1 if value == Qt.Checked else 0
            # check states when filtering data by proxy
            self.proxyCheckStates = []
            for i, el in enumerate(self.checkStates):
                # append filtered data (first column data)
                self.proxyCheckStates.append(self.searchProxy.index(i, 0, QtCore.QModelIndex()).data(QtCore.Qt.DisplayRole))
            if self.flag:
                if len(win.main_window.searchBox.text()) > 0:
                    # row index for row color delegate func when filtering (searching)
                    self.idx = self.proxyCheckStates.index(self._data[row][0])
                else:
                    # row index when proxy don't applied
                    self.idx = row
                # emit signal for brush selected row
                upd.row_brush_signal.emit(self.idx, value)
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

    def insertRows(self, rows, parent=QtCore.QModelIndex, position=0):
        self.beginInsertRows(parent(), position, len(rows) - 1)
        self._data.extend(rows)
        self.endInsertRows()
        return True

    def removeRows(self, position, rows, parent=QtCore.QModelIndex()):
        start, end = position, rows
        self.beginRemoveRows(parent, start, end)
        del self._data[-1]
        self.endRemoveRows()
        return True

    def updateTable(self):
        index_1 = self.index(0, 0)
        index_2 = self.index(0, -1)
        self.searchProxy.dataChanged.emit(index_1, index_2, [QtCore.Qt.DisplayRole])

    def flags(self, index):
        if index.column() == 0:
            return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsUserCheckable
        return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        super(Ui, self).__init__()

        self.main_window = MainWin()
        self.main_window.setupUi(self)
        self.table = self.main_window.tableView
        self.table.horizontalHeader().setStretchLastSection(True)

        # Row color delegates
        self.silver = RowColorDelegateSilver()
        self.white = RowColorDelegateWhite()

        # Buttons connections
        self.main_window.open_action.triggered.connect(self.loadCsv)
        self.main_window.export_action.triggered.connect(self.writeCsv)
        self.main_window.about_action.triggered.connect(lambda x: about.exec_())
        self.main_window.status_box.setChecked(True)
        self.main_window.resetButton.clicked.connect(lambda : self.resetRowBrush(True))
        self.main_window.resetButton.setDisabled(True)
        self.main_window.start_btn.setDisabled(True)
        self.tableResizeAccept = False
        upd.row_brush_signal.connect(lambda row, state: self.rowColor(row, state))
        self.main_window.searchBox.textEdited.connect(self.search)
        self.main_window.searchBox.setDisabled(True)
        self.regeditPath = "Software\\QtProject\\OrganizationDefaults\\FileDialog"

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F and event.modifiers() == Qt.ControlModifier:
            self.main_window.searchBox.setFocus(QtCore.Qt.ShortcutFocusReason)

    def search(self):
        """
        search register names by key input
        """
        self.model.searchProxy.setFilterKeyColumn(4)
        self.model.searchProxy.setFilterRegExp(self.main_window.searchBox.text())
        self.resetRowBrush(False)

        # correct painting selected rows for proxy data
        if self.main_window.searchBox.text():
            for i, row in enumerate(self.model.checkStates):
                if self.model.checkStates.get(self.model.searchProxy.index(i, 0, QtCore.QModelIndex()).data(QtCore.Qt.DisplayRole)) == 1:
                    self.rowColor(i, True)
        # for all data
        else:
            for i, row in enumerate(self.model.checkStates):
                self.rowColor(i, self.model.checkStates[row])

        # hide rows with minuses by default
        if self.main_window.status_box.checkState():
            self.hide_minuses()

    def resetRowBrush(self, flag):
        """
        return default color and uncheck all checboxes for all rows
        :param flag: if flag is True, reset checkboxes, else reset only row color
        :return:
        """
        for i, row in enumerate(self.model.checkStates):
            self.rowColor(i, 0)
            if flag:
                if self.model.checkStates[row] == 1:
                    self.model.checkStates[row] = 0

    def rowColor(self, row, state):
        """
        painting selected rows to silver color
        :param row: selected row
        :param state: if False - reset row color to white
        """
        if state:
            self.table.setItemDelegateForRow(row, self.silver)
        else:
            self.table.setItemDelegateForRow(row, self.white)

    def loadCsv(self, fileName):
        """
        loading Modbus RTU map.txt
        """
        # parsing last path for our file, if not, creating new
        self.key = reg.CreateKey(reg.HKEY_CURRENT_USER, self.regeditPath)
        try:
            self.lastPath = reg.QueryValueEx(self.key, "lastPath")
        except:
            self.lastPath = f"{QtCore.QDir.homePath()}\Modbus RTU Map.txt", '',

        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open Modbus RTU map",
                               f'{self.lastPath[0]}', "*.txt")
        if fileName:
            # set current path as last for next time
            reg.SetValueEx(self.key, 'lastPath', 0, reg.REG_SZ, f'{fileName}')
            self.lastPath = reg.QueryValueEx(self.key, "lastPath")

            # our table model object
            self.model = TableModel()
            self.loadCsvOnOpen(fileName)
            self.model.searchProxy.setFilterCaseSensitivity(0)
            # set proxy model
            self.model.searchProxy.setSourceModel(self.model)
            self.table.setModel(self.model.searchProxy)
            # insert new data from modbus map to table
            self.model.insertRows(self.rows)
            # hide rows with minuses
            self.main_window.status_box.stateChanged.connect(self.hide_minuses)

            if not self.tableResizeAccept:
                self.table.resizeColumnsToContents()
                self.tableResizeAccept = True
            self.model.flag = True
            self.hide_minuses()
            self.resetRowBrush(True)
            # activate all buttons
            if self.rows:
                self.main_window.start_btn.setDisabled(False)
                self.main_window.resetButton.setDisabled(False)
                self.main_window.searchBox.setDisabled(False)
                self.rows.clear()
                self.ints.clear()
        reg.CloseKey(self.key)

    def loadCsvOnOpen(self, fileName):
        if child.client:
            child.stop(True)
        if fileName:
            print(fileName + " loaded")
            f = open(fileName, 'r')
            with f:
                self.fileName = fileName
                self.fname = os.path.splitext(str(fileName))[0].split("///")[-1]
                self.setWindowTitle('AR-CON ModBus   ' + self.fname)
                # modbus registers
                self.regs = []
                # registers with "INT"(signed) data
                self.ints = []
                # parsed data form modbus map
                self.rows = []
                self.max_len = 0
                # json filename
                self.name = ()
                # iterating and parsing rows from map
                for row in csv.reader(f, delimiter="|"):
                    if len(row) == 1:
                        self.name += row[0],
                        continue
                    if row == []:
                        continue
                    if row[1] == ' INT ':
                        self.ints.append(int(row[0].strip(' ').strip()))
                    if row[4].strip()[0] == '=':
                        self.ints.append(row[4].strip(' ='))
                    if row[0] == '     ':
                        continue
                    row[4] = row[4].strip(' ')
                    if len(row[4]) > self.max_len:
                        self.max_len = len(row[4])
                    row.append('')
                    self.model.checkStates.setdefault(row[0], 0)
                    self.rows.append(row)
                    # implementing cortage with register and data type
                    cort = (int(row[0]), row[1].strip())
                    self.regs.append(cort)

            self.main_window.name_label.setText(self.name[0])
            # dict for INT data. Keys - registers with INT (signed) data type, values - states (str)
            self.dict = {}
            current_digit = None
            for el in self.ints:
                if isinstance(el, int):
                    current_digit = el
                    self.dict[el] = []
                else:
                    self.dict[current_digit].append(el)


    def writeCsv(self, fileName):
        """
        saving data from table to txt file, delimiters is "|"
        :param fileName: path and filename
        """
        # stop modbus client
        if child.client:
            child.stop(True)

        # get checked(selected) rows for save
        checkedRows = [i for i, row in enumerate(self.model.checkStates) if self.model.checkStates[row] == 1]
        # if no checked save all data
        rows = [range(self.model.rowCount(0)), checkedRows][any(checkedRows)]
        jsonName = self.name[0].split(':')[1].strip(' .json')

        if checkedRows:
            jsonName += ' selected rows'

        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save data",
                                                  (f'{self.lastPath[0].strip(".txt")} {jsonName}'),
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
                row_data.clear()

    def hide_minuses(self):
        for row in range(self.model.searchProxy.rowCount()):
            if self.main_window.status_box.checkState():
                if self.model.searchProxy.index(row, 3, QtCore.QModelIndex()).data(QtCore.Qt.DisplayRole) == " - ":
                    self.main_window.tableView.setRowHidden(row, True)
            else:
                self.main_window.tableView.setRowHidden(row, False)


class ChildWindow(QtWidgets.QDialog, Ui):
    """
    Implementing dialog window with communication settings.
    Also with modbus client methods
    """
    def __init__(self):
        super(Ui, self).__init__()
        QtWidgets.QDialog.__init__(self, None, QtCore.Qt.WindowFlags(QtCore.Qt.WA_DeleteOnClose))
        self.settings = SettingsWin()
        self.settings.setupUi(self)

        self.cb_data = []
        self.client = None

        win.main_window.slave_id_sb.clear()
        win.main_window.slave_id_sb.setValue(1)

        win.main_window.scan_rate_sb.clear()
        win.main_window.scan_rate_sb.setValue(500)

        self.settings.buttonBox.accepted.connect(self.accept_data)
        self.settings.buttonBox.rejected.connect(self.reject_data)
        win.main_window.start_btn.clicked.connect(self.first_start)

        speeds = ('1200', '2400', '4800', '9600', '19200', '38400', '57600', '115200')
        self.settings.baud_set.clear()
        self.settings.baud_set.addItems(speeds)
        self.settings.baud_set.setCurrentIndex(5)

        parity = ('N', 'E')
        self.settings.parity_set.clear()
        self.settings.parity_set.addItems(parity)

        stop_bits = ('1', '2')
        self.settings.sb_set.clear()
        self.settings.sb_set.addItems(stop_bits)

        self.accept_data()
        self.update_ports()

    def update_ports(self):
        """
        updating active com ports
        """
        self.settings.com_set.clear()
        info_list = QSerialPortInfo()
        serial_list = info_list.availablePorts()
        self.serial_ports = [(f'{port.portName()} {port.description()}') for port in serial_list]
        self.settings.com_set.addItems(self.serial_ports)
        self.cb_data[0] = self.settings.com_set.currentText()
        self.set_text()

    def set_text(self):
        self.com_port_label_text = (f'{self.cb_data[0]} ({self.cb_data[1]}-{self.cb_data[2]}-{self.cb_data[3]})')
        win.main_window.com_port_label.setText(self.com_port_label_text)

    def accept_data(self):
        """
        save settings
        """
        self.cb_data.clear()
        self.cb_data.append(self.settings.com_set.currentText())
        self.cb_data.append(self.settings.baud_set.currentText())
        self.cb_data.append(self.settings.parity_set.currentText())
        self.cb_data.append(self.settings.sb_set.currentText())
        self.set_text()
        self.close()

    def thread_start(self):
        """
        implementing thread for modbus client
        """
        self.threadd = QtCore.QThread(parent=self)
        worker.moveToThread(self.threadd)
        self.threadd.started.connect(worker.run)

    def reject_data(self):
        """
        reject settings
        """
        try:
            self.settings.com_set.setCurrentText(self.cb_data[0])
            self.settings.baud_set.setCurrentText(self.cb_data[1])
            self.settings.parity_set.setCurrentText(self.cb_data[2])
            self.settings.sb_set.setCurrentText(self.cb_data[3])
        except:
            pass
        self.close()

    def first_start(self):
        """
        first starting starts with Qthread implement
        next starts with start/stop the Qthread
        """
        self.thread_start()
        self.start()

    def start(self):
        """
        connecting to modbus slave device and starting read data from this one
        """
        # implementing QTimer for update the table when data changed every 400 ms
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
        win.main_window.start_btn.clicked.connect(lambda x: self.stop(True))
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
        win.main_window.status_label.setText('Connected')
        win.main_window.status_label.setStyleSheet('color: rgb(0, 128, 0);')

    def stop(self, flag):
        """
        stop modbus client
        :param flag: setting status text ("Stopped") when stops by the stop button,
        if not the flag the error text will be set
        """
        self.timer.stop()
        self._isRunning = False
        try:
            self.client.socket.close()
        except:
            pass
        win.main_window.status_label.setStyleSheet('color: rgb(230, 0, 0);')
        if flag:
            win.main_window.status_label.setText('Stopped')
        win.main_window.start_btn.setText('Connect')
        win.main_window.start_btn.clicked.disconnect()
        win.main_window.start_btn.clicked.connect(self.start)
        self.threadd.terminate()
        self.client = None


class AboutWindow(QtWidgets.QDialog):
    def __init__(self):
        QtWidgets.QDialog.__init__(self, None, QtCore.Qt.WindowFlags(QtCore.Qt.WA_DeleteOnClose))
        self.about = AboutWin()
        self.about.setupUi(self)
        self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint)
        self.about.ok_btn.clicked.connect(self.close)


class Upd(QtCore.QObject):
    """
    implementing signals
    """
    update_signal = QtCore.pyqtSignal()
    stop_signal = QtCore.pyqtSignal()
    err_signal = QtCore.pyqtSignal(bool)
    row_brush_signal = QtCore.pyqtSignal(int, int)


class ReadHoldingRegisters(QtCore.QObject):
    """
    Modbus client for read holding registers. It can read more than 125 registers, 1000 for example :)
    """
    def __init__(self):
        super(ReadHoldingRegisters, self).__init__()
        self.upd = Upd()
        self.upd.update_signal.connect(child.connected)
        self.upd.stop_signal.connect(child.stop)
        self.upd.err_signal.connect(lambda x: child.stop(x))

    def new_it(self, response):
        """
        writing a new data to table with model.setData func
        :param response: the readed data from device
        """
        row = 0
        for value in response:
            idx = win.model.createIndex(row, 5)
            win.model.setData(idx, value)
            row += 1
        if child._isRunning:
            self.upd.update_signal.emit()

    def run(self):
        count = 125  # max count modbus registers for read once
        self.res = None
        slaveID = int(win.main_window.slave_id_sb.value())  # device' slave id
        reg_blocks = win.regs[-1][0] // count  # the count of blocks with 125 registers
        last_block = win.regs[-1][0] - reg_blocks * count  # the last block with a amount of the residual registers

        try:
            while child._isRunning:
                adr = 0  # the starting register
                mb_data = []  # the readed data from device
                response = []  # the parsed data
                time.sleep((int(win.main_window.scan_rate_sb.value())) / 1000)  # scan rate

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

                # parsing data by data type
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

        except pymodbus.exceptions.ConnectionException:
            win.main_window.status_label.setText(
            f"{child.settings.com_set.currentText().split(' ')[0].strip()}: N/A")
        except AttributeError:
            if type(self.res) == pymodbus.exceptions.ModbusIOException:
                win.main_window.status_label.setText("Timeout")
        finally:
            self.upd.err_signal.emit(False)


def child_exec():
    if child.client:
        child.stop(True)
    else:
        child.update_ports()
    child.show()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon('icon.ico'))
    upd = Upd()

    win = Ui()
    child = ChildWindow()
    about = AboutWindow()
    worker = ReadHoldingRegisters()
    win.show()
    btn = win.main_window.set_btn
    btn.clicked.connect(child_exec)
    sys.exit(app.exec_())

