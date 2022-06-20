import sys
import os
import winreg as reg

from pymodbus.exceptions import ModbusIOException, ConnectionException
from PyQt5 import QtGui, QtCore, QtWidgets
from pymodbus.client.sync import ModbusSerialClient
from PyQt5.QtSerialPort import QSerialPortInfo

import csv
import time
import struct

from MainWindow import Main
from DialogWindow import Settings
from AboutWindow import About


class TableModel(QtCore.QAbstractTableModel):
    """
    Custom model for our table
    """
    def __init__(self):
        super(TableModel, self).__init__()
        self.searchProxy = QtCore.QSortFilterProxyModel(self)
        self._data = []
        # check states for 0(first) column
        self.checkStates = []
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
                if col == 0:
                    return QtCore.Qt.Checked if self.checkStates[row] == 1 else QtCore.Qt.Unchecked

            if role == QtCore.Qt.BackgroundRole:
                if self.checkStates[row] == 1:
                    return QtGui.QColor("red").lighter(180)
                return QtGui.QColor("white")

    def setData(self, index, value, role=QtCore.Qt.EditRole):
        row = index.row()
        col = index.column()

        if role == QtCore.Qt.EditRole:
            if value:
                self._data[row][col] = value
            self.dataChanged.emit(index, index, (role,))
            return True

        if role == QtCore.Qt.CheckStateRole:
            self.checkStates[row] = 1 if value == QtCore.Qt.Checked else 0
            self.updateTable()
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

    def updateTable(self):
        index_1 = self.index(0, 0)
        index_2 = self.index(0, -1)
        self.searchProxy.dataChanged.emit(index_1, index_2, (QtCore.Qt.DisplayRole,))

    def flags(self, index):
        if index.column() == 0:
            return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsUserCheckable
        return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable | QtCore.Qt.ItemIsUserCheckable


class MainWin(QtWidgets.QMainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        super(MainWin, self).__init__()

        self.mainWin = Main()
        self.mainWin.setupUi(self)
        self.table = self.mainWin.tableView
        self.table.horizontalHeader().setStretchLastSection(True)

        # Buttons connections
        self.mainWin.open_action.triggered.connect(self.loadCsv)
        self.mainWin.export_action.triggered.connect(self.writeCsv)
        self.mainWin.about_action.triggered.connect(lambda x: about.exec_())
        self.mainWin.statusBox.setChecked(True)
        self.mainWin.resetButton.clicked.connect(self.resetRowBrush)
        self.mainWin.resetButton.setDisabled(True)
        self.mainWin.startBtn.setDisabled(True)
        self.tableResizeAccept = False
        self.mainWin.searchBox.textEdited.connect(self.search)
        self.mainWin.searchBox.setDisabled(True)
        self.regeditPath = "Software\\QtProject\\OrganizationDefaults\\FileDialog"

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_F and event.modifiers() == QtCore.Qt.ControlModifier:
            self.mainWin.searchBox.setFocus(QtCore.Qt.ShortcutFocusReason)

    def search(self):
        """
        search register names by key input
        """
        self.model.searchProxy.setFilterKeyColumn(4)
        self.model.searchProxy.setFilterRegExp(self.mainWin.searchBox.text())

        # hide rows with minuses by default
        if self.mainWin.statusBox.checkState():
            self.hideMinusRows()

    def resetRowBrush(self):
        """
        return default color and uncheck all checboxes for all rows
        """
        for i in range(len(self.model.checkStates)):
            self.model.checkStates[i] = 0
        self.model.updateTable()


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
            self.mainWin.statusBox.stateChanged.connect(self.hideMinusRows)

            if not self.tableResizeAccept:
                self.table.resizeColumnsToContents()
                self.tableResizeAccept = True
            self.model.flag = True
            self.hideMinusRows()
            self.resetRowBrush()
            # activate all buttons
            if self.rows:
                self.mainWin.startBtn.setDisabled(False)
                self.mainWin.resetButton.setDisabled(False)
                self.mainWin.searchBox.setDisabled(False)
                self.rows.clear()
                self.ints.clear()
        reg.CloseKey(self.key)

    def loadCsvOnOpen(self, fileName):
        if dialog.client:
            dialog.stop(True)
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
                self.maxLen = 0
                # json filename
                self.name = ()
                self.model.checkStates = []
                # iterating and parsing rows from map
                for row in (csv.reader(f, delimiter="|")):
                    self.model.checkStates.append(0)
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
                    if len(row[4]) > self.maxLen:
                        self.maxLen = len(row[4])
                    row.append('')
                    self.rows.append(row)
                    # implementing cortage with register and data type
                    cort = (int(row[0]), row[1].strip())
                    self.regs.append(cort)

            self.mainWin.name_label.setText(self.name[0])
            # dict for INT data. Keys - registers with INT (signed) data type, values - states (str)
            self.intDataType = {}
            num = None
            for el in self.ints:
                if isinstance(el, int):
                    num = el
                    self.intDataType[el] = []
                else:
                    self.intDataType[num].append(el)


    def writeCsv(self, fileName):
        """
        saving data from table to txt file, delimiters is "|"
        :param fileName: path and filename
        """
        # stop modbus client
        if dialog.client:
            dialog.stop(True)

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
                    rowData = []
                    for col in range(self.model.columnCount(0)):
                        rowData.append(
                            self.model.index(row, col, QtCore.QModelIndex()).data(QtCore.Qt.DisplayRole))
                        if rowData == '':
                            continue
                    rowData[4] = rowData[4].ljust(self.maxLen + 1, ' ')
                    writer.writerow(rowData)
                rowData.clear()

    def hideMinusRows(self):
        for row in range(self.model.searchProxy.rowCount()):
            if self.mainWin.statusBox.checkState():
                if self.model.searchProxy.index(row, 3, QtCore.QModelIndex()).data(QtCore.Qt.DisplayRole) == " - ":
                    self.mainWin.tableView.setRowHidden(row, True)
            else:
                self.mainWin.tableView.setRowHidden(row, False)


class DialogWin(QtWidgets.QDialog, Main):
    """
    Implementing dialog window with communication settings.
    Also with modbus client methods
    """
    def __init__(self):
        super(Main, self).__init__()
        QtWidgets.QDialog.__init__(self, None, QtCore.Qt.WindowFlags(QtCore.Qt.WA_DeleteOnClose))
        self.settings = Settings()
        self.settings.setupUi(self)

        self.cb_data = []
        self.client = None

        main.mainWin.slaveIdSb.clear()
        main.mainWin.slaveIdSb.setValue(1)

        main.mainWin.scanRateSb.clear()
        main.mainWin.scanRateSb.setValue(500)

        self.settings.buttonBox.accepted.connect(self.acceptData)
        self.settings.buttonBox.rejected.connect(self.rejectData)
        main.mainWin.startBtn.clicked.connect(self.firstStart)

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

        self.acceptData()
        self.update_ports()

    def update_ports(self):
        """
        updating active com ports
        """
        self.settings.com_set.clear()
        self.serialPorts = [(f'{port.portName()} {port.description()}') for port in QSerialPortInfo().availablePorts()]
        self.settings.com_set.addItems(self.serialPorts)
        self.cb_data[0] = self.settings.com_set.currentText()
        self.set_text()

    def set_text(self):
        self.com_port_label_text = (f'{self.cb_data[0]} ({self.cb_data[1]}-{self.cb_data[2]}-{self.cb_data[3]})')
        main.mainWin.com_port_label.setText(self.com_port_label_text)

    def acceptData(self):
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

    def rejectData(self):
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

    def firstStart(self):
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
        self.connect()
        self.isRunning = True
        self.threadd.start()
        main.mainWin.startBtn.setText('Stop')
        main.mainWin.startBtn.clicked.disconnect()
        main.mainWin.startBtn.clicked.connect(lambda x: self.stop(True))
        main.mainWin.statusLabel.setText('Connection...')

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
        main.mainWin.statusLabel.setText('Connected')
        main.mainWin.statusLabel.setStyleSheet('color: rgb(0, 128, 0);')

    def stop(self, flag):
        """
        stop modbus client
        :param flag: setting status text ("Stopped") when stops by the stop button,
        if not the flag the error text will be set
        """
        self.isRunning = False
        try:
            self.client.socket.close()
        except:
            pass
        main.mainWin.statusLabel.setStyleSheet('color: rgb(230, 0, 0);')
        if flag:
            main.mainWin.statusLabel.setText('Stopped')
        main.mainWin.startBtn.setText('Connect')
        main.mainWin.startBtn.clicked.disconnect()
        main.mainWin.startBtn.clicked.connect(self.start)
        self.threadd.terminate()
        self.client = None


class AboutWin(QtWidgets.QDialog):
    def __init__(self):
        QtWidgets.QDialog.__init__(self, None, QtCore.Qt.WindowFlags(QtCore.Qt.WA_DeleteOnClose))
        self.about = About()
        self.about.setupUi(self)
        self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint)
        self.about.ok_btn.clicked.connect(self.close)


class Signals(QtCore.QObject):
    """
    implementing signals
    """
    updateSignal = QtCore.pyqtSignal()
    stopSignal = QtCore.pyqtSignal()
    errorSignal = QtCore.pyqtSignal(bool)


class ReadHoldingRegisters(QtCore.QObject):
    """
    Modbus client for read holding registers. It can read more than 125 registers, 1000 for example :)
    """
    def __init__(self):
        super(ReadHoldingRegisters, self).__init__()
        self.sig = Signals()
        self.sig.updateSignal.connect(lambda: main.model.updateTable())
        self.sig.stopSignal.connect(lambda: dialog.stop)
        self.sig.errorSignal.connect(lambda x: dialog.stop(x))

    def run(self):
        count = 125  # max count modbus registers for read once
        response = None # responsed data from slave device
        slaveID = int(main.mainWin.slaveIdSb.value())  # device' slave id
        regBlocks = main.regs[-1][0] // count  # the count of blocks with 125 registers
        lastBlock = main.regs[-1][0] - regBlocks * count  # the last block with a amount of the residual registers

        try:
            while dialog.isRunning:
                adr = 0  # the starting register
                modbusData = []  # the readed data from device
                result = []  # the parsed data
                for i in range(regBlocks):
                    response = dialog.client.read_holding_registers(address=adr, count=count, unit=slaveID)
                    adr += count
                    modbusData.extend(response.registers)

                response = dialog.client.read_holding_registers(address=adr, count=lastBlock, unit=slaveID)
                modbusData.extend(response.registers)

                # parsing data by data type
                for i in range(len(main.regs)):
                    register, dataType = main.regs[i][0], main.regs[i][1]

                    if dataType == 'MEA':
                        registerData = struct.pack('>HH', modbusData[register], modbusData[register - 1])
                        floatData = [struct.unpack('>f', registerData)]
                        result.append(str(round(floatData[0][0], 3)))

                    elif dataType == 'BIN':
                        registerData = modbusData[register - 1]
                        ones = []
                        for x in range(16):
                            if (registerData >> x) & 0x1:
                                ones.append(' ' + (str(x)))
                        ones.reverse()
                        registerData = f"{modbusData[register - 1]:>016b}"
                        for j in range(len(ones)):
                            registerData += f' {ones[j]}'
                        if registerData[15] == '1':
                            registerData = f'ДА   {registerData}'
                        elif registerData[15] == '0':
                            registerData = f'НЕТ  {registerData}'
                        if registerData[14] == '1' and registerData[20] == '1':
                            registerData += ' (Тревога)'
                        elif registerData[13] == '1' and registerData[20] == '1':
                            registerData += ' (Предупреждение)'
                        result.append(registerData)

                    elif dataType == 'INT':
                        registerData = modbusData[register - 1]
                        try:
                            if int(registerData) > int(len(main.intDataType[register][-1])):
                                continue
                            else:
                                registerData = str(main.intDataType[register][registerData]).replace(': ', ' - ')
                        except:
                            pass
                        result.append(registerData)
                row = 0
                for value in result:
                    idx = main.model.createIndex(row, 5)
                    main.model.setData(idx, value)
                    row += 1

                self.sig.updateSignal.emit()
                time.sleep((int(main.mainWin.scanRateSb.value())) / 1000)  # scan rate

        except ConnectionException:
            main.mainWin.statusLabel.setText(
            f"{dialog.settings.com_set.currentText().split(' ')[0].strip()}: N/A")
        except AttributeError:
            if type(response) == ModbusIOException:
                main.mainWin.statusLabel.setText("Timeout")
        finally:
            self.sig.errorSignal.emit(False)


def childExec():
    if dialog.client:
        dialog.stop(True)
    else:
        dialog.update_ports()
    dialog.show()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon('icon.ico'))
    sig = Signals()

    main = MainWin()
    dialog = DialogWin()
    about = AboutWin()
    worker = ReadHoldingRegisters()
    main.show()
    btn = main.mainWin.setBtn
    btn.clicked.connect(childExec)
    sys.exit(app.exec_())

