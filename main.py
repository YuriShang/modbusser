import sys
import os
import pkgutil
import winreg as reg
import threading

from pymodbus.exceptions import ModbusIOException, ConnectionException
from pymodbus.pdu import ExceptionResponse
from PyQt5 import QtGui, QtCore, QtWidgets
from pymodbus.client.sync import ModbusSerialClient
import serial.tools.list_ports as SerialPortInfo

import csv
import time
import struct

from MainWindow import Main
from DialogWindow import Dialog
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

            if role == QtCore.Qt.CheckStateRole:
                if col == 0:
                    return QtCore.Qt.Checked if self.checkStates[row] == 1 else QtCore.Qt.Unchecked

            if role == QtCore.Qt.BackgroundRole:
                if self.checkStates[row] == 1:
                    color = QtGui.QColor("grey")
                    color.setAlpha(123)
                    return color
                return QtGui.QColor("white")

    def setData(self, index, value, role=QtCore.Qt.EditRole):
        row = index.row()
        col = index.column()

        if role == QtCore.Qt.EditRole:
            if value:
                self._data[row][col] = value
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
        return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsUserCheckable


class MainWin(QtWidgets.QMainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        super(MainWin, self).__init__()

        self.mainWin = Main()
        self.mainWin.setupUi(self)
        self.table = self.mainWin.tableView
        self.table.horizontalHeader().setStretchLastSection(True)

        # Buttons connections
        self.mainWin.openAction.triggered.connect(self.loadCsv)
        self.mainWin.exportAction.triggered.connect(self.writeCsv)
        self.mainWin.exportAction.setDisabled(True)
        self.mainWin.about_action.triggered.connect(lambda x: about.exec_())
        self.mainWin.statusBox.setChecked(True)
        self.mainWin.resetButton.clicked.connect(self.resetRowBrush)
        self.mainWin.resetButton.setDisabled(True)
        self.mainWin.startBtn.setDisabled(True)
        self.tableResizeAccept = False
        self.mainWin.searchBox.textEdited.connect(self.search)
        self.mainWin.searchBox.setDisabled(True)
        self.mainWin.startBtn.clicked.connect(worker.start)
        self.regeditPath = "Software\\QtProject\\OrganizationDefaults\\FileDialog"

    def keyPressEvent(self, event):
        """
        handling pressed keys for "open", "save", "search" methods
        """
        if event.key() == QtCore.Qt.Key_F and event.modifiers() == QtCore.Qt.ControlModifier:
            self.mainWin.searchBox.setFocus(QtCore.Qt.ShortcutFocusReason)
        if event.key() == QtCore.Qt.Key_S and event.modifiers() == QtCore.Qt.ControlModifier and \
            self.mainWin.exportAction.isEnabled():
                self.writeCsv(self.fileName)
        if event.key() == QtCore.Qt.Key_O and event.modifiers() == QtCore.Qt.ControlModifier:
            self.loadCsv()

    def search(self):
        """
        search register names by keyboard input
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

    def loadCsv(self):
        """
        loading Modbus RTU map.txt
        """
        # parsing last path for our file, if not, creating new
        self.key = reg.CreateKey(reg.HKEY_CURRENT_USER, self.regeditPath)
        try:
            self.lastPath = reg.QueryValueEx(self.key, "lastPath")
        except:
            self.lastPath = f"{QtCore.QDir.homePath()}\Modbus RTU Map.txt", '',

        self.fileName, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open Modbus RTU map",
                                                                 f'{self.lastPath[0]}', "*.txt")
        if self.fileName:
            # set current path as last for next time
            reg.SetValueEx(self.key, 'lastPath', 0, reg.REG_SZ, f'{self.fileName}')
            self.lastPath = reg.QueryValueEx(self.key, "lastPath")

            # our table model object
            self.model = TableModel()
            self.loadCsvOnOpen(self.fileName)
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
                self.mainWin.exportAction.setDisabled(False)
                self.rows.clear()
                self.ints.clear()
        reg.CloseKey(self.key)

    def loadCsvOnOpen(self, fileName):
        if worker.client:
            worker.stop(True)
        if fileName:
            print(fileName + " loaded")
            f = open(fileName, 'r')
            with f:
                self.fname = os.path.splitext(str(fileName))[0].split("///")[-1]
                self.setWindowTitle('AR-CON ModBus   ' + self.fname)
                self.regs = []  # modbus registers
                self.ints = []  # registers with "INT"(signed) data
                self.rows = []  # parsed data form modbus map
                self.maxLen = 0
                self.name = ()  # json filename, will be inserted to main win
                self.model.checkStates = [] # reset table checkbox states
                for row in (csv.reader(f, delimiter="|")):  # iterating and parsing rows from loaded map
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
                    row[4] = row[4].strip(' ').replace('\x04', 'Ohm')
                    self.model.checkStates.append(0)
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
        if worker.client:
            worker.stop(True)

        # get checked(selected) rows for save
        checkedRows = []

        for row in range(self.model.searchProxy.rowCount()):
            if self.model.searchProxy.index(row, 0, QtCore.QModelIndex()).data(QtCore.Qt.CheckStateRole):
                checkedRows.append(row)

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

    def closeEvent(self, event):
        """
        safe app closing
        """
        print("Quit...")
        worker.stop(True)


class DialogWin(QtWidgets.QDialog, Main):
    """
    Implementing settings window with communication dialog.
    """

    def __init__(self):
        QtWidgets.QDialog.__init__(self, None, QtCore.Qt.WindowFlags(QtCore.Qt.WA_DeleteOnClose))
        self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint)
        super(Main, self).__init__()
        self.dialog = Dialog()
        self.dialog.setupUi(self)

        self.cbData = []

        main.mainWin.slaveIdSb.clear()
        main.mainWin.slaveIdSb.setValue(1)

        main.mainWin.scanRateSb.clear()
        main.mainWin.scanRateSb.setValue(600)

        self.dialog.buttonBox.accepted.connect(self.acceptData)
        self.dialog.buttonBox.rejected.connect(self.rejectData)

        self.dialog.comSet.setEditable(True)

        speeds = ('1200', '2400', '4800', '9600', '19200', '38400', '57600', '115200')
        self.dialog.baudSet.clear()
        self.dialog.baudSet.addItems(speeds)
        self.dialog.baudSet.setCurrentIndex(5)
        self.dialog.baudSet.setEditable(True)

        parity = ('N', 'E')
        self.dialog.paritySet.clear()
        self.dialog.paritySet.addItems(parity)

        stopBits = ('1', '2')
        self.dialog.sbSet.clear()
        self.dialog.sbSet.addItems(stopBits)

        self.acceptData()
        self.updatePorts()

    def updatePorts(self):
        """
        updating active com ports
        """
        comSetData = [self.dialog.comSet.itemText(i) for i in range(self.dialog.comSet.count())]
        self.serialPorts = [(f'{port} {name.split(" (")[0]}') for port, name, _ in SerialPortInfo.comports()]
        if comSetData != self.serialPorts:
            self.dialog.comSet.clear()
            self.dialog.comSet.addItems(self.serialPorts)
            self.cbData[0] = self.dialog.comSet.currentText()
            self.setText()

    def setText(self):
        self.comLabelText = (f'{self.cbData[0]} ({self.cbData[1]}-{self.cbData[2]}-{self.cbData[3]})')
        main.mainWin.comLabel.setText(self.comLabelText)

    def acceptData(self):
        """
        save dialog
        """
        self.cbData.clear()
        self.cbData.append(self.dialog.comSet.currentText())
        self.cbData.append(self.dialog.baudSet.currentText())
        self.cbData.append(self.dialog.paritySet.currentText())
        self.cbData.append(self.dialog.sbSet.currentText())
        self.setText()
        self.close()

    def rejectData(self):
        """
        reject dialog
        """
        try:
            self.dialog.comSet.setCurrentText(self.cbData[0])
            self.dialog.baudSet.setCurrentText(self.cbData[1])
            self.dialog.paritySet.setCurrentText(self.cbData[2])
            self.dialog.sbSet.setCurrentText(self.cbData[3])
        except:
            pass
        self.close()


class ModbusClient(QtCore.QObject):
    """
    Modbus client for read holding registers. It can read more than 125 registers, 1000 for example :)
    """

    def __init__(self):
        super(ModbusClient, self).__init__()
        self.sig = Signals()
        self.sig.updateSignal.connect(lambda: main.model.updateTable())
        self.sig.stopSignal.connect(lambda x: self.stop(x))
        self.sig.connectionSignal.connect(lambda: self.connected())
        self.sig.comErrSignal.connect(self.comError)
        self.sig.timeoutErrSignal.connect(self.timeoutError)
        self.sig.dataErrSignal.connect(self.illegalDataError)
        self.client = None
        self.isRunning = False
        self.connection = False

    def connected(self):
        self.connection = True
        main.mainWin.statusLabel.setText('Connected')
        main.mainWin.statusLabel.setStyleSheet('color: rgb(0, 128, 0);')

    def start(self):
        """
        connecting to modbus slave device and starting read data from this one
        """
        self.isRunning = True
        self.threadd = threading.Thread(target=self.run)
        self.threadd.start()
        main.mainWin.startBtn.setText('Stop')
        main.mainWin.startBtn.clicked.disconnect()
        main.mainWin.startBtn.clicked.connect(lambda: self.stop(True))
        main.mainWin.statusLabel.setText('Connection...')

    def stop(self, flag):
        """
        stop modbus client
        :param flag: setting status text ("Stopped") when stops by the stop button,
        if not the flag the error text will be set
        """
        self.isRunning = False
        self.connection = False
        if flag:
            main.mainWin.statusLabel.setText('Stopped')
        main.mainWin.statusLabel.setStyleSheet('color: rgb(230, 0, 0);')
        main.mainWin.startBtn.setText('Connect')
        main.mainWin.startBtn.clicked.disconnect()
        main.mainWin.startBtn.clicked.connect(self.start)
        try:
            self.client.socket.close()
        except:
            pass
        self.client = None

    def comError(self):
        main.mainWin.statusLabel.setText(f"{settings.dialog.comSet.currentText().split(' ')[0]}: N/A"),

    def timeoutError(self):
        main.mainWin.statusLabel.setText("Timeout")

    def illegalDataError(self):
        main.mainWin.statusLabel.setText("Illegal data")

    def run(self):
        """
        Main runner function. It does requests to modbus slave device, adaptation and insert data to the Table
        """
        count = 125  # max count modbus registers for read once
        response = None  # responsed data from slave device
        slaveID = int(main.mainWin.slaveIdSb.value())  # device' slave id
        regBlocks = main.regs[-1][0] // count  # the count of blocks with 125 registers
        lastBlock = main.regs[-1][0] - regBlocks * count  # the last block with a amount of the residual registers

        self.client = ModbusSerialClient(
            method='rtu',
            port=settings.dialog.comSet.currentText().split(' ')[0].strip(),
            baudrate=int(settings.dialog.baudSet.currentText()),
            timeout=1,
            parity=settings.dialog.paritySet.currentText(),
            stopbits=int(settings.dialog.sbSet.currentText()),
            bytesize=8)
        self.client.connect()

        try:
            while self.isRunning:
                adr = 0  # start register
                modbusData = []  # readed data from device
                result = []  # parsed data

                for i in range(regBlocks):
                    response = self.client.read_holding_registers(address=adr, count=count, unit=slaveID)
                    adr += count
                    modbusData.extend(response.registers)

                response = self.client.read_holding_registers(address=adr, count=lastBlock, unit=slaveID)
                modbusData.extend(response.registers)

                # parsing data by type
                for i in range(len(main.regs)):
                    register, dataType = main.regs[i][0], main.regs[i][1]
                    if dataType == 'MEA':
                        delta = 0
                        if register == len(modbusData):
                            delta += 1
                            lastBlock += 1
                        registerData = struct.pack('>HH', modbusData[register - delta],
                                                   modbusData[register - delta - 1])
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
                        if registerData[9] == '1' and registerData[15] == '1' and int(registerData, 2) <= 65:
                            registerData += ' (Тревога)'
                        elif registerData[8] == '1' and registerData[15] == '1' and int(registerData, 2) <= 129:
                            registerData += ' (Предупреждение)'
                        for j in range(len(ones)):
                            registerData += f' {ones[j]}'
                        if registerData[15] == '1':
                            registerData = f'ДА   {registerData}'
                        elif registerData[15] == '0':
                            registerData = f'НЕТ  {registerData}'
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

                for row, value in enumerate(result):
                    idx = main.model.createIndex(row, 5)
                    main.model.setData(idx, value)
                if not self.connection:
                    self.sig.connectionSignal.emit()
                self.sig.updateSignal.emit()
                time.sleep((int(main.mainWin.scanRateSb.value())) / 1000)  # scan rate

        except ConnectionException:
            self.sig.comErrSignal.emit()
        except AttributeError:
            if type(response) == ModbusIOException:
                self.sig.timeoutErrSignal.emit()
            elif type(response) == ExceptionResponse:
                self.sig.dataErrSignal.emit()
        else:
            self.sig.stopSignal.emit(True)
        finally:
            self.sig.stopSignal.emit(False)


class AboutWin(QtWidgets.QDialog):
    def __init__(self):
        QtWidgets.QDialog.__init__(self, None, QtCore.Qt.WindowFlags(QtCore.Qt.WA_DeleteOnClose))
        self.about = About()
        self.about.setupUi(self)
        self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint)
        self.about.ok_btn.clicked.connect(self.close)


class Signals(QtCore.QObject):
    """
    implementing signals for thread safety
    """
    updateSignal = QtCore.pyqtSignal() # updating table model
    stopSignal = QtCore.pyqtSignal(bool) # stopping runner
    connectionSignal = QtCore.pyqtSignal() # signal of a valid connection
    comErrSignal = QtCore.pyqtSignal() # signal emit when handled COM port error, for example - port is not aviable
    timeoutErrSignal = QtCore.pyqtSignal() # connection timeout, device is n/a
    dataErrSignal = QtCore.pyqtSignal() # illegal data, for example - count of a requesting registers too much


def settingsExec():
    """
    running settings windows and updating active com ports if it need
    """
    if worker.client:
        worker.stop(True)
    else:
        settings.updatePorts()
    settings.exec_()


def resourcePath(relativePath):
    """
    get absolute path to resource
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        basePath = sys._MEIPASS
    except Exception:
        basePath = os.path.abspath(".")
    return f"{basePath}{relativePath}"


def setIcon(filePath):
    """
    set app icon
    """
    filePath = resourcePath(filePath)
    app.setWindowIcon(QtGui.QIcon(filePath))
    about.about.pict.setPixmap(QtGui.QPixmap(filePath))


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    sig = Signals()
    worker = ModbusClient()
    main = MainWin()
    settings = DialogWin()
    about = AboutWin()
    setIcon("\icon\icon.ico")
    main.show()
    settingBtn = main.mainWin.setBtn
    settingBtn.clicked.connect(settingsExec)
    sys.exit(app.exec_())

