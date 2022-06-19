# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main_win.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class MainWin(object):
    def setupUi(self, MODBUSSER):
        MODBUSSER.setObjectName("AR-CON ModBus")
        MODBUSSER.resize(787, 484)
        QtCore.QMetaObject.connectSlotsByName(MODBUSSER)
        self.centralwidget = QtWidgets.QWidget(MODBUSSER)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.start_btn = QtWidgets.QPushButton(self.centralwidget)
        self.start_btn.setMaximumSize(QtCore.QSize(150, 16777215))
        self.start_btn.setObjectName("start_btn")
        self.gridLayout_2.addWidget(self.start_btn, 1, 16, 1, 1)
        self.set_btn = QtWidgets.QPushButton(self.centralwidget)
        #self.set_btn.setMaximumSize(QtCore.QSize(100, 16777215))
        self.set_btn.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.set_btn.setCheckable(False)
        self.set_btn.setObjectName("set_btn")
        self.gridLayout_2.addWidget(self.set_btn, 1, 1, 1, 1)
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.gridLayout_2.addWidget(self.line, 2, 0, 1, 17)
        self.label = QtWidgets.QLabel(self.centralwidget)
        #self.label.setMaximumSize(QtCore.QSize(100, 16777215))
        self.label.setObjectName("label")
        self.label.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.gridLayout_2.addWidget(self.label, 1, 14, 1, 1)
        self.resetButton = QtWidgets.QPushButton(self.centralwidget)
        #self.resetButton.setMaximumSize(QtCore.QSize(87, 16777215))
        self.resetButton.setObjectName("resetButton")
        self.gridLayout_2.addWidget(self.resetButton, 1, 0, 1, 1)
        self.slave_id_sb = QtWidgets.QSpinBox(self.centralwidget)
        self.slave_id_sb.setMaximumSize(QtCore.QSize(60, 16777215))
        self.slave_id_sb.setObjectName("slave_id_sb")
        self.slave_id_sb.setAlignment(QtCore.Qt.AlignRight)
        self.gridLayout_2.addWidget(self.slave_id_sb, 1, 15, 1, 1)
        self.name_label = QtWidgets.QLabel(self.centralwidget)
        self.name_label.setText("JSON file name:")
        self.name_label.setObjectName("name_label")
        self.gridLayout_2.addWidget(self.name_label, 9, 0, 1, 3)
        self.com_port_label = QtWidgets.QLabel(self.centralwidget)
        self.com_port_label.setText("")
        self.com_port_label.setObjectName("com_port_label")
        self.gridLayout_2.addWidget(self.com_port_label, 10, 0, 1, 3, QtCore.Qt.AlignBottom)
        self.tableView = QtWidgets.QTableView(self.centralwidget)
        self.tableView.setObjectName("tableView")
        self.gridLayout_2.addWidget(self.tableView, 0, 0, 1, 17)
        self.status_box = QtWidgets.QCheckBox(self.centralwidget)
        self.status_box.setObjectName("status_box")
        self.gridLayout_2.addWidget(self.status_box, 1, 2, 1, 1)
        self.status_label = QtWidgets.QLabel(self.centralwidget)
        self.status_label.setAlignment(QtCore.Qt.AlignHCenter)
        self.status_label.setMaximumSize(QtCore.QSize(150, 16777215))
        self.status_label.setStyleSheet('color: rgb(255, 0, 0);')
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(False)
        font.setWeight(35)
        font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        self.status_label.setFont(font)
        self.status_label.setObjectName("status_label")
        self.gridLayout_2.addWidget(self.status_label, 9, 16, 1, 1)
        self.scan_rate_sb = QtWidgets.QSpinBox(self.centralwidget)
        self.scan_rate_sb.setObjectName("scan_rate_sb")
        self.scan_rate_sb.setRange(100, 5000)
        self.scan_rate_sb.setMaximumSize(QtCore.QSize(60, 16777215))
        self.scan_rate_sb.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.gridLayout_2.addWidget(self.scan_rate_sb, 1, 13, 1, 1)
        self.scan_rate_label = QtWidgets.QLabel(self.centralwidget)
        self.scan_rate_label.setObjectName("scan_ra te_label")
        self.scan_rate_label.setMaximumSize(QtCore.QSize(90, 16777215))
        self.scan_rate_label.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.gridLayout_2.addWidget(self.scan_rate_label, 1, 12, 1, 1)

        self.searchLabel = QtWidgets.QLabel(self.centralwidget)
        self.searchLabel.setObjectName("searchLabel")
        self.searchLabel.setMaximumSize(QtCore.QSize(90, 16777215))
        self.searchLabel.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.gridLayout_2.addWidget(self.searchLabel, 1, 10, 1, 1)

        self.searchBox = QtWidgets.QLineEdit(self.centralwidget)
        self.searchBox.setObjectName("searchBox")
        self.searchBox.setMaximumSize(QtCore.QSize(90, 16777215))
        self.gridLayout_2.addWidget(self.searchBox, 1, 11, 1, 1)

        self.con_indicator = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(8)
        MODBUSSER.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MODBUSSER)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 787, 21))
        self.menubar.setObjectName("menubar")
        self.file_menu = QtWidgets.QMenu(self.menubar)
        self.file_menu.setObjectName("file_menu")
        self.about_menu = QtWidgets.QMenu(self.menubar)
        self.about_menu.setObjectName("about_menu")
        MODBUSSER.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MODBUSSER)
        self.statusbar.setObjectName("statusbar")
        MODBUSSER.setStatusBar(self.statusbar)
        self.open_action = QtWidgets.QAction(MODBUSSER)
        self.open_action.setObjectName("open_action")
        self.export_action = QtWidgets.QAction(MODBUSSER)
        self.export_action.setObjectName("export_action")
        self.about_action = QtWidgets.QAction(MODBUSSER)
        self.about_action.setObjectName("about_action")
        self.file_menu.addAction(self.open_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.export_action)
        self.about_menu.addAction(self.about_action)
        self.menubar.addAction(self.file_menu.menuAction())
        self.menubar.addAction(self.about_menu.menuAction())

        self.retranslateUi(MODBUSSER)
        QtCore.QMetaObject.connectSlotsByName(MODBUSSER)

    def retranslateUi(self, MODBUSSER):
        _translate = QtCore.QCoreApplication.translate
        MODBUSSER.setWindowTitle(_translate("AR-CON ModBus", "AR-CON ModBus"))
        self.start_btn.setText(_translate("AR-CON ModBus", "Connect"))
        self.set_btn.setText(_translate("AR-CON ModBus", "Settings"))
        self.label.setText(_translate("AR-CON ModBus", "      Slave ID"))
        self.resetButton.setText(_translate("AR-CON ModBus", "Reset"))
        self.status_box.setText(_translate("AR-CON ModBus", "Status"))
        self.status_label.setText(_translate("AR-CON ModBus", "No connection"))
        self.scan_rate_label.setText(_translate("AR-CON ModBus", "ScanRate"))
        self.file_menu.setTitle(_translate("AR-CON ModBus", "File"))
        self.about_menu.setTitle(_translate("AR-CON ModBus", "Help"))
        self.open_action.setText(_translate("AR-CON ModBus", "Open"))
        self.export_action.setText(_translate("AR-CON ModBus", "Save"))
        self.about_action.setText(_translate("AR-CON ModBus", "About"))
        self.searchLabel.setText(_translate("AR-CON ModBus", "Search"))

