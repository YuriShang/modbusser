# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'second.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class SettingsWin(object):
    def setupUi(self, set_win):
        set_win.setObjectName("set_win")
        set_win.resize(400, 300)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(set_win.sizePolicy().hasHeightForWidth())
        set_win.setSizePolicy(sizePolicy)
        set_win.setMinimumSize(QtCore.QSize(400, 300))
        set_win.setMaximumSize(QtCore.QSize(400, 300))
        self.gridLayout = QtWidgets.QGridLayout(set_win)
        self.gridLayout.setObjectName("gridLayout")
        self.parity_label = QtWidgets.QLabel(set_win)
        self.parity_label.setObjectName("parity_label")
        self.gridLayout.addWidget(self.parity_label, 2, 0, 1, 2)
        self.sb_set = QtWidgets.QComboBox(set_win)
        self.sb_set.setObjectName("sb_set")
        self.gridLayout.addWidget(self.sb_set, 3, 2, 1, 3)
        self.sb_label = QtWidgets.QLabel(set_win)
        self.sb_label.setObjectName("sb_label")
        self.gridLayout.addWidget(self.sb_label, 3, 0, 1, 2)
        self.com_label = QtWidgets.QLabel(set_win)
        self.com_label.setObjectName("com_label")
        self.gridLayout.addWidget(self.com_label, 0, 0, 1, 1)
        self.baud_set = QtWidgets.QComboBox(set_win)
        self.baud_set.setObjectName("baud_set")
        self.gridLayout.addWidget(self.baud_set, 1, 2, 1, 3)
        self.baud_label = QtWidgets.QLabel(set_win)
        self.baud_label.setObjectName("baud_label")
        self.gridLayout.addWidget(self.baud_label, 1, 0, 1, 2)
        self.com_set = QtWidgets.QComboBox(set_win)
        self.com_set.setObjectName("com_set")
        self.gridLayout.addWidget(self.com_set, 0, 2, 1, 3)
        self.parity_set = QtWidgets.QComboBox(set_win)
        self.parity_set.setObjectName("parity_set")
        self.gridLayout.addWidget(self.parity_set, 2, 2, 1, 3)
        self.buttonBox = QtWidgets.QDialogButtonBox(set_win)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.gridLayout.addWidget(self.buttonBox, 9, 1, 1, 3)

        self.retranslateUi(set_win)
        QtCore.QMetaObject.connectSlotsByName(set_win)


    def retranslateUi(self, set_win):
        _translate = QtCore.QCoreApplication.translate
        set_win.setWindowTitle(_translate("set_win", "Настройки COM"))
        self.parity_label.setText(_translate("set_win", "Четность:"))
        self.sb_label.setText(_translate("set_win", "Стопбиты:"))
        self.com_label.setText(_translate("set_win", "COM:"))
        self.baud_label.setText(_translate("set_win", "Скорость:"))

    