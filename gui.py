# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'gui.ui'
#
# Created by: PyQt5 UI code generator 5.15.0
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(569, 307)
        MainWindow.setStyleSheet("background-color:rgba(0,0,0,0.5); font-weight:bold;")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(10, 40, 391, 26))
        self.lineEdit.setStyleSheet("background-color:rgb(255, 255, 255)")
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(410, 40, 151, 26))
        self.pushButton.setStyleSheet("background-color:rgb(255, 255, 255)")
        self.pushButton.setObjectName("pushButton")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(10, 80, 391, 26))
        self.lineEdit_2.setStyleSheet("background-color:rgb(255, 255, 255)")
        self.lineEdit_2.setText("")
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(410, 80, 151, 26))
        self.pushButton_2.setStyleSheet("background-color:rgb(255, 255, 255)")
        self.pushButton_2.setObjectName("pushButton_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_3.setGeometry(QtCore.QRect(10, 120, 391, 26))
        self.lineEdit_3.setStyleSheet("background-color:rgb(255, 255, 255)")
        self.lineEdit_3.setText("")
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(410, 117, 141, 31))
        self.label.setStyleSheet("color:rgb(255, 255, 255);\n"
"background-color:none;\n"
"font-weight:bold;")
        self.label.setObjectName("label")
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setEnabled(True)
        self.checkBox.setGeometry(QtCore.QRect(170, 160, 121, 31))
        self.checkBox.setTabletTracking(False)
        self.checkBox.setAutoFillBackground(False)
        self.checkBox.setStyleSheet("color:rgb(230, 75, 238);\n"
"background-color:none;\n"
"font-weiht:bold;\n"
"font-size:25px;")
        self.checkBox.setChecked(True)
        self.checkBox.setObjectName("checkBox")
        self.checkBox_2 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_2.setGeometry(QtCore.QRect(300, 160, 131, 31))
        self.checkBox_2.setStyleSheet("color:rgb(230, 75, 238);\n"
"background-color:none;\n"
"font-weiht:bold;\n"
"font-size:25px;")
        self.checkBox_2.setObjectName("checkBox_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(270, 210, 121, 31))
        self.pushButton_3.setStyleSheet("background-color: rgb(138, 226, 52);\n"
"color:black;\n"
"font-weight:bold;")
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(400, 210, 131, 31))
        self.pushButton_4.setStyleSheet("background-color: rgb(239, 41, 41);\n"
"color:black;\n"
"font-weight:bold;")
        self.pushButton_4.setObjectName("pushButton_4")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(70, 260, 461, 20))
        self.label_2.setStyleSheet("color:rgb(252, 175, 62);\n"
"font: italic 11pt \"DejaVu Serif\";\n"
"")
        self.label_2.setObjectName("label_2")
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(40, 220, 201, 23))
        self.progressBar.setStyleSheet("background-color:rgb(243, 243, 243)")
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "List Files to Excel"))
        self.pushButton.setText(_translate("MainWindow", "Select Files Path"))
        self.pushButton_2.setText(_translate("MainWindow", "Select Excel Path"))
        self.label.setText(_translate("MainWindow", "Excel File Name"))
        self.checkBox.setText(_translate("MainWindow", "Images"))
        self.checkBox_2.setText(_translate("MainWindow", "Audios"))
        self.pushButton_3.setText(_translate("MainWindow", "Convert"))
        self.pushButton_4.setText(_translate("MainWindow", "Cancel"))
        self.label_2.setText(_translate("MainWindow", "Developed by: Venkata Ramana P <github.com/itsmepvr>"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
