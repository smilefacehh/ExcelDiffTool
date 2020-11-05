# -*- coding: utf-8 -*-
# @Date: 2020-11-03
# @Autor: rafaellu

# Form implementation generated from reading ui file 'MainWindow.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QApplication
from PyQt5.QtGui import QColor, QFont, QTextCursor
import sys
import os
from excel_diff import *
import images


class TextEdit(QtWidgets.QTextEdit):
    def __init__(self,*args,**kw):
        super(TextEdit, self).__init__(*args,**kw)
        self.setAcceptDrops(True)

    def canInsertFromMimeData(self,mimeData):
        if mimeData.hasUrls():
            return True

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super(TextEdit, self).dragEnterEvent(event)

    def dragMoveEvent(self, event):
        super(TextEdit, self).dragMoveEvent(event)


class Ui_MainWindow(object):

    def __init__(self):
        self.excel_name_path1 = {}
        self.excel_name_path2 = {}
        self.log_path = 'log.txt'

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1002, 628)
        MainWindow.setMinimumSize(QtCore.QSize(1002, 628))
        # icon = QtGui.QIcon()
        # icon.addPixmap(QtGui.QPixmap("emma1.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(QtGui.QIcon(':/emma2.ico'))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setContentsMargins(20, 20, 20, -1)
        self.horizontalLayout.setSpacing(30)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setMinimumSize(QtCore.QSize(520, 0))
        self.textEdit.setObjectName("textEdit")
        self.textEdit.setAcceptDrops(False)
        self.horizontalLayout.addWidget(self.textEdit)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setContentsMargins(-1, -1, -1, 0)
        self.verticalLayout.setSpacing(20)
        self.verticalLayout.setObjectName("verticalLayout")
        self.textEdit_2 = TextEdit(self.centralwidget)
        self.textEdit_2.setMaximumSize(QtCore.QSize(400, 3720))
        self.textEdit_2.setObjectName("textEdit_2")
        self.textEdit_2.dropEvent = self.textEdit2DropEvent
        self.verticalLayout.addWidget(self.textEdit_2)
        self.textEdit_3 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_3.setMaximumSize(QtCore.QSize(400, 3720))
        self.textEdit_3.setObjectName("textEdit_3")
        self.textEdit_3.dropEvent = self.textEdit3DropEvent
        self.verticalLayout.addWidget(self.textEdit_3)
        self.verticalLayout_2.addLayout(self.verticalLayout)
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setMinimumSize(QtCore.QSize(400, 50))
        self.pushButton.setMaximumSize(QtCore.QSize(400, 50))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.okButtonClick)
        self.verticalLayout_2.addWidget(self.pushButton)
        self.horizontalLayout.addLayout(self.verticalLayout_2)
        self.gridLayout.addLayout(self.horizontalLayout, 0, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Emma表格对比工具"))
        self.textEdit_2.setPlaceholderText(_translate("MainWindow", "拖入文件或目录1 - 旧\n备注：拖入多个表格时，确保文件名称对应"))
        self.textEdit_3.setPlaceholderText(_translate("MainWindow", "拖入文件或目录2 - 新\n备注：拖入多个表格时，确保文件名称对应"))
        self.pushButton.setText(_translate("MainWindow", "对比"))

    def textEdit2DropEvent(self, event):
        self.textEdit_2.setText('')
        self.excel_name_path1 = {}
        name_path = {}
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                filepath = url.toLocalFile()
                if os.path.isfile(filepath) and '.xls' in filepath:
                    name_path[filepath[filepath.rindex('/') + 1:]] = filepath
                elif os.path.exists(filepath):
                    for file in os.listdir(filepath):
                        if os.path.isfile(filepath + '/' + file) and '.xls' in file:
                            name_path[file] = filepath + '/' + file
            
            name_path = sorted(name_path.items())
            for pair in name_path:
                self.excel_name_path1[pair[0]] = pair[1]
                self.textEdit_2.setText(self.textEdit_2.toPlainText() + pair[1] + '\n')
            event.acceptProposedAction()
        

    def textEdit3DropEvent(self, event):
        self.textEdit_3.setText('')
        self.excel_name_path2 = {}
        name_path = {}
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                filepath = url.toLocalFile()
                if os.path.isfile(filepath) and '.xls' in filepath:
                    name_path[filepath[filepath.rindex('/') + 1:]] = filepath
                elif os.path.exists(filepath):
                    for file in os.listdir(filepath):
                        if os.path.isfile(filepath + '/' + file) and '.xls' in file:
                            name_path[file] = filepath + '/' + file
            
            name_path = sorted(name_path.items())
            for pair in name_path:
                self.excel_name_path2[pair[0]] = pair[1]
                self.textEdit_3.setText(self.textEdit_3.toPlainText() + pair[1] + '\n')
            event.acceptProposedAction()

    def checkFileValid(self):
        """ 注意：只有一个文件的时候，允许文件名不一致 """
        if not self.excel_name_path1 or not self.excel_name_path2:
            QMessageBox.critical(self.centralwidget, '错误', '请拖入对比文件或目录', QMessageBox.Ok)
            return False
        elif len(self.excel_name_path1) != len(self.excel_name_path2):
            QMessageBox.critical(self.centralwidget, '错误', '对比文件数量不一致', QMessageBox.Ok)
            return False
        elif len(self.excel_name_path1) > 1 and self.excel_name_path1.keys() != self.excel_name_path2.keys():
            QMessageBox.critical(self.centralwidget, '错误', '对比文件名称不一致，请检查', QMessageBox.Ok)
            return False
        return True
    
    def okButtonClick(self):
        if not self.checkFileValid():
            return
        self.pushButton.setText(QtCore.QCoreApplication.translate("MainWindow", "对比中…"))
        self.textEdit.setText('')
        QApplication.processEvents() # 立即刷新
 
        ed_log = {}
        if len(self.excel_name_path1) == 1:
            ex_path1 = list(self.excel_name_path1.values())[0]
            ex_path2 = list(self.excel_name_path2.values())[0]
            ed = ExcelDiff(ex_path1, ex_path2, self)
            ed.run()
            if ed.log:
                ed_log[ex_path2] = ed.log
        else:
            for k,v in self.excel_name_path1.items():
                ex_path1 = v
                ex_path2 = self.excel_name_path2[k]
                ed = ExcelDiff(ex_path1, ex_path2, self)
                ed.run()
                if ex_path2 not in ed_log:
                    ed_log[ex_path2] = {}
                if ed.log:
                    ed_log[ex_path2] = ed.log

        self.print_write_log(self.log_path, ed_log)
        self.log('<br><b>End!</b>')
        self.pushButton.setText(QtCore.QCoreApplication.translate("MainWindow", "对比"))

    def log(self, info):
        if self.textEdit.toPlainText() == '':
            self.textEdit.setText(info)
        else:
            self.textEdit.setText(self.textEdit.toHtml() + info)
        cursor = self.textEdit.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.textEdit.setTextCursor(cursor)
        QApplication.processEvents()

    def print_write_log(self, path, ed_log):
        """
        @brief 打印、写日志
        """
        with open(path, 'w', encoding='utf-8') as f:
            for k,v in ed_log.items():

                f.write(k + '\n') # 表格路径
                self.log('<b><font color="#DC143C" size="12">' + k + '</font></b>')

                for sheet_name, modify in v.items():
                    f.write(sheet_name + '\n') # 页签名称
                    self.log('<br><br><b>' + sheet_name + '</b>')

                    f.write(','.join(modify['标题']) + '\n') # 标题
                    self.log('<font color="#000000">' + ','.join(modify['标题']) + '</font>')
                    
                    if len(modify['修改']) > 0:
                        f.write('\n修改\n')
                        self.log('<br>修改')

                        for mod in modify['修改']:
                            f.write(mod + '\n')
                            self.log('<font color="#FF8C00">' + mod + '</font>')

                    if len(modify['新增']) > 0:
                        f.write('\n新增\n')
                        self.log('<br>新增')
                        
                        for mod in modify['新增']:
                            f.write(mod + '\n')
                            self.log('<font color="#0000FF">' + mod + '</font>')

                    if len(modify['删除']) > 0:
                        f.write('\n删除\n')
                        self.log('<br>删除')
                        
                        for mod in modify['删除']:
                            f.write(mod + '\n')
                            self.log('<font color="#FF0000">' + mod + '</font>')

                    if len(modify['行变化']) > 0:
                        f.write('\n行变化\n')
                        self.log('<br>行变化')
                        
                        for mod in modify['行变化']:
                            f.write(mod + '\n')
                            self.log('<font color="#FFC0CB">' + mod + '</font>')
                self.log('<br>===============================================================<br>')


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())