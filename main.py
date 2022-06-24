import sys

from PyQt5 import QtWidgets, QtCore
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtGui import QTextCursor
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from uidesi import Ui_mainWindow
from checkEquipmentCode import checkCode

class MainWindow(QtWidgets.QMainWindow, Ui_mainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent=parent)
        self.setupUi(self)
        self.lineEdit.setFocusPolicy(QtCore.Qt.NoFocus)  # 禁止手动修改
        self.textEdit.setFocusPolicy(QtCore.Qt.NoFocus)  # 禁止手动修改
        self.thread = None

    @pyqtSlot(bool)
    def on_pushButton_2_clicked(self, checked):
        self.textEdit.moveCursor(QTextCursor.End)
        # self.textEdit.append('点击按钮了')
        if len(self.lineEdit.text()) == 0:
            self.show_message()
        else:
            self.mainLogic()

    def mainLogic(self):
        self.thread = checkCode(self.lineEdit.text())
        self.thread.signal_toTextEdit.connect(self.call_back_toTextEdit)  # 进程连接回传到GUI的事件
        self.thread.signal_toProgressBar.connect(self.call_back_toProgressBar)
        self.thread.start()  # 开始线程

    def call_back_toTextEdit(self, msg):
        self.textEdit.append(str(msg))  # 将线程的参数传入textedit

    def call_back_toProgressBar(self, msg):
        self.progressBar.setValue(int(msg))

    @pyqtSlot(bool)
    def on_pushButton_clicked(self):
        fname = QFileDialog.getOpenFileName(self, '选择文件', './', 'Excel工作表(*.xlsx)')
        self.lineEdit.setText(fname[0])

    def show_message(self):
        QMessageBox.critical(self, "错误", "请选择需要检测的excel文件")

    def textEdit_append(self, newText):
        self.textEdit.append(newText)

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())