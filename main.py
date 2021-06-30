from PyQt5.QtCore import QMetaObject, QSize, QThread, pyqtSignal
import sys
import os

from PyQt5 import QtWidgets, QtGui
from PyQt5 import QtCore
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QFileDialog, QLabel

import read_excel
from dohody import Ui_MainWindow
from read_excel import Read


class MyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowIcon(QtGui.QIcon('img/roskazna.png'))
        self.ui.open_file_one.clicked.connect(self.open_file)
        self.ui.open_file_two.clicked.connect(self.open_file)
        self.filename = ''
        self.filename_one = ''
        self.check_one = False
        self.filename_two = ''
        self.check_two = False

    def open_file(self):
        self.filename = QFileDialog.getOpenFileName(
            None, 'Открыть', os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'), 'All Files(*.xlsx *.xls)')
        sender = self.sender()
        if str(self.filename) in "('', '')":
            self.ui.statusbar.showMessage('Файл не выбран')
        else:
            if sender.text() == 'Загрузить таблицу 1':
                self.ui.status_one.setPixmap(QPixmap("img/good.png"))
                self.filename_one = self.filename
                # read = Read()
                # read.read_excel(self.filename_one[0])
                self.check_one = True
                self.check_two = False
                self.new_thread()
            else:
                self.ui.status_two.setPixmap(QPixmap("img/good.png"))
                self.filename_two = self.filename
                # read = Read()
                # read.read_excel(self.filename_two[0])
                self.check_one = False
                self.check_two = True
                self.new_thread()
        #     self.new_thread()

    def new_thread(self):
        self.my_thread = Read(my_window=self)
        self.my_thread.start()


app = QtWidgets.QApplication([])
application = MyWindow()
application.setWindowTitle("Конвертер excel Доходы")
application.show()

sys.exit(app.exec())