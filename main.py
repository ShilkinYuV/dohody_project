from PyQt5.QtCore import QMetaObject, QSize, QThread, pyqtSignal
import sys
import os

from PyQt5 import QtWidgets, QtGui
from PyQt5 import QtCore
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QFileDialog, QLabel

from dohody import Ui_MainWindow


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
        self.filename_two = ''

    def open_file(self):
        self.filename = QFileDialog.getOpenFileName(
            None, 'Открыть', os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'), 'All Files(*.xlsx)')
        sender = self.sender()
        if str(self.filename) in "('', '')":
            self.ui.statusbar.showMessage('Файл не выбран')
        else:
            if sender.text() == 'Загрузить таблицу 1':
                self.ui.status_one.setPixmap(QPixmap("img/good.png"))
                self.filename_one = self.filename
                print(self.filename_one[0])
            else:
                self.ui.status_two.setPixmap(QPixmap("img/good.png"))
                self.filename_two = self.filename
                print(self.filename_two[0])
        #     self.new_thread()


app = QtWidgets.QApplication([])
application = MyWindow()
application.setWindowTitle("Конвертер excel Доходы")
application.show()

sys.exit(app.exec())