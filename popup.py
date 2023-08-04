from PyQt5 import uic
from PyQt5.QtWidgets import *


class MyDialog(QFrame):
    # button_clicked = pyqtSignal(str)
    # def __init__(self, rm, multimeter, parent=None):
    def __init__(self, parent=None):
        super(MyDialog, self).__init__(parent)
        uic.loadUi("UI/message.ui", self)
        # self.setWindowIcon(QIcon('2.jpg'))
        self.setFixedSize(self.size())
        self.show()



if __name__ == "__main__":
    app = QApplication([])
    window = MyDialog()
    window.show()
    app.exec_()