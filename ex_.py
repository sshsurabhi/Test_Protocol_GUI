import sys
import random
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QVBoxLayout, QWidget
from PyQt5.QtGui import QColor

NUMS = [2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47, 53, 59, 61, 67, 71, 73, 79, 83, 89, 97, 101, 103, 107, 109, 113, 127, 131, 137, 139, 149, 151, 157, 163, 167, 173, 179, 181, 191, 193, 197, 199]

class PrimeNumberChecker(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Prime Number Checker")
        self.setGeometry(100, 100, 300, 150)

        layout = QVBoxLayout()

        self.line_edit = QLineEdit()
        layout.addWidget(self.line_edit)

        self.start_button = QPushButton("START")
        self.start_button.clicked.connect(self.start_button_clicked)
        layout.addWidget(self.start_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def start_button_clicked(self):
        random_number = random.randint(1, 1000)
        self.line_edit.setText(str(random_number))

        if random_number in NUMS:
            self.line_edit.setStyleSheet("background-color: lightgreen;")
        else:
            self.line_edit.setStyleSheet("background-color: red;")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PrimeNumberChecker()
    window.show()
    sys.exit(app.exec_())
