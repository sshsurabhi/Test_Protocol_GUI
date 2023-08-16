import sys
from PyQt5 import QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLineEdit, QLabel
from configparser import ConfigParser

class PowerSupplyApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.config = ConfigParser()
        self.config.read('config.ini')  # Update with your actual config file path

        layout = QVBoxLayout()
        self.line_edit = QLineEdit()
        self.button = QPushButton('B')
        self.label = QLabel('Sel')


        layout.addWidget(self.button)
        layout.addWidget(self.label)
        layout.addWidget(self.line_edit)

        self.button.clicked.connect(self.update_values)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)


        self.channel_sett = self.config.get('Power Supplies', 'Channel_set')
        self.volatge_sett = self.config.get('Power Supplies', 'Voltage_set')
        self.current_sett = self.config.get('Power Supplies', 'Current_set')

        print(self.channel_sett, self.volatge_sett, self.current_sett)

        self.line_edit.returnPressed.connect(self.line_insert)

    def update_values(self):

        if self.button.text()=='B':
            self.label.setText('C')
            self.line_edit.setText(self.channel_sett)

        elif self.button.text()=='P':
            self.label.setText('V')
            self.line_edit.setText(self.volatge_sett)

        elif self.button.text()=='S':
            self.label.setText('I')
            self.line_edit.setText(self.current_sett)
        
        elif self.button.text()=='Save':
            with open('config.ini', 'w') as config_file:
                self.config.write(config_file)


    def line_insert(self):
        if self.label.text()=='C':

            self.channel_sett = str(self.line_edit.text())
            self.config.set('Power Supplies', 'Channel_set', self.channel_sett)
            self.button.setText('P')

        elif self.label.text()=='V':

            self.volatge_sett = str(self.line_edit.text())
            self.config.set('Power Supplies', 'Voltage_set', self.volatge_sett)
            self.button.setText('S')

        elif self.label.text()=='I':

            self.current_sett= str(self.line_edit.text())
            self.config.set('Power Supplies', 'Current_set', self.current_sett)
            self.button.setText('Save')

        else:
            print('false')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PowerSupplyApp()
    window.show()
    sys.exit(app.exec_())
