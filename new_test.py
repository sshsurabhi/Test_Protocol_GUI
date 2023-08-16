import datetime, os, sys, time, openpyxl, configparser, serial
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
##################################################################################################################################
class App(QMainWindow):
    def __init__(self):
        super(App, self).__init__()
        uic.loadUi("UI/ascha.ui", self)
        self.setWindowIcon(QIcon('images_/icons/2.jpg'))
        self.setFixedSize(self.size())
        self.setStatusTip('Moewe Optik Gmbh. 2023')
        self.show()
        ########################################################################################################
        self.serial_port = None
        self.thread = None

        # self.serial_thread = SerialPortThread()
        # self.serial_thread.com_ports_available.connect(self.update_com_ports)
        # self.serial_thread.start()

        self.baudrate_box.addItems(['9600','57600','115200'])
        self.baudrate_box.setCurrentText('115200')
        # self.connect_button.clicked.connect(self.connect_or_disconnect_serial_port)
        # self.refresh_button.clicked.connect(self.refresh_connect)
        ########################################################################################################
        self.commands = ['i2c:scan', 'i2c:read:53:04:FC', 'i2c:write:53:', 'i2c:read:53:20:00', 'i2c:write:73:04', 'i2c:scan','i2c:write:21:0300','i2c:write:21:0100','i2c:write:21:01FF', 'i2c:write:73:01',
                    'i2c:scan', 'i2c:write:4F:06990918', 'i2c:write:4F:01F8', 'i2c:read:4F:1E:00']
        self.start_button.clicked.connect(self.connect)

        ########################################################################################################
        self.rm = visa.ResourceManager()
        self.multimeter = None
        self.powersupply = None

        self.value_edit.returnPressed.connect(self.load_voltage_current)
    def on_button_click(self, file_path):
        if file_path:
            pixmap = QPixmap(file_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.setScaledContents(True)
            self.image_label.setFixedSize(pixmap.width(), pixmap.height())

            if self.start_button.text() == 'STEP4':
                reply = self.show_good_message('Check All your environment Correct. We cannot change later')                
                if reply == QMessageBox.Yes:                    
                    self.start_button.setText('NEXT')
                    self.on_button_click('images_/images/On_Devices.jpg')
                    self.info_label.setText('Power ON of the Powersupply and also Multimeter...\n and wait for 12 seconds.\n Press "Netzteil ON"')                    
                else:
                    self.on_button_click('images_/icons/next_1.jpg')
                    
            # if self.start_button.text() == 'JUMPER OK':
            #     reply = self.show_good_message('')
            #     if reply == QMessageBox.Yes:
            #         self.start_button.setText('STROM')
            #         self.start_button.setEnabled(True)
            #         self.powersupply.write('OUTPut '+self.PS_channel+',ON')
            #         self.on_button_click('images_/images/PP16.jpg')
            #         self.info_label.setText('Press STROM button')
            #     else:
            #         self.on_button_click('images_/images/PP17.jpg')

    def connect(self):
        if self.start_button.text() == 'START':
            self.on_button_click('images_/images/board_on_mat_.jpg')
            self.info_label.setText("Place the board on ESD Matte See the Image on the right.\n\nCheck all your environment with the image\n\nCheck All connections.Please close if anything not correct.\n\nIf correct, \n\n''Press STEP1 Button.'' ")
            self.start_button.setText("STEP1")
        elif self.start_button.text() == 'STEP1':
            self.on_button_click('images_/images/PP4_.jpg')
            self.info_label.setText('Check all the "4" screws\n\n of the board, (See the image). Fit all the 4 screws (4x M2,5x5 Torx)\n\ Press "STEP2".')
            self.start_button.setText("STEP2")
        elif self.start_button.text() == 'STEP2':
            self.on_button_click('images_/images/board_on_mat.jpg')
            self.info_label.setText('AFter fitting the 4 Screws, \n\n Place back the Board on ESD Matte. \\n then, Press "STEP3"')
            self.start_button.setText("STEP3")
        elif self.start_button.text()=='STEP3':
            self.on_button_click('images_/images/board_with_cabels.jpg')
            self.info_label.setText('Connect the Power Cables to the board (see image).\n\n\n Press STEP4.')
            self.start_button.setText("STEP4")
        elif self.start_button.text()=='STEP4':
            self.on_button_click('images_/icons/next.jpg')
        elif self.start_button.text()=='NEXT':
            self.start_button.setEnabled(False)
            self.show_good_message('Wait for 5 seconds. Untill the Powersupply and Multimeter get SET')
            self.start_button.setText('MULTI ON')
            self.info_label.setText("Press MULTI ON.\n\n You can see MULTIMETER Name on TextBox.")
            self.on_button_click('images_/images/PP9.jpg')
        elif self.start_button.text()=='MULTI ON':
            self.connect_multimeter()
        elif self.start_button.text()=='POWER ON':
            self.connect_powersupply()
        elif self.start_button.text()=='STROM-I':
            self.calc_voltage_before_jumper()
        elif self.start_button.text()=='SPANNUNG':
            # QMessageBox.information(self, 'Information', 'Supplying Current is either more or less. So please Swith OFF the PowerSupply, and Put back all the Euipment back.')
            # self.powersupply.write('OUTPut '+self.PS_channel+',OFF')
            self.info_label.setText('Press "POWER OFF" button')
            self.on_button_click('images_/images/Start2.png')
            self.start_button.setText('POWER OFF')
            voltage_before_jumper = self.multimeter.query('MEAS:VOLT:DC?')
            self.voltage_before_jumper = voltage_before_jumper
            self.result_edit.setText(str(voltage_before_jumper))
        elif self.start_button.text()=='POWER OFF':
            self.powersupply.write('OUTPut '+self.PS_channel+',OFF')
            self.info_label.setText('wait 10 seconds')
            self.start_button.setText('Close J')
            self.on_button_click('images_/images/close_jumper.jpg')
        elif self.start_button.text()=='Close J':
            self.start_button.setEnabled(False)
            self.show_good_message('CLOSE the Jumper with Soldering. \n If You Close then Press YES')
            self.start_button.setText('STROM')

        # elif self.

            

        

            
        ########################################################################################################
    def load_voltage_current(self):
        if (self.vals_button.text() == 'CH' and self.value_edit.text() in ['ch1', 'CH1', 'Ch1', 'cH1']):
            self.powersupply.write('INSTrument CH1')
            self.PS_channel = self.value_edit.text()
            self.vals_button.setText('V')
            self.info_label.setText('Write 30 in the Yellow Box next to "V" \n \n Press "Enter"\n You can check the value in the powersupply.')
            self.value_edit.clear()
            self.on_button_click('images_/images/PP7_1.jpg')
            self.value_edit.setValidator(QRegExpValidator(QRegExp(r'^\d+(\.\d+)?$')))
        elif (self.vals_button.text() == 'CH' and self.value_edit.text() in ['ch2', 'CH2', 'Ch2', 'cH2']):
            self.powersupply.write('INSTrument CH2')
            self.PS_channel = self.value_edit.text()
            self.vals_button.setText('V')
            self.info_label.setText('Write 30 in the Yellow Box next to "V" \n \n Press "Enter"\n You can check the value in the powersupply.')
            self.value_edit.clear()
            self.on_button_click('images_/images/PP7_1.jpg')
            self.value_edit.setValidator(QRegExpValidator(QRegExp(r'^\d+(\.\d+)?$')))
        elif (self.vals_button.text() == 'CH' and self.value_edit.text() in ['ch3', 'CH3', 'Ch3', 'cH3']):
            self.powersupply.write('INSTrument CH3')
            self.PS_channel = self.value_edit.text()
            self.vals_button.setText('V')
            self.info_label.setText('Write 30 in the Yellow Box next to "V" \n \n Press "Enter"\n You can check the value in the powersupply.\n You can see the "Channel Selection" in the powersupply.')
            self.value_edit.clear()
            self.on_button_click('images_/images/PP7_1.jpg')
            self.value_edit.setValidator(QRegExpValidator(QRegExp(r'^\d+(\.\d+)?$')))
        elif self.vals_button.text() == 'V':
            self.max_voltage =  self.value_edit.text()
            self.powersupply.write(self.PS_channel+':VOLTage ' + self.max_voltage)
            # self.textBrowser.append(self.powersupply.query(self.PS_channel+':VOLTage?'))
            max_voltage = self.max_voltage
            self.vals_button.setText('I')
            self.info_label.setText('Enter 0.5 in the box next to I\n\n Press "Enter".\n Check the value change in the Powersupply.')
            self.value_edit.clear()
            self.on_button_click('images_/images/PP7_2.jpg')
        elif self.vals_button.text() == 'I':
            self.max_current = self.value_edit.text()
            self.powersupply.write('CH1:CURRent ' + self.max_current)
            # self.textBrowser.append(self.powersupply.query(self.PS_channel+':CURRent?'))
            self.info_label.setText('Enter 0.05 in the box next to I')
            self.powersupply.write('OUTPut '+self.PS_channel+',ON')
            self.value_edit.setEnabled(False)
            self.start_button.setEnabled(True)
            self.info_label.setText('Press STROM-I') # modify here'
            self.start_button.setText('STROM-I')
            self.value_edit.setStyleSheet("")
            self.value_edit.clear()
            self.on_button_click('images_/images/PP8.jpg')
        # elif self.vals_button.text() == 'Tolz V':
        #     self.volt_toleranz = self.value_edit.text()
        #     self.textBrowser.append(self.volt_toleranz)
        #     self.value_edit.clear()
        #     self.info_label.setText('Enter 0.5 in the box next to I')
        #     self.vals_button.setText('Tolz I')
        #     self.on_button_click('images_/images/PP8_1.jpg')
        # elif self.vals_button.text() == 'Tolz I':
        #     self.curr_toleranz = self.value_edit.text()
        #     self.textBrowser.append(self.curr_toleranz)
        #     self.value_edit.setEnabled(False)
        #     self.start_button.setEnabled(True)
        #     self.info_label.setText('Press MESS V_I')
        #     self.start_button.setText('MESS V_I')
        #     self.on_button_click('images_/images/PP17.jpg')
        else:
            self.textBrowser.append('Wrong Input')

    def calc_voltage_before_jumper(self):
        current = float(self.powersupply.query('MEASure:CURRent? '+self.PS_channel))        
        self.result_edit.setText('Current: '+str(current))
        self.textBrowser.append('Current: '+str(current))
        if self.start_button.text() == 'STROM-I':
            if 0.04 <= current <= 0.06:
                self.start_button.setText('SPANNUNG')
                self.start_button.setEnabled(False)
                self.current_before_jumper = current
                self.on_button_click('images_/images/R709_before_jumper.jpg')
                self.info_label.setText('Press SPANNUNG to Calculate initial VOLTAGE at R709.\n \n Calculate Voltage at the Component \n Shown in the figure.')
            else:
                self.current_before_jumper = current
                self.start_button.setText('SPANNUNG')
                self.info_label.setText('Press SPANNUNG to Calculate initial VOLTAGE at R709.\n \n Calculate Voltage at the Component \n Shown in the figure.')
                self.on_button_click('images_/images/R709_before_jumper.jpg')
        elif self.start_button.text() == 'STROM':
            if 0.09 <= current <= 0.15:
                self.start_button.setEnabled(False)
                self.AC_DC_box.setEnabled(True)
                self.current_after_jumper = current
                self.on_button_click('images_/images/sel_DC_in_multimeter.jpg')
                QMessageBox.information(self, "Information", "Select DCV in the box near AC/DC")
                self.info_label.setText('\n \n \n \n Select DCV from AC/DC..!')
            else:
                QMessageBox.information(self, 'Information', 'Supplying Current is either more or less. So please Swith OFF the PowerSupply, and Put back all the Euipment back.')
                self.powersupply.write('OUTPut '+self.PS_channel+',OFF')

    def connect_multimeter(self):
        if not self.multimeter:
            try:
                self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
                self.textBrowser.append(self.multimeter.query('*IDN?'))
                self.on_button_click('images_/images/Power_ON_PS.jpg')
                self.start_button.setText('POWER ON')
                self.info_label.setText('Press POWER ON button.\n \n It connects the powersupply...!' )
            except visa.errors.VisaIOError:
                self.textBrowser.append('Multimeter has not been presented')
        else:
            self.multimeter.close()
            self.multimeter = None
            self.textBrowser.append(self.multimeter.query('*IDN?'))

    def connect_powersupply(self):
        if not self.powersupply:
            try:
                self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
                self.textBrowser.setText(self.powersupply.resource_name)
                self.start_button.setEnabled(False)
                self.value_edit.setStyleSheet("background-color: lightyellow;")
                self.info_label.setText('Write CH1 in the Yellow Box (Highlighted)\n \n next to CH \n\n Press "ENTER"')
                self.value_edit.setEnabled(True)
                self.vals_button.setText('CH')
                self.on_button_click('images_/images/PP7.jpg')
            except visa.errors.VisaIOError:
                QMessageBox.information(self, "PowerSupply Connection", "PowerSupply is not present at the given IP Address.")
                self.textBrowser.setText('Powersupply has not been presented.')
        else:
            self.powersupply.close()
            self.powersupply = None
            self.PS_button.setText('PS ON')
            self.textBrowser.setText('Netzteil Disconnected')

    def show_good_message(self, message):
        self.timer1 = QTimer()
        self.timer1.timeout.connect(self.enable_button)
        self.timer1.start(10000)
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText(message)
        msgBox.setWindowTitle("Congratulations!")
        self.title_label.setText('Powersupply Test')
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)        
        return msgBox.exec_()
    def enable_button(self):
        self.timer1.stop()
        self.start_button.setEnabled(True)
def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())
if __name__ == '__main__':
    main()
