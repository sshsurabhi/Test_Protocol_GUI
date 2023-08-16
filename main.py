import datetime, os, sys, time, openpyxl, configparser, serial
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
##################################################################################################################################
class WorkerThread(QThread):
    process_completed = pyqtSignal()
    result_signal = pyqtSignal(str)
    response_signal = pyqtSignal(str)
    def __init__(self, commands, serial_port):
        super().__init__()
        self.commands = commands
        self.serial_port = serial_port
    def run(self):
        if self.serial_port.is_open:
            self.serial_port.timeout = 5
            self.next_command_index = 0
            for command in self.commands:
                if self.next_command_index < (len(self.commands)+1):
                    self.serial_port.write(command.encode())
                    # self.textBrowser.append('Processed Command: '+command + ' ')
                    response = self.serial_port.readline().decode('ascii')
                    self.result_signal.emit((f"Response: {response}"))
                    self.next_command_index += 1
                    current_date = datetime.datetime.now()
                    decimal_date = int(current_date.strftime('%Y%m%d'))
                    hex_date = hex(decimal_date)[2:].upper().zfill(8)
                    if command == self.commands[1]:
                        ading = response.split(':')[3][:-1]
                        self.commands[2] = self.commands[2]+ading+'01'+hex_date+'2A0101030000FFFF'
                        self.response_signal.emit(ading)
                    if command == self.commands[2]:
                        start_time = time.time()
                        while time.time() - start_time < 5:
                            if response:
                                break
                            response = self.serial_port.readline().decode('ascii')
                            self.result_signal.emit(response)
            self.serial_port.close()
            self.process_completed.emit()
        else:
            self.result_signal.emit('Serial Port Closed')
    def on_button_clicked(self):
        self.result_signal.emit("Button clicked")
##################################################################################################################################
class SerialPortThread(QThread):
    com_ports_available = pyqtSignal(list)
    def run(self):
        com_ports = []
        for i in range(256):
            try:
                s = serial.Serial('COM'+str(i))
                com_ports.append('COM'+str(i))
                s.close()
            except serial.SerialException:
                pass
        self.com_ports_available.emit(com_ports)
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
        self.serial_thread = SerialPortThread()
        self.serial_thread.com_ports_available.connect(self.update_com_ports)
        self.serial_thread.start()
        self.baudrate_box.addItems(['9600','57600','115200'])
        self.baudrate_box.setCurrentText('115200')
        self.connect_button.clicked.connect(self.connect_or_disconnect_serial_port)
        self.refresh_button.clicked.connect(self.refresh_connect)
        ########################################################################################################
        self.commands = ['i2c:scan', 'i2c:read:53:04:FC', 'i2c:write:53:', 'i2c:read:53:20:00', 'i2c:write:73:04', 'i2c:scan','i2c:write:21:0300','i2c:write:21:0100','i2c:write:21:01FF', 'i2c:write:73:01',
                    'i2c:scan', 'i2c:write:4F:06990918', 'i2c:write:4F:01F8', 'i2c:read:4F:1E:00']
        self.start_button.clicked.connect(self.connect)
        # self.timer = QTimer(self)
        # self.timer.timeout.connect(self.update_time_label)
        # self.timer.start(1000)

    
        self.rm = visa.ResourceManager()
        self.multimeter = None
        self.powersupply = None
        self.AC_DC_box.addItems(['<select>', 'DCV', 'ACV'])
        # self.AC_DC_box.currentTextChanged.connect(self.handleSelectionChange)
        self.AC_DC_box.currentTextChanged.connect(self.selct_AC_DC_box)
        self.test_button.clicked.connect(self.on_cal_voltage_current)
        self.PS_channel = None
        self.max_voltage = 0
        self.max_current = 0
        self.max_volt_tolz = 0
        self.max_current_tolz = 0
        self.value_edit.returnPressed.connect(self.load_voltage_current)
        ########################################################################################################
        self.AC_DC_box.setEnabled(False)
        self.test_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.value_edit.setEnabled(False)
        self.connect_button.setEnabled(False)
        self.refresh_button.setEnabled(False)
        self.version_edit.setEnabled(False)
        self.result_edit.setEnabled(False)
        self.firstMessage()
        ########################################################################################################
        self.save_button.clicked.connect(self.create_ini_file)
        self.current_before_jumper = 0
        self.current_after_jumper = 0
        self.voltage_before_jumper = 0
        self.dcv_bw_gnd_n_r709 = 0
        self.dcv_bw_gnd_n_r700 = 0
        self.acv_bw_gnd_n_r709 = 0
        self.acv_bw_gnd_n_r700 = 0
        self.dcv_bw_gnd_n_c443 = 0
        self.dcv_bw_gnd_n_c442 = 0
        self.dcv_bw_gnd_n_c441 = 0
        self.dcv_bw_gnd_n_c412 = 0
        self.dcv_bw_gnd_n_c430 = 0
        self.acv_bw_gnd_n_c443 = 0
        self.acv_bw_gnd_n_c442 = 0
        self.acv_bw_gnd_n_c441 = 0
        self.acv_bw_gnd_n_c412 = 0
        self.acv_bw_gnd_n_c430 = 0
        self.uid = 0
        self.ic704_register_reading = 0
        ########################################################################################################
    def connect_multimeter(self):
        if not self.multimeter:
            try:
                self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
                self.textBrowser.append(self.multimeter.query('*IDN?'))
                self.on_button_click('images_/images/R709_before_jumper.jpg')
                self.start_button.setText('SPANNUNG')
                self.info_label.setText('Check VOLTAGE at the component R709.\nIt should be between 3.28V to 3.8V.\n If it is other value,\nthen please pack the board back.' )
            except visa.errors.VisaIOError:
                self.textBrowser.append('Multimeter has not been presented')
        else:
            self.multimeter.close()
            self.multimeter = None
            self.textBrowser.append(self.multimeter.query('*IDN?'))
    ########################################################################################################
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
    ########################################################################################################
    def firstMessage(self):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QIcon('images_/icons/icon.png'))
        msgBox.setText("Welcome to the testing world.")
        msgBox.setInformativeText("Press OK if you are ready.")
        msgBox.setWindowTitle("Message")
        self.on_button_click('images_/images/Welcome.jpg')
        msgBox.setStandardButtons(QMessageBox.Ok)
        self.info_label.setText('Now,\n  Press START Button  ')
        ret_value = msgBox.exec_()
        if ret_value == QMessageBox.Ok:
            self.secondMessage()
    ########################################################################################################
    def secondMessage(self):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QIcon('images_/icons/icon.png'))
        msgBox.setText("Press the PowerON buttons of PowerSupply and Multimeter to avoid delay.\n\nSet all the environment as shown in the image.\n\nRead the information carefully everytime.\n")
        msgBox.setInformativeText("Then press the button.")
        msgBox.setWindowTitle("Message")
        msgBox.setStandardButtons(QMessageBox.Ok)
        ret_value = msgBox.exec_()
        if ret_value == QMessageBox.Ok:
            self.title_label.setText('Preparation Test')
            self.info_label.setText('Press START Button')
            self.on_button_click('images_/images/PP1.jpg')
    ########################################################################################################
    def on_button_click(self, file_path):
        if file_path:
            pixmap = QPixmap(file_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.setScaledContents(True)
            self.image_label.setFixedSize(pixmap.width(), pixmap.height())

            if self.start_button.text() == 'STEP4':
                reply = self.show_good_message()                
                if reply == QMessageBox.Yes:                    
                    self.start_button.setText('Netzteil ON')
                    self.on_button_click('images_/images/img4.jpg')
                    self.info_label.setText('Power ON of the Powersupply and also Multimeter...\n and wait for 5 seconds.\n Press "Netzteil ON"')                    
                else:
                    self.on_button_click('images_/icons/3.jpg')

            if self.start_button.text() == 'JUMPER OK':
                reply = self.jumper_close()
                if reply == QMessageBox.Yes:
                    self.start_button.setText('STROM')
                    self.start_button.setEnabled(True)
                    self.powersupply.write('OUTPut '+self.PS_channel+',ON')
                    self.on_button_click('images_/images/PP16.jpg')
                    self.info_label.setText('Press STROM button')
                else:
                    self.on_button_click('images_/images/PP17.jpg')
    ########################################################################################################
    def show_good_message(self):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText("Good that you have correct Environment. Stage 1 has been Completed. Do you want to continue?")
        msgBox.setWindowTitle("Congratulations!")
        self.title_label.setText('Powersupply Test')
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)        
        return msgBox.exec_()
    ########################################################################################################
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
            self.on_button_click('images_/icons/2.jpg')
        elif self.start_button.text()== 'Netzteil ON':
            self.connect_powersupply()
        elif self.start_button.text() == 'MESS I':
            self.calc_voltage_before_jumper()
        elif self.start_button.text() == 'JUMPER OK':
            self.on_button_click('images_/Power_ON_PS.jpg')
        elif self.start_button.text()== 'MULTI ON':
            self.connect_multimeter()
        elif self.start_button.text()== 'SPANNUNG':
            self.info_label.setText('Press "MESS I" button')
            self.on_button_click('images_/images/PP7_4.jpg')
            self.start_button.setText('MESS I')
            voltage_before_jumper = self.multimeter.query('MEAS:VOLT:DC?')
            self.voltage_before_jumper = voltage_before_jumper
            self.result_edit.setText(str(voltage_before_jumper))
            self.textBrowser.append(('PowerSUpply \n'+self.powersupply.query('MEASure:VOLTage? '+self.PS_channel)+'V'))
        elif self.start_button.text()== 'STROM':
            self.calc_voltage_before_jumper()
        elif (self.start_button.text()== 'Messung' and self.AC_DC_box.text == 'DCV'):
            self.mess_with_multimeter()
            self.start_button.setText('Messung')
        elif self.start_button.text()== 'Messung':
            self.mess_with_multimeter()
        elif self.start_button.text()=='Power OFF':
            self.info_label.setText('Switch OFF both PowerSupply & Multimeter.')
            self.on_button_click('images_/images/Start2.png')
            QMessageBox.information(self, "Important", "Read and watch the image clearly and do the process carefully.\n Press 'OFF ALL' button")
            self.start_button.setText('OFF ALL')
        elif self.start_button.text()=='OFF ALL':
            self.info_label.setText('Take FPGA Module and do as shown in the figure. Do all necessary things. \n then Press FPGA')
            self.on_button_click('images_/images/FPGA_.jpg')
            QMessageBox.information(self, "Important", "Read and watch the image clearly and do the process carefully.")
            self.start_button.setText('FPGA')
        elif self.start_button.text()=='FPGA':
            self.info_label.setText('FIX FPGA Module to the board as shown.\The Press JETZT button')
            self.on_button_click('images_/images/FPGA_2.jpg')
            QMessageBox.information(self, "Important", "Read and watch the image clearly and do the process carefully.")
            self.start_button.setText('JETZT')     
        elif self.start_button.text()=='JETZT':
            self.info_label.setText('Switch ON Powersupply, and Multimeter')
            self.on_button_click('images_/images/board_with_cabels.jpg')
            QMessageBox.information(self, "Important", "Read and watch the image clearly and do the process carefully.")
            self.start_button.setText('MESS AN')
        elif self.start_button.text()=='MESS AN':
            self.info_label.setText('Switch ON Powersupply, and Multimeter and wait 5 seconds.\Then press SERIAL TEST button')
            self.on_button_click('images_/images/Start2.png')
            QMessageBox.information(self, "Important", "Read and watch the image clearly and do the process carefully.")
            self.connect_button.setEnabled(True)
            self.start_button.setEnabled(False)
        elif self.start_button.text()=='SERIAL TEST':
            self.start_process()
            self.info_label.setText('Wait 10 seconds ')
            self.on_button_click('images_/images/Start2.png')
            QMessageBox.information(self, "Important", "Read and watch the image clearly and do the process carefully.")
        else:
            self.start_button.setText("Start")
            self.textBrowser.append('disconnected')
    ########################################################################################################
    def on_cal_voltage_current(self):
        if self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'GO':
            ret_volt = self.multimeter.query('*IDN?')
            self.info_label.setText('See the Component in the Image and\n\n\n check the voltage at the same component.')
            self.textBrowser.append(str(ret_volt))
            self.test_button.setText('R709')
            self.on_button_click('images_/images/R709.jpg')        
        elif self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'R709':
            ret_volt = self.multimeter.query('MEAS:VOLT:DC?')
            if 3.28 <= float(ret_volt) <= 3.38:
                self.textBrowser.append("DC Voltage at R709:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
                self.dcv_bw_gnd_n_r709 = ret_volt
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                self.dcv_bw_gnd_n_r709 = str(ret_volt)+' Negative Value'
                self.dcv_bw_gnd_n_r709 = ret_volt
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('R700')
            self.on_button_click('images_/images/R700.jpg')
        elif self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'R700':
            ret_volt = self.multimeter.query('MEAS:VOLT:DC?')
            if 4.95 <= float(ret_volt) <= 5.05:
                self.textBrowser.append("DC Voltage at R700:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
                self.dcv_bw_gnd_n_r700 = ret_volt
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                self.dcv_bw_gnd_n_r700 = ret_volt
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('R709')
            self.test_button.setEnabled(False)
            self.AC_DC_box.setEnabled(True)
            self.on_button_click('images_/images/R709.jpg')
            self.info_label.setText('Change the selection in AC/DC Box. \n\n\n Now we have to calculate AC Voltage at the last two components... \n\n\n So, Select ACV in AC/DC Box')
            QMessageBox.information(self, "Information", "Select ACV in the box near AC/DC")
        elif self.AC_DC_box.currentText() == 'ACV' and self.test_button.text() == 'R709':
            ret_volt = self.multimeter.query('MEAS:VOLT:AC?')
            if float(ret_volt) <= 0.01:
                self.textBrowser.append("AC Voltage at R709:"+ str(ret_volt))
                self.acv_bw_gnd_n_r709  = ret_volt
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                self.acv_bw_gnd_n_r709  = ret_volt
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('R700')
            self.on_button_click('images_/images/R709.jpg')
        elif self.AC_DC_box.currentText() == 'ACV' and self.test_button.text() == 'R700':
            ret_volt = self.multimeter.query('MEAS:VOLT:AC?')
            if float(ret_volt) <= 0.01:
                self.textBrowser.append("AC Voltage at R700:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setEnabled(False)
            self.AC_DC_box.setEnabled(True)
            self.info_label.setText('Select DCV in AC DC Box')
            QMessageBox.information(self, "Information", "Select DCV in the box near AC/DC")
            self.test_button.setText('C443')
            self.on_button_click('images_/images/C443.jpg')
        elif self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'C443':
            ret_volt = self.multimeter.query('MEAS:VOLT:DC?')
            if 11.95 <= float(ret_volt) <= 12.05:
                self.textBrowser.append("DC Voltage at C443:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('C442')
            self.on_button_click('images_/images/C442.jpg')
        elif self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'C442':
            ret_volt = self.multimeter.query('MEAS:VOLT:DC?')
            if 4.95 <= float(ret_volt) <= 5.05:
                self.textBrowser.append("DC Voltage at C442:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('C441')
            self.on_button_click('images_/images/C441.jpg')
        elif self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'C441':
            ret_volt = self.multimeter.query('MEAS:VOLT:DC?')
            if 4.95 <= float(ret_volt) <= 5.05:
                self.textBrowser.append("DC Voltage at C441:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('C412')
            self.on_button_click('images_/images/C412.jpg')
        elif self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'C412':
            ret_volt = self.multimeter.query('MEAS:VOLT:DC?')
            if 4.98 <= float(ret_volt) <= 5.02:
                self.textBrowser.append("DC Voltage at C412:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('C430')
            self.on_button_click('images_/images/C430.jpg')
        elif self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'C430':
            ret_volt = self.multimeter.query('MEAS:VOLT:DC?')
            if 2.028 <= float(ret_volt) <= 2.068:
                self.textBrowser.append("DC Voltage at C430:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('C443')
            self.on_button_click('images_/images/C443.jpg')
            self.test_button.setEnabled(False)
            self.AC_DC_box.setEnabled(True)
            self.info_label.setText('\n \n \n \n Select DCV in AC DC Box')        
        elif self.AC_DC_box.currentText() == 'ACV' and self.test_button.text() == 'C443':
            ret_volt = self.multimeter.query('MEAS:VOLT:AC?')
            if float(ret_volt) < 0.01:
                self.textBrowser.append("DC Voltage at C443:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('C442')
            self.on_button_click('images_/images/C442.jpg')
        elif self.AC_DC_box.currentText() == 'ACV' and self.test_button.text() == 'C442':
            ret_volt = self.multimeter.query('MEAS:VOLT:AC?')
            if float(ret_volt) < 0.05:
                self.textBrowser.append("DC Voltage at C442:"+ str(ret_volt))
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('C441')
            self.on_button_click('images_/images/C441.jpg')
        elif self.AC_DC_box.currentText() == 'ACV' and self.test_button.text() == 'C441':
            ret_volt = self.multimeter.query('MEAS:VOLT:AC?')
            if float(ret_volt) < 0.05:
                self.textBrowser.append("DC Voltage at C441:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('C412')
            self.on_button_click('images_/images/C412.jpg')
        elif self.AC_DC_box.currentText() == 'ACV' and self.test_button.text() == 'C412':
            ret_volt = self.multimeter.query('MEAS:VOLT:AC?')
            if float(ret_volt) < 0.001:
                self.textBrowser.append("DC Voltage at C412:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setText('C430')
            self.on_button_click('images_/images/C430.jpg')
        elif self.AC_DC_box.currentText() == 'ACV' and self.test_button.text() == 'C430':
            ret_volt = self.multimeter.query('MEAS:VOLT:AC?')
            if float(ret_volt) < 0.001:
                self.textBrowser.append("DC Voltage at C412:"+ str(ret_volt))
                self.result_edit.setStyleSheet("background-color: green;")
                self.result_edit.setText(ret_volt)
            else:
                self.result_edit.setStyleSheet("background-color: red;")
                self.result_edit.setText(ret_volt)
                QMessageBox.information(self, "Status", "Voltage is diferred"+str(ret_volt))
            self.test_button.setEnabled(False)
            self.info_label.setText('\n \n \n \n SWITCH OFF the POWERSUPPLY and Assemble SDIO board, spacers and screws')
            self.start_button.setEnabled(True)
            self.start_button.setText('Power OFF')
            self.on_button_click('images_/images/img8.jpg')
            QMessageBox.information(self, "Status", "Assemble SDIO board, spacers and screws")
        else:
            QMessageBox.information(self, "Status", "Wrong Testing")
    ########################################################################################################
    def selct_AC_DC_box(self):
        self.test_button.setEnabled(True)
        self.AC_DC_box.setEnabled(False)
    ########################################################################################################
    def update_time_label(self):
        current_time = QTime.currentTime().toString(Qt.DefaultLocaleLongDate)
        current_date = QDate.currentDate().toString(Qt.DefaultLocaleLongDate)
        self.time_label.setText(f"{current_time} - {current_date}")
        # return current_date, current_time
    ########################################################################################################
    def update_com_ports(self, com_ports):
        self.port_box.clear()
        self.port_box.addItems(com_ports)
    ########################################################################################################
    def connect_or_disconnect_serial_port(self):
        if self.serial_port is None:
            com_port = self.port_box.currentText()            # Get the selected com port and baud rate
            baud_rate = int(self.baudrate_box.currentText())
            self.serial_port = serial.Serial(com_port, baud_rate, timeout=1)            # Create a new serial port object and open it
            self.port_box.setEnabled(False)  # Disable the combo boxes and change the button text
            self.start_button.setText('SERIAL TEST')
            self.start_button.setEnabled(True)
            self.connect_button.setEnabled(True)
            self.baudrate_box.setEnabled(False)
            self.connect_button.setText('Disconnect')
            self.textBrowser.append('Serial Communication Connected')
            self.refresh_button.setEnabled(False)
        else:
            self.serial_port.close()            # Close the serial port
            self.serial_port = None
            self.connect_button.setEnabled(True)
            self.port_box.setEnabled(True)            # Enable the combo boxes and change the button text
            self.baudrate_box.setEnabled(True)
            self.refresh_button.setEnabled(True)
            self.connect_button.setText('Connect')
            self.textBrowser.append('Communication Disconnected')
    ########################################################################################################
    def refresh_connect(self):
        self.serial_thread.quit()
        self.serial_thread.wait()
        self.serial_thread.start()
    ########################################################################################################
    def on_widget_button_clicked(self, message):
        self.textBrowser.append(message)
    def update_lineinsert(self, response):
            self.id_Edit.setText(response)
    ########################################################################################################
    def start_process(self):
        if self.thread is None or not self.thread.isRunning():
            QMessageBox.information(self, "Process Started", "Process has been started.")
            self.thread = WorkerThread(self.commands, self.serial_port)
            self.thread.result_signal.connect(self.on_widget_button_clicked)
            self.thread.process_completed.connect(self.process_completed)
            self.thread.response_signal.connect(self.update_lineinsert)
            self.thread.start()
        else:
            QMessageBox.warning(self, "Process In Progress", "Process is already running.")
    ########################################################################################################
    def process_completed(self):
        QMessageBox.information(self, "Process Completed", "Process has been completed.")
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
            self.info_label.setText('Press MULTI ON') # modify here'
            self.start_button.setText('MULTI ON')
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
    ########################################################################################################
    def calc_voltage_before_jumper(self):
        current = float(self.powersupply.query('MEASure:CURRent? '+self.PS_channel))        
        self.result_edit.setText('Current: '+str(current))
        self.textBrowser.append('Current: '+str(current))
        if self.start_button.text() == 'MESS I':
            if 0.04 <= current <= 0.06:
                self.start_button.setText('JUMPER OK')
                self.start_button.setEnabled(False)
                self.current_before_jumper = current
                self.powersupply.write('OUTPut '+self.PS_channel+',OFF')
                self.on_button_click('images_/images/close_jumper.jpg')
                self.info_label.setText('Close the JUMPER')
            else:
                QMessageBox.information(self, 'Information', 'Supplying Current is either more or less. So please Swith OFF the PowerSupply, and Put back all the Euipment back.')
                self.powersupply.write('OUTPut '+self.PS_channel+',OFF')
                self.start_button.setText('JUMPER OK')
                self.start_button.setEnabled(False)
                self.info_label.setText('Close the JUMPER')
                self.on_button_click('images_/images/close_jumper.jpg')
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
    ########################################################################################################
    def jumper_close(self):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText("Close the Jumper Before proceed further. If closed the Press Yes")
        msgBox.setWindowTitle("IMPORTANT!")
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        return msgBox.exec_()
    ########################################################################################################
    def closeEvent(self, event):
        if not self.powersupply is None:
            try:
                response = self.powersupply.query(f"OUTP? CH{self.PS_channel}")
                print(response)
                if not response == None:
                    self.powersupply.write('OUTPut '+self.PS_channel+',OFF')
                else:
                    pass
            except visa.errors.InvalidSession:
                print('Error in PS')
        else:
            pass
        event.accept()
    ########################################################################################################
    def create_ini_file(self):
        config = configparser.ConfigParser()
        config.add_section('Powersupply Test')
        config.add_section('I2C Test')
        power_supply_values = {
            "Supply Voltage" : self.max_voltage,
            "Supply Current" : self.max_current,
            "current_before_jumper" : self.current_before_jumper,
            "voltage_before_jumper" : self.voltage_before_jumper,
            "current_after_jumper" : self.current_after_jumper,
            "dcv_bw_gnd_n_r709" : self.dcv_bw_gnd_n_r709,
            "dcv_bw_gnd_n_r700" : self.dcv_bw_gnd_n_r700,
            "acv_bw_gnd_n_r709" : self.acv_bw_gnd_n_r709,
            "acv_bw_gnd_n_r700" : self.acv_bw_gnd_n_r700,
            "dcv_bw_gnd_n_c443" : self.dcv_bw_gnd_n_c443,
            "dcv_bw_gnd_n_c442" : self.dcv_bw_gnd_n_c442,
            "dcv_bw_gnd_n_c441" : self.dcv_bw_gnd_n_c441,
            "dcv_bw_gnd_n_c412" : self.dcv_bw_gnd_n_c412,
            "dcv_bw_gnd_n_c430" : self.dcv_bw_gnd_n_c430,
            "acv_bw_gnd_n_c443" : self.acv_bw_gnd_n_c443,
            "acv_bw_gnd_n_c442" : self.acv_bw_gnd_n_c442,
            "acv_bw_gnd_n_c441" : self.acv_bw_gnd_n_c441,
            "acv_bw_gnd_n_c412" : self.acv_bw_gnd_n_c412,
            "acv_bw_gnd_n_c430" : self.acv_bw_gnd_n_c430
        }
        I2C_Values = {
            "UID" : self.uid,
            "ic704_register_reading" : self.ic704_register_reading
        }
        for key, value in power_supply_values.items():
            config.set('Powersupply Test', key, str(value))
        for key, value in I2C_Values.items():
            config.set('I2C Test', key, str(value))
        with open('conf_igg.ini', 'w') as configfile:
            config.write(configfile)
    ########################################################################################################
def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())
if __name__ == '__main__':
    main()
