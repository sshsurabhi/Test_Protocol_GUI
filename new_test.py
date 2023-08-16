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
        ########################################################################################################
        # self.timer = QTimer(self)
        # self.timer.timeout.connect(self.update_time_label)
        # self.timer.start(1000) 
        ########################################################################################################
        self.rm = visa.ResourceManager()
        self.multimeter = None
        self.powersupply = None
        self.AC_DC_box.addItems(['<select>', 'DCV', 'ACV'])
        self.AC_DC_box.currentTextChanged.connect(self.selct_AC_DC_box)
        self.test_button.clicked.connect(self.on_cal_voltage_current)
        ########################################################################################################
        self.config_file = configparser.ConfigParser()
        self.config_file.read('conf_igg.ini')
        self.PS_channel = self.config_file.get('Power Supplies', 'Channel_set')
        self.max_voltage = self.config_file.get('Power Supplies', 'Voltage_set')
        self.max_current = self.config_file.get('Power Supplies', 'Current_set')
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
        # self.save_button.clicked.connect(self.create_ini_file)
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
                reply = self.show_good_message('Check All your environment Correct. We cannot change later')                
                if reply == QMessageBox.Yes:                    
                    self.start_button.setText('NEXT')
                    self.on_button_click('images_/images/On_Devices.jpg')
                    self.info_label.setText('Power ON of the Powersupply and also Multimeter...\n and wait for 12 seconds.\n Press "Netzteil ON"')                    
                else:
                    self.on_button_click('images_/icons/next_1.jpg')

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
            self.show_good_message('Wait for 10 seconds. Untill the Powersupply and Multimeter get SET')
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
            reply = self.show_good_message('CLOSE the Jumper with Soldering. \n If You Close then Press YES')
            if reply == QMessageBox.Yes:                    
                self.start_button.setText('STROM')
                self.on_button_click('images_/images/PP8.jpg')
                self.powersupply.write('OUTPut '+self.PS_channel+',ON')
                self.info_label.setText('Press STROM button...\n and and Calculate the supply current\n after closed the JUMPER')                    
            else:
                self.on_button_click('images_/images/close_jumper.jpg')
        elif self.start_button.text()=='STROM':
            self.calc_voltage_before_jumper()
            self.on_button_click('images_/images/R709_before_jumper.jpg')

            

        # elif self.

            

        

            
        ########################################################################################################
    def load_voltage_current(self):
        if self.vals_button.text() == 'CH':
            self.PS_channel = str(self.value_edit.text())
            self.powersupply.write('INSTrument '+self.PS_channel)
            self.config_file.set('Power Supplies', 'Channel_set', self.PS_channel)
            self.vals_button.setText('V')
            self.value_edit.setText(self.max_voltage)
            self.info_label.setText('Write 30 in the Yellow Box next to "V" \n \n Press "Enter"\n You can check the value in the powersupply.')
            self.on_button_click('images_/images/PP7_1.jpg')
            self.value_edit.setValidator(QRegExpValidator(QRegExp(r'^\d+(\.\d+)?$')))
        elif self.vals_button.text() == 'V':
            self.max_voltage =  str(self.value_edit.text())
            self.config_file.set('Power Supplies', 'Voltage_set', self.max_voltage)
            self.powersupply.write(self.PS_channel+':VOLTage ' + self.max_voltage)
            self.vals_button.setText('I')
            self.value_edit.setText(self.max_current)
            self.info_label.setText('Enter 0.5 in the box next to I\n\n Press "Enter".\n Check the value change in the Powersupply.')
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
                self.value_edit.setText(self.PS_channel)
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
