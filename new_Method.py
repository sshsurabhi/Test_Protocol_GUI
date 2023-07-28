import sys, serial, datetime, time
from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import pyvisa as visa



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
                    # self.text_browser.append('Processed Command: '+command + '\n')
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

class SerialPortThread(QThread):        #    Thread to update com ports in the network
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

class App(QMainWindow):
    def __init__(self):
        super(App, self).__init__()
        uic.loadUi("UI/ascha.ui", self)
        self.show()
#############################################################################################################################################################
        self.serial_port = None
        self.thread = None
        self.serial_thread = SerialPortThread()
        self.serial_thread.com_ports_available.connect(self.update_com_ports)
        self.serial_thread.start()
        self.baudrate_box.addItems(['9600','57600','115200'])
        self.baudrate_box.setCurrentText('115200')
        self.connect_button.clicked.connect(self.connect_or_disconnect_serial_port)
        self.refresh_button.clicked.connect(self.refresh_connect)
###################################################################################################
        self.start_button.clicked.connect(self.connect)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time_label)
        self.timer.start(1000) 
        self.rm = visa.ResourceManager()
        self.MM_button.clicked.connect(self.connect_multimeter)
        self.multimeter = None
        self.PS_button.clicked.connect(self.connect_powersupply)
        self.powersupply = None
        self.AC_DC_box.addItems(['<select>', 'DCV', 'ACV'])
        self.AC_DC_box.currentTextChanged.connect(self.handleSelectionChange)
        self.test_button.clicked.connect(self.on_cal_voltage_current)     
        self.DC_values = []
        self.AC_values = []



        self.PS_channel = None
        self.max_voltage = 0
        self.max_current = 0
        self.max_volt_tolz = 0
        self.max_current_tolz = 0
        self.value_edit.returnPressed.connect(self.load_voltage_current)
##########################################################################################################################################################
        self.AC_DC_box.setEnabled(False)
        self.test_button.setEnabled(False)
        self.MM_button.setEnabled(False)
        self.PS_button.setEnabled(False)
        self.Volt_button.setEnabled(False)
        self.current_button.setEnabled(False)
        self.value_edit.setEnabled(False)
        self.connect_button.setEnabled(False)
        self.refresh_button.setEnabled(False)
        self.version_edit.setEnabled(False)
        self.firstMessage()
##########################################################################################################################################################

    def connect_multimeter(self):
        if not self.multimeter:
            try:
                self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::5024::SOCKET')
                self.multimeter.read_termination = '\n'
                self.multimeter.write_termination = '\n'
                self.textBrowser.append(self.multimeter.read())
                self.MM_button.setText('MM OFF')
                self.Volt_button.setEnabled(True)
                self.start_button.setText('SPANNUNG')
            except visa.errors.VisaIOError:
                # QMessageBox.information(self, "Multimeter Connection", "Multimeter is not present at the given IP address.")
                self.textBrowser.append('Multimeter has not been presented')
        else:
            self.multimeter.close()
            self.multimeter = None
            self.MM_button.setText('MM ON')
            self.textBrowser.append('Multimeter Disconnected')

    def connect_powersupply(self):
        if not self.powersupply:
            try:
                self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
                self.powersupply.read_termination = '\n'
                self.powersupply.write_termination = '\n'
                self.textBrowser.setText(self.powersupply.query('*IDN?'))
                self.info_label.setText('Write CH1 in the box next to CH')
                self.PS_button.setText('PS OFF')
                self.PS_button.setEnabled(False)
                self.value_edit.setEnabled(True)
                self.vals_button.setText('CH')
            except visa.errors.VisaIOError:
                QMessageBox.information(self, "PowerSupply Connection", "PowerSupply is not present at the given IP Address.")
                self.textBrowser.setText('Powersupply has not been presented.')
        else:
            self.powersupply.close()
            self.powersupply = None
            self.PS_button.setText('PS ON')
            self.textBrowser.setText('Netzteil Disconnected')
    def calc_voltage(self):
        self.multimeter.query('*')


    def firstMessage(self):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QIcon('images_/icon.png'))
        msgBox.setText("Welcome to the testing world.")
        msgBox.setInformativeText("Press OK if you are ready.")
        msgBox.setWindowTitle("Message")
        msgBox.setStandardButtons(QMessageBox.Ok)
        self.info_label.setText('\n Press START Button \n')
        msgBox.exec_()
                                                                                                                                                                        # ret_value = msgBox.exec_()
                                                                                                                                                                        # if ret_value == QMessageBox.Ok:
                                                                                                                                                                        #     self.secondMessage()

    def secondMessage(self):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QIcon('images_/icon.png'))
        msgBox.setText("Set all the environment as shown in the image. \n \n \n Read the information carefully everytime.")
        msgBox.setInformativeText("Then press the button.")
        msgBox.setWindowTitle("Message")
        msgBox.setStandardButtons(QMessageBox.Ok)
        self.on_button_click('images_/1.jpg')
        self.title_label.setText('Preparation Test')
        self.info_label.setText('\n Press START Button \n')
        msgBox.exec_()

    def on_button_click(self, file_path):
        if file_path:
            pixmap = QPixmap(file_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.setScaledContents(True)
            self.image_label.setFixedSize(pixmap.width(), pixmap.height())

            if self.start_button.text() == 'Step4':
                reply = self.show_good_message()
                if reply == QMessageBox.Yes:
                    self.start_button.setText('Step5')
                else:
                    self.on_button_click('images_/Start4.png')
                    self.on_button_click('Start4.png')

    def show_good_message(self):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText("Good that you have correct Environment.\n Stage 1 has been Completed.\n Do you want to continue?")
        msgBox.setWindowTitle("Congratulations!")
        self.title_label.setText('Powersupply Test')
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        return msgBox.exec_()


    def connect(self):
            if self.start_button.text() == 'START':
                self.on_button_click('images_/Start.png')
                self.info_label.setText('Place the board on ESD Matte and See the Image on the right side.\n Check all your environment according to the image.\n Check All connections.Please close if anything not correct. If correct, then \nplease press the step1 Button.')
                self.start_button.setText("Step1")
            elif self.start_button.text() == 'Step1':
                self.on_button_click('images_/Start1.png')
                self.info_label.setText('Check all the "4" screws are fitted according to the image.\n If not then please fit 4x "M2,5x5" Torx and \n M 2.5 Screws on X200 & X300 \n please press the button Step2.')
                self.start_button.setText("Step2")
            elif self.start_button.text() == 'Step2':
                self.on_button_click('images_/Start.png')
                self.info_label.setText('Place back the Board on ESD Matte.')
                self.start_button.setText("Step3")
            elif self.start_button.text()=='Step3':
                self.on_button_click('images_/Start3.png')
                self.info_label.setText('Connect the other Cable to the board.\n as shown in the Image on the right.')
                self.start_button.setText("Step4")
            elif self.start_button.text()=='Step4':
                self.on_button_click('images_/Start2.png')
                self.info_label.setText('\n Press Netzteil ON Button')
                self.start_button.setText('Step5')
                self.start_button.setText('Netzteil ON')
                self.PS_button.setEnabled(True)
            elif self.start_button.text()== 'Netzteil ON':
                self.info_label.setText('Press Netzteil ON button')
                self.connect_powersupply()
            elif self.start_button.text()== 'MULTI ON':
                self.connect_multimeter()
            elif self.start_button.text()== 'SPANNUNG':
                self.textBrowser.append(self.powersupply.query(self.PS_channel+':VOLTage?'))
                self.start_button.setText('STROM')
            elif self.start_button.text()== 'STROM':
                self.textBrowser.append(self.powersupply.query(self.PS_channel+':CURRent?'))
            else:
                self.start_button.setText("Start")
                self.textBrowser.append('disconnected')

    def handleSelectionChange(self, text):
        if text == 'DCV':
            self.selected_command = 'MEAS:VOLT:DC?'
            self.test_button.setEnabled(True)
            self.on_button_click('images_/3.jpg')
            self.test_button.setText('R709')
            self.info_label.setText('Press R709')
            self.AC_DC_box.setEnabled(False)  # Disable AC_DC_box after selecting 'DCV'
        elif text == 'ACV':
            self.selected_command = 'MEAS:VOLT:AC?'
            self.test_button.setEnabled(True)
            self.test_button.setText('GO')
            self.AC_DC_box.setEnabled(True)  # Enable AC_DC_box after selecting 'ACV'
            self.info_label.setText('Ready to proceed with AC voltage measurement.')
        else:
            self.selected_command = None
            self.test_button.setEnabled(False)
            self.info_label.setText('Select DC in the combobox to proceed further.')

    def on_cal_voltage_current(self):
        if self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'GO':
            self.multimeter.query('*IDN?')
            self.info_label.setText('Press R709')
            
        elif self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'R709':
            self.result_edit.setText(self.multimeter.query('MEASure:VOLTage:DC?'))
            self.textBrowser.setText(str(float(self.powersupply.query('MEASure:VOLTage? CH1'))))

        elif self.AC_DC_box.currentText() == 'ACV' and self.test_button.text() == 'R709':
            self.result_edit.setText(self.multimeter.query('MEAS:VOLT:AC?'))

    def update_time_label(self):
        current_time = QTime.currentTime().toString(Qt.DefaultLocaleLongDate)
        current_date = QDate.currentDate().toString(Qt.DefaultLocaleLongDate)
        self.time_label.setText(f"{current_time} - {current_date}")
        # return current_date, current_time
    def update_com_ports(self, com_ports):
        self.port_box.clear()
        self.port_box.addItems(com_ports)
    def connect_or_disconnect_serial_port(self):
        if self.serial_port is None:
            com_port = self.port_box.currentText()            # Get the selected com port and baud rate
            baud_rate = int(self.baud_rates_combo.currentText())
            self.serial_port = serial.Serial(com_port, baud_rate, timeout=1)            # Create a new serial port object and open it
            self.port_box.setEnabled(False)  # Disable the combo boxes and change the button text
            self.baud_rates_combo.setEnabled(False)
            self.connect_button.setText('Disconnect')
            self.text_browser.append('Serial Communication Connected \n')
            self.test_button.setEnabled(True)
            self.next_button.setEnabled(True)
            self.refresh_button.setEnabled(False)
        else:
            self.serial_port.close()            # Close the serial port
            self.serial_port = None
            self.connect_button.setEnabled(True)
            self.port_box.setEnabled(True)            # Enable the combo boxes and change the button text
            self.baud_rates_combo.setEnabled(True)
            self.refresh_button.setEnabled(True)
            self.next_button.setEnabled(False)
            self.test_button.setEnabled(False)
            self.connect_button.setText('Connect')
            self.text_browser.append('Communication Disconnected \n')
    def refresh_connect(self):
        self.serial_thread.quit()
        self.serial_thread.wait()
        self.serial_thread.start()

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
    def process_completed(self):
        QMessageBox.information(self, "Process Completed", "Process has been completed.")

    def load_voltage_current(self):
        if (self.vals_button.text() == 'CH' and self.value_edit.text() in ['ch1', 'CH1', 'Ch1', 'cH1']):
            self.powersupply.write('INSTrument CH1')
            self.PS_channel = self.value_edit.text()
            self.vals_button.setText('V')
            self.info_label.setText('Write 30 in the box  next to V')
            self.value_edit.clear()
            

        elif (self.vals_button.text() == 'CH' and self.value_edit.text() in ['ch2', 'CH2', 'Ch2', 'cH2']):
            self.powersupply.write('INSTrument CH2')
            self.PS_channel = self.value_edit.text()
            self.vals_button.setText('V')
            self.info_label.setText('Enter 30 in the box  next to V')
            self.value_edit.clear()

        elif (self.vals_button.text() == 'CH' and self.value_edit.text() in ['ch3', 'CH3', 'Ch3', 'cH3']):
            self.powersupply.write('INSTrument CH3')
            self.PS_channel = self.value_edit.text()
            self.vals_button.setText('V')
            self.info_label.setText('Enter 30 in the box  next to V')
            self.value_edit.clear()

        elif self.vals_button.text() == 'V':
            self.value_edit.setValidator(QRegExpValidator(QRegExp(r'^\d+(\.\d+)?$')))
            self.powersupply.write(self.PS_channel+':VOLTage ' + self.value_edit.text())
            print(self.powersupply.query(self.PS_channel+':VOLTage?'))
            self.vals_button.setText('I')
            self.info_label.setText('Enter 0.5 in the box  next to I')
            self.value_edit.clear()

        elif self.vals_button.text() == 'I':
            self.powersupply.write('CH1:CURRent ' + str(float(self.value_edit.text())))
            print(self.powersupply.query(self.PS_channel+':CURRent?'))
            self.vals_button.setText('Tolz V')
            self.info_label.setText('Enter 0.05 in the box  next to I')
            self.value_edit.clear()

        elif self.vals_button.text() == 'Tolz V':
            self.volt_toleranz = self.value_edit.text()
            print(self.volt_toleranz)
            self.value_edit.clear()
            self.info_label.setText('Enter 0.03 in the box  next to I')
            self.vals_button.setText('Tolz I')

        elif self.vals_button.text() == 'Tolz I':
            self.curr_toleranz = self.value_edit.text()
            self.value_edit.setEnabled(False)
            self.MM_button.setEnabled(True)
            self.info_label.setText('Press MULTI ON Button')
            self.start_button.setText('MULTI ON')
            self.powersupply.write('OUTPut CH1,ON')
        else:
            self.textBrowser.append('Wrong Input')

def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
