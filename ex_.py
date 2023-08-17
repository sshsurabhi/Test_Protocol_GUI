# import configparser

# def create_config_file():
#     config = configparser.ConfigParser()

#     config.add_section('Powersupply Test')
#     powersupply_values = [
#         "DCV b/w GND - R709",
#         "DCV b/w GND - R700",
#         "ACV b/w GND - R709",
#         "ACV b/w GND - R700",
#         "DCV b/w ISOGND - C443",
#         "DCV b/w ISOGND - C442",
#         "DCV b/w ISOGND - C441",
#         "DCV b/w ISOGND - C412",
#         "DCV 2.048V b/w ISOGND - C430",
#         "ACV b/w ISOGND - C443",
#         "ACV b/w ISOGND - C442",
#         "ACV b/w ISOGND - C441",
#         "ACV b/w ISOGND - C412"
#     ]
#     for value in powersupply_values:
#         config.set('Powersupply Test', value, '')

#     config.add_section('I2C Test')
#     i2c_values = [
#         "UID",
#         "IC704 result registers Reading"
#     ]
#     for value in i2c_values:
#         config.set('I2C Test', value, '')

#     with open('conf_ig.ini', 'w') as configfile:
#         config.write(configfile)

# def update_value_in_section(section, key, new_value):
#     config = configparser.ConfigParser()
#     config.read('conf_ig.ini')

#     if config.has_section(section) and config.has_option(section, key):
#         config.set(section, key, str(new_value))  # Convert float to string

#         with open('conf_ig.ini', 'w') as configfile:
#             config.write(configfile)

# if __name__ == "__main__":
#     create_config_file()
#     print("conf_ig.ini file created successfully.")

#     update_value_in_section('Powersupply Test', 'DCV b/w GND - R709', 53.5)
#     print("Value updated in 'DCV b/w GND - R709' successfully.")


# import configparser
# import openpyxl

# def convert_ini_to_excel(ini_file, excel_file):
#     config = configparser.ConfigParser()
#     config.read(ini_file)

#     workbook = openpyxl.Workbook()
#     for section in config.sections():
#         worksheet = workbook.create_sheet(title=section)
#         for key, value in config.items(section):
#             worksheet.append([key, value])

#     workbook.remove(workbook['Sheet'])  # Remove default sheet
#     workbook.save(excel_file)

# if __name__ == "__main__":
#     ini_file = 'conf_ig.ini'
#     excel_file = 'conf_ig.xlsx'

#     convert_ini_to_excel(ini_file, excel_file)
#     print(f"INI file '{ini_file}' converted to Excel file '{excel_file}' successfully.")





# import configparser

# def print_ini_file(file_path):
#     config = configparser.ConfigParser()
#     config.read(file_path)

#     for section in config.sections():
#         print(f"[{section}]")
#         for key, value in config.items(section):
#             print(f"{key} = {value}")
#         print()

# if __name__ == "__main__":
#     ini_file = 'conf_ig.ini'
#     print_ini_file(ini_file)


# import configparser
# import openpyxl

# def convert_ini_to_excel(ini_file, excel_file):
#     config = configparser.ConfigParser()
#     config.read(ini_file)

#     workbook = openpyxl.Workbook()

#     for section in config.sections():
#         worksheet = workbook.create_sheet(title=section)
#         worksheet.append(["Section", "Key", "Value"])  # Add headers

#         for key, value in config.items(section):
#             worksheet.append([section, key, value])

#     workbook.remove(workbook['Sheet'])  # Remove default sheet
#     workbook.save(excel_file)

# if __name__ == "__main__":
#     ini_file = 'conf_ig.ini'
#     excel_file = 'conf_ig.xlsx'

#     convert_ini_to_excel(ini_file, excel_file)
#     print(f"INI file '{ini_file}' converted to Excel file '{excel_file}' successfully.")



# import configparser
# import openpyxl

# def convert_ini_to_excel(ini_file, excel_file):
#     config = configparser.ConfigParser()
#     config.read(ini_file)

#     workbook = openpyxl.Workbook()
#     worksheet = workbook.active
#     worksheet.title = 'All Sections'

#     worksheet.append(["Section", "Key", "Value"])  # Add headers

#     for section in config.sections():
#         for key, value in config.items(section):
#             worksheet.append([section, key, value])

#     workbook.save(excel_file)

# if __name__ == "__main__":
#     ini_file = 'conf_ig.ini'
#     excel_file = 'conf_ig.xlsx'

#     convert_ini_to_excel(ini_file, excel_file)
#     print(f"INI file '{ini_file}' converted to Excel file '{excel_file}' successfully.")


# import configparser
# X = 25
# Y = 54
# YX = 75
# XY = 857
# def create_ini_file():
#     config = configparser.ConfigParser()

#     config.add_section('Powersupply Test')
#     powersupply_values = {
#         "DCV b/w GND - R709" : X,
#         "DCV b/w GND - R700" : Y,
#         "ACV b/w GND - R709" : XY,
#         "ACV b/w GND - R700" : YX
#     }
#     # Loop through the dictionary to set values for each key
#     for key, value in powersupply_values.items():
#         config.set('Powersupply Test', key, str(value))

#     config.add_section('I2C Test')
#     i2c_values = [
#         "UID",
#         "IC704 result registers Reading"
#     ]
#     for value in i2c_values:
#         config.set('I2C Test', value, '')

#     with open('conf_igg.ini', 'w') as configfile:
#         config.write(configfile)

# if __name__ == "__main__":
#     create_ini_file()
#     print("conf_ig.ini file created successfully.")


# import time
# import pyvisa as visa

# # Connect to the multimeter
# rm = visa.ResourceManager()
# dmm = rm.open_resource('TCPIP0::192.168.222.207::INSTR')  # Replace with your actual device address

# # Set up the multimeter
# # dmm.write('CONF:VOLT:DC 10,0.001')  # Configure DC voltage measurement range and resolution

# # Initialize variables to store voltage readings
# num_readings = 13
# voltage_readings = []

# # Read voltage every 10 seconds for a total of num_readings times
# for i in range(num_readings):
#     voltage = float(dmm.query('MEAS:VOLT:DC?'))
#     voltage_readings.append(voltage)
#     print(f'Voltage {i+1}: {voltage} V')
#     if i < num_readings - 1:
#         time.sleep(10)

# # Close the connection to the multimeter
# dmm.close()
# rm.close()

# # Store the readings in separate variables (var1, var2, var3, etc.)
# var_names = [f'var{i+1}' for i in range(num_readings)]
# readings_dict = {var_name: voltage for var_name, voltage in zip(var_names, voltage_readings)}

# # Print the stored readings
# for var_name, voltage in readings_dict.items():
#     print(f'{var_name}: {voltage} V')



# import sys
# from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QMessageBox
# from PyQt5.QtCore import QTimer

# class App1(QMainWindow):
#     def __init__(self):
#         super(App1, self).__init__()
#         self.setWindowTitle("APP1")
#         self.setGeometry(100, 100, 300, 200)

#         self.press_button = QPushButton("Press", self)
#         self.press_button.setGeometry(100, 100, 100, 30)
#         self.press_button.clicked.connect(self.handle_button_click)

#     def handle_button_click(self):
#         self.press_button.setEnabled(False)  # Disable the button
#         self.show_waiting_popup()

#     def show_waiting_popup(self):
#         self.timer = QTimer(self)
#         self.timer.timeout.connect(self.enable_button)
#         self.timer.start(5000)  # Start a timer for 5 seconds

#         waiting_popup = QMessageBox()
#         waiting_popup.setWindowTitle("Waiting")
#         waiting_popup.setText("Please wait for 5 seconds...")
#         waiting_popup.exec_()

#     def enable_button(self):
#         self.timer.stop()
#         self.press_button.setEnabled(True)  # Enable the button

# def main():
#     app = QApplication(sys.argv)
#     window = App1()
#     window.show()
#     sys.exit(app.exec_())

# if __name__ == "__main__":
#     main()

# import pyvisa as visa
# import time

# # Connect to the multimeter
# rm = visa.ResourceManager()
# dmm = rm.open_resource('TCPIP0::192.168.222.207::INSTR')  # Replace with your actual device address
# time.sleep(2)
# voltage = float(dmm.query('MEAS:VOLT:DC?'))
# print(voltage)
# time.sleep(2)
# v = float(dmm.query('MEAS:VOLT:AC?'))
# print(v)
# time.sleep(2)
# volt = float(dmm.query('MEAS:VOLT:AC?'))
# print(volt)
# time.sleep(2)
# dc_volt = float(dmm.query('MEAS:VOLT:DC?'))
# print(dc_volt)
# time.sleep(2)
# D_volt = float(dmm.query('MEAS:VOLT:DC?'))
# print(D_volt)
# time.sleep(2)
# ac_volt = float(dmm.query('MEAS:VOLT:AC?'))
# print(ac_volt)
import sys
import os
from PyQt5.QtWidgets import QApplication, QLabel, QMainWindow, QVBoxLayout, QWidget
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import QTimer, Qt

class ImageChangerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Image Changer")
        self.setGeometry(100, 100, 400, 300)

        # Create QLabel widget to display images
        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)

        # Set up layout
        layout = QVBoxLayout()
        layout.addWidget(self.image_label)

        central_widget = QWidget(self)
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # List of image filenames
        self.image_filenames = ["images_/images/C443.jpg", "images_/images/C412.jpg", "images_/images/C430.jpg"]
        self.current_image_index = 0

        # Create a timer to control image changes
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.change_image)
        self.timer.start(2000)  # Change image every 2 seconds

    def change_image(self):
        if self.current_image_index < len(self.image_filenames):
            image_filename = self.image_filenames[self.current_image_index]
            pixmap = QPixmap(image_filename)
            self.image_label.setPixmap(pixmap)
            self.current_image_index += 1
        else:
            self.timer.stop()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ImageChangerApp()
    window.show()
    sys.exit(app.exec_())
