# -*- coding: utf-8 -*-

import sys
import os
import xmlrpc.client
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.label_1 = QtWidgets.QLabel(self.centralwidget)
        self.label_1.setGeometry(QtCore.QRect(160, 40, 421, 31))
        self.label_1.setMinimumSize(QtCore.QSize(421, 31))
        self.label_1.setMaximumSize(QtCore.QSize(421, 16777215))
        self.label_1.setStyleSheet("color: rgb(237, 51, 59);")
        self.label_1.setObjectName("label_1")


        self.lineEdit_url = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_url.setGeometry(QtCore.QRect(270, 120, 251, 31))
        self.lineEdit_url.setStyleSheet("background-color: rgb(204, 229, 255);")
        self.lineEdit_url.setObjectName("lineEdit_url")

        self.lineEdit_db = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_db.setGeometry(QtCore.QRect(270, 180, 251, 31))
        self.lineEdit_db.setStyleSheet("background-color: rgb(204, 229, 255);")
        self.lineEdit_db.setObjectName("lineEdit_db")

        self.lineEdit_user = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_user.setGeometry(QtCore.QRect(270, 240, 251, 31))
        self.lineEdit_user.setStyleSheet("background-color: rgb(204, 229, 255);")
        self.lineEdit_user.setObjectName("lineEdit_user")

        self.lineEdit_password = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_password.setGeometry(QtCore.QRect(270, 300, 251, 31))
        self.lineEdit_password.setStyleSheet("background-color: rgb(204, 229, 255);")
        self.lineEdit_password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.lineEdit_password.setObjectName("lineEdit_password")

        self.lineEdit_file_path = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_file_path.setGeometry(QtCore.QRect(270, 360, 251, 31))
        self.lineEdit_file_path.setStyleSheet("background-color: rgb(143, 240, 164);")
        self.lineEdit_file_path.setPlaceholderText("/home/your/file/path")
        self.lineEdit_file_path.setObjectName("lineEdit_file_path")

        self.label_url = QtWidgets.QLabel(self.centralwidget)
        self.label_url.setGeometry(QtCore.QRect(160, 120, 71, 31))
        self.label_url.setStyleSheet("font: 57 14pt \"Ubuntu\";\n"
                                     "font: 11pt \"Ubuntu\";")
        self.label_url.setObjectName("label_url")

        self.label_db = QtWidgets.QLabel(self.centralwidget)
        self.label_db.setGeometry(QtCore.QRect(160, 180, 91, 31))
        self.label_db.setObjectName("label_db")

        self.label_user = QtWidgets.QLabel(self.centralwidget)
        self.label_user.setGeometry(QtCore.QRect(160, 240, 121, 31))
        self.label_user.setObjectName("label_user")

        self.label_password = QtWidgets.QLabel(self.centralwidget)
        self.label_password.setGeometry(QtCore.QRect(160, 300, 81, 31))
        self.label_password.setObjectName("label_password")


        # App Name Label and ComboBox
        self.label_app_name = QtWidgets.QLabel(self.centralwidget)
        self.label_app_name.setGeometry(QtCore.QRect(160, 410, 91, 31))
        self.label_app_name.setText("App Name")
        self.label_app_name.setObjectName("label_app_name")

        self.comboBox_app_name = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_app_name.setGeometry(QtCore.QRect(270, 410, 251, 31))
        self.comboBox_app_name.setObjectName("comboBox_app_name")
        self.comboBox_app_name.addItems(["Select App", "Contact", "Attendance"])


        # Interval Label
        self.label_interval = QtWidgets.QLabel(self.centralwidget)
        self.label_interval.setGeometry(QtCore.QRect(600, 120, 100, 20))
        self.label_interval.setText("Select Interval")
        self.label_interval.setObjectName("label_interval")
        self.label_interval.setStyleSheet("color: #2E67F8;")

        # Radio Buttons for interval selection
        self.radio_10min = QtWidgets.QRadioButton("10 Minutes", self.centralwidget)
        self.radio_10min.setGeometry(QtCore.QRect(600, 145, 100, 20))
        self.radio_1hr = QtWidgets.QRadioButton("1 Hour", self.centralwidget)
        self.radio_1hr.setGeometry(QtCore.QRect(600, 165, 100, 20))
        self.radio_6hr = QtWidgets.QRadioButton("6 Hours", self.centralwidget)
        self.radio_6hr.setGeometry(QtCore.QRect(600, 185, 100, 20))
        self.radio_12hr = QtWidgets.QRadioButton("12 Hours", self.centralwidget)
        self.radio_12hr.setGeometry(QtCore.QRect(600, 205, 100, 20))
        self.radio_24hr = QtWidgets.QRadioButton("24 Hours", self.centralwidget)
        self.radio_24hr.setGeometry(QtCore.QRect(600, 225, 100, 20))

        # Group the radio buttons
        self.interval_button_group = QtWidgets.QButtonGroup(self.centralwidget)
        self.interval_button_group.addButton(self.radio_10min)
        self.interval_button_group.addButton(self.radio_1hr)
        self.interval_button_group.addButton(self.radio_6hr)
        self.interval_button_group.addButton(self.radio_12hr)
        self.interval_button_group.addButton(self.radio_24hr)

        # Initialize the timer
        self.timer = QtCore.QTimer(self.centralwidget)
        self.timer.timeout.connect(self.send_data)


        # Buttons for connect and send data
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(160, 490, 181, 41))
        self.pushButton.setStyleSheet("background-color: rgb(53, 132, 228);\n"
                                             "background-color: rgb(98, 160, 234);")
        self.pushButton.setObjectName("pushButton")

        self.pushButton1 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton1.setGeometry(QtCore.QRect(380, 490, 177, 41))
        self.pushButton1.setStyleSheet("background-color: rgb(143, 240, 164);")
        self.pushButton1.setObjectName("pushButton1")

        
        self.label_file_path = QtWidgets.QLabel(self.centralwidget)
        self.label_file_path.setGeometry(QtCore.QRect(160, 360, 81, 31))
        self.label_file_path.setObjectName("label_file_path")

        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(160, 490, 181, 41))
        self.pushButton.setStyleSheet("background-color: rgb(53, 132, 228);\n"
                                             "background-color: rgb(98, 160, 234);")
        
        icon_path = os.path.join(os.path.dirname(__file__), "connect-36.png")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(icon_path), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton.setIcon(icon)
        self.pushButton.setObjectName("pushButton")

        self.pushButton1 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton1.setGeometry(QtCore.QRect(380, 490, 177, 41))
        self.pushButton1.setStyleSheet("background-color: rgb(143, 240, 164);")

        icon_path = os.path.join(os.path.dirname(__file__), "send-36.png")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(icon_path), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton1.setIcon(icon)
        self.pushButton1.setObjectName("pushButton1")

        
        self.status_label = QtWidgets.QLabel(self.centralwidget)
        self.status_label.setGeometry(QtCore.QRect(600, 40, 150, 31))
        self.status_label.setAlignment(QtCore.Qt.AlignRight)
        self.status_label.setObjectName("status_label")
        self.update_status_label(connected=False)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 22))
        self.menubar.setObjectName("menubar")
        self.menuLIS_Connection = QtWidgets.QMenu(self.menubar)
        self.menuLIS_Connection.setObjectName("menuLIS_Connection")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menubar.addAction(self.menuLIS_Connection.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        

        MainWindow.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint | QtCore.Qt.WindowCloseButtonHint)


        # Connect buttons to functions
        self.pushButton.clicked.connect(self.start_timer_based_on_interval)
        self.pushButton1.clicked.connect(self.send_data)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label_1.setText(_translate("MainWindow", "Give proper credential to connect with remote server."))
        self.label_url.setText(_translate("MainWindow", "URL"))
        self.label_db.setText(_translate("MainWindow", "Database"))
        self.label_user.setText(_translate("MainWindow", "User"))
        self.label_password.setText(_translate("MainWindow", "Password"))
        self.label_file_path.setText(_translate("MainWindow", "File Path"))
        self.pushButton.setText(_translate("MainWindow", "Save and Connect"))
        self.pushButton1.setText(_translate("MainWindow", "Send Data to server"))
        self.menuLIS_Connection.setTitle(_translate("MainWindow", "Odoo Database Connection"))

    def connect_to_server(self):
        url = self.lineEdit_url.text()
        db = self.lineEdit_db.text()
        username = self.lineEdit_user.text()
        password = self.lineEdit_password.text()

        common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))

        try:
            uid = common.authenticate(db, username, password, {})
            if uid:
                self.show_message("Authentication success")
                self.uid = uid
                self.url = url
                self.db = db
                self.username = username
                self.password = password
                self.update_status_label(connected=True)
            else:
                self.show_message("Authentication Failed", is_error=True)
                self.update_status_label(connected=False)
        except Exception as e:
            self.show_message("Error: " + str(e), is_error=True)
            self.update_status_label(connected=False)



    def send_data(self):
        selected_app = self.comboBox_app_name.currentText()
        if selected_app == "Contact":
            self.send_contact_data()
        elif selected_app == "Attendance":
            self.send_attendance_data()
        else:
            self.show_message("Please select an app from the dropdown.", is_error=True)
    
    # for attendance data
    def send_attendance_data(self):
        file_path = self.lineEdit_file_path.text()
        if not file_path:
            self.show_message("Please provide the path to the Excel file.", is_error=True)
            return

        try:
            # Load Excel file
            df = pd.read_excel(file_path)
            if 'Employee' not in df.columns or 'Check In' not in df.columns or 'Check Out' not in df.columns:
                return  # Exit the function silently if columns are missing

            models = xmlrpc.client.ServerProxy(f'{self.url}/xmlrpc/2/object')

            # Fetch employee names and IDs from Odoo
            employees = models.execute_kw(
                self.db, self.uid, self.password,
                'hr.employee', 'search_read',
                [[['name', 'in', df['Employee'].tolist()]]],
                {'fields': ['id', 'name']}
            )
            
            # Create a mapping of employee names to IDs
            employee_map = {emp['name']: emp['id'] for emp in employees}

            # Loop through the rows in the Excel file
            for _, row in df.iterrows():
                try:
                    employee_name = row['Employee']
                    employee_id = employee_map.get(employee_name)
                    check_in = row['Check In'].strftime("%Y-%m-%d %H:%M:%S")
                    check_out = row['Check Out'].strftime("%Y-%m-%d %H:%M:%S")

                    if not employee_id:
                        continue  # Skip if employee not found

                    # Check if a record with the same employee and check-in/check-out times already exists
                    existing_records = models.execute_kw(
                        self.db, self.uid, self.password,
                        'hr.attendance', 'search',
                        [[['employee_id', '=', employee_id],
                        ['check_in', '=', check_in],
                        ['check_out', '=', check_out]]]
                    )

                    # If no existing record is found, create a new one
                    if not existing_records:
                        data = {
                            'employee_id': employee_id,
                            'check_in': check_in,
                            'check_out': check_out
                        }
                        result = models.execute_kw(self.db, self.uid, self.password, 'hr.attendance', 'create', [data])
                        print(f"Attendance created with ID: {result}")
                    else:
                        print(f"Record for Employee '{employee_name}' at {check_in} already exists. Skipping.")

                except Exception as e:
                    continue  # Skip this row if an error occurs

        except FileNotFoundError:
            return  # Exit silently if file not found
        except Exception as e:
            return  # Exit silently if any other exception occurs


    #For creating contact
    def send_contact_data(self):
        # Get file path from the input field
        file_path = self.lineEdit_file_path.text()
        if not file_path:
            self.show_message("Please provide the path to the Excel file.", is_error=True)
            return

        try:
            # Load the Excel file
            df = pd.read_excel(file_path)
            if 'Name' not in df.columns or 'Phone' not in df.columns:
                self.show_message("The Excel file must contain 'Name' and 'Phone' columns.", is_error=True)
                return

            models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(self.url))

            # Loop through the rows in the Excel file
            for index, row in df.iterrows():
                data = {
                    'name': row['Name'],
                    'phone': str(row['Phone'])
                }
                # Sending the data to create res.partner record
                result = models.execute_kw(self.db, self.uid, self.password, 'res.partner', 'create', [data])
                print(f"Contact created with ID: {result}")

            self.show_message("Contacts created successfully!")
            
        except FileNotFoundError:
            self.show_message("File not found. Please check the file path.", is_error=True)
        except Exception as e:
            self.show_message("Error occurred: " + str(e), is_error=True)

    def show_message(self, message, is_error=False):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical if is_error else QMessageBox.Information)
        msg.setText(message)
        msg.setWindowTitle("Error" if is_error else "Status")
        msg.exec_()

    def update_status_label(self, connected):
        if connected:
            self.status_label.setText('<span style="color: green;">Connected</span>')
        else:
            self.status_label.setText('<span style="color: red;">Disconnected</span>')
    
    def start_timer_based_on_interval(self):
        # Connect to the server first
        self.connect_to_server()

        # Determine the interval based on the selected radio button
        if self.radio_10min.isChecked():
            interval_ms = 6 * 1000  # 10 minutes in milliseconds
            interval_text = "10 minutes"
        elif self.radio_1hr.isChecked():
            interval_ms = 60 * 60 * 1000  # 1 hour in milliseconds
            interval_text = "1 hour"
        elif self.radio_6hr.isChecked():
            interval_ms = 6 * 60 * 60 * 1000  # 6 hours in milliseconds
            interval_text = "6 hours"
        elif self.radio_12hr.isChecked():
            interval_ms = 12 * 60 * 60 * 1000  # 12 hours in milliseconds
            interval_text = "12 hours"
        elif self.radio_24hr.isChecked():
            interval_ms = 24 * 60 * 60 * 1000  # 24 hours in milliseconds
            interval_text = "24 hours"
        else:
            self.show_message("Please select a valid interval.", is_error=True)
            return

        # Set the timer interval and start it
        self.timer.setInterval(interval_ms)
        self.timer.timeout.connect(self.send_data)
        self.timer.start()

        # Initial call to send_data and feedback message to the user
        self.send_data()
        self.show_message(f"Data will be sent automatically every {interval_text}.")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)

    MainWindow.show()
    sys.exit(app.exec_())
