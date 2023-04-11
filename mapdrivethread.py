# Standard library imports
import datetime
import os
import re
import signal
import string
import subprocess
import sys
from subprocess import CREATE_NO_WINDOW


# Third-party imports
from PyQt5.QtCore import pyqtSlot, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QLabel, QLineEdit, QMessageBox, 
                              QPushButton, QTextEdit, QVBoxLayout, QWidget, QTableWidget, QHeaderView)
# Define the MapDriveThread class
class MapDriveThread(QThread):
    output_signal = pyqtSignal(str)

    # Constructor for the custom QThread class
    def __init__(self, ip, shared, drive_letter):
        QThread.__init__(self) # Call the parent QThread constructor
        self.ip = ip # Set the ip attribute to the provided ip
        self.shared = shared # Set the shared attribute to the provided shared folder
        self.drive_letter = drive_letter # Set the drive_letter attribute to the provided drive letter

    # The run method is executed when the QThread is started
    def run(self):
        drive_path = "{}{}".format("\\\\{}\\".format(self.ip), self.shared)

        command = 'net use {} "{}" /persistent:yes'.format(self.drive_letter, drive_path)
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=True, creationflags=CREATE_NO_WINDOW)

        output = result.stdout + result.stderr
        self.output_signal.emit(output)
