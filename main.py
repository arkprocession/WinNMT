# Standard library imports
import datetime
import os
import re
import signal
import string
import subprocess
import sys
from subprocess import CREATE_NO_WINDOW
from mapdrivethread import MapDriveThread
from sharefolderthread import ShareFolderThread
from ui import NetworkDriveMapper

# Third-party imports
from PyQt5.QtCore import pyqtSlot, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QLabel, QLineEdit, QMessageBox, 
                              QPushButton, QTextEdit, QVBoxLayout, QWidget, QTableWidget, QHeaderView)





















if __name__ == "__main__":
    app = QApplication(sys.argv)

    main_win = NetworkDriveMapper()
    main_win.show()

    sys.exit(app.exec_())
