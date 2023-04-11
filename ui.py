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

# Third-party imports
from PyQt5.QtCore import pyqtSlot, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QLabel, QLineEdit, QMessageBox, 
                              QPushButton, QTextEdit, QVBoxLayout, QWidget, QTableWidget, QHeaderView)


# Define the NetworkDriveMapper class
class NetworkDriveMapper(QWidget):

    def __init__(self):
        super().__init__()

        # Configure the window
        self.configure_window()

        # Set the style sheet for the application
        self.set_style_sheet()

        # Create and configure labels, input fields, and buttons
        self.create_widgets()

        # Create and set the layout for the widgets
        self.create_layout()

        # Connect signals and slots for the widgets
        self.connect_signals_and_slots()


    # Method to configure the window's properties
    def configure_window(self):
        self.setWindowTitle("Network Mapping Tool")
        self.setFixedSize(1000, 500)


    # Method to set the style sheet for the application
    def set_style_sheet(self):
        self.setStyleSheet("""
            QWidget {
                font-family: "Segoe UI";
                font-size: 14px;
            }
            QPushButton {
                padding: 5px;
            }
        """)


    # Method to create and configure the widget elements in the window
    def create_widgets(self):
    
        # Create and configure the "Select directory to share" label and input field
        self.adv_shared_path_label = QLabel("Select directory to share:")
        self.adv_shared_path_input = QLineEdit()
        self.adv_shared_path_input.setPlaceholderText("Enter directory path")
        
        # Create and configure the "Browse" button
        self.browse_button = QPushButton("Browse")
        self.browse_button.setIcon(QIcon("browse_icon.png"))

        # Create and configure the "Share/Create" button
        self.add_adv_shared_button = QPushButton("Share/Create")
        self.add_adv_shared_button.setIcon(QIcon("share_icon.png"))

        # Create and configure a separator line
        self.separator_line = QFrame()
        self.separator_line.setFrameShape(QFrame.HLine)
        self.separator_line.setFrameShadow(QFrame.Sunken)

        # Create and configure the "IP address" label and input field
        self.ip_label = QLabel("IP address:")
        self.ip_input = QLineEdit()
        self.ip_input.setPlaceholderText("Enter IP address")

        # Create and configure the "Shared folder" label and input field
        self.shared_label = QLabel("Shared folder:")
        self.shared_input = QLineEdit()
        self.shared_input.setPlaceholderText("Enter shared folder name")
        
        # Create and configure the "Drive letter" label and dropdown
        self.drive_label = QLabel("Drive letter:")
        self.drive_dropdown = QComboBox()
        self.drive_dropdown.addItems(self.get_available_drives())
        
        # Create and configure the "Map network drive" button
        self.connect_button = QPushButton("Map network drive")
        self.connect_button.setIcon(QIcon("connect_icon.png"))

        # Create and configure the "Refresh" button
        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.setIcon(QIcon("refresh_icon.png"))
        self.refresh_button.setToolTip("Refresh drive letters")
        
        # Create and configure the "Stop" button
        self.stop_button = QPushButton("Stop")
        self.stop_button.setIcon(QIcon("stop_icon.png"))
        self.stop_button.setEnabled(False)

        
        #Create and configure the "Log" label
        self.log_label = QLabel("Log:")
        
        # Create and configure the log widget (read-only text area)
        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)
        
        # Create and configure the "Mapped Drives" label
        self.mapped_drives_label = QLabel("Mapped Drives:")

        # Create and configure the mapped drives table
        self.mapped_drives_table = QTableWidget()
        self.mapped_drives_table.setColumnCount(3)
        self.mapped_drives_table.setHorizontalHeaderLabels(["Drive Letter", "Remote Path", "Status"])
        self.mapped_drives_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.mapped_drives_table.verticalHeader().setVisible(False)
        self.mapped_drives_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.mapped_drives_table.setEditTriggers(QTableWidget.NoEditTriggers)

        # Create and configure the "Disconnect" button
        self.disconnect_button = QPushButton("Disconnect")
        self.disconnect_button.setIcon(QIcon("disconnect_icon.png"))

        # Create and configure the "Search" label and input field
        self.search_label = QLabel("Search:")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search for mapped drives")

        # Create and configure the "Help" button
        self.help_button = QPushButton("Help")
        self.help_button.setIcon(QIcon("help_icon.png"))
        
        # Create and configure the "Clear Log" button
        self.clear_log_button = QPushButton("Clear Log")
        self.clear_log_button.setIcon(QIcon("clear_icon.png"))
         
         
    # Create a layout for the widgets   
    def create_layout(self):
        
        layout = QGridLayout()

        # Add widgets to the layout
        layout.addWidget(self.adv_shared_path_label, 0, 0)  # "Select directory to share" label
        layout.addWidget(self.adv_shared_path_input, 0, 1)  # Directory input field
        layout.addWidget(self.browse_button, 0, 2)  # Browse button
        layout.addWidget(self.add_adv_shared_button, 0, 3)  # Share/Create button
        layout.addWidget(self.separator_line, 1, 0, 1, 4)  # Separator line
        layout.addWidget(self.ip_label, 2, 0)  # IP address label
        layout.addWidget(self.ip_input, 2, 1)  # IP address input field
        layout.addWidget(self.shared_label, 3, 0)  # Shared folder label
        layout.addWidget(self.shared_input, 3, 1)  # Shared folder input field
        layout.addWidget(self.drive_label, 4, 0)  # Drive letter label
        layout.addWidget(self.drive_dropdown, 4, 1)  # Drive letter dropdown
        layout.addWidget(self.connect_button, 2, 2, 2, 1)  # Map network drive button
        layout.addWidget(self.refresh_button, 4, 2)  # Refresh button
        layout.addWidget(self.stop_button, 4, 3)  # Stop button
        layout.addWidget(self.log_label, 5, 0)  # Log label
        layout.addWidget(self.log_widget, 6, 0, 1, 4)  # Log widget (text area)
        layout.addWidget(self.mapped_drives_label, 7, 0)  # Mapped drives label
        layout.addWidget(self.mapped_drives_table, 8, 0, 1, 3)  # Mapped drives table
        layout.addWidget(self.disconnect_button, 8, 3)  # Disconnect button
        layout.addWidget(self.search_label, 9, 0)  # Search label
        layout.addWidget(self.clear_log_button, 9, 3)  # Clear Log button

        # Set the layout for the window
        self.setLayout(layout)
    
  
    # Define the connected signals here
    def connect_signals_and_slots(self):
        # Connect the Browse button to a slot
        self.browse_button.clicked.connect(self.browse_folder)

        # Connect the Share/Create button to a slot
        self.add_adv_shared_button.clicked.connect(self.start_share_folder_thread) # Updated

        # Connect the Map Network Drive button to a slot
        self.connect_button.clicked.connect(self.connect_drive_thread) # Updated

        # Connect the Refresh button to a slot
        self.refresh_button.clicked.connect(self.refresh_drive_letters)

        # Connect the Stop button to a slot
        self.stop_button.clicked.connect(self.stop_map_drive_thread) # Added
        
        # Connect the Clear Log button to a slot
        self.clear_log_button.clicked.connect(self.clear_log)
    
    
    def clear_log(self):
        self.log_widget.clear()

    
    #Browse Method
    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Directory")
        self.adv_shared_path_input.setText(folder)
        self.adv_shared_path_input.setFocus()    


    #Get Available Drives 
    def get_available_drives(self):
        used_drives = [d.split()[0] for d in os.popen("net use").readlines() if d.startswith("  ") and ":" in d]
        unavailable_drives = [c + ":" for c in string.ascii_uppercase if os.path.exists(c + ":")]
        available_drives = [c + ":" for c in string.ascii_uppercase if c + ":" not in used_drives and c + ":" not in unavailable_drives]
        return available_drives
        
        
    #Refresh Drive letters in the dropdown list  
    def refresh_drive_letters(self):
        self.drive_dropdown.clear()
        self.drive_dropdown.addItems(self.get_available_drives())
        
        
    # Log a message with a timestamp
    def log_message(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"{timestamp}: {message}"
        self.log_widget.append(formatted_message)
        
        
    #Reset user input fields    
    def reset_fields(self):
        self.adv_shared_path_input.clear()
        self.ip_input.clear()
        self.shared_input.clear()  
 

    # Start the folder sharing thread
    def start_share_folder_thread(self):
        self.add_adv_shared_button.setEnabled(False)
        self.connect_button.setEnabled(False)

        folder_path = self.adv_shared_path_input.text().replace("/", "\\")
        if not re.match(r"^[a-zA-Z]:\\", folder_path):
            self.log_message("Please enter a valid directory path with a root drive.")
            self.share_folder_thread_finished()
            return

        self.share_folder_thread = ShareFolderThread(folder_path)
        self.share_folder_thread.output_signal.connect(self.handle_share_folder_output)
        self.share_folder_thread.finished.connect(self.share_folder_thread_finished)
        self.share_folder_thread.start()
        self.log_message("Sharing folder...")


    # Handle output from the share folder thread
    @pyqtSlot(str)
    def handle_share_folder_output(self, output):
        # Check output for specific messages and handle accordingly
        if "Folder is already shared." in output:
            self.log_widget.append("Folder is already shared.")
        elif "FailedToUpdateACL" in output:
            self.log_widget.append("Failed to update the ACL for full control.")
        elif "FolderCreatedAndShared" in output:
            self.log_widget.append("Folder created and shared successfully.")
        elif "FolderShared" in output:
            self.log_message("Folder has been shared successfully.")
            self.reset_fields()

        else:
            self.log_widget.append(f"Failed to share the folder:\n\n{output}")


    # Function to re-enable buttons after folder sharing thread is finished
    @pyqtSlot() 
    def share_folder_thread_finished(self):
        self.add_adv_shared_button.setEnabled(True)
        self.connect_button.setEnabled(True)


    # Start the drive mapping thread
    def connect_drive_thread(self):
        # Disable buttons during drive mapping process
        self.add_adv_shared_button.setEnabled(False)
        self.connect_button.setEnabled(False)
        
        # Get IP address, shared folder, and drive letter from user input
        ip = self.ip_input.text()
        shared = self.shared_input.text()
        drive_letter = self.drive_dropdown.currentText()

        # Check if user input is valid
        if not ip:
            self.log_message("Please enter an IP address.")
            self.map_drive_thread_finished()
            return

        if not shared:
            self.log_message("Please enter a shared folder.")
            self.map_drive_thread_finished()
            return

        # Initialize the MapDriveThread with IP, shared folder, and drive letter
        self.map_drive_thread = MapDriveThread(ip, shared, drive_letter)
        
        # Connect signals for output handling and thread completion
        self.map_drive_thread.output_signal.connect(self.handle_map_drive_output)
        self.map_drive_thread.finished.connect(self.map_drive_thread_finished)
        
        # Start the drive mapping thread
        self.map_drive_thread.start()
        self.log_message("Mapping network drive...")
        self.map_drive_thread_running = True
        self.stop_button.setEnabled(True)


    # Function to re-enable buttons after drive mapping thread is finished
    @pyqtSlot()
    def map_drive_thread_finished(self):
        self.add_adv_shared_button.setEnabled(True)
        self.connect_button.setEnabled(True)
        self.stop_button.setEnabled(False)


     # Handle output from the map drive thread
    @pyqtSlot(str)
    def handle_map_drive_output(self, output):
        # Check output for specific messages and handle accordingly
        if "The command completed successfully." in output:
            self.log_message("Network drive mapped successfully.")
            self.reset_fields()
            self.refresh_drive_letters()

        else:
            if "System error 85" in output:
                self.log_widget.append("The local device name is already in use.")

            else:
                self.log_widget.append(f"Failed to map network drive: {output}")
    
    
    #Stop the map drive thread
    def stop_map_drive_thread(self):
        if self.map_drive_thread_running:
            self.map_drive_thread.terminate()
            self.map_drive_thread.wait()
            self.map_drive_thread_running = False
            self.log_message("Stopped mapping network drive.")
            self.map_drive_thread_finished()