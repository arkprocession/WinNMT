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
import win32net

# Third-party imports
from PyQt5.QtCore import pyqtSlot, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QLabel, QLineEdit, QMessageBox, 
                              QPushButton, QTextEdit, QVBoxLayout, QWidget, QTableWidget, QHeaderView, QTableWidgetItem, QGroupBox)
                              



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

        # Call the retrieve_mapped_drives and retrieve_shared_folders methods on program start
        self.retrieve_shared_folders()
        self.retrieve_mapped_drives()

    # Method to configure the window's properties
    def configure_window(self):
        self.setWindowTitle("Network Mapping Tool")
        self.setFixedSize(1000, 1000)


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
        self.drive_label = QLabel("Name:")
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
        
        # Create and configure the log widget (read-only text area)
        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)

        # Create and configure the mapped drives table
        self.shared_drives_table = QTableWidget()
        self.shared_drives_table.setColumnCount(2)
        self.shared_drives_table.setHorizontalHeaderLabels(["Name", "Remote Path"])
        self.shared_drives_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.shared_drives_table.verticalHeader().setVisible(False)
        self.shared_drives_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.shared_drives_table.setEditTriggers(QTableWidget.NoEditTriggers)

        # Create and configure the "Disconnect" button
        self.disconnect_button = QPushButton("Unshare")
        self.disconnect_button.setIcon(QIcon("disconnect_icon.png"))
        
        # Create and configure the "Clear Log" button
        self.clear_log_button = QPushButton("Clear Log")
        self.clear_log_button.setIcon(QIcon("clear_icon.png"))
        
        self.retrieve_shared_button = QPushButton("Refresh Shared Folders")
        
        
          # Create and configure the mapped drives table
        self.mapped_drives_table = QTableWidget()
        self.mapped_drives_table.setColumnCount(2)
        self.mapped_drives_table.setHorizontalHeaderLabels(["Name", "Remote Path"])
        self.mapped_drives_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.mapped_drives_table.verticalHeader().setVisible(False)
        self.mapped_drives_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.mapped_drives_table.setEditTriggers(QTableWidget.NoEditTriggers)


        # Create and configure the "Retrieve Mapped Drives" button
        self.retrieve_mapped_drives_button = QPushButton("Refresh Mapped Drives")
        
        # Create and configure the "Disconnect Mapped Drive" button
        self.disconnect_mapped_drive_button = QPushButton("Unmap")
         
         
    def create_layout(self):

        layout = QVBoxLayout()

        # Share folder section
        share_folder_layout = QGridLayout()
        share_folder_group = QGroupBox("Share Folder")
        share_folder_layout.addWidget(self.adv_shared_path_label, 0, 0)
        share_folder_layout.addWidget(self.adv_shared_path_input, 0, 1)
        share_folder_layout.addWidget(self.browse_button, 0, 2)
        share_folder_layout.addWidget(self.add_adv_shared_button, 0, 3)
        share_folder_group.setLayout(share_folder_layout)
        layout.addWidget(share_folder_group)

        # Map network drive section
        map_network_drive_layout = QGridLayout()
        map_network_drive_group = QGroupBox("Map Network Drive")
        map_network_drive_layout.addWidget(self.ip_label, 0, 0)
        map_network_drive_layout.addWidget(self.ip_input, 0, 1)
        map_network_drive_layout.addWidget(self.shared_input, 1, 1)
        map_network_drive_layout.addWidget(self.drive_label, 2, 0)
        map_network_drive_layout.addWidget(self.drive_dropdown, 2, 1)
        map_network_drive_layout.addWidget(self.connect_button, 0, 2, 2, 1)
        map_network_drive_layout.addWidget(self.refresh_button, 2, 2)
        map_network_drive_layout.addWidget(self.stop_button, 2, 3)
        map_network_drive_group.setLayout(map_network_drive_layout)
        layout.addWidget(map_network_drive_group)



        # Shared folder section
        shared_drives_layout = QVBoxLayout()
        shared_drives_group = QGroupBox("Shared Folders")
        shared_drives_layout.addWidget(self.shared_drives_table)
        shared_drives_layout.addWidget(self.retrieve_shared_button)
        shared_drives_layout.addWidget(self.disconnect_button)
        shared_drives_group.setLayout(shared_drives_layout)
        layout.addWidget(shared_drives_group)

        # Mapped drives section
        mapped_drives_layout = QVBoxLayout()
        mapped_drives_group = QGroupBox("Mapped Drives")
        mapped_drives_layout.addWidget(self.mapped_drives_table)
        mapped_drives_layout.addWidget(self.retrieve_mapped_drives_button)
        mapped_drives_layout.addWidget(self.disconnect_mapped_drive_button)
        mapped_drives_group.setLayout(mapped_drives_layout)
        layout.addWidget(mapped_drives_group)

        # Log section
        log_layout = QVBoxLayout()
        log_group = QGroupBox("Log")
        log_layout.addWidget(self.log_widget)
        log_layout.addWidget(self.clear_log_button)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        
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
        
        self.retrieve_shared_button.clicked.connect(self.retrieve_shared_folders)

        self.disconnect_button.clicked.connect(self.on_disconnect_button_clicked)
        
        
        # Connect the Retrieve Mapped Drives button to a slot
        self.retrieve_mapped_drives_button.clicked.connect(self.retrieve_mapped_drives)
        # Connect the Disconnect Mapped Drive button to a slot
        self.disconnect_mapped_drive_button.clicked.connect(self.on_disconnect_mapped_drive_button_clicked)
    
    
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
    # Call the retrieve_shared_folders method to update the table
        self.retrieve_shared_folders()

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
        
        # Call the retrieve_mapped_drives method to update the table
        self.retrieve_mapped_drives()        
    
    
    #Stop the map drive thread
    def stop_map_drive_thread(self):
        if self.map_drive_thread_running:
            self.map_drive_thread.terminate()
            self.map_drive_thread.wait()
            self.map_drive_thread_running = False
            self.log_message("Stopped mapping network drive.")
            self.map_drive_thread_finished()
            
            
    def retrieve_shared_folders(self):
        # Retrieve shared folder information
        server = None  # Use the local computer
        level = 2  # Level 2 contains the share name, type, and other details
        resume_handle = 0

        shared_folders, _, _ = win32net.NetShareEnum(server, level, resume_handle)

        # Filter out IPC$ shares
        shared_folders = [folder for folder in shared_folders if folder["netname"] != "IPC$"]

        if not shared_folders:
            self.log_message("No shared folders found.")
            self.shared_drives_table.setRowCount(0)  # Clear the table
            return

        self.shared_drives_table.setRowCount(len(shared_folders))

        for index, folder in enumerate(shared_folders):
            share_name, local_path = folder["netname"], folder["path"]

            self.shared_drives_table.setItem(index, 0, QTableWidgetItem(share_name))
            self.shared_drives_table.setItem(index, 1, QTableWidgetItem(local_path))
            self.shared_drives_table.setItem(index, 2, QTableWidgetItem("Shared"))

                
                
    @pyqtSlot()
    def on_disconnect_button_clicked(self):
        selected_rows = self.shared_drives_table.selectionModel().selectedRows()
        if not selected_rows:
            self.log_message("No shared folder selected for disconnecting.")
            return

        for row in selected_rows:
            shared_folder = self.shared_drives_table.item(row.row(), 0).text()
            result = subprocess.run(f"net share {shared_folder} /delete", shell=True, text=True, capture_output=True)
            if result.returncode == 0:
                self.log_message(f"Shared folder '{shared_folder}' has been disconnected.")
                self.shared_drives_table.removeRow(row.row())
            else:
                self.log_message(f"Failed to disconnect shared folder '{shared_folder}': {result.stderr}")
                
                
     # Method to retrieve mapped network drives
    def retrieve_mapped_drives(self):
        server = None  # Use the local computer
        level = 1  # Level 1 contains the local and remote path

        mapped_drives, _, _ = win32net.NetUseEnum(server, level)

        if not mapped_drives:
            self.log_message("No mapped drives found.")
            self.mapped_drives_table.setRowCount(0)  # Clear the table
            return

        self.mapped_drives_table.setRowCount(len(mapped_drives))

        for index, drive in enumerate(mapped_drives):
            local_name, remote_name = drive["local"], drive["remote"]

            self.mapped_drives_table.setItem(index, 0, QTableWidgetItem(local_name))
            self.mapped_drives_table.setItem(index, 1, QTableWidgetItem(remote_name))

                
    @pyqtSlot()
    def on_disconnect_mapped_drive_button_clicked(self):
        selected_rows = self.mapped_drives_table.selectionModel().selectedRows()
        if not selected_rows:
            self.log_message("No mapped drive selected for disconnecting.")
            return

        for row in selected_rows:
            mapped_drive = self.mapped_drives_table.item(row.row(), 0).text()
            result = subprocess.run(f"net use {mapped_drive} /delete", shell=True, text=True, capture_output=True)
            if result.returncode == 0:
                self.log_message(f"Mapped drive '{mapped_drive}' has been disconnected.")
                self.mapped_drives_table.removeRow(row.row())
            else:
                self.log_message(f"Failed to disconnect mapped drive '{mapped_drive}': {result.stderr}")

                
    def closeEvent(self, event):
        confirm_box = QMessageBox()
        confirm_box.setWindowTitle("Confirm Exit")
        confirm_box.setText("Are you sure you want to exit?")
        confirm_box.setIcon(QMessageBox.Question)
        confirm_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        confirm_box.setDefaultButton(QMessageBox.No)
        reply = confirm_box.exec_()
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
