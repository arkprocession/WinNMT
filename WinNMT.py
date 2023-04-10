import os
import string
import sys
import subprocess
from subprocess import CREATE_NO_WINDOW
import re
import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QFileDialog, QComboBox, QGridLayout, QFrame, QTextEdit
from PyQt5.QtCore import pyqtSlot, QThread, pyqtSignal, Qt
import signal
from PyQt5.QtGui import QIcon

class MapDriveThread(QThread):
    output_signal = pyqtSignal(str)

    def __init__(self, ip, shared, drive_letter):
        QThread.__init__(self)
        self.ip = ip
        self.shared = shared
        self.drive_letter = drive_letter

    def run(self):
        drive_path = "{}{}".format("\\\\{}\\".format(self.ip), self.shared)

        command = 'net use {} "{}" /persistent:yes'.format(self.drive_letter, drive_path)
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=True, creationflags=CREATE_NO_WINDOW)


        output = result.stdout + result.stderr
        self.output_signal.emit(output)


class ShareFolderThread(QThread):
    output_signal = pyqtSignal(str)

    def __init__(self, folder_path):
        QThread.__init__(self)
        self.folder_path = folder_path

    def run(self):
        # PowerShell script to share the folder
        powershell_script = f"""
        $FolderPath = "{self.folder_path}"
        $ShareName = "{os.path.basename(self.folder_path)}"
        $ShareDescription = "{os.path.basename(self.folder_path)} shared folder"

        if (!(Test-Path $FolderPath)) {{
            New-Item -ItemType Directory -Path $FolderPath
        }}

        $ShareName = (Generate-UniqueShareName -ShareName $ShareName)

        New-SmbShare -Name $ShareName -Path $FolderPath -FullAccess "Everyone"

        $Acl = Get-Acl $FolderPath
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
        $Acl.AddAccessRule($AccessRule)
        Set-Acl -Path $FolderPath -AclObject $Acl

        Write-Host "FolderShared"
        """

        # Add the Generate-UniqueShareName function to the PowerShell script
        powershell_script = f"""
        function Generate-UniqueShareName([string]$ShareName) {{
            $ShareNameBase = $ShareName
            $Counter = 1

            while ((Get-SmbShare -ErrorAction SilentlyContinue | Where-Object {{ $_.Name -eq $ShareName }})) {{
                $ShareName = "$ShareNameBase" + "_" + $Counter
                $Counter += 1
            }}

            return $ShareName
        }}

        {powershell_script}
        """

        result = subprocess.run(["powershell", "-Command", powershell_script], capture_output=True, text=True, creationflags=CREATE_NO_WINDOW)


        self.output_signal.emit(result.stdout)

class NetworkDriveMapper(QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Network Mapping Tool")
        self.setFixedSize(1000, 500)

        # Set style sheet for the application
        self.setStyleSheet("""
            QWidget {
                font-family: "Segoe UI";
                font-size: 14px;
            }
            QPushButton {
                padding: 5px;
            }
        """)

        # Create labels and input fields
        self.adv_shared_path_label = QLabel("Select directory to share:")
        self.adv_shared_path_input = QLineEdit()
        self.adv_shared_path_input.setPlaceholderText("Enter directory path")

        self.browse_button = QPushButton("Browse")
        self.browse_button.setIcon(QIcon("browse_icon.png"))

        self.add_adv_shared_button = QPushButton("Share/Create")
        self.add_adv_shared_button.setIcon(QIcon("share_icon.png"))

        self.separator_line = QFrame()
        self.separator_line.setFrameShape(QFrame.HLine)
        self.separator_line.setFrameShadow(QFrame.Sunken)

        self.ip_label = QLabel("IP address:")
        self.ip_input = QLineEdit()
        self.ip_input.setPlaceholderText("Enter IP address")

        self.shared_label = QLabel("Shared folder:")
        self.shared_input = QLineEdit()
        self.shared_input.setPlaceholderText("Enter shared folder name")

        self.drive_label = QLabel("Drive letter:")
        self.drive_dropdown = QComboBox()
        self.drive_dropdown.addItems(self.get_available_drives())

        self.connect_button = QPushButton("Map network drive")
        self.connect_button.setIcon(QIcon("connect_icon.png"))

        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.setIcon(QIcon("refresh_icon.png"))
        self.refresh_button.setToolTip("Refresh drive letters")

        # Create log widget
        self.log_label = QLabel("Log:")
        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)

        # Add the Stop button
        self.stop_button = QPushButton("Stop")
        self.stop_button.setIcon(QIcon("stop_icon.png"))
        self.stop_button.setEnabled(False)

        # Create layout
        layout = QGridLayout()
        layout.addWidget(self.adv_shared_path_label, 0, 0)
        layout.addWidget(self.adv_shared_path_input, 0, 1, 1, 2)
        layout.addWidget(self.browse_button, 0, 3)
        layout.addWidget(self.add_adv_shared_button, 0, 4)

        layout.addWidget(self.separator_line, 1, 0, 1, 5)

        layout.addWidget(self.ip_label, 2, 0)
        layout.addWidget(self.ip_input, 2, 1, 1, 2)

        layout.addWidget(self.shared_label, 3, 0)
        layout.addWidget(self.shared_input, 3, 1, 1, 2)

        layout.addWidget(self.drive_label, 4, 0)
        layout.addWidget(self.drive_dropdown, 4, 1)
        layout.addWidget(self.refresh_button, 4, 2)

        layout.addWidget(self.connect_button, 5, 1)  # Move "Map network drive" button below dropdown
        layout.addWidget(self.stop_button, 5, 2)  # Move "Stop" button below refresh button

        layout.addWidget(self.log_label, 6, 0)
        layout.addWidget(self.log_widget, 7, 0, 1, 5)

        self.setLayout(layout)
  
        # Connect signals and slots
        self.browse_button.clicked.connect(self.browse_folder)
        self.add_adv_shared_button.clicked.connect(self.start_share_folder_thread)
        self.connect_button.clicked.connect(self.connect_drive_thread)
        self.stop_button.clicked.connect(self.stop_map_drive_thread)
        self.refresh_button.clicked.connect(self.refresh_drive_letters)  # Connect the Refresh button to the refresh_drive_letters function
    
    def log_message(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"{timestamp}: {message}"
        self.log_widget.append(formatted_message)

    

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Directory")

    def get_available_drives(self):
        used_drives = [d.split()[0] for d in os.popen("net use").readlines() if d.startswith("  ") and ":" in d]
        unavailable_drives = [c + ":" for c in string.ascii_uppercase if os.path.exists(c + ":")]
        available_drives = [c + ":" for c in string.ascii_uppercase if c + ":" not in used_drives and c + ":" not in unavailable_drives]
        return available_drives
        
    def refresh_drive_letters(self):
        self.drive_dropdown.clear()
        self.drive_dropdown.addItems(self.get_available_drives())
        
        
        
    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Directory")
        self.adv_shared_path_input.setText(folder)
        self.adv_shared_path_input.setFocus()    


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

        
        
    @pyqtSlot() 
    def share_folder_thread_finished(self):
        self.add_adv_shared_button.setEnabled(True)
        self.connect_button.setEnabled(True)

    def connect_drive_thread(self):
        self.add_adv_shared_button.setEnabled(False)
        self.connect_button.setEnabled(False)

        ip = self.ip_input.text()
        shared = self.shared_input.text()
        drive_letter = self.drive_dropdown.currentText()

        if not ip:
            self.log_message("Please enter an IP address.")
            self.map_drive_thread_finished()
            return

        if not shared:
            self.log_message("Please enter a shared folder.")
            self.map_drive_thread_finished()
            return

        self.map_drive_thread = MapDriveThread(ip, shared, drive_letter)
        self.map_drive_thread.output_signal.connect(self.handle_map_drive_output)
        self.map_drive_thread.finished.connect(self.map_drive_thread_finished)
        self.map_drive_thread.start()
        self.log_message("Mapping network drive...")
        self.map_drive_thread_running = True
        self.stop_button.setEnabled(True)


    @pyqtSlot()
    def map_drive_thread_finished(self):
        self.add_adv_shared_button.setEnabled(True)
        self.connect_button.setEnabled(True)
        self.stop_button.setEnabled(False)

    @pyqtSlot(str)
    def handle_share_folder_output(self, output):
        if "FolderAlreadyShared" in output:
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
           

    @pyqtSlot(str)
    def handle_map_drive_output(self, output):
        if "The command completed successfully." in output:
            self.log_message("Network drive mapped successfully.")
            self.reset_fields()
            self.refresh_drive_letters()

        else:
            if "System error 85" in output:
                self.log_widget.append("The local device name is already in use.")

            else:
                self.log_widget.append(f"Failed to map network drive: {output}")

    def reset_fields(self):
        self.adv_shared_path_input.clear()
        self.ip_input.clear()
        self.shared_input.clear()
    
        
    def stop_map_drive_thread(self):
        if self.map_drive_thread_running:
            self.map_drive_thread.terminate()
            self.map_drive_thread.wait()
            self.map_drive_thread_running = False
            self.log_message("Stopped mapping network drive.")
            self.map_drive_thread_finished()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    main_win = NetworkDriveMapper()
    main_win.show()

    sys.exit(app.exec_())
