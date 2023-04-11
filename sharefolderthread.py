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

# Define the ShareFolderThread class
class ShareFolderThread(QThread):
    output_signal = pyqtSignal(str)

    # Constructor for the custom QThread class
    def __init__(self, folder_path):
        QThread.__init__(self)  # Call the parent QThread constructor
        self.folder_path = folder_path  # Set the folder_path attribute to the provided folder path

    # The run method is executed when the QThread is started
    def run(self):
        # PowerShell script to share the folder
        powershell_script = f"""
        $FolderPath = "{self.folder_path}"
        $ShareName = "{os.path.basename(self.folder_path)}"
        $ShareDescription = "{os.path.basename(self.folder_path)} shared folder"

        $ExistingShare = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object {{ $_.Path -eq $FolderPath }}

        if ($ExistingShare) {{
            Write-Host "Folder is already shared."
        }} else {{
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
        }}

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
