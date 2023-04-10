# WinNMT (Windows Advanced Share & Network Mapping Tool) (GUI)

WinNMT (Windows Advanced Share & Network Mapping Tool) (GUI)


This is a Python script that provides a Graphical User Interface (GUI) for mapping network drives and sharing folders on Windows operating system.

The script uses the PyQt5 library to create a user-friendly interface with input fields for the IP address, shared folder name, and drive letter. It also includes buttons for browsing for a folder, sharing a folder, mapping a network drive, and refreshing the available drive letters.

The mapping of network drives is executed using the net use command, while the sharing of folders is performed using PowerShell. The script also includes a function to retrieve the available drive letters on the system by parsing the output of the net use command.

The GUI logs the output of the operations performed, making it easy to track any issues that may arise. The script includes methods for handling user input, starting and stopping threads, and resetting input fields after an operation has been executed.

Overall, this script offers a convenient and efficient way to map network drives and share folders on Windows operating system.
