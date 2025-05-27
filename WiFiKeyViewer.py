# The code was originally written by Reza Torabi and later developed further by Phoenix Marie.
import ctypes
import sys
import subprocess
import os
import qrcode
import shutil # For high-level file operations (e.g., copying, archiving)
import pandas as pd # For data manipulation and analysis
import openpyxl # For reading and writing Excel files (.xlsx)
from datetime import datetime # For timestamping backup files
import xml.etree.ElementTree as ET # For parsing XML Wi-Fi profiles
import re # For regex in password strength

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton, QLabel, QTableWidget,
    QTableWidgetItem, QLineEdit, QMessageBox, QFileDialog, QWidget, QHBoxLayout,
    QMenu, QAction, QStatusBar, QTabWidget, QInputDialog, QProgressDialog # Added QTabWidget, QInputDialog, QProgressDialog
)
from PyQt5.QtGui import QClipboard, QPixmap, QIcon
from PyQt5.QtCore import Qt, QPoint, QTimer, QThread, pyqtSignal # Added QTimer, QThread, pyqtSignal for background tasks


# Worker thread for long-running operations to keep UI responsive
class Worker(QThread):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            result = self.func(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class WifiProfileExtractor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Wi-Fi Profile Management App")
        self.setGeometry(100, 100, 1300, 800) # Adjusted window size for new tabs and buttons

        self.init_ui()
        self.check_admin_privileges()

        # Initialize status bar
        self.statusBar = self.statusBar()
        self.statusBar.showMessage("Ready")

        # Setup timer for auto-refresh (optional, can be configured by user)
        self.auto_refresh_timer = QTimer(self)
        self.auto_refresh_timer.timeout.connect(self.extract_wifi_profiles)
        # self.auto_refresh_timer.start(300000) # Auto-refresh every 5 minutes (300000 ms)

    # UI Initialization
    def init_ui(self):
        """Initializes the user interface elements with tabs."""
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout()

        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # Tab 1: Saved Profiles
        self.saved_profiles_tab = QWidget()
        self.setup_saved_profiles_tab(self.saved_profiles_tab)
        self.tab_widget.addTab(self.saved_profiles_tab, "Saved Profiles")

        # Tab 2: Available Networks
        self.available_networks_tab = QWidget()
        self.setup_available_networks_tab(self.available_networks_tab)
        self.tab_widget.addTab(self.available_networks_tab, "Available Networks")

        self.central_widget.setLayout(main_layout)

    def setup_saved_profiles_tab(self, tab_widget):
        """Sets up the UI for the Saved Profiles tab."""
        layout = QVBoxLayout(tab_widget)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search Wi-Fi Profile")
        self.search_box.textChanged.connect(self.search_profiles)
        layout.addWidget(self.search_box)

        self.table = QTableWidget(0, 7) # Added a column for Password Strength
        self.table.setHorizontalHeaderLabels(
            ["Profile Name", "Password", "Authentication", "Encryption", "SSID Visibility", "Status", "Strength"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        layout.addWidget(self.table)

        self.setup_buttons(layout)

        # Connect table selection change to update button states
        self.table.itemSelectionChanged.connect(self.update_button_states)

    def setup_available_networks_tab(self, tab_widget):
        """Sets up the UI for the Available Networks tab."""
        layout = QVBoxLayout(tab_widget)

        self.available_networks_table = QTableWidget(0, 5)
        self.available_networks_table.setHorizontalHeaderLabels(
            ["SSID", "Signal (dBm)", "Authentication", "Encryption", "BSSID"]
        )
        self.available_networks_table.horizontalHeader().setStretchLastSection(True)
        self.available_networks_table.verticalHeader().setVisible(False)
        self.available_networks_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.available_networks_table.setSelectionMode(QTableWidget.SingleSelection)
        layout.addWidget(self.available_networks_table)

        scan_button_layout = QHBoxLayout()
        self.scan_networks_button = QPushButton("Scan Available Networks")
        self.scan_networks_button.setIcon(QIcon(":/icons/scan.png")) # Placeholder icon
        self.scan_networks_button.clicked.connect(self.scan_available_networks)
        scan_button_layout.addWidget(self.scan_networks_button)

        self.connect_available_button = QPushButton("Connect to Selected Network")
        self.connect_available_button.setIcon(QIcon(":/icons/connect.png")) # Placeholder icon
        self.connect_available_button.clicked.connect(self.connect_to_available_network)
        self.connect_available_button.setEnabled(False) # Disabled until selection
        scan_button_layout.addWidget(self.connect_available_button)

        layout.addLayout(scan_button_layout)

        self.available_networks_table.itemSelectionChanged.connect(
            lambda: self.connect_available_button.setEnabled(self.available_networks_table.currentRow() != -1)
        )


    def setup_buttons(self, layout):
        """Sets up the action buttons across multiple rows."""
        button_layout_row1 = QHBoxLayout()
        button_layout_row2 = QHBoxLayout()
        button_layout_row3 = QHBoxLayout()

        # Row 1: Extraction, Refresh, Export
        self.extract_button = QPushButton("Extract Wi-Fi Profiles")
        self.extract_button.setIcon(QIcon(":/icons/extract.png"))
        self.extract_button.clicked.connect(self.extract_wifi_profiles)
        button_layout_row1.addWidget(self.extract_button)

        self.refresh_button = QPushButton("Refresh Table")
        self.refresh_button.setIcon(QIcon(":/icons/refresh.png"))
        self.refresh_button.clicked.connect(self.extract_wifi_profiles)
        button_layout_row1.addWidget(self.refresh_button)

        self.export_txt_button = QPushButton("Export to Text")
        self.export_txt_button.setIcon(QIcon(":/icons/save_txt.png"))
        self.export_txt_button.clicked.connect(self.export_to_file)
        self.export_txt_button.setEnabled(False)
        button_layout_row1.addWidget(self.export_txt_button)

        self.export_excel_button = QPushButton("Export to Excel")
        self.export_excel_button.setIcon(QIcon(":/icons/save_excel.png"))
        self.export_excel_button.clicked.connect(self.export_to_excel)
        self.export_excel_button.setEnabled(False)
        button_layout_row1.addWidget(self.export_excel_button)

        self.export_xml_button = QPushButton("Export Selected to XML") # New: Export Selected to XML
        self.export_xml_button.setIcon(QIcon(":/icons/save_xml.png"))
        self.export_xml_button.clicked.connect(self.export_selected_profile_to_xml)
        self.export_xml_button.setEnabled(False)
        button_layout_row1.addWidget(self.export_xml_button)

        self.import_xml_button = QPushButton("Import Profile from XML") # New: Import Profile from XML
        self.import_xml_button.setIcon(QIcon(":/icons/import_xml.png"))
        self.import_xml_button.clicked.connect(self.import_profile_from_xml)
        button_layout_row1.addWidget(self.import_xml_button)


        # Row 2: Copy, QR, Connect, Disconnect, Delete, Rename
        self.copy_button = QPushButton("Copy Selected")
        self.copy_button.setIcon(QIcon(":/icons/copy.png"))
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        self.copy_button.setEnabled(False)
        button_layout_row2.addWidget(self.copy_button)

        self.qr_button = QPushButton("Generate QR Code")
        self.qr_button.setIcon(QIcon(":/icons/qr.png"))
        self.qr_button.clicked.connect(self.generate_qr_code)
        self.qr_button.setEnabled(False)
        button_layout_row2.addWidget(self.qr_button)

        self.connect_button = QPushButton("Connect to Profile")
        self.connect_button.setIcon(QIcon(":/icons/connect.png"))
        self.connect_button.clicked.connect(self.connect_to_profile)
        self.connect_button.setEnabled(False)
        button_layout_row2.addWidget(self.connect_button)

        self.disconnect_button = QPushButton("Disconnect Current")
        self.disconnect_button.setIcon(QIcon(":/icons/disconnect.png"))
        self.disconnect_button.clicked.connect(self.disconnect_current_connection)
        self.disconnect_button.setEnabled(True)
        button_layout_row2.addWidget(self.disconnect_button)

        self.delete_button = QPushButton("Delete Profile")
        self.delete_button.setIcon(QIcon(":/icons/delete.png"))
        self.delete_button.clicked.connect(self.delete_selected_profile)
        self.delete_button.setEnabled(False)
        button_layout_row2.addWidget(self.delete_button)

        self.rename_button = QPushButton("Rename Profile") # New: Rename Profile
        self.rename_button.setIcon(QIcon(":/icons/rename.png"))
        self.rename_button.clicked.connect(self.rename_selected_profile)
        self.rename_button.setEnabled(False)
        button_layout_row2.addWidget(self.rename_button)


        # Row 3: Utility, About, Dark Mode
        self.backup_logs_button = QPushButton("Backup Logs")
        self.backup_logs_button.setIcon(QIcon(":/icons/backup.png"))
        self.backup_logs_button.clicked.connect(self.backup_logs)
        self.backup_logs_button.setEnabled(False)
        button_layout_row3.addWidget(self.backup_logs_button)

        self.show_summary_button = QPushButton("Show Summary")
        self.show_summary_button.setIcon(QIcon(":/icons/summary.png"))
        self.show_summary_button.clicked.connect(self.show_summary)
        self.show_summary_button.setEnabled(False)
        button_layout_row3.addWidget(self.show_summary_button)

        self.about_button = QPushButton("About")
        self.about_button.setIcon(QIcon(":/icons/about.png"))
        self.about_button.clicked.connect(self.show_about)
        button_layout_row3.addWidget(self.about_button)

        self.dark_mode_button = QPushButton("Toggle Dark Mode")
        self.dark_mode_button.setIcon(QIcon(":/icons/dark_mode.png"))
        self.dark_mode_button.clicked.connect(self.toggle_dark_mode)
        button_layout_row3.addWidget(self.dark_mode_button)

        layout.addLayout(button_layout_row1)
        layout.addLayout(button_layout_row2)
        layout.addLayout(button_layout_row3)

    def update_button_states(self):
        """Updates the enabled/disabled state of buttons based on table selection and data presence."""
        has_selection = self.table.currentRow() != -1
        self.copy_button.setEnabled(has_selection)
        self.qr_button.setEnabled(has_selection)
        self.connect_button.setEnabled(has_selection)
        self.delete_button.setEnabled(has_selection)
        self.rename_button.setEnabled(has_selection) # Enable rename button

        has_data = self.table.rowCount() > 0
        self.export_txt_button.setEnabled(has_data)
        self.export_excel_button.setEnabled(has_data)
        self.export_xml_button.setEnabled(has_selection) # Export XML only for selected
        self.show_summary_button.setEnabled(has_data)
        self.backup_logs_button.setEnabled(os.path.exists("wifi_profiles_log.txt"))


    # Utility and Logging
    def log_error(self, message):
        """Logs an error message to a file with a timestamp and displays it in the status bar."""
        error_log_file = "app_error_log.txt"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(error_log_file, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] ERROR: {message}\n")
        self.statusBar.showMessage(f"Error: {message}", 5000)

    def check_admin_privileges(self):
        """Checks if the application has administrative privileges."""
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if not is_admin:
                QMessageBox.warning(
                    self,
                    "Warning",
                    "This application may require administrative privileges to extract all Wi-Fi passwords and perform some network operations (connect/disconnect/delete/rename).",
                )
                self.statusBar.showMessage("Admin privileges recommended for full functionality.", 5000)
        except Exception as e:
            self.log_error(f"Could not check admin privileges: {e}")
            self.statusBar.showMessage("Admin privilege check failed.", 3000)

    # --- Core Wi-Fi Profile Management ---
    def extract_wifi_profiles(self):
        """Extracts Wi-Fi profiles and their details using netsh commands."""
        self.statusBar.showMessage("Extracting Wi-Fi profiles...", 0)
        self.progress_dialog = QProgressDialog("Extracting profiles...", "Cancel", 0, 0, self)
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setCancelButton(None) # No cancel button for this process
        self.progress_dialog.show()

        worker = Worker(self._extract_profiles_background)
        worker.finished.connect(self._handle_extraction_finished)
        worker.error.connect(self._handle_extraction_error)
        worker.start()

    def _extract_profiles_background(self):
        """Background task for extracting profiles."""
        result = subprocess.check_output(
            ["netsh", "wlan", "show", "profiles"], encoding="utf-8"
        )
        profiles = [
            line.split(":")[1].strip()
            for line in result.split("\n") if "All User Profile" in line
        ]

        extracted_data = []
        for profile in profiles:
            details = self.get_profile_details(profile)
            extracted_data.append((profile, details))
        return extracted_data

    def _handle_extraction_finished(self, extracted_data):
        """Handles the result of background profile extraction."""
        self.table.setRowCount(0)
        for profile, details in extracted_data:
            self.add_table_row(profile, details)

        self.progress_dialog.close()
        self.update_button_states()
        self.auto_save_logs()
        self.statusBar.showMessage(f"Extracted {len(extracted_data)} Wi-Fi profiles.", 3000)

    def _handle_extraction_error(self, error_message):
        """Handles errors from background profile extraction."""
        self.progress_dialog.close()
        QMessageBox.critical(self, "Error", f"Failed to extract profiles: {error_message}")
        self.log_error(f"Failed to extract profiles: {error_message}")
        self.statusBar.showMessage("Failed to extract profiles.", 5000)


    def get_profile_details(self, profile_name):
        """Retrieves detailed information for a given Wi-Fi profile, including password strength."""
        details = {
            "password": "No Password",
            "auth": "N/A",
            "encryption": "N/A",
            "ssid_visibility": "N/A",
            "status": "Disconnected",
            "strength": "N/A" # New field for password strength
        }
        try:
            result = subprocess.check_output(
                ["netsh", "wlan", "show", "profile", profile_name, "key=clear"],
                encoding="utf-8"
            )
            for line in result.split("\n"):
                if "Key Content" in line:
                    password = line.split(":")[1].strip()
                    details["password"] = password
                    details["strength"] = self.check_password_strength(password) # Calculate strength
                elif "Authentication" in line:
                    details["auth"] = line.split(":")[1].strip()
                elif "Cipher" in line:
                    details["encryption"] = line.split(":")[1].strip()
                elif "SSID name" in line and "broadcast" in line:
                    details["ssid_visibility"] = line.split(":")[1].strip()

            connected_result = subprocess.check_output(
                ["netsh", "wlan", "show", "interfaces"], encoding="utf-8"
            )
            if profile_name in connected_result:
                details["status"] = "Connected"
        except subprocess.CalledProcessError:
            details["password"] = "Access Denied (Run as Administrator)"
            details["strength"] = "N/A" # Cannot determine strength without access
        except Exception as e:
            self.log_error(f"Error getting details for {profile_name}: {e}")
        return details

    def add_table_row(self, profile_name, details):
        """Adds a new row to the table with profile information."""
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        for i, value in enumerate([
            profile_name, details["password"], details["auth"],
            details["encryption"], details["ssid_visibility"], details["status"],
            details["strength"] # Add strength to the table
        ]):
            self.table.setItem(row_position, i, QTableWidgetItem(value))

    def check_password_strength(self, password):
        """Evaluates password strength (basic implementation)."""
        if password == "No Password" or "Access Denied" in password:
            return "N/A"
        if not password:
            return "Empty"

        score = 0
        if len(password) >= 8:
            score += 1
        if re.search(r"[a-z]", password):
            score += 1
        if re.search(r"[A-Z]", password):
            score += 1
        if re.search(r"\d", password):
            score += 1
        if re.search(r"[!@#$%^&*(),.?\":{}|<>]", password):
            score += 1

        if score == 5:
            return "Very Strong"
        elif score == 4:
            return "Strong"
        elif score == 3:
            return "Moderate"
        elif score == 2:
            return "Weak"
        else:
            return "Very Weak"

    # --- Network Operations ---
    def scan_available_networks(self):
        """Scans for available Wi-Fi networks and displays them."""
        self.statusBar.showMessage("Scanning for available networks...", 0)
        self.progress_dialog = QProgressDialog("Scanning networks...", "Cancel", 0, 0, self)
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()

        worker = Worker(self._scan_networks_background)
        worker.finished.connect(self._handle_scan_finished)
        worker.error.connect(self._handle_scan_error)
        worker.start()

    def _scan_networks_background(self):
        """Background task for scanning available networks."""
        output = subprocess.check_output(["netsh", "wlan", "show", "networks", "mode=bssid"], encoding="utf-8", stderr=subprocess.STDOUT)
        networks = []
        current_network = {}
        for line in output.splitlines():
            line = line.strip()
            if line.startswith("SSID"):
                if current_network: # Save previous network if exists
                    networks.append(current_network)
                current_network = {"SSID": line.split(":", 1)[1].strip().strip('"')}
            elif "Signal" in line:
                current_network["Signal"] = line.split(":", 1)[1].strip()
            elif "Authentication" in line:
                current_network["Authentication"] = line.split(":", 1)[1].strip()
            elif "Encryption" in line:
                current_network["Encryption"] = line.split(":", 1)[1].strip()
            elif "BSSID" in line:
                current_network["BSSID"] = line.split(":", 1)[1].strip()
        if current_network: # Add the last network
            networks.append(current_network)
        return networks

    def _handle_scan_finished(self, networks_data):
        """Handles the result of background network scan."""
        self.available_networks_table.setRowCount(0)
        for net in networks_data:
            row_position = self.available_networks_table.rowCount()
            self.available_networks_table.insertRow(row_position)
            self.available_networks_table.setItem(row_position, 0, QTableWidgetItem(net.get("SSID", "N/A")))
            self.available_networks_table.setItem(row_position, 1, QTableWidgetItem(net.get("Signal", "N/A")))
            self.available_networks_table.setItem(row_position, 2, QTableWidgetItem(net.get("Authentication", "N/A")))
            self.available_networks_table.setItem(row_position, 3, QTableWidgetItem(net.get("Encryption", "N/A")))
            self.available_networks_table.setItem(row_position, 4, QTableWidgetItem(net.get("BSSID", "N/A")))
        self.progress_dialog.close()
        self.statusBar.showMessage(f"Found {len(networks_data)} available networks.", 3000)

    def _handle_scan_error(self, error_message):
        """Handles errors from background network scan."""
        self.progress_dialog.close()
        QMessageBox.critical(self, "Error Scanning Networks", f"Failed to scan networks: {error_message}")
        self.log_error(f"Failed to scan networks: {error_message}")
        self.statusBar.showMessage("Failed to scan networks.", 5000)


    def connect_to_available_network(self):
        """Attempts to connect to a selected available Wi-Fi network."""
        selected_row = self.available_networks_table.currentRow()
        if selected_row != -1:
            ssid = self.available_networks_table.item(selected_row, 0).text()
            password, ok = QInputDialog.getText(self, "Connect to Network", f"Enter password for '{ssid}':", QLineEdit.Password)

            if ok and password is not None:
                self.statusBar.showMessage(f"Attempting to connect to '{ssid}'...", 0)
                try:
                    # Create a temporary profile XML
                    profile_xml = f"""<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>{ssid}</name>
    <SSIDConfig>
        <SSID>
            <hex>{ssid.encode('utf-8').hex()}</hex>
            <name>{ssid}</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>WPA2PSK</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>{password}</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
    <MacRandomization>
        <enableRandomization>false</enableRandomization>
    </MacRandomization>
</WLANProfile>"""
                    temp_xml_path = f"{ssid}_temp_profile.xml"
                    with open(temp_xml_path, "w", encoding="utf-8") as f:
                        f.write(profile_xml)

                    # Add the profile
                    subprocess.check_output(["netsh", "wlan", "add", "profile", f"filename=\"{temp_xml_path}\""], encoding="utf-8", stderr=subprocess.STDOUT)

                    # Connect to the profile
                    subprocess.check_output(["netsh", "wlan", "connect", f"name=\"{ssid}\""], encoding="utf-8", stderr=subprocess.STDOUT)

                    QMessageBox.information(self, "Success", f"Attempted to connect to '{ssid}'. Please check your network status.")
                    self.statusBar.showMessage(f"Attempted to connect to '{ssid}'.", 3000)
                    self.extract_wifi_profiles() # Refresh saved profiles table
                except subprocess.CalledProcessError as e:
                    error_message = e.output.strip()
                    QMessageBox.critical(self, "Connection Error",
                                         f"Failed to connect to '{ssid}':\n{error_message}\n\n"
                                         "Ensure the network is in range and you have permissions (Administrator privileges may be required).")
                    self.log_error(f"Failed to connect to '{ssid}': {error_message}")
                    self.statusBar.showMessage("Failed to connect to network.", 5000)
                except Exception as e:
                    QMessageBox.critical(self, "Connection Error", f"An unexpected error occurred: {e}")
                    self.log_error(f"Unexpected error connecting to '{ssid}': {e}")
                    self.statusBar.showMessage("Unexpected error connecting to network.", 5000)
                finally:
                    if os.path.exists(temp_xml_path):
                        os.remove(temp_xml_path) # Clean up temp file
            else:
                self.statusBar.showMessage("Connection cancelled.", 2000)
        else:
            QMessageBox.warning(self, "No Selection", "Please select an available network to connect to.")
            self.statusBar.showMessage("No network selected for connection.", 3000)


    def connect_to_profile(self):
        """Attempts to connect to the selected Wi-Fi profile."""
        selected_row = self.table.currentRow()
        if selected_row != -1:
            profile_name = self.table.item(selected_row, 0).text()
            self.statusBar.showMessage(f"Attempting to connect to '{profile_name}'...", 0)
            try:
                subprocess.check_output(
                    ["netsh", "wlan", "connect", f"name=\"{profile_name}\""],
                    encoding="utf-8",
                    stderr=subprocess.STDOUT
                )
                QMessageBox.information(self, "Success", f"Attempted to connect to '{profile_name}'. Please check your network status.")
                self.statusBar.showMessage(f"Attempted to connect to '{profile_name}'.", 3000)
                self.extract_wifi_profiles() # Refresh status
            except subprocess.CalledProcessError as e:
                error_message = e.output.strip()
                QMessageBox.critical(self, "Connection Error",
                                     f"Failed to connect to '{profile_name}':\n{error_message}\n\n"
                                     "Ensure the network is in range and you have permissions (Administrator privileges may be required).")
                self.log_error(f"Failed to connect to '{profile_name}': {error_message}")
                self.statusBar.showMessage("Failed to connect to profile.", 5000)
            except Exception as e:
                QMessageBox.critical(self, "Connection Error", f"An unexpected error occurred: {e}")
                self.log_error(f"Unexpected error connecting to '{profile_name}': {e}")
                self.statusBar.showMessage("Unexpected error connecting to profile.", 5000)
        else:
            QMessageBox.warning(self, "No Selection", "Please select a profile to connect to.")
            self.statusBar.showMessage("No profile selected for connection.", 3000)

    def disconnect_current_connection(self):
        """Disconnects from the currently active Wi-Fi connection."""
        self.statusBar.showMessage("Attempting to disconnect from current Wi-Fi...", 0)
        reply = QMessageBox.question(
            self, 'Disconnect Wi-Fi',
            "Are you sure you want to disconnect from the current Wi-Fi network?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            try:
                subprocess.check_output(
                    ["netsh", "wlan", "disconnect"],
                    encoding="utf-8",
                    stderr=subprocess.STDOUT
                )
                QMessageBox.information(self, "Success", "Disconnected from current Wi-Fi network.")
                self.statusBar.showMessage("Disconnected from current Wi-Fi.", 3000)
                self.extract_wifi_profiles() # Refresh status
            except subprocess.CalledProcessError as e:
                error_message = e.output.strip()
                QMessageBox.critical(self, "Disconnection Error",
                                     f"Failed to disconnect: {error_message}\n\n"
                                     "You might need Administrator privileges or no network is currently connected.")
                self.log_error(f"Failed to disconnect: {error_message}")
                self.statusBar.showMessage("Failed to disconnect.", 5000)
            except Exception as e:
                QMessageBox.critical(self, "Disconnection Error", f"An unexpected error occurred: {e}")
                self.log_error(f"Unexpected error disconnecting: {e}")
                self.statusBar.showMessage("Unexpected error disconnecting.", 5000)

    def delete_selected_profile(self):
        """Deletes the selected Wi-Fi profile from the system."""
        selected_row = self.table.currentRow()
        if selected_row != -1:
            profile_name = self.table.item(selected_row, 0).text()
            reply = QMessageBox.question(
                self, 'Delete Profile',
                f"Are you sure you want to delete the Wi-Fi profile '{profile_name}' from your system? This action cannot be undone.",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.statusBar.showMessage(f"Attempting to delete profile '{profile_name}'...", 0)
                try:
                    subprocess.check_output(
                        ["netsh", "wlan", "delete", "profile", f"name=\"{profile_name}\""],
                        encoding="utf-8",
                        stderr=subprocess.STDOUT
                    )
                    QMessageBox.information(self, "Success", f"Profile '{profile_name}' deleted successfully.")
                    self.statusBar.showMessage(f"Profile '{profile_name}' deleted.", 3000)
                    self.extract_wifi_profiles() # Refresh table after deletion
                except subprocess.CalledProcessError as e:
                    error_message = e.output.strip()
                    QMessageBox.critical(self, "Error Deleting Profile",
                                         f"Failed to delete profile '{profile_name}':\n{error_message}\n\n"
                                         "This usually requires Administrator privileges.")
                    self.log_error(f"Failed to delete profile '{profile_name}': {error_message}")
                    self.statusBar.showMessage("Failed to delete profile.", 5000)
                except Exception as e:
                    QMessageBox.critical(self, "Error Deleting Profile", f"An unexpected error occurred: {e}")
                    self.log_error(f"Unexpected error deleting profile '{profile_name}': {e}")
                    self.statusBar.showMessage("Unexpected error deleting profile.", 5000)
        else:
            QMessageBox.warning(self, "No Selection", "Please select a row to delete.")
            self.statusBar.showMessage("No profile selected for deletion.", 3000)

    def rename_selected_profile(self):
        """Renames the selected Wi-Fi profile."""
        selected_row = self.table.currentRow()
        if selected_row != -1:
            old_profile_name = self.table.item(selected_row, 0).text()
            new_profile_name, ok = QInputDialog.getText(self, "Rename Profile",
                                                       f"Enter new name for '{old_profile_name}':",
                                                       QLineEdit.Normal, old_profile_name)

            if ok and new_profile_name and new_profile_name != old_profile_name:
                self.statusBar.showMessage(f"Attempting to rename profile '{old_profile_name}' to '{new_profile_name}'...", 0)
                try:
                    # Export the profile to a temp XML
                    temp_xml_path = f"{old_profile_name}_temp_export.xml"
                    subprocess.check_output(
                        ["netsh", "wlan", "export", "profile", f"name=\"{old_profile_name}\"", f"folder=\"{os.getcwd()}\"", "key=clear"],
                        encoding="utf-8",
                        stderr=subprocess.STDOUT
                    )

                    # Modify the XML to change the profile name
                    tree = ET.parse(temp_xml_path)
                    root = tree.getroot()
                    # Find the <name> tag and update its text
                    name_tag = root.find('{http://www.microsoft.com/networking/WLAN/profile/v1}name')
                    if name_tag is not None:
                        name_tag.text = new_profile_name
                        tree.write(temp_xml_path, encoding="utf-8", xml_declaration=True)
                    else:
                        raise ValueError("Could not find <name> tag in profile XML.")

                    # Delete the old profile
                    subprocess.check_output(
                        ["netsh", "wlan", "delete", "profile", f"name=\"{old_profile_name}\""],
                        encoding="utf-8",
                        stderr=subprocess.STDOUT
                    )

                    # Add the new profile (with the new name)
                    subprocess.check_output(
                        ["netsh", "wlan", "add", "profile", f"filename=\"{temp_xml_path}\""],
                        encoding="utf-8",
                        stderr=subprocess.STDOUT
                    )

                    QMessageBox.information(self, "Success", f"Profile '{old_profile_name}' successfully renamed to '{new_profile_name}'.")
                    self.statusBar.showMessage(f"Profile renamed to '{new_profile_name}'.", 3000)
                    self.extract_wifi_profiles() # Refresh table
                except subprocess.CalledProcessError as e:
                    error_message = e.output.strip()
                    QMessageBox.critical(self, "Error Renaming Profile",
                                         f"Failed to rename profile '{old_profile_name}':\n{error_message}\n\n"
                                         "This usually requires Administrator privileges.")
                    self.log_error(f"Failed to rename profile '{old_profile_name}': {error_message}")
                    self.statusBar.showMessage("Failed to rename profile.", 5000)
                except Exception as e:
                    QMessageBox.critical(self, "Error Renaming Profile", f"An unexpected error occurred: {e}")
                    self.log_error(f"Unexpected error renaming profile '{old_profile_name}': {e}")
                    self.statusBar.showMessage("Unexpected error renaming profile.", 5000)
                finally:
                    if os.path.exists(temp_xml_path):
                        os.remove(temp_xml_path) # Clean up temp file
            elif ok and new_profile_name == old_profile_name:
                self.statusBar.showMessage("Profile name not changed.", 2000)
            else:
                self.statusBar.showMessage("Rename cancelled.", 2000)
        else:
            QMessageBox.warning(self, "No Selection", "Please select a row to rename.")
            self.statusBar.showMessage("No profile selected for renaming.", 3000)


    # --- Import/Export XML ---
    def import_profile_from_xml(self):
        """Imports a Wi-Fi profile from an XML file."""
        path, _ = QFileDialog.getOpenFileName(self, "Select Wi-Fi Profile XML", "", "XML Files (*.xml);;All Files (*)")
        if path:
            self.statusBar.showMessage(f"Importing profile from '{os.path.basename(path)}'...", 0)
            try:
                subprocess.check_output(
                    ["netsh", "wlan", "add", "profile", f"filename=\"{path}\""],
                    encoding="utf-8",
                    stderr=subprocess.STDOUT
                )
                QMessageBox.information(self, "Success", f"Profile from '{os.path.basename(path)}' imported successfully!")
                self.statusBar.showMessage("Profile imported successfully.", 3000)
                self.extract_wifi_profiles() # Refresh table
            except subprocess.CalledProcessError as e:
                error_message = e.output.strip()
                QMessageBox.critical(self, "Error Importing Profile",
                                     f"Failed to import profile from '{os.path.basename(path)}':\n{error_message}\n\n"
                                     "Ensure the XML is valid and you have Administrator privileges.")
                self.log_error(f"Failed to import profile from '{path}': {error_message}")
                self.statusBar.showMessage("Failed to import profile.", 5000)
            except Exception as e:
                QMessageBox.critical(self, "Error Importing Profile", f"An unexpected error occurred: {e}")
                self.log_error(f"Unexpected error importing profile from '{path}': {e}")
                self.statusBar.showMessage("Unexpected error importing profile.", 5000)
        else:
            self.statusBar.showMessage("Import cancelled.", 2000)

    def export_selected_profile_to_xml(self):
        """Exports the selected Wi-Fi profile to an XML file."""
        selected_row = self.table.currentRow()
        if selected_row != -1:
            profile_name = self.table.item(selected_row, 0).text()
            path, _ = QFileDialog.getSaveFileName(self, f"Export Profile '{profile_name}' to XML",
                                                  f"{profile_name}.xml", "XML Files (*.xml);;All Files (*)")
            if path:
                self.statusBar.showMessage(f"Exporting profile '{profile_name}' to XML...", 0)
                try:
                    subprocess.check_output(
                        ["netsh", "wlan", "export", "profile", f"name=\"{profile_name}\"", f"folder=\"{os.path.dirname(path)}\"", "key=clear"],
                        encoding="utf-8",
                        stderr=subprocess.STDOUT
                    )
                    # The command saves it with the profile name, so we might need to move/rename if path is different
                    # For simplicity, let's assume the user saves to the desired filename.
                    # If the user specified a different filename than profile_name.xml, we need to rename the exported file.
                    exported_filename = os.path.join(os.path.dirname(path), f"{profile_name}.xml")
                    if exported_filename != path:
                        shutil.move(exported_filename, path)

                    QMessageBox.information(self, "Success", f"Profile '{profile_name}' exported to XML successfully!")
                    self.statusBar.showMessage("Profile exported to XML.", 3000)
                except subprocess.CalledProcessError as e:
                    error_message = e.output.strip()
                    QMessageBox.critical(self, "Error Exporting Profile",
                                         f"Failed to export profile '{profile_name}' to XML:\n{error_message}\n\n"
                                         "This usually requires Administrator privileges.")
                    self.log_error(f"Failed to export profile '{profile_name}' to XML: {error_message}")
                    self.statusBar.showMessage("Failed to export profile to XML.", 5000)
                except Exception as e:
                    QMessageBox.critical(self, "Error Exporting Profile", f"An unexpected error occurred: {e}")
                    self.log_error(f"Unexpected error exporting profile '{profile_name}' to XML: {e}")
                    self.statusBar.showMessage("Unexpected error exporting profile to XML.", 5000)
            else:
                self.statusBar.showMessage("Export cancelled.", 2000)
        else:
            QMessageBox.warning(self, "No Selection", "Please select a row to export to XML.")
            self.statusBar.showMessage("No profile selected for XML export.", 3000)


    # --- Standard Export/Utility Functions (Modified to use status bar) ---
    def auto_save_logs(self):
        """Automatically saves extracted profiles and passwords to a log file."""
        try:
            log_file_path = "wifi_profiles_log.txt"
            with open(log_file_path, "w", encoding="utf-8") as file:
                for row in range(self.table.rowCount()):
                    profile = self.table.item(row, 0).text()
                    password = self.table.item(row, 1).text()
                    file.write(f"Profile: {profile}, Password: {password}\n")
            self.update_button_states()
        except Exception as e:
            self.log_error(f"Error saving log: {e}")

    def generate_qr_code(self):
        """Generates a QR code for the selected Wi-Fi profile's credentials."""
        selected_row = self.table.currentRow()
        if selected_row != -1:
            profile = self.table.item(selected_row, 0).text()
            password = self.table.item(selected_row, 1).text()
            if password == "No Password" or "Access Denied" in password:
                QMessageBox.warning(self, "QR Code Error", "Cannot generate QR code for profiles without a readable password.")
                self.statusBar.showMessage("Cannot generate QR code without password.", 3000)
                return

            qr_data = f"WIFI:S:{profile};T:WPA;P:{password};;"
            qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
            qr.add_data(qr_data)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")

            file_name = f"{profile}_qr.png"
            try:
                img.save(file_name)
                QMessageBox.information(self, "QR Code Generated", f"QR Code saved as {file_name}")
                self.statusBar.showMessage(f"QR Code saved as {file_name}", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Error Saving QR", f"Failed to save QR code: {e}")
                self.log_error(f"Failed to save QR code: {e}")
                self.statusBar.showMessage("Failed to save QR code.", 5000)
        else:
            QMessageBox.warning(self, "No Selection", "Please select a row to generate QR code.")
            self.statusBar.showMessage("No profile selected for QR code.", 3000)

    def export_to_file(self):
        """Exports the table content to a plain text file."""
        path, _ = QFileDialog.getSaveFileName(self, "Save Profiles to Text", "", "Text Files (*.txt);;All Files (*)")
        if path:
            self.statusBar.showMessage("Exporting to text file...", 0)
            try:
                with open(path, "w", encoding="utf-8") as file:
                    for row in range(self.table.rowCount()):
                        row_data = []
                        for col in range(self.table.columnCount()):
                            item = self.table.item(row, col)
                            if item:
                                row_data.append(item.text())
                            else:
                                row_data.append("")
                        file.write(", ".join(row_data) + "\n")
                QMessageBox.information(self, "Success", "Profiles exported to text file successfully!")
                self.statusBar.showMessage("Profiles exported to text file.", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export to text file: {e}")
                self.log_error(f"Failed to export to text file: {e}")
                self.statusBar.showMessage("Failed to export to text file.", 5000)

    def export_to_excel(self):
        """Exports the table content to an Excel (.xlsx) file using openpyxl."""
        path, _ = QFileDialog.getSaveFileName(self, "Save Profiles to Excel", "", "Excel Files (*.xlsx);;All Files (*)")
        if path:
            self.statusBar.showMessage("Exporting to Excel file...", 0)
            try:
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.title = "Wi-Fi Profiles"

                headers = [self.table.horizontalHeaderItem(col).text() for col in range(self.table.columnCount())]
                sheet.append(headers)

                for row_idx in range(self.table.rowCount()):
                    row_data = []
                    for col_idx in range(self.table.columnCount()):
                        item = self.table.item(row_idx, col_idx)
                        row_data.append(item.text() if item else "")
                    sheet.append(row_data)

                workbook.save(path)
                QMessageBox.information(self, "Success", "Profiles exported to Excel successfully!")
                self.statusBar.showMessage("Profiles exported to Excel.", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export to Excel: {e}")
                self.log_error(f"Failed to export to Excel: {e}")
                self.statusBar.showMessage("Failed to export to Excel.", 5000)

    def backup_logs(self):
        """Creates a timestamped backup of the wifi_profiles_log.txt file using shutil."""
        log_file = "wifi_profiles_log.txt"
        if not os.path.exists(log_file):
            QMessageBox.warning(self, "Backup Error", "No log file found to backup.")
            self.statusBar.showMessage("No log file found for backup.", 3000)
            return

        backup_dir = "wifi_backups"
        os.makedirs(backup_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"wifi_profiles_log_backup_{timestamp}.txt"
        backup_path = os.path.join(backup_dir, backup_filename)

        self.statusBar.showMessage("Backing up log file...", 0)
        try:
            shutil.copy2(log_file, backup_path)
            QMessageBox.information(self, "Backup Success", f"Log file backed up to:\n{backup_path}")
            self.statusBar.showMessage("Log file backed up successfully.", 3000)
        except Exception as e:
            QMessageBox.critical(self, "Backup Error", f"Failed to backup log file: {e}")
            self.log_error(f"Failed to backup log file: {e}")
            self.statusBar.showMessage("Failed to backup log file.", 5000)

    def show_summary(self):
        """Analyzes the extracted data using pandas and displays a summary."""
        if self.table.rowCount() == 0:
            QMessageBox.information(self, "Summary", "No data available to generate a summary. Please extract profiles first.")
            self.statusBar.showMessage("No data for summary.", 3000)
            return

        self.statusBar.showMessage("Generating summary...", 0)
        data = []
        headers = [self.table.horizontalHeaderItem(col).text() for col in range(self.table.columnCount())]
        for row_idx in range(self.table.rowCount()):
            row_data = {}
            for col_idx, header in enumerate(headers):
                item = self.table.item(row_idx, col_idx)
                row_data[header] = item.text() if item else ""
            data.append(row_data)

        df = pd.DataFrame(data)

        summary_text = "--- Wi-Fi Profile Summary ---\n\n"
        summary_text += f"Total Profiles: {len(df)}\n\n"

        if "Authentication" in df.columns:
            auth_summary = df["Authentication"].value_counts().to_string()
            summary_text += "Authentication Types:\n"
            summary_text += auth_summary + "\n\n"

        if "Encryption" in df.columns:
            encryption_summary = df["Encryption"].value_counts().to_string()
            summary_text += "Encryption Types:\n"
            summary_text += encryption_summary + "\n\n"

        if "Status" in df.columns:
            status_summary = df["Status"].value_counts().to_string()
            summary_text += "Connection Status:\n"
            summary_text += status_summary + "\n\n"

        if "Strength" in df.columns:
            strength_summary = df["Strength"].value_counts().to_string()
            summary_text += "Password Strength Distribution:\n"
            summary_text += strength_summary + "\n\n"

        QMessageBox.information(self, "Wi-Fi Profile Summary", summary_text)
        self.statusBar.showMessage("Summary generated.", 3000)


    def search_profiles(self):
        """Filters the table rows based on the search query."""
        query = self.search_box.text().lower()
        for row in range(self.table.rowCount()):
            profile_item = self.table.item(row, 0)
            if profile_item:
                self.table.setRowHidden(row, query not in profile_item.text().lower())
        self.statusBar.showMessage("Search complete.", 1000)

    def clear_table(self):
        """Clears all rows from the table and disables relevant buttons."""
        self.table.setRowCount(0)
        self.update_button_states()
        self.statusBar.showMessage("Table cleared.", 2000)

    def copy_to_clipboard(self):
        """Copies the selected row's profile name and password to the clipboard."""
        selected_row = self.table.currentRow()
        if selected_row != -1:
            profile = self.table.item(selected_row, 0).text()
            password = self.table.item(selected_row, 1).text()
            clipboard = QApplication.clipboard()
            clipboard.setText(f"Profile: {profile}\nPassword: {password}")
            QMessageBox.information(self, "Copied", "Profile information copied to clipboard!")
            self.statusBar.showMessage("Profile info copied to clipboard.", 3000)
        else:
            QMessageBox.warning(self, "No Selection", "Please select a row to copy.")
            self.statusBar.showMessage("No profile selected for copy.", 3000)

    def toggle_dark_mode(self):
        """Toggles between light and dark mode themes."""
        if self.styleSheet():
            self.setStyleSheet("")
            self.statusBar.showMessage("Light mode enabled.", 2000)
        else:
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #2e2e2e;
                    color: #f0f0f0;
                }
                QWidget {
                    background-color: #2e2e2e;
                    color: #f0f0f0;
                }
                QTabWidget::pane { /* The tab widget frame */
                    border: 1px solid #505050;
                    background-color: #2e2e2e;
                }
                QTabWidget::tab-bar {
                    left: 5px; /* move to the right */
                }
                QTabBar::tab {
                    background: #4a4a4a;
                    color: #f0f0f0;
                    border: 1px solid #505050;
                    border-bottom-color: #2e2e2e; /* same as pane color */
                    border-top-left-radius: 4px;
                    border-top-right-radius: 4px;
                    min-width: 8ex;
                    padding: 5px;
                }
                QTabBar::tab:selected, QTabBar::tab:hover {
                    background: #555555;
                }
                QTabBar::tab:selected {
                    border-color: #505050;
                    border-bottom-color: #555555; /* same as selected tab color */
                }
                QTableWidget {
                    background-color: #3c3c3c;
                    color: #f0f0f0;
                    gridline-color: #505050;
                    selection-background-color: #555555;
                    border: 1px solid #505050;
                }
                QTableWidget::item {
                    border-bottom: 1px solid #505050;
                }
                QHeaderView::section {
                    background-color: #4a4a4a;
                    color: #f0f0f0;
                    padding: 4px;
                    border: 1px solid #505050;
                }
                QPushButton {
                    background-color: #555555;
                    color: #f0f0f0;
                    border: 1px solid #666666;
                    border-radius: 5px;
                    padding: 8px 15px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #666666;
                }
                QPushButton:pressed {
                    background-color: #444444;
                }
                QPushButton:disabled {
                    background-color: #3a3a3a;
                    color: #999999;
                    border: 1px solid #4a4a4a;
                }
                QLineEdit {
                    background-color: #3c3c3c;
                    color: #f0f0f0;
                    border: 1px solid #505050;
                    border-radius: 5px;
                    padding: 5px;
                }
                QMessageBox {
                    background-color: #2e2e2e;
                    color: #f0f0f0;
                }
                QMessageBox QPushButton {
                    background-color: #555555;
                    color: #f0f0f0;
                    border: 1px solid #666666;
                    border-radius: 5px;
                    padding: 5px 10px;
                }
                QMessageBox QPushButton:hover {
                    background-color: #666666;
                }
                QProgressDialog {
                    background-color: #2e2e2e;
                    color: #f0f0f0;
                    border: 1px solid #505050;
                }
                QProgressDialog QLabel {
                    color: #f0f0f0;
                }
                QProgressDialog QProgressBar {
                    background-color: #3c3c3c;
                    border: 1px solid #505050;
                    border-radius: 5px;
                    text-align: center;
                    color: #f0f0f0;
                }
                QProgressDialog QProgressBar::chunk {
                    background-color: #555555;
                    border-radius: 5px;
                }
            """)
            self.statusBar.showMessage("Dark mode enabled.", 2000)

    def show_about(self):
        """Displays information about the application."""
        QMessageBox.information(
            self, "About", "Wi-Fi Profile Management App\nVersion 2.3\n"
            "Developed with Python, PyQt5, shutil, openpyxl, pandas, and xml.etree.ElementTree.\n\n"
            "Features:\n"
            "- Extract & View Saved Wi-Fi Profiles (requires admin for passwords)\n"
            "- Scan & View Available Networks\n"
            "- Connect to Saved or Available Networks (requires admin for available)\n"
            "- Delete, Rename Wi-Fi Profiles (requires admin)\n"
            "- Import/Export Wi-Fi Profiles to XML (requires admin)\n"
            "- Search, Copy, QR Code Generation\n"
            "- Export to Text & Excel\n"
            "- Backup Logs\n"
            "- Data Summary with Password Strength Analysis\n"
            "- Context Menu for table actions\n"
            "- Status Bar for real-time feedback\n"
            "- Dark Mode Toggle\n"
            "- Responsive UI with background processing for long tasks"
        )
        self.statusBar.showMessage("About information displayed.", 2000)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WifiProfileExtractor()
    window.show()
    sys.exit(app.exec_())

 
