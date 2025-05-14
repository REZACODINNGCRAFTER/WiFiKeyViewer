import ctypes
import sys
import subprocess
import os
import qrcode
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton, QLabel, QTableWidget,
    QTableWidgetItem, QLineEdit, QMessageBox, QFileDialog, QWidget, QHBoxLayout
)
from PyQt5.QtGui import QClipboard, QPixmap
from PyQt5.QtCore import Qt


class WifiProfileExtractor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Wi-Fi Profile Password Extraction App")
        self.setGeometry(100, 100, 1000, 650)
        self.init_ui()
        self.check_admin_privileges()

    # UI Initialization
    def init_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout()

        self.setup_search_box(layout)
        self.setup_table(layout)
        self.setup_buttons(layout)

        self.central_widget.setLayout(layout)

    def setup_search_box(self, layout):
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search Wi-Fi Profile")
        self.search_box.textChanged.connect(self.search_profiles)
        layout.addWidget(self.search_box)

    def setup_table(self, layout):
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Profile Name", "Password", "Authentication", "Encryption", "SSID Visibility", "Status"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

    def setup_buttons(self, layout):
        button_layout = QHBoxLayout()

        self.extract_button = QPushButton("Extract Wi-Fi Profiles")
        self.extract_button.clicked.connect(self.extract_wifi_profiles)
        button_layout.addWidget(self.extract_button)

        self.export_button = QPushButton("Export to File")
        self.export_button.clicked.connect(self.export_to_file)
        self.export_button.setEnabled(False)
        button_layout.addWidget(self.export_button)

        self.clear_button = QPushButton("Clear Table")
        self.clear_button.clicked.connect(self.clear_table)
        button_layout.addWidget(self.clear_button)

        self.copy_button = QPushButton("Copy Selected")
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        button_layout.addWidget(self.copy_button)

        self.qr_button = QPushButton("Generate QR Code")
        self.qr_button.clicked.connect(self.generate_qr_code)
        self.qr_button.setEnabled(False)
        button_layout.addWidget(self.qr_button)

        self.about_button = QPushButton("About")
        self.about_button.clicked.connect(self.show_about)
        button_layout.addWidget(self.about_button)

        self.dark_mode_button = QPushButton("Toggle Dark Mode")
        self.dark_mode_button.clicked.connect(self.toggle_dark_mode)
        button_layout.addWidget(self.dark_mode_button)

        layout.addLayout(button_layout)

    # Functionality
    def extract_wifi_profiles(self):
        try:
            result = subprocess.check_output(
                ["netsh", "wlan", "show", "profiles"], encoding="utf-8"
            )
            profiles = [
                line.split(":")[1].strip()
                for line in result.split("\n") if "All User Profile" in line
            ]

            self.table.setRowCount(0)
            for profile in profiles:
                details = self.get_profile_details(profile)
                self.add_table_row(profile, details)

            self.export_button.setEnabled(len(profiles) > 0)
            self.qr_button.setEnabled(len(profiles) > 0)
            self.auto_save_logs(profiles)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract profiles: {e}")

    def get_profile_details(self, profile_name):
        details = {
            "password": "No Password",
            "auth": "N/A",
            "encryption": "N/A",
            "ssid_visibility": "N/A",
            "status": "Disconnected"
        }
        try:
            result = subprocess.check_output(
                ["netsh", "wlan", "show", "profile", profile_name, "key=clear"],
                encoding="utf-8"
            )
            for line in result.split("\n"):
                if "Key Content" in line:
                    details["password"] = line.split(":")[1].strip()
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
            details["password"] = "Access Denied"
        return details

    def add_table_row(self, profile_name, details):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        for i, value in enumerate([
            profile_name, details["password"], details["auth"],
            details["encryption"], details["ssid_visibility"], details["status"]
        ]):
            self.table.setItem(row_position, i, QTableWidgetItem(value))

    def auto_save_logs(self, profiles):
        try:
            with open("wifi_profiles_log.txt", "w") as file:
                for row in range(self.table.rowCount()):
                    profile = self.table.item(row, 0).text()
                    password = self.table.item(row, 1).text()
                    file.write(f"Profile: {profile}, Password: {password}\n")
        except Exception as e:
            print(f"Error saving log: {e}")

    def generate_qr_code(self):
        selected_row = self.table.currentRow()
        if selected_row != -1:
            profile = self.table.item(selected_row, 0).text()
            password = self.table.item(selected_row, 1).text()
            qr_data = f"WIFI:S:{profile};T:WPA;P:{password};;"
            qr = qrcode.QRCode()
            qr.add_data(qr_data)
            qr.make(fit=True)
            img = qr.make_image(fill="black", back_color="white")
            img.save(f"{profile}_qr.png")
            QMessageBox.information(self, "QR Code", f"QR Code saved as {profile}_qr.png")
        else:
            QMessageBox.warning(self, "No Selection", "Please select a row to generate QR code.")

    def check_admin_privileges(self):
        """Check if the application has administrative privileges."""
        try:

            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if not is_admin:
                QMessageBox.warning(
                    self,
                    "Warning",
                    "This application may require administrative privileges to function fully.",
                )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to check admin privileges: {e}")
    # Additional Methods (Export, Search, Dark Mode, etc.)
    def export_to_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Text Files (*.txt)")
        if path:
            try:
                with open(path, "w") as file:
                    for row in range(self.table.rowCount()):
                        file.write(", ".join([
                            self.table.item(row, col).text()
                            for col in range(self.table.columnCount())
                        ]) + "\n")
                QMessageBox.information(self, "Success", "Profiles exported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export: {e}")

    def search_profiles(self):
        query = self.search_box.text().lower()
        for row in range(self.table.rowCount()):
            profile_item = self.table.item(row, 0)
            self.table.setRowHidden(row, query not in profile_item.text().lower())

    def clear_table(self):
        self.table.setRowCount(0)
        self.export_button.setEnabled(False)
        self.qr_button.setEnabled(False)

    def copy_to_clipboard(self):
        selected_row = self.table.currentRow()
        if selected_row != -1:
            profile = self.table.item(selected_row, 0).text()
            password = self.table.item(selected_row, 1).text()
            clipboard = QApplication.clipboard()
            clipboard.setText(f"Profile: {profile}\nPassword: {password}")
            QMessageBox.information(self, "Copied", "Profile information copied to clipboard!")
        else:
            QMessageBox.warning(self, "No Selection", "Please select a row to copy.")

    def toggle_dark_mode(self):
        self.setStyleSheet("" if self.styleSheet() else "background-color: #2e2e2e; color: white;")

    def show_about(self):
        QMessageBox.information(
            self, "About", "Wi-Fi Profile Password Extraction App\nVersion 2.0\nDeveloped with Python and PyQt5."
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WifiProfileExtractor()
    window.show()
    sys.exit(app.exec_())
