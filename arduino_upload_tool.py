import os
import sys
import time
import zipfile
import subprocess
import tempfile
import json
import urllib.request
import ssl

import requests
import certifi

from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout,
    QDialog, QLabel, QProgressBar, QTextEdit, QHBoxLayout,
    QMessageBox, QComboBox, QFileDialog, QListWidget, QListWidgetItem,
    QInputDialog, QRadioButton, QButtonGroup, QFrame, QSpacerItem,
    QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from serial.tools import list_ports  # pip install pyserial


# ---------------------------------------------------------
# CONFIG
# ---------------------------------------------------------
ARDUINO_CLI_URL = "https://downloads.arduino.cc/arduino-cli/arduino-cli_latest_Windows_64bit.zip"
TEMP_DIR = tempfile.gettempdir()
DOWNLOAD_PATH = os.path.join(TEMP_DIR, "arduino-cli.zip")
INSTALL_DIR = os.path.expanduser("~/.arduino-cli")
CLI_NAME = "arduino-cli.exe"

RECENT_FILE = os.path.join(os.path.expanduser("~"), ".arduino_recent_projects.json")
MAX_RECENT = 10

EDITOR_PREF_FILE = os.path.join(os.path.expanduser("~"), ".arduino_editor_preferences.json")


# ---------------------------------------------------------
# EDITOR PREFERENCE HELPERS
# ---------------------------------------------------------
def load_editor_preference():
    if not os.path.exists(EDITOR_PREF_FILE):
        return None
    try:
        with open(EDITOR_PREF_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def save_editor_preference(name, path):
    data = {
        "default_editor_name": name,
        "default_editor_path": path
    }
    try:
        with open(EDITOR_PREF_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------
# EDITOR DETECTION
# ---------------------------------------------------------
def detect_editors():
    editors = []

    # Arduino IDE (common paths)
    arduino_paths = [
        r"C:\Program Files\Arduino IDE\Arduino IDE.exe",
        r"C:\Program Files\Arduino IDE\arduino-ide.exe",
        r"C:\Program Files (x86)\Arduino\arduino.exe",
        r"C:\Program Files\Arduino\arduino.exe",
    ]
    for p in arduino_paths:
        if os.path.exists(p):
            editors.append(("Arduino IDE", p))
            break

    # VS Code
    vscode_paths = [
        rf"C:\Users\{os.getlogin()}\AppData\Local\Programs\Microsoft VS Code\Code.exe",
        r"C:\Program Files\Microsoft VS Code\Code.exe",
    ]
    for p in vscode_paths:
        if os.path.exists(p):
            editors.append(("VS Code", p))
            break

    # Notepad++
    notepadpp_paths = [
        r"C:\Program Files\Notepad++\notepad++.exe",
        r"C:\Program Files (x86)\Notepad++\notepad++.exe",
    ]
    for p in notepadpp_paths:
        if os.path.exists(p):
            editors.append(("Notepad++", p))
            break

    # Notepad (always available)
    editors.append(("Notepad", "notepad.exe"))

    return editors


# ---------------------------------------------------------
# EDITOR SELECTION DIALOG
# ---------------------------------------------------------
class EditorSelectionDialog(QDialog):
    def __init__(self, parent=None, ino_path=""):
        super().__init__(parent)
        self.setWindowTitle("Select Editor")
        self.setModal(True)

        self.ino_path = ino_path
        self.selected_editor_name = None
        self.selected_editor_path = None
        self.set_as_default = False

        self.editors = detect_editors()
        self.custom_editor_path = None

        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        # Section: Select an editor
        title1 = QLabel("Select an editor")
        title1.setStyleSheet("font-weight: bold;")
        layout.addWidget(title1)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.HLine)
        layout.addWidget(sep1)

        self.editor_group = QButtonGroup(self)
        self.editor_group.setExclusive(True)

        # Radio buttons for detected editors
        for name, path in self.editors:
            rb = QRadioButton(name)
            rb.editor_name = name
            rb.editor_path = path
            self.editor_group.addButton(rb)
            layout.addWidget(rb)

        # Preselect VS Code if present
        buttons = self.editor_group.buttons()
        vs_code_button = next((b for b in buttons if b.editor_name == "VS Code"), None)
        if vs_code_button:
            vs_code_button.setChecked(True)
        elif buttons:
            buttons[0].setChecked(True)

        # Custom editor
        self.custom_radio = QRadioButton("Custom editor…")
        self.custom_radio.editor_name = "Custom"
        self.custom_radio.editor_path = None
        self.editor_group.addButton(self.custom_radio)

        custom_layout = QHBoxLayout()
        custom_layout.addWidget(self.custom_radio)

        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.browse_custom_editor)
        custom_layout.addWidget(self.browse_button)

        layout.addLayout(custom_layout)

        layout.addSpacing(10)

        # Section: Default editor behavior
        title2 = QLabel("Default editor behavior")
        title2.setStyleSheet("font-weight: bold;")
        layout.addWidget(title2)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.HLine)
        layout.addWidget(sep2)

        self.behavior_group = QButtonGroup(self)
        self.behavior_group.setExclusive(True)

        self.use_once_radio = QRadioButton("Use this editor once")
        self.set_default_radio = QRadioButton("Set this editor as default")

        self.behavior_group.addButton(self.use_once_radio)
        self.behavior_group.addButton(self.set_default_radio)

        self.use_once_radio.setChecked(True)

        layout.addWidget(self.use_once_radio)
        layout.addWidget(self.set_default_radio)

        layout.addItem(QSpacerItem(0, 10, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch(1)

        ok_btn = QPushButton("OK")
        cancel_btn = QPushButton("Cancel")

        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)

    def browse_custom_editor(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select editor executable",
            "",
            "Executable files (*.exe);;All files (*)"
        )
        if path:
            self.custom_editor_path = path
            self.custom_radio.setChecked(True)
            self.custom_radio.editor_path = path

    def accept(self):
        checked_button = self.editor_group.checkedButton()
        if not checked_button:
            QMessageBox.warning(self, "No editor selected", "Please select an editor.")
            return

        name = checked_button.editor_name
        path = checked_button.editor_path

        if name == "Custom" and not path:
            QMessageBox.warning(self, "No custom editor", "Please browse for a custom editor executable.")
            return

        self.selected_editor_name = name
        self.selected_editor_path = path
        self.set_as_default = self.set_default_radio.isChecked()

        super().accept()


# ---------------------------------------------------------
# ROBUST DOWNLOAD FUNCTION
# ---------------------------------------------------------
def robust_download(url, dest, progress_callback, log_callback, max_retries=5):
    attempt = 1
    while attempt <= max_retries:
        try:
            msg = f"Downloading… attempt {attempt}"
            progress_callback(msg, 10)
            log_callback(msg)

            try:
                r = requests.get(url, stream=True, verify=certifi.where(), timeout=10)
            except requests.exceptions.SSLError:
                msg = "SSL failed — retrying without verification…"
                progress_callback(msg, 20)
                log_callback(msg)
                r = requests.get(url, stream=True, verify=False, timeout=10)

            total = int(r.headers.get("content-length", 0))
            downloaded = 0

            with open(dest, "wb") as f:
                for chunk in r.iter_content(chunk_size=1024 * 256):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total > 0:
                            pct = int((downloaded / total) * 30) + 10
                            progress_callback("Downloading…", pct)

            log_callback(f"Download completed. Saved to: {dest}")
            return

        except Exception as e:
            msg = f"Download failed, retrying… ({attempt}/{max_retries}) - {e}"
            progress_callback(msg, 10)
            log_callback(msg)
            attempt += 1
            time.sleep(1)

    raise Exception("Download failed after multiple retries.")




# ---------------------------------------------------------
# INSTALLER THREAD
# ---------------------------------------------------------
class InstallerThread(QThread):
    progress = pyqtSignal(str, int)
    log = pyqtSignal(str)
    test_result = pyqtSignal(bool, str)

    def run(self):
        # Step 1: Download
        try:
            robust_download(
                ARDUINO_CLI_URL,
                DOWNLOAD_PATH,
                self.progress.emit,
                self.log.emit
            )
        except Exception as e:
            msg = f"Download failed: {e}"
            self.progress.emit(msg, 0)
            self.log.emit(msg)
            self.test_result.emit(False, msg)
            return

        # Step 2: Unzip
        self.progress.emit("Unzipping…", 30)
        self.log.emit("Unzipping archive…")
        os.makedirs(INSTALL_DIR, exist_ok=True)
        try:
            with zipfile.ZipFile(DOWNLOAD_PATH, "r") as zip_ref:
                zip_ref.extractall(INSTALL_DIR)
            self.log.emit(f"Extracted to {INSTALL_DIR}")
        except Exception as e:
            msg = f"Unzip failed: {e}"
            self.progress.emit(msg, 0)
            self.log.emit(msg)
            self.test_result.emit(False, msg)
            return

        # Step 3: Install
        self.progress.emit("Installing…", 45)
        self.log.emit("Installation step complete.")

        # Step 4: Update PATH
        self.progress.emit("Updating PATH…", 50)
        self.log.emit("Updating PATH environment variable…")
        self.update_path(INSTALL_DIR)

        # Step 5: Test installation
        self.progress.emit("Testing Arduino CLI…", 55)
        self.log.emit("Starting post-installation tests…")
        ok, message = self.test_installation()
        self.log.emit(message)

        if ok:
            self.progress.emit("Installation successful!", 60)
        else:
            self.progress.emit("Installation completed with issues.", 60)

        self.test_result.emit(ok, message)

        self.log.emit("Starting Downloads of WinLibs")
        self.get_sw()




    def get_sw(self):
        try:
            self.log.emit("Fetching latest WinLibs release info...")
    
            # Ignore SSL certificate if needed
            ssl._create_default_https_context = ssl._create_unverified_context
    
            api_url = "https://api.github.com/repos/brechtsanders/winlibs_mingw/releases/latest"
            req = urllib.request.Request(api_url, headers={"User-Agent": "Mozilla/5.0"})
            data = urllib.request.urlopen(req).read()
            release = json.loads(data)
    
            # Find UCRT64 asset
            asset_url = None
            for asset in release["assets"]:
                name = asset["name"].lower()
                if "ucrt" in name and name.endswith(".zip"):
                    asset_url = asset["browser_download_url"]
                    break
                
            if not asset_url:
                self.log.emit("Could not find UCRT64 asset in latest release.")
                return
    
            self.log.emit(f"Downloading:\n{asset_url}")
            zip_path = "winlibs.zip"
    
            # Smooth progress download (60 → 90)
            self.progress.emit("Downloading WinLibs…", 60)
            ok = self.download_with_progress(asset_url, zip_path, 60, 90)
            if not ok:
                self.log.emit("Failed to download WinLibs.")
                return
    
            self.log.emit("Download complete. Extracting…")
            self.progress.emit("Extracting WinLibs…", 90)
    
            extract_dir = os.path.abspath("winlibs")
            os.makedirs(extract_dir, exist_ok=True)
    
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
    
            os.remove(zip_path)
    
            # Find /bin folder
            bin_path = None
            for root, dirs, files in os.walk(extract_dir):
                if root.endswith("bin"):
                    bin_path = root
                    break
                
            if bin_path:
                os.environ["PATH"] += os.pathsep + bin_path
                self.log.emit(f"Added to PATH:\n{bin_path}")
    
                # Rename mingw32-make.exe → make.exe
                make_src = os.path.join(bin_path, "mingw32-make.exe")
                make_dst = os.path.join(bin_path, "make.exe")
    
                if os.path.exists(make_src):
                    try:
                        if os.path.exists(make_dst):
                            os.remove(make_dst)
                        os.rename(make_src, make_dst)
                        self.log.emit("Renamed mingw32-make.exe → make.exe")
                    except Exception as e:
                        self.log.emit(f"Failed to rename mingw32-make.exe: {e}")
                else:
                    self.log.emit("mingw32-make.exe not found — cannot rename.")
            else:
                self.log.emit("Could not locate bin folder!")
    
            self.progress.emit("WinLibs installation complete!", 100)
            self.log.emit("WinLibs installation complete.")
    
        except Exception as e:
            self.log.emit(f"Error: {e}")



    def download_with_progress(self, url, dest, start_pct=60, end_pct=90):
        try:
            with urllib.request.urlopen(url) as response:
                total = int(response.headers.get("Content-Length", 0))
                downloaded = 0

                chunk_size = 1024 * 256  # 256 KB
                with open(dest, "wb") as f:
                    while True:
                        chunk = response.read(chunk_size)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)

                        if total > 0:
                            pct = start_pct + int((downloaded / total) * (end_pct - start_pct))
                            self.progress.emit("Downloading WinLibs…", pct)

            return True
        except Exception as e:
            self.log.emit(f"WinLibs download error: {e}")
            return False


    def update_path(self, path):
        current_path = os.environ.get("PATH", "")
        if path not in current_path:
            os.system(f'setx PATH "%PATH%;{path}"')
            self.log.emit(f"Added {path} to PATH.")
        else:
            self.log.emit(f"{path} already in PATH.")

    def test_installation(self):
        cli_path = os.path.join(INSTALL_DIR, CLI_NAME)

        if not os.path.exists(cli_path):
            return False, f"{CLI_NAME} not found after installation."

        try:
            self.log.emit("Running 'arduino-cli version'…")
            result = subprocess.run(
                [cli_path, "version"],
                capture_output=True,
                text=True
            )
            self.log.emit("stdout:\n" + result.stdout)
            self.log.emit("stderr:\n" + result.stderr)
            if result.returncode != 0:
                return False, "arduino-cli version failed."
        except Exception as e:
            return False, f"Error running version: {e}"

        try:
            self.log.emit("Running 'arduino-cli core update-index'…")
            result = subprocess.run(
                [cli_path, "core", "update-index"],
                capture_output=True,
                text=True
            )
            self.log.emit("stdout:\n" + result.stdout)
            self.log.emit("stderr:\n" + result.stderr)
            if result.returncode != 0:
                return False, "core update-index failed."
        except Exception as e:
            return False, f"Error updating core index: {e}"

        return True, "Arduino CLI is fully functional."


# ---------------------------------------------------------
# WIZARD DIALOG
# ---------------------------------------------------------
class WizardDialog(QDialog):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

        self.setWindowTitle("Arduino CLI Installer Wizard")
        self.setFixedSize(600, 400)

        self.status_label = QLabel("Starting…")
        self.status_label.setStyleSheet("font-weight: bold;")
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)

        self.details = QTextEdit()
        self.details.setReadOnly(True)
        self.details.setVisible(False)

        self.toggle_details_btn = QPushButton("Show details")
        self.toggle_details_btn.clicked.connect(self.toggle_details)

        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.close)

        top_layout = QVBoxLayout()
        top_layout.addWidget(self.status_label)
        top_layout.addWidget(self.progress)
        top_layout.addWidget(self.details)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.toggle_details_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)

        main_layout = QVBoxLayout()
        main_layout.addLayout(top_layout)
        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

        self.thread = InstallerThread()
        self.thread.progress.connect(self.update_status)
        self.thread.log.connect(self.append_log)
        self.thread.test_result.connect(self.handle_test_result)
        self.thread.start()

    def update_status(self, message, value):
        self.status_label.setText(message)
        self.progress.setValue(value)

    def append_log(self, text):
        self.details.append(text)

    def handle_test_result(self, ok, message):
        if ok:
            self.status_label.setText("✅ Installation successful!")
            self.status_label.setStyleSheet("color: green; font-weight: bold;")
            self.main_window.check_cli_installed()
        else:
            self.status_label.setText("❌ Installation completed with issues.")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
        self.append_log("Test result: " + message)

    def toggle_details(self):
        visible = not self.details.isVisible()
        self.details.setVisible(visible)
        self.toggle_details_btn.setText("Hide details" if visible else "Show details")


# ---------------------------------------------------------
# MAIN WINDOW
# ---------------------------------------------------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Arduino CLI Installer & Tools")

        self.current_project_path = None
        self.cli_version = "Unknown"

        # ---------- TOP: CLI STATUS + INSTALL BUTTON ----------
        self.status_label = QLabel("Arduino CLI Status: Checking…")
        self.status_label.setStyleSheet("font-weight: bold;")

        self.install_btn = QPushButton("Install / Update Arduino CLI")
        self.install_btn.clicked.connect(self.open_wizard)

        top_row = QHBoxLayout()
        top_row.addWidget(self.status_label)
        top_row.addStretch()
        top_row.addWidget(self.install_btn)

        # ---------- PROJECT BUTTONS ----------
        self.create_project_btn = QPushButton("Create New Project")
        self.create_project_btn.clicked.connect(self.create_new_project)

        self.select_project_btn = QPushButton("Select Project")
        self.select_project_btn.clicked.connect(self.open_project)

        self.show_editor_btn = QPushButton("Show Project in Editor")
        self.show_editor_btn.clicked.connect(self.open_project_in_editor)

        project_buttons = QHBoxLayout()
        project_buttons.addWidget(self.create_project_btn)
        project_buttons.addWidget(self.select_project_btn)
        project_buttons.addWidget(self.show_editor_btn)

        # ---------- CURRENT PROJECT ----------
        self.current_project_label = QLabel("Current project: None")
        self.current_project_label.setStyleSheet("font-style: italic;")

        # ---------- RECENT PROJECTS ----------
        self.recent_list = QListWidget()
        self.recent_list.itemDoubleClicked.connect(self.open_recent_project)

        # ---------- BOARD + PORT ----------
        self.board_label = QLabel("Board:")
        self.board_combo = QComboBox()

        self.port_label = QLabel("Port:")
        self.port_combo = QComboBox()

        board_port_row = QHBoxLayout()
        board_port_row.addWidget(self.board_label)
        board_port_row.addWidget(self.board_combo)
        board_port_row.addStretch()
        board_port_row.addWidget(self.port_label)
        board_port_row.addWidget(self.port_combo)

        # ---------- UPLOAD BUTTON ----------
        self.upload_btn = QPushButton("Upload to Board")
        self.upload_btn.clicked.connect(self.upload_to_board)

        # ---------- TERMINAL ----------
        self.terminal = QTextEdit()
        self.terminal.setReadOnly(True)
        self.terminal.setMinimumHeight(200)

        # ---------- MAIN LAYOUT ----------
        layout = QVBoxLayout()
        layout.addLayout(top_row)
        layout.addLayout(project_buttons)
        layout.addWidget(self.current_project_label)
        layout.addWidget(QLabel("Recent Projects:"))
        layout.addWidget(self.recent_list)
        layout.addLayout(board_port_row)
        layout.addWidget(self.upload_btn)
        layout.addWidget(QLabel("Terminal Output:"))
        layout.addWidget(self.terminal)

        self.setLayout(layout)

        # ---------- INIT ----------
        self.check_cli_installed()
        self.populate_ports()
        self.populate_boards()
        self.refresh_recent_list()

    # ---------- HELPERS ----------
    def log_terminal(self, text):
        self.terminal.append(text)
        self.terminal.moveCursor(self.terminal.textCursor().End)

    def open_wizard(self):
        dialog = WizardDialog(self)
        dialog.exec_()

    def cli_path(self):
        return os.path.join(INSTALL_DIR, CLI_NAME)

    # ---------- CLI STATUS ----------
    def check_cli_installed(self):
        cli_path = self.cli_path()

        if not os.path.exists(cli_path):
            self.status_label.setText("Arduino CLI Status: ❌ Not Installed")
            return False

        try:
            result = subprocess.run(
                [cli_path, "version"],
                capture_output=True,
                text=True
            )
            version = result.stdout.strip()
            self.cli_version = version
            self.status_label.setText(f"Arduino CLI Status: ✅ Installed ({version})")
            return True
        except Exception as e:
            self.status_label.setText("Arduino CLI Status: ❌ Error Running CLI")
            self.log_terminal(f"Error checking CLI: {e}")
            return False

    # ---------- PORTS ----------
    def populate_ports(self):
        self.port_combo.clear()
        ports = list_ports.comports()
        for p in ports:
            self.port_combo.addItem(f"{p.device} - {p.description}", p.device)
        if self.port_combo.count() == 0:
            self.port_combo.addItem("No ports detected", "")

    # ---------- BOARDS ----------
    def populate_boards(self):
        self.board_combo.clear()
        cli_ok = self.check_cli_installed()
        boards = []

        # Preferred common boards first
        common_boards = [
            ("Arduino Uno", "arduino:avr:uno"),
            ("Arduino Nano", "arduino:avr:nano"),
            ("Arduino Mega 2560", "arduino:avr:mega"),
            ("Arduino Duemilanove", "arduino:avr:diecimila"),
            ("Arduino Leonardo", "arduino:avr:leonardo"),
        ]

        for name, fqbn in common_boards:
            boards.append((name, fqbn))

        # Add autodetected boards from CLI
        if cli_ok:
            try:
                result = subprocess.run(
                    [self.cli_path(), "board", "listall"],
                    capture_output=True,
                    text=True
                )
                self.log_terminal("arduino-cli board listall:\n" + result.stdout + result.stderr)

                if result.returncode == 0:
                    for line in result.stdout.splitlines():
                        line = line.strip()
                        if not line or line.startswith("Board Name") or line.startswith("-"):
                            continue

                        parts = [p for p in line.split(" ") if p]
                        if len(parts) >= 2:
                            fqbn = parts[-1]
                            name = " ".join(parts[:-1])

                            # Avoid duplicates
                            if not any(fqbn == existing_fqbn for _, existing_fqbn in boards):
                                boards.append((name, fqbn))

            except Exception as e:
                self.log_terminal(f"Error listing boards: {e}")

        # Populate dropdown with "Name (fqbn)"
        for name, fqbn in boards:
            self.board_combo.addItem(f"{name} ({fqbn})", fqbn)

    # ---------- RECENT PROJECTS ----------
    def load_recent_projects(self):
        if not os.path.exists(RECENT_FILE):
            return []
        try:
            with open(RECENT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []

    def save_recent_projects(self, projects):
        try:
            with open(RECENT_FILE, "w", encoding="utf-8") as f:
                json.dump(projects, f, indent=2)
        except Exception:
            pass

    def add_recent_project(self, path):
        projects = self.load_recent_projects()
        projects = [p for p in projects if p != path]
        projects.insert(0, path)
        projects = projects[:MAX_RECENT]
        self.save_recent_projects(projects)
        self.refresh_recent_list()

    def refresh_recent_list(self):
        self.recent_list.clear()
        for p in self.load_recent_projects():
            item = QListWidgetItem(os.path.basename(p) or p)
            item.setToolTip(p)
            item.setData(Qt.UserRole, p)
            self.recent_list.addItem(item)

    def open_recent_project(self, item):
        folder = item.data(Qt.UserRole)
        self.open_project_folder(folder)

    # ---------- PROJECT MANAGEMENT ----------
    def create_new_project(self):
        choice = QMessageBox.question(
            self,
            "Create New Project",
            "Yes = Example sketch (Blink + floats)\nNo = Empty sketch",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        example = (choice == QMessageBox.Yes)

        folder = QFileDialog.getExistingDirectory(self, "Select Project Folder")
        if not folder:
            return

        project_name = os.path.basename(folder.rstrip("/\\"))
        if not project_name:
            project_name = "ArduinoProject"

        ino_path = os.path.join(folder, project_name + ".ino")

        if example:
            sketch = """\
void setup() {
    Serial.begin(9600);
    pinMode(LED_BUILTIN, OUTPUT);
    delay(500);
    Serial.println("Sending float values...");
}

void loop() {
    float a = 1.23;
    float b = 4.56;
    float c = 7.89;

    Serial.print(a, 3);
    Serial.print(", ");
    Serial.print(b, 3);
    Serial.print(", ");
    Serial.println(c, 3);

    digitalWrite(LED_BUILTIN, HIGH);
    delay(500);
    digitalWrite(LED_BUILTIN, LOW);
    delay(500);
}
"""
        else:
            sketch = "void setup() {}\n\nvoid loop() {}\n"

        try:
            with open(ino_path, "w", encoding="utf-8") as f:
                f.write(sketch)
            self.open_project_folder(folder)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create project:\n{e}")
            self.log_terminal(f"Failed to create project: {e}")

    def open_project(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Project Folder")
        if folder:
            self.open_project_folder(folder)

    def open_project_folder(self, folder):
        if not os.path.exists(folder):
            QMessageBox.warning(self, "Error", "Project folder does not exist.")
            return

        ino_files = [f for f in os.listdir(folder) if f.endswith(".ino")]
        if not ino_files:
            QMessageBox.warning(self, "Error", "No .ino file found in this folder.")
            return

        ino_path = os.path.join(folder, ino_files[0])
        self.current_project_path = folder
        self.current_project_label.setText(f"Current project: {ino_path}")
        self.add_recent_project(folder)
        self.log_terminal(f"Opened project:\n{ino_path}")

    # ---------- OPEN IN EDITOR ----------
    def open_project_in_editor(self):
        if not self.current_project_path:
            QMessageBox.warning(self, "No Project", "No project is currently selected.")
            return

        ino_files = [f for f in os.listdir(self.current_project_path) if f.endswith(".ino")]
        if not ino_files:
            QMessageBox.warning(self, "Error", "No .ino file found in this project.")
            return

        ino_path = os.path.join(self.current_project_path, ino_files[0])

        # SHIFT override: if default exists and SHIFT is NOT pressed, use default directly
        pref = load_editor_preference()
        shift_pressed = QApplication.keyboardModifiers() & Qt.ShiftModifier

        if pref and not shift_pressed:
            editor_path = pref.get("default_editor_path")
            editor_name = pref.get("default_editor_name", "Editor")
            if editor_path:
                self.open_in_editor(editor_path, ino_path, editor_name)
                return

        # No default or SHIFT pressed → show dialog
        dialog = EditorSelectionDialog(self, ino_path=ino_path)
        if dialog.exec_() == QDialog.Accepted:
            editor_name = dialog.selected_editor_name
            editor_path = dialog.selected_editor_path

            if dialog.set_as_default:
                save_editor_preference(editor_name, editor_path)

            self.open_in_editor(editor_path, ino_path, editor_name)

    def open_in_editor(self, editor_path, ino_path, editor_name="Editor"):
        try:
            subprocess.Popen([editor_path, ino_path])
            self.log_terminal(f"Opened project in {editor_name}: {ino_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open editor:\n{e}")
            self.log_terminal(f"Failed to open editor: {e}")

    # ---------- UPLOAD ----------
    def upload_to_board(self):
        self.terminal.clear()

        if not self.check_cli_installed():
            QMessageBox.warning(self, "Arduino CLI", "Arduino CLI must be installed first.")
            return

        if not self.current_project_path:
            QMessageBox.warning(self, "No Project", "No project is open.")
            return

        project_dir = self.current_project_path
        ino_files = [f for f in os.listdir(project_dir) if f.endswith(".ino")]
        if not ino_files:
            QMessageBox.warning(self, "Error", "No .ino file found in the project.")
            return

        ino_path = os.path.join(project_dir, ino_files[0])

        port = self.port_combo.currentData()
        if not port:
            QMessageBox.warning(self, "Port", "No valid COM port selected.")
            return

        fqbn = self.board_combo.currentData()
        if not fqbn:
            QMessageBox.warning(self, "Board", "No valid board selected.")
            return

        cli_path = self.cli_path()

        # Compile
        self.log_terminal(f"Compiling {ino_path} for {fqbn} in {project_dir}…")
        try:
            compile_proc = subprocess.run(
                [cli_path, "compile", "--fqbn", fqbn, "."],
                cwd=project_dir,
                capture_output=True,
                text=True
            )
        except Exception as e:
            self.log_terminal(f"Compile error: {e}")
            QMessageBox.critical(self, "Compile Error", f"Failed to run compile:\n{e}")
            return

        self.log_terminal("Compile stdout:\n" + compile_proc.stdout)
        self.log_terminal("Compile stderr:\n" + compile_proc.stderr)

        if compile_proc.returncode != 0:
            QMessageBox.critical(
                self,
                "Compile Error",
                f"Compile failed:\n{compile_proc.stdout}\n{compile_proc.stderr}"
            )
            return

        # Upload
        self.log_terminal(f"Uploading to {port}…")
        try:
            upload_proc = subprocess.run(
                [cli_path, "upload", "-p", port, "--fqbn", fqbn, "."],
                cwd=project_dir,
                capture_output=True,
                text=True
            )
        except Exception as e:
            self.log_terminal(f"Upload error: {e}")
            QMessageBox.critical(self, "Upload Error", f"Failed to run upload:\n{e}")
            return

        self.log_terminal("Upload stdout:\n" + upload_proc.stdout)
        self.log_terminal("Upload stderr:\n" + upload_proc.stderr)

        if upload_proc.returncode != 0:
            QMessageBox.critical(
                self,
                "Upload Error",
                f"Upload failed:\n{upload_proc.stdout}\n{upload_proc.stderr}"
            )
        else:
            QMessageBox.information(
                self,
                "Upload Successful",
                f"Sketch uploaded to {port} successfully."
            )


# ---------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(900, 700)
    window.show()
    sys.exit(app.exec_())
