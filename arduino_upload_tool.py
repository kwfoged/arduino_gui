import os
import sys
import time
import zipfile
import subprocess
import tempfile
import json
import urllib.request
import ssl
import shutil

import requests
import certifi

from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout,
    QDialog, QLabel, QProgressBar, QTextEdit, QHBoxLayout,
    QMessageBox, QComboBox, QFileDialog,
    QRadioButton, QButtonGroup, QFrame, QSpacerItem,
    QSizePolicy, QLineEdit, QGroupBox, QTreeWidget,
    QTreeWidgetItem, QStyle
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPoint, QSize, QTimer
from PyQt5.QtGui import QIcon

from serial.tools import list_ports  # pip install pyserial


# ---------------------------------------------------------
# CONFIG
# ---------------------------------------------------------
ARDUINO_CLI_URL = "https://downloads.arduino.cc/arduino-cli/arduino-cli_latest_Windows_64bit.zip"
TEMP_DIR = tempfile.gettempdir()
DOWNLOAD_PATH = os.path.join(TEMP_DIR, "arduino-cli.zip")
INSTALL_DIR = os.path.expanduser("~/.arduino-cli")
CLI_NAME = "arduino-cli.exe"

WINLIBS_DIR = os.path.join(os.path.expanduser("~"), ".winlibs")
WINLIBS_ZIP = os.path.join(TEMP_DIR, "winlibs.zip")

SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".arduino_gui_settings.json")
MAX_RECENT = 10

DEFAULT_SETTINGS = {
    "recent_projects": [],
    "editor": None,
    "gui": {
        "window_size": [900, 700],
        "window_pos": None,
        "last_board": "arduino:avr:uno",
        "last_port": None,
        "sim_time": 10
    }
}


# ---------------------------------------------------------
# SETTINGS HELPERS
# ---------------------------------------------------------
def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        return json.loads(json.dumps(DEFAULT_SETTINGS))

    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return json.loads(json.dumps(DEFAULT_SETTINGS))

    if not isinstance(data, dict):
        data = {}

    for key, default_value in DEFAULT_SETTINGS.items():
        if key not in data:
            data[key] = json.loads(json.dumps(default_value))

    if not isinstance(data.get("gui"), dict):
        data["gui"] = json.loads(json.dumps(DEFAULT_SETTINGS["gui"]))
    else:
        for key, default_value in DEFAULT_SETTINGS["gui"].items():
            if key not in data["gui"]:
                data["gui"][key] = default_value

    if not isinstance(data.get("recent_projects"), list):
        data["recent_projects"] = []

    return data


def save_settings(settings):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------
# EDITOR DETECTION
# ---------------------------------------------------------
def detect_editors():
    editors = []

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

    vscode_paths = [
        rf"C:\Users\{os.getlogin()}\AppData\Local\Programs\Microsoft VS Code\Code.exe",
        r"C:\Program Files\Microsoft VS Code\Code.exe",
    ]
    for p in vscode_paths:
        if os.path.exists(p):
            editors.append(("VS Code", p))
            break

    notepadpp_paths = [
        r"C:\Program Files\Notepad++\notepad++.exe",
        r"C:\Program Files (x86)\Notepad++\notepad++.exe",
    ]
    for p in notepadpp_paths:
        if os.path.exists(p):
            editors.append(("Notepad++", p))
            break

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

        title1 = QLabel("Select an editor")
        title1.setStyleSheet("font-weight: bold;")
        layout.addWidget(title1)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.HLine)
        layout.addWidget(sep1)

        self.editor_group = QButtonGroup(self)
        self.editor_group.setExclusive(True)

        for name, path in self.editors:
            rb = QRadioButton(name)
            rb.editor_name = name
            rb.editor_path = path
            self.editor_group.addButton(rb)
            layout.addWidget(rb)

        buttons = self.editor_group.buttons()
        vs_code_button = next((b for b in buttons if b.editor_name == "VS Code"), None)
        if vs_code_button:
            vs_code_button.setChecked(True)
        elif buttons:
            buttons[0].setChecked(True)

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
# INSTALLER THREAD (UPDATED FOR ~/.winlibs)
# ---------------------------------------------------------
class InstallerThread(QThread):
    progress = pyqtSignal(str, int)
    log = pyqtSignal(str)
    test_result = pyqtSignal(bool, str)

    def run(self):
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

        self.progress.emit("Installing…", 45)
        self.log.emit("Installation step complete.")

        self.progress.emit("Updating PATH…", 50)
        self.log.emit("Updating PATH environment variable…")
        self.update_path(INSTALL_DIR)

        self.progress.emit("Testing Arduino CLI…", 55)
        self.log.emit("Starting post-installation tests…")
        ok, message = self.test_installation()
        self.log.emit(message)

        if ok:
            self.progress.emit("Installation successful!", 60)
        else:
            self.progress.emit("Installation completed with issues.", 100)

        self.test_result.emit(ok, message)

        self.log.emit("Starting Downloads of WinLibs")
        self.get_sw()

    def download_with_progress(self, url, dest, start_pct=60, end_pct=90):
        try:
            if os.path.exists(dest):
                try:
                    os.remove(dest)
                    self.log.emit("Removed old winlibs.zip")
                except Exception as e:
                    self.log.emit(f"Could not remove old winlibs.zip: {e}")

            self.progress.emit("Starting WinLibs download…", start_pct)

            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as response:
                total = int(response.headers.get("Content-Length", 0))
                downloaded = 0

                chunk_size = 1024 * 256
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

    def get_sw(self):
        try:
            self.log.emit("Fetching latest WinLibs release info...")

            ssl._create_default_https_context = ssl._create_unverified_context

            api_url = "https://api.github.com/repos/brechtsanders/winlibs_mingw/releases/latest"
            req = urllib.request.Request(api_url, headers={"User-Agent": "Mozilla/5.0"})
            data = urllib.request.urlopen(req).read()
            release = json.loads(data)

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

            ok = self.download_with_progress(asset_url, WINLIBS_ZIP, 60, 90)
            if not ok:
                self.log.emit("Failed to download WinLibs.")
                return

            self.log.emit("Download complete. Extracting…")
            self.progress.emit("Extracting WinLibs…", 90)

            os.makedirs(WINLIBS_DIR, exist_ok=True)

            with zipfile.ZipFile(WINLIBS_ZIP, 'r') as zip_ref:
                zip_ref.extractall(WINLIBS_DIR)

            os.remove(WINLIBS_ZIP)

            bin_path = None
            for root, dirs, files in os.walk(WINLIBS_DIR):
                if root.endswith("bin"):
                    bin_path = root
                    break

            if bin_path:
                os.environ["PATH"] += os.pathsep + bin_path
                self.log.emit(f"Added to PATH:\n{bin_path}")

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

        # Start installer thread
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
        self.main_window.check_winlibs_installed()

    def toggle_details(self):
        visible = not self.details.isVisible()
        self.details.setVisible(visible)
        self.toggle_details_btn.setText("Hide details" if visible else "Show details")


# ---------------------------------------------------------
# SIMULATOR THREAD (UPDATED FOR ~/.winlibs)
# ---------------------------------------------------------
class SimulatorThread(QThread):
    output = pyqtSignal(str)

    def __init__(self, project_path, sim_seconds):
        super().__init__()
        self.project_path = project_path
        self.sim_seconds = sim_seconds

    def find_winlibs_make(self):
        # First try PATH
        make_path = shutil.which("make")
        if make_path:
            return make_path

        # Then try ~/.winlibs
        for root, dirs, files in os.walk(WINLIBS_DIR):
            if "make.exe" in files:
                return os.path.join(root, "make.exe")

        return None

    def run(self):
        make_path = self.find_winlibs_make()
        if not make_path:
            self.output.emit(
                "'make' not found.\n"
                "WinLibs may not be installed or PATH not updated.\n"
            )
            return

        self.output.emit(f"Using make at: {make_path}\n")
        self.output.emit("Building simulator (make)...\n")

        try:
            proc = subprocess.Popen(
                [make_path],
                cwd=self.project_path,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )
        except Exception as e:
            self.output.emit(f"Error starting make: {e}\n")
            return

        for line in proc.stdout:
            self.output.emit(line)
        proc.wait()

        if proc.returncode != 0:
            self.output.emit("\nBuild failed.\n")
            return

        exe = os.path.join(self.project_path, "sim.exe")
        if not os.path.exists(exe):
            self.output.emit("sim.exe not found after build.\n")
            return

        self.output.emit(f"\nRunning simulator for {self.sim_seconds} seconds...\n")

        start_time = time.time()

        try:
            proc = subprocess.Popen(
                [exe],
                cwd=self.project_path,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )
        except Exception as e:
            self.output.emit(f"Error starting simulator: {e}\n")
            return

        while True:
            if proc.poll() is not None:
                break

            line = proc.stdout.readline()
            if line:
                self.output.emit(line)

            if time.time() - start_time >= self.sim_seconds:
                proc.terminate()
                self.output.emit("\nSimulation time limit reached.\n")
                break

        proc.wait()
        self.output.emit(f"\nSimulator exited with code {proc.returncode}\n")


# ---------------------------------------------------------
# GROUPBOX WITH STANDARD ICON
# ---------------------------------------------------------
def make_icon_group(parent, title: str, standard_pixmap: QStyle.StandardPixmap, inner_layout):
    box = QGroupBox()
    outer = QVBoxLayout(box)

    header = QHBoxLayout()
    icon_label = QLabel()
    icon = parent.style().standardIcon(standard_pixmap)
    icon_label.setPixmap(icon.pixmap(16, 16))

    text_label = QLabel(title)
    text_label.setStyleSheet("font-weight: bold;")

    header.addWidget(icon_label)
    header.addWidget(text_label)
    header.addStretch()

    outer.addLayout(header)
    outer.addLayout(inner_layout)

    return box


# ---------------------------------------------------------
# MAIN WINDOW
# ---------------------------------------------------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Arduino CLI Installer & Tools")

        self.settings = load_settings()

        self.current_project_path = None
        self.cli_version = "Unknown"
        self.sim_thread = None

        # -------------------------
        # Installed Software Status
        # -------------------------
        self.status_label = QLabel("Arduino CLI Status: Checking…")
        self.status_label.setStyleSheet("font-weight: bold;")

        self.winlibs_status_label = QLabel("WinLibs Status: Checking…")
        self.winlibs_status_label.setStyleSheet("font-weight: bold;")

        self.install_btn = QPushButton("Install / Update Arduino CLI")
        self.install_btn.clicked.connect(self.open_wizard)

        # -------------------------
        # Project Section
        # -------------------------
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

        self.current_project_label = QLabel("Current project: None")
        self.current_project_label.setStyleSheet("font-style: italic;")

        # Recent projects
        self.recent_list = QTreeWidget()
        self.recent_list.setColumnCount(2)
        self.recent_list.setHeaderLabels(["Project Name", "Path"])
        self.recent_list.itemDoubleClicked.connect(self.open_recent_project)

        self.recent_toggle_btn = QPushButton("▼ Recent Projects")
        self.recent_toggle_btn.setCheckable(True)
        self.recent_toggle_btn.setChecked(True)
        self.recent_toggle_btn.clicked.connect(self.toggle_recent_projects)

        self.recent_container = QWidget()
        recent_layout = QVBoxLayout()
        recent_layout.setContentsMargins(0, 0, 0, 0)
        recent_layout.addWidget(self.recent_list)
        self.recent_container.setLayout(recent_layout)

        # -------------------------
        # Upload Section
        # -------------------------
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

        self.upload_btn = QPushButton("Upload to Board")
        self.upload_btn.clicked.connect(self.upload_to_board)

        # -------------------------
        # Simulation Section
        # -------------------------
        self.run_sim_btn = QPushButton("Run Simulator")
        self.run_sim_btn.clicked.connect(self.run_simulator)

        self.sim_time_label = QLabel("Sim Time (sec):")
        self.sim_time_edit = QLineEdit(str(self.settings["gui"].get("sim_time", 10)))
        self.sim_time_edit.setFixedWidth(60)

        # -------------------------
        # Terminal Output
        # -------------------------
        self.terminal = QTextEdit()
        self.terminal.setReadOnly(True)
        self.terminal.setMinimumHeight(200)

        # -------------------------
        # Group Layouts
        # -------------------------
        sw_layout = QVBoxLayout()
        sw_layout.addWidget(self.status_label)
        sw_layout.addWidget(self.winlibs_status_label)
        sw_layout.addWidget(self.install_btn)
        sw_group = make_icon_group(self, "Installed Software Status", QStyle.SP_ComputerIcon, sw_layout)

        project_layout = QVBoxLayout()
        project_layout.addLayout(project_buttons)
        project_layout.addWidget(self.current_project_label)
        project_layout.addWidget(self.recent_toggle_btn)
        project_layout.addWidget(self.recent_container)
        project_group = make_icon_group(self, "Project", QStyle.SP_DirIcon, project_layout)

        upload_layout = QVBoxLayout()
        upload_layout.addLayout(board_port_row)
        upload_layout.addWidget(self.upload_btn)
        upload_group = make_icon_group(self, "Upload Code", QStyle.SP_ArrowUp, upload_layout)

        sim_layout = QVBoxLayout()
        sim_layout.addWidget(self.sim_time_label)
        sim_layout.addWidget(self.sim_time_edit)
        sim_layout.addWidget(self.run_sim_btn)
        sim_group = make_icon_group(self, "Simulation", QStyle.SP_MediaPlay, sim_layout)

        upload_sim_row = QHBoxLayout()
        upload_sim_row.addWidget(upload_group)
        upload_sim_row.addWidget(sim_group)

        terminal_layout = QVBoxLayout()
        terminal_layout.addWidget(self.terminal)
        terminal_group = make_icon_group(self, "Terminal Output", QStyle.SP_ComputerIcon, terminal_layout)

        # -------------------------
        # Main Layout
        # -------------------------
        layout = QVBoxLayout()
        layout.addWidget(sw_group)
        layout.addWidget(project_group)
        layout.addLayout(upload_sim_row)
        layout.addWidget(terminal_group)
        self.setLayout(layout)

        # -------------------------
        # Initialize State
        # -------------------------
        self.check_cli_installed()
        self.check_winlibs_installed()
        self.ensure_winlibs_in_path()   # <-- IMPORTANT FIX
        self.populate_ports()
        self.populate_boards()
        self.refresh_recent_list()
        self.restore_gui_state()

        # Auto-refresh COM ports every 2 seconds
        self.port_timer = QTimer(self)
        self.port_timer.timeout.connect(self.populate_ports)
        self.port_timer.start(2000)


    # ---------------------------------------------------------
    # ENSURE WINLIBS IS IN PATH
    # ---------------------------------------------------------
    def ensure_winlibs_in_path(self):
        if not os.path.exists(WINLIBS_DIR):
            return

        for root, dirs, files in os.walk(WINLIBS_DIR):
            if "make.exe" in files:
                bin_path = root
                if bin_path not in os.environ["PATH"]:
                    os.environ["PATH"] += os.pathsep + bin_path
                    self.log_terminal(f"WinLibs added to PATH: {bin_path}")
                return

    # ---------------------------------------------------------
    # GUI STATE
    # ---------------------------------------------------------
    def restore_gui_state(self):
        gui = self.settings.get("gui", {})
        size = gui.get("window_size", [900, 700])
        if isinstance(size, list) and len(size) == 2:
            self.resize(QSize(size[0], size[1]))

        pos = gui.get("window_pos")
        if isinstance(pos, list) and len(pos) == 2:
            self.move(QPoint(pos[0], pos[1]))

        last_board = gui.get("last_board")
        if last_board:
            idx = self.board_combo.findData(last_board)
            if idx >= 0:
                self.board_combo.setCurrentIndex(idx)

        last_port = gui.get("last_port")
        if last_port:
            idx = self.port_combo.findData(last_port)
            if idx >= 0:
                self.port_combo.setCurrentIndex(idx)

        sim_time = gui.get("sim_time")
        if sim_time:
            self.sim_time_edit.setText(str(sim_time))

    def save_gui_state(self):
        gui = self.settings.get("gui", {})
        size = self.size()
        pos = self.pos()
        gui["window_size"] = [size.width(), size.height()]
        gui["window_pos"] = [pos.x(), pos.y()]
        gui["last_board"] = self.board_combo.currentData()
        gui["last_port"] = self.port_combo.currentData()
        try:
            gui["sim_time"] = int(self.sim_time_edit.text())
        except ValueError:
            gui["sim_time"] = 10
        self.settings["gui"] = gui
        save_settings(self.settings)

    def closeEvent(self, event):
        self.save_gui_state()
        super().closeEvent(event)

    # ---------------------------------------------------------
    # TERMINAL LOGGING
    # ---------------------------------------------------------
    def log_terminal(self, text):
        self.terminal.append(text)
        self.terminal.moveCursor(self.terminal.textCursor().End)

    # ---------------------------------------------------------
    # INSTALLER WIZARD
    # ---------------------------------------------------------
    def open_wizard(self):
        dialog = WizardDialog(self)
        dialog.exec_()

    # ---------------------------------------------------------
    # CLI PATH
    # ---------------------------------------------------------
    def cli_path(self):
        return os.path.join(INSTALL_DIR, CLI_NAME)

    # ---------------------------------------------------------
    # CLI STATUS CHECK
    # ---------------------------------------------------------
    def check_cli_installed(self):
        cli_path = self.cli_path()

        if not os.path.exists(cli_path):
            self.status_label.setText("Arduino CLI Status: ❌ Not Installed")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
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
            self.status_label.setStyleSheet("color: green; font-weight: bold;")
            return True
        except Exception as e:
            self.status_label.setText("Arduino CLI Status: ❌ Error Running CLI")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
            self.log_terminal(f"Error checking CLI: {e}")
            return False

    # ---------------------------------------------------------
    # WINLIBS STATUS CHECK
    # ---------------------------------------------------------
    def check_winlibs_installed(self):
        # 1. Try PATH first
        make_path = shutil.which("make")

        # 2. If not found, search recursively inside ~/.winlibs
        if not make_path:
            for root, dirs, files in os.walk(WINLIBS_DIR):
                if "make.exe" in files:
                    make_path = os.path.join(root, "make.exe")
                    break

        # 3. Update GUI
        if make_path:
            self.winlibs_status_label.setText(f"WinLibs Status: ✅ Installed ({make_path})")
            self.winlibs_status_label.setStyleSheet("color: green; font-weight: bold;")
        else:
            self.winlibs_status_label.setText("WinLibs Status: ❌ Not Installed")
            self.winlibs_status_label.setStyleSheet("color: red; font-weight: bold;")

    # ---------------------------------------------------------
    # PORTS
    # ---------------------------------------------------------
    def populate_ports(self):
        self.port_combo.clear()
        try:
            ports = list_ports.comports()
            for p in ports:
                self.port_combo.addItem(f"{p.device} ({p.description})", p.device)
        except Exception as e:
            self.log_terminal(f"Error listing ports: {e}")

    # ---------------------------------------------------------
    # BOARDS
    # ---------------------------------------------------------
    def populate_boards(self):
        self.board_combo.clear()
    
        boards = [
            ("Arduino Uno", "arduino:avr:uno"),
            ("Arduino Nano", "arduino:avr:nano"),
            ("Arduino Mega 2560", "arduino:avr:mega"),
            ("Arduino Duemilanove", "arduino:avr:diecimila"),
            ("Arduino Leonardo", "arduino:avr:leonardo"),
        ]
    
        # Add common boards first
        for name, fqbn in boards:
            self.board_combo.addItem(f"{name} ({fqbn})", fqbn)
    
        # Optional: dynamically load ALL boards from Arduino CLI
        cli = self.cli_path()
        if os.path.exists(cli):
            try:
                result = subprocess.run(
                    [cli, "board", "listall"],
                    capture_output=True,
                    text=True
                )
                for line in result.stdout.splitlines():
                    if ":" in line and " " in line:
                        parts = line.strip().split()
                        fqbn = parts[-1]
                        name = " ".join(parts[:-1])
                        if fqbn not in [b[1] for b in boards]:
                            self.board_combo.addItem(f"{name} ({fqbn})", fqbn)
            except Exception:
                pass


    # ---------------------------------------------------------
    # RECENT PROJECTS
    # ---------------------------------------------------------
    def refresh_recent_list(self):
        self.recent_list.clear()

        recent = self.settings.get("recent_projects", [])
        cleaned = []

        for entry in recent:
            if isinstance(entry, str):
                path = entry
                name = os.path.basename(path.rstrip("/\\"))
                entry = {"name": name, "path": path}

            if not isinstance(entry, dict):
                continue

            name = entry.get("name")
            path = entry.get("path")

            if not name or not path:
                continue

            if not os.path.exists(path):
                continue

            cleaned.append(entry)
            item = QTreeWidgetItem([name, path])
            self.recent_list.addTopLevelItem(item)

        self.settings["recent_projects"] = cleaned
        save_settings(self.settings)

    def add_recent_project(self, name, path):
        recent = self.settings.get("recent_projects", [])

        recent = [r for r in recent if r.get("path") != path]
        recent.insert(0, {"name": name, "path": path})
        recent = recent[:MAX_RECENT]

        self.settings["recent_projects"] = recent
        save_settings(self.settings)
        self.refresh_recent_list()

    # ---------------------------------------------------------
    # COLLAPSIBLE RECENT PROJECTS
    # ---------------------------------------------------------
    def toggle_recent_projects(self):
        visible = self.recent_toggle_btn.isChecked()
        self.recent_container.setVisible(visible)
        self.recent_toggle_btn.setText("▼ Recent Projects" if visible else "► Recent Projects")

    # ---------------------------------------------------------
    # PROJECT MANAGEMENT
    # ---------------------------------------------------------
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

        project_name = os.path.basename(folder.rstrip("/\\")) or "ArduinoProject"

        ino_path = os.path.join(folder, project_name + ".ino")
        main_cpp_path = os.path.join(folder, "main.cpp")
        main_h_path = os.path.join(folder, "main.h")
        sim_dir = os.path.join(folder, "sim")
        sim_h_path = os.path.join(sim_dir, "sim_arduino.h")
        sim_cpp_path = os.path.join(sim_dir, "sim_arduino.cpp")
        makefile_path = os.path.join(folder, "Makefile")

        os.makedirs(sim_dir, exist_ok=True)

        # -------------------------
        # main.h
        # -------------------------
        main_h = """#pragma once

void sim_setup();
void sim_loop();
"""

        # -------------------------
        # main.cpp + .ino
        # -------------------------
        if example:
            main_cpp = r"""#include "main.h"

#ifdef SIMULATION
#include "sim/sim_arduino.h"
#else
#include <Arduino.h>
#endif

void sim_setup() {
    Serial.begin(115200);
    Serial.println("Sending float values...");
}

void sim_loop() {
    float a = 1.23f;
    float b = 4.56f;
    float c = 7.89f;

    Serial.print(a);
    Serial.print(", ");
    Serial.print(b);
    Serial.print(", ");
    Serial.println(c);

    digitalWrite(LED_BUILTIN, 1);
    delay(500);
    digitalWrite(LED_BUILTIN, 0);
    delay(500);
}

#ifdef SIMULATION
int main() {
    sim_setup();
    for (int i = 0; i < 10; ++i) {
        sim_loop();
    }
    return 0;
}
#endif
"""
            ino = f"""#include "main.h"

void setup() {{
    sim_setup();
}}

void loop() {{
    sim_loop();
}}
"""
        else:
            main_cpp = r"""#include "main.h"

#ifdef SIMULATION
#include "sim/sim_arduino.h"
#else
#include <Arduino.h>
#endif

void sim_setup() {
    // TODO: add setup simulation / Arduino code
}

void sim_loop() {
    // TODO: add loop simulation / Arduino code
}

#ifdef SIMULATION
int main() {
    sim_setup();
    while (true) {
        sim_loop();
    }
    return 0;
}
#endif
"""
            ino = f"""#include "main.h"

void setup() {{
    sim_setup();
}}

void loop() {{
    sim_loop();
}}
"""

        # -------------------------
        # sim_arduino.h
        # -------------------------
        sim_h = """#pragma once
#include <iostream>
#include <chrono>
#include <thread>

void delay(unsigned long ms);
void digitalWrite(int pin, int value);

struct SerialSim {
    void begin(unsigned long baud) {
        std::cout << "[Serial.begin] baud=" << baud << std::endl;
    }

    template<typename T>
    void print(const T& v) { std::cout << v; }

    template<typename T>
    void println(const T& v) { std::cout << v << std::endl; }

    void println() { std::cout << std::endl; }

    void write(uint8_t b) {
        std::cout << "[Serial.write] " << (int)b << std::endl;
    }
};


extern SerialSim Serial;

#define LED_BUILTIN 13
"""

        # -------------------------
        # sim_arduino.cpp
        # -------------------------
        sim_cpp = """#include "sim_arduino.h"

SerialSim Serial;

void delay(unsigned long ms) {
    std::this_thread::sleep_for(std::chrono::milliseconds(ms));
}

void digitalWrite(int pin, int value) {
    std::cout << "[digitalWrite] pin " << pin << " = " << value << std::endl;
}
"""

        # -------------------------
        # Makefile
        # -------------------------
        makefile = (
            "CXX = g++\n"
            "CXXFLAGS = -std=c++17 -O2 -DSIMULATION\n"
            "TARGET = sim.exe\n\n"
            "SRC = main.cpp sim/sim_arduino.cpp\n\n"
            "all: $(TARGET)\n\n"
            "$(TARGET): $(SRC)\n"
            "\t$(CXX) $(CXXFLAGS) -o $(TARGET) $(SRC)\n\n"
            "clean:\n"
            "\t-del $(TARGET) 2> NUL\n"
        )

        # -------------------------
        # Write all files
        # -------------------------
        try:
            with open(main_h_path, "w", encoding="utf-8") as f:
                f.write(main_h)
            with open(main_cpp_path, "w", encoding="utf-8") as f:
                f.write(main_cpp)
            with open(ino_path, "w", encoding="utf-8") as f:
                f.write(ino)
            with open(sim_h_path, "w", encoding="utf-8") as f:
                f.write(sim_h)
            with open(sim_cpp_path, "w", encoding="utf-8") as f:
                f.write(sim_cpp)
            with open(makefile_path, "w", encoding="utf-8") as f:
                f.write(makefile)

            self.current_project_path = folder
            self.current_project_label.setText(f"Current project: {folder}")
            self.add_recent_project(project_name, folder)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create project:\n{e}")
            self.log_terminal(f"Failed to create project: {e}")

    # ---------------------------------------------------------
    # OPEN EXISTING PROJECT
    # ---------------------------------------------------------
    def open_project(self):
        folder = QFileDialog.getExistingDirectory(self, "Select existing project folder")
        if not folder:
            return

        self.current_project_path = folder
        name = os.path.basename(folder)
        self.current_project_label.setText(f"Current project: {folder}")

        self.add_recent_project(name, folder)

    def open_recent_project(self, item, column):
        path = item.text(1)
        if not os.path.exists(path):
            QMessageBox.warning(self, "Missing project", "This project folder no longer exists.")
            self.refresh_recent_list()
            return

        self.current_project_path = path
        self.current_project_label.setText(f"Current project: {path}")

    # ---------------------------------------------------------
    # EDITOR OPEN
    # ---------------------------------------------------------
    def open_project_in_editor(self):
        if not self.current_project_path:
            QMessageBox.warning(self, "No project", "No current project selected.")
            return

        ino_files = [f for f in os.listdir(self.current_project_path) if f.endswith(".ino")]
        if not ino_files:
            QMessageBox.warning(self, "No .ino file", "No .ino file found in the current project.")
            return

        ino_path = os.path.join(self.current_project_path, ino_files[0])

        editor_pref = self.settings.get("editor")
        if editor_pref and os.path.exists(editor_pref.get("path", "")):
            editor_path = editor_pref["path"]
            try:
                subprocess.Popen([editor_path, ino_path])
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to open editor: {e}")
            return

        dialog = EditorSelectionDialog(self, ino_path)
        if dialog.exec_() == QDialog.Accepted:
            name = dialog.selected_editor_name
            path = dialog.selected_editor_path
            if dialog.set_as_default:
                self.settings["editor"] = {"name": name, "path": path}
                save_settings(self.settings)
            try:
                subprocess.Popen([path, ino_path])
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to open editor: {e}")

    # ---------------------------------------------------------
    # UPLOAD TO BOARD
    # ---------------------------------------------------------
    def upload_to_board(self):
        if not self.current_project_path:
            QMessageBox.warning(self, "No project", "No current project selected.")
            return

        cli_path = self.cli_path()
        if not os.path.exists(cli_path):
            QMessageBox.warning(self, "CLI not installed", "Arduino CLI is not installed.")
            return

        board = self.board_combo.currentData()
        port = self.port_combo.currentData()
        if not board or not port:
            QMessageBox.warning(self, "Missing info", "Please select board and port.")
            return

        ino_files = [f for f in os.listdir(self.current_project_path) if f.endswith(".ino")]
        if not ino_files:
            QMessageBox.warning(self, "No .ino file", "No .ino file found in the current project.")
            return

        ino_path = os.path.join(self.current_project_path, ino_files[0])

        cmd = [
            cli_path, "compile", "--fqbn", board, self.current_project_path
        ]
        self.log_terminal("Running: " + " ".join(cmd))
        try:
            result = subprocess.run(cmd, capture_output=True, text=True)
            self.log_terminal(result.stdout)
            self.log_terminal(result.stderr)
            if result.returncode != 0:
                QMessageBox.warning(self, "Compile error", "Compilation failed. See terminal output.")
                return
        except Exception as e:
            self.log_terminal(f"Error compiling: {e}")
            return

        cmd = [
            cli_path, "upload", "-p", port, "--fqbn", board, self.current_project_path
        ]
        self.log_terminal("Running: " + " ".join(cmd))
        try:
            result = subprocess.run(cmd, capture_output=True, text=True)
            self.log_terminal(result.stdout)
            self.log_terminal(result.stderr)
            if result.returncode != 0:
                QMessageBox.warning(self, "Upload error", "Upload failed. See terminal output.")
            else:
                QMessageBox.information(self, "Success", "Upload completed.")
        except Exception as e:
            self.log_terminal(f"Error uploading: {e}")

    # ---------------------------------------------------------
    # RUN SIMULATOR
    # ---------------------------------------------------------
    def run_simulator(self):
        if not self.current_project_path:
            QMessageBox.warning(self, "No project", "No current project selected.")
            return

        try:
            sim_seconds = int(self.sim_time_edit.text())
        except ValueError:
            QMessageBox.warning(self, "Invalid time", "Simulation time must be an integer.")
            return

        if self.sim_thread and self.sim_thread.isRunning():
            QMessageBox.warning(self, "Simulator running", "Simulator is already running.")
            return

        self.sim_thread = SimulatorThread(self.current_project_path, sim_seconds)
        self.sim_thread.output.connect(self.log_terminal)
        self.sim_thread.start()


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
