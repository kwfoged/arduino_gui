import sys
import struct
import time
import csv
import os

from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import QTextCursor
import pyqtgraph as pg
import serial
import serial.tools.list_ports


MAX_CHANNELS = 10
FLOAT_SIZE = 4

# Embedded config (no external file)
DEFAULT_CONFIG = {
    "theme": "dark"
}

HEADER1 = 0xAA
HEADER2 = 0x55

DARK_STYLESHEET = """
    QWidget {
        background-color: #1e1e1e;
        color: #d4d4d4;
        font-family: Segoe UI, sans-serif;
        font-size: 10pt;
    }
    QGroupBox {
        border: 1px solid #444;
        margin-top: 10px;
        padding-top: 15px;
        font-weight: bold;
    }
    QGroupBox:title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 0 5px;
    }
    QPushButton {
        background-color: #2d2d2d;
        border: 1px solid #555;
        padding: 5px;
        border-radius: 4px;
    }
    QPushButton:hover {
        background-color: #3a3a3a;
    }
    QPushButton:pressed {
        background-color: #444;
    }
    QLineEdit, QTextEdit, QSpinBox, QDoubleSpinBox, QComboBox {
        background-color: #2a2a2a;
        border: 1px solid #555;
        padding: 3px;
        border-radius: 3px;
        color: #e0e0e0;
    }
    QComboBox QAbstractItemView {
        background-color: #2a2a2a;
        selection-background-color: #444;
        color: #e0e0e0;
    }
    QCheckBox {
        spacing: 6px;
    }
    QCheckBox::indicator {
        width: 14px;
        height: 14px;
    }
    QCheckBox::indicator:unchecked {
        border: 1px solid #777;
        background-color: #2a2a2a;
    }
    QCheckBox::indicator:checked {
        border: 1px solid #aaa;
        background-color: #007acc;
    }
    QTabWidget::pane {
        border: 1px solid #444;
        background-color: #1e1e1e;
    }
    QTabBar::tab {
        background-color: #2a2a2a;
        padding: 6px;
        border: 1px solid #444;
        border-bottom: none;
    }
    QTabBar::tab:selected {
        background-color: #3a3a3a;
    }
"""


# ============================================================
#   SERIAL READER THREAD
# ============================================================

class SerialReader(QtCore.QThread):
    data_received = QtCore.pyqtSignal(object)   # (timestamp_or_None, [values])

    def __init__(self, port, baudrate, mode):
        super().__init__()
        self.port = port
        self.baudrate = baudrate
        self.mode = mode  # "binary_ts", "binary_no_ts", "ascii"
        self.running = True

    def run(self):
        try:
            ser = serial.Serial(self.port, self.baudrate, timeout=1)
        except Exception as e:
            print("Serial open error:", e)
            return

        if self.mode == "ascii":
            self.run_ascii(ser)
        else:
            self.run_binary_framed(ser)

    # ---------------- ASCII MODE ----------------
    def run_ascii(self, ser):
        num_channels = 0
        while self.running:
            try:
                line = ser.readline()
            except Exception:
                break

            if not line:
                continue

            try:
                text = line.decode(errors="ignore").strip()
            except Exception:
                continue

            if not text:
                continue

            parts = text.split(",")
            try:
                values = [float(p.strip()) for p in parts if p.strip() != ""]
            except ValueError:
                continue

            if not values:
                continue

            if num_channels == 0:
                num_channels = min(len(values), MAX_CHANNELS)

            self.data_received.emit((None, values[:num_channels]))

        ser.close()

    # ---------------- BINARY FRAMED MODE ----------------
    def run_binary_framed(self, ser):
        while self.running:
            b = ser.read(1)
            if not b or b[0] != HEADER1:
                continue

            b = ser.read(1)
            if not b or b[0] != HEADER2:
                continue

            nf_raw = ser.read(1)
            if not nf_raw:
                continue

            num_floats = nf_raw[0]
            if num_floats == 0 or num_floats > MAX_CHANNELS + 1:
                continue

            needed = num_floats * FLOAT_SIZE
            raw = ser.read(needed)
            if len(raw) != needed:
                continue

            try:
                floats = list(struct.unpack(f"{num_floats}f", raw))
            except struct.error:
                continue

            if self.mode == "binary_ts":
                timestamp = floats[0]
                values = floats[1:]
            else:
                timestamp = None
                values = floats

            self.data_received.emit((timestamp, values))

        ser.close()

    def stop(self):
        self.running = False
        self.wait()


# ============================================================
#   MAIN WINDOW
# ============================================================

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Real-Time Serial 10-Channel Float Plotter")
        self.resize(1600, 900)

        self.time_data = []
        self.ch_data = [[] for _ in range(MAX_CHANNELS)]
        self.start_time = None
        self.num_channels = 0
        self.buffer_limit = 2000

        self.recording = False
        self.record_time_data = []
        self.record_ch_data = [[] for _ in range(MAX_CHANNELS)]

        self.sample_index = 0
        self.serial_thread = None

        self.current_theme = DEFAULT_CONFIG["theme"]
        self.current_data_mode = "binary_no_ts"

        self._build_ui()
        self._build_menu()
        self.apply_theme(self.current_theme)

        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.update_plot)

    # --------------------------------------------------------
    #   BUILD USER INTERFACE
    # --------------------------------------------------------
    def _build_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)

        layout = QtWidgets.QVBoxLayout(central)
        self.tabs = QtWidgets.QTabWidget()
        layout.addWidget(self.tabs)

        # =====================================================
        #   PLOTTER TAB
        # =====================================================
        plotter_tab = QtWidgets.QWidget()
        plotter_layout = QtWidgets.QHBoxLayout(plotter_tab)

        # ---------------- LEFT PANEL ----------------
        left = QtWidgets.QVBoxLayout()

        # Start/Stop buttons
        acq_box = QtWidgets.QHBoxLayout()
        self.start_btn = QtWidgets.QPushButton("Start")
        self.stop_btn = QtWidgets.QPushButton("Stop")
        self.stop_btn.setEnabled(False)
        acq_box.addWidget(self.start_btn)
        acq_box.addWidget(self.stop_btn)
        left.addLayout(acq_box)

        # ---------------- SETUP GROUP ----------------
        setup_group = QtWidgets.QGroupBox("Setup")
        setup_layout = QtWidgets.QFormLayout()

        self.port_box = QtWidgets.QComboBox()
        for p in serial.tools.list_ports.comports():
            self.port_box.addItem(p.device)
        setup_layout.addRow("COM Port:", self.port_box)

        self.baud_box = QtWidgets.QComboBox()
        for b in [9600, 19200, 38400, 57600, 115200, 230400]:
            self.baud_box.addItem(str(b))
        setup_layout.addRow("Baudrate:", self.baud_box)

        self.data_type_box = QtWidgets.QComboBox()
        self.data_type_box.addItem("Binary (serial.write) with timestamp")
        self.data_type_box.addItem("Binary (serial.write) no timestamp")
        self.data_type_box.addItem("ASCII (serial.print)")
        setup_layout.addRow("Incoming Data Type:", self.data_type_box)

        setup_group.setLayout(setup_layout)
        left.addWidget(setup_group)

        # ---------------- PLOTTING GROUP ----------------
        plot_group = QtWidgets.QGroupBox("Plotting")
        plot_layout = QtWidgets.QFormLayout()

        self.window_box = QtWidgets.QSpinBox()
        self.window_box.setRange(1, 120)
        self.window_box.setValue(10)
        plot_layout.addRow("Window (sec):", self.window_box)

        self.channel_checkboxes = []
        self.channel_names = []
        self.channel_scales = []
        self.channel_offsets = []
        self.channel_color_labels = []

        self.colors = [
            "yellow", "cyan", "magenta", "green", "red",
            "blue", "orange", "purple", "lime", "pink"
        ]

        def make_color_label(color):
            lbl = QtWidgets.QLabel()
            lbl.setFixedSize(16, 16)
            lbl.setStyleSheet(f"background-color: {color}; border: 1px solid black;")
            return lbl

        for i in range(MAX_CHANNELS):
            row = QtWidgets.QHBoxLayout()

            color_label = make_color_label(self.colors[i])
            self.channel_color_labels.append(color_label)
            row.addWidget(color_label)

            cb = QtWidgets.QCheckBox(f"CH{i+1}")
            cb.setChecked(True)
            row.addWidget(cb)

            name_edit = QtWidgets.QLineEdit(f"CH{i+1}")
            name_edit.setFixedWidth(100)
            row.addWidget(name_edit)

            scale = QtWidgets.QDoubleSpinBox()
            scale.setRange(-1e6, 1e6)
            scale.setValue(1.0)
            scale.setDecimals(4)
            row.addWidget(scale)

            offset = QtWidgets.QDoubleSpinBox()
            offset.setRange(-1e6, 1e6)
            offset.setValue(0.0)
            offset.setDecimals(4)
            row.addWidget(offset)

            self.channel_checkboxes.append(cb)
            self.channel_names.append(name_edit)
            self.channel_scales.append(scale)
            self.channel_offsets.append(offset)

            plot_layout.addRow(row)

        plot_group.setLayout(plot_layout)
        left.addWidget(plot_group)

        # =====================================================
        #   DATA SAVING (Aligned Layout)
        # =====================================================
        save_group = QtWidgets.QGroupBox("Data Saving")
        save_layout = QtWidgets.QFormLayout()

        # -------- Row 1: Buffer Size + Save Live Data --------
        buffer_row = QtWidgets.QHBoxLayout()

        buffer_row.addWidget(QtWidgets.QLabel("Buffer Size:"))

        self.buffer_box = QtWidgets.QSpinBox()
        self.buffer_box.setRange(10, 500000)
        self.buffer_box.setValue(2000)
        buffer_row.addWidget(self.buffer_box)

        buffer_row.addStretch()

        self.save_live_btn = QtWidgets.QPushButton("Save Live Data")
        self.save_live_btn.setEnabled(False)
        buffer_row.addWidget(self.save_live_btn)

        save_layout.addRow(buffer_row)

        # -------- Row 2: Start/Stop Recording + Save Rec Data --------
        rec_row = QtWidgets.QHBoxLayout()

        self.rec_start_btn = QtWidgets.QPushButton("Start Recording")
        rec_row.addWidget(self.rec_start_btn)

        self.rec_stop_btn = QtWidgets.QPushButton("Stop Recording")
        self.rec_stop_btn.setEnabled(False)
        rec_row.addWidget(self.rec_stop_btn)

        rec_row.addStretch()

        self.save_rec_btn = QtWidgets.QPushButton("Save Rec Data")
        self.save_rec_btn.setEnabled(False)
        rec_row.addWidget(self.save_rec_btn)

        save_layout.addRow(rec_row)

        save_group.setLayout(save_layout)
        left.addWidget(save_group)

        left.addStretch()
        plotter_layout.addLayout(left, 3)

        # ---------------- PLOT AREA ----------------
        self.plot_widget = pg.PlotWidget()
        self.plot_widget.showGrid(x=True, y=True)
        self.curves = [self.plot_widget.plot(pen=c) for c in self.colors]
        plotter_layout.addWidget(self.plot_widget, 7)

        self.tabs.addTab(plotter_tab, "Plotter")

        # =====================================================
        #   MONITOR TAB
        # =====================================================
        monitor_tab = QtWidgets.QWidget()
        monitor_layout = QtWidgets.QVBoxLayout(monitor_tab)

        options_row = QtWidgets.QHBoxLayout()

        self.monitor_autoscroll_cb = QtWidgets.QCheckBox("Auto-scroll")
        self.monitor_autoscroll_cb.setChecked(True)
        options_row.addWidget(self.monitor_autoscroll_cb)

        self.monitor_names_cb = QtWidgets.QCheckBox("Show channel names")
        self.monitor_names_cb.setChecked(True)
        options_row.addWidget(self.monitor_names_cb)

        self.monitor_color_cb = QtWidgets.QCheckBox("Color-coded")
        self.monitor_color_cb.setChecked(True)
        options_row.addWidget(self.monitor_color_cb)

        self.monitor_filter_cb = QtWidgets.QCheckBox("Filter channels")
        self.monitor_filter_cb.setChecked(False)
        options_row.addWidget(self.monitor_filter_cb)

        self.monitor_full_packet_cb = QtWidgets.QCheckBox("Show full package")
        self.monitor_full_packet_cb.setChecked(False)
        options_row.addWidget(self.monitor_full_packet_cb)

        self.monitor_clear_btn = QtWidgets.QPushButton("Clear Monitor")
        options_row.addWidget(self.monitor_clear_btn)

        options_row.addStretch()
        monitor_layout.addLayout(options_row)

        self.monitor = QtWidgets.QTextEdit()
        self.monitor.setReadOnly(True)
        monitor_layout.addWidget(self.monitor)

        self.tabs.addTab(monitor_tab, "Monitor")

        # ---------------- CONNECTIONS ----------------
        self.start_btn.clicked.connect(self.start_reading)
        self.stop_btn.clicked.connect(self.stop_reading)

        self.rec_start_btn.clicked.connect(self.start_recording)
        self.rec_stop_btn.clicked.connect(self.stop_recording)

        self.save_live_btn.clicked.connect(self.save_live_data)
        self.save_rec_btn.clicked.connect(self.save_rec_data)

        self.monitor_clear_btn.clicked.connect(self.monitor.clear)

    # --------------------------------------------------------
    #   MENU BAR
    # --------------------------------------------------------
    def _build_menu(self):
        menubar = self.menuBar()

        setup_menu = menubar.addMenu("Setup")
        about_menu = menubar.addMenu("About")

        self.dark_action = QtWidgets.QAction("Dark Mode", self, checkable=True)
        self.normal_action = QtWidgets.QAction("Normal Mode", self, checkable=True)

        theme_group = QtWidgets.QActionGroup(self)
        theme_group.addAction(self.dark_action)
        theme_group.addAction(self.normal_action)
        self.dark_action.setChecked(True)

        self.dark_action.triggered.connect(lambda: self.set_theme("dark"))
        self.normal_action.triggered.connect(lambda: self.set_theme("normal"))

        setup_menu.addAction(self.dark_action)
        setup_menu.addAction(self.normal_action)

        help_action = QtWidgets.QAction("Help", self)
        help_action.triggered.connect(self.show_help)
        about_menu.addAction(help_action)

    # --------------------------------------------------------
    #   THEME + CONFIG
    # --------------------------------------------------------
    def set_theme(self, theme):
        self.current_theme = theme
        self.apply_theme(theme)

    def apply_theme(self, theme):
        app = QtWidgets.QApplication.instance()
        if theme == "dark":
            app.setStyleSheet(DARK_STYLESHEET)
            self.dark_action.setChecked(True)
        else:
            app.setStyleSheet("")
            self.normal_action.setChecked(True)

    def show_help(self):
        text = (
            "Real-Time Serial Plotter\n"
            "Framed binary + ASCII support\n\n"
            "Binary protocol:\n"
            "  0xAA 0x55 <num_floats> <float...>\n\n"
            "Modes:\n"
            "  • Binary with timestamp: first float = time (s)\n"
            "  • Binary no timestamp: all floats are channels\n"
            "  • ASCII: CSV line, time generated on PC\n"
        )
        QtWidgets.QMessageBox.information(self, "Help", text)

    # --------------------------------------------------------
    #   CHANNEL ENABLE/DISABLE VISUALS
    # --------------------------------------------------------
    def set_channel_enabled(self, index, enabled):
        if enabled:
            self.channel_color_labels[index].setStyleSheet(
                f"background-color: {self.colors[index]}; border: 1px solid black;"
            )
            self.channel_checkboxes[index].setStyleSheet("")
            self.channel_names[index].setStyleSheet("")
            self.channel_scales[index].setStyleSheet("")
            self.channel_offsets[index].setStyleSheet("")
        else:
            self.channel_color_labels[index].setStyleSheet(
                "background-color: #555555; border: 1px solid #333333;"
            )
            grey = "color: #666666;"
            self.channel_checkboxes[index].setStyleSheet(grey)
            self.channel_names[index].setStyleSheet(grey)
            self.channel_scales[index].setStyleSheet(grey)
            self.channel_offsets[index].setStyleSheet(grey)

    # --------------------------------------------------------
    #   START / STOP ACQUISITION
    # --------------------------------------------------------
    def start_reading(self):
        port = self.port_box.currentText()
        baud = int(self.baud_box.currentText())
        self.buffer_limit = self.buffer_box.value()

        idx = self.data_type_box.currentIndex()
        if idx == 0:
            self.current_data_mode = "binary_ts"
        elif idx == 1:
            self.current_data_mode = "binary_no_ts"
        else:
            self.current_data_mode = "ascii"

        # Reset buffers
        self.time_data.clear()
        for ch in self.ch_data:
            ch.clear()

        self.start_time = None
        self.num_channels = 0

        self.recording = False
        self.record_time_data.clear()
        self.record_ch_data = [[] for _ in range(MAX_CHANNELS)]
        self.rec_start_btn.setEnabled(True)
        self.rec_stop_btn.setEnabled(False)

        self.save_live_btn.setEnabled(False)
        self.save_rec_btn.setEnabled(False)

        self.sample_index = 0

        for i in range(MAX_CHANNELS):
            self.channel_checkboxes[i].setEnabled(True)
            self.channel_scales[i].setEnabled(True)
            self.channel_offsets[i].setEnabled(True)
            self.channel_names[i].setEnabled(True)
            self.channel_checkboxes[i].setChecked(True)
            self.set_channel_enabled(i, True)

        self.serial_thread = SerialReader(port, baud, self.current_data_mode)
        self.serial_thread.data_received.connect(self.handle_data)
        self.serial_thread.start()

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.timer.start(50)

    def stop_reading(self):
        if self.serial_thread:
            self.serial_thread.stop()
            self.serial_thread = None

        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.timer.stop()

        self.recording = False
        self.rec_start_btn.setEnabled(True)
        self.rec_stop_btn.setEnabled(False)

    # --------------------------------------------------------
    #   RECORDING
    # --------------------------------------------------------
    def start_recording(self):
        self.recording = True
        self.record_time_data.clear()
        self.record_ch_data = [[] for _ in range(MAX_CHANNELS)]
        self.rec_start_btn.setEnabled(False)
        self.rec_stop_btn.setEnabled(True)
        self.save_rec_btn.setEnabled(False)
        self.sample_index = 0

    def stop_recording(self):
        self.recording = False
        self.rec_start_btn.setEnabled(True)
        self.rec_stop_btn.setEnabled(False)
        self.save_rec_btn.setEnabled(len(self.record_time_data) > 0)

    # --------------------------------------------------------
    #   HANDLE INCOMING DATA
    # --------------------------------------------------------
    def handle_data(self, payload):
        timestamp, values = payload

        # Auto-detect channels
        if self.num_channels == 0:
            self.num_channels = min(len(values), MAX_CHANNELS)
            for i in range(MAX_CHANNELS):
                active = i < self.num_channels
                self.channel_checkboxes[i].setEnabled(active)
                self.channel_scales[i].setEnabled(active)
                self.channel_offsets[i].setEnabled(active)
                self.channel_names[i].setEnabled(active)
                self.channel_checkboxes[i].setChecked(active)
                self.set_channel_enabled(i, active)

        # Determine time
        if timestamp is not None and self.current_data_mode == "binary_ts":
            t = float(timestamp)
        else:
            if self.start_time is None:
                self.start_time = time.perf_counter()
            t = time.perf_counter() - self.start_time

        self.sample_index += 1

        # Store data
        self.time_data.append(t)
        for i in range(self.num_channels):
            self.ch_data[i].append(values[i])

        # Trim buffer
        if len(self.time_data) > self.buffer_limit:
            self.time_data = self.time_data[-self.buffer_limit:]
            for i in range(self.num_channels):
                self.ch_data[i] = self.ch_data[i][-self.buffer_limit:]

        # Recording
        if self.recording:
            self.record_time_data.append(t)
            for i in range(self.num_channels):
                self.record_ch_data[i].append(values[i])

        # Enable save buttons
        self.save_live_btn.setEnabled(len(self.time_data) > 0)
        self.save_rec_btn.setEnabled(len(self.record_time_data) > 0)

        # =====================================================
        #   MONITOR OUTPUT
        # =====================================================

        # FULL PACKAGE MODE
        if self.monitor_full_packet_cb.isChecked():
            header_text = "AA 55"

            nf = len(values) + (
                1 if timestamp is not None and self.current_data_mode == "binary_ts" else 0
            )

            raw_floats = []
            if timestamp is not None and self.current_data_mode == "binary_ts":
                raw_floats.append(timestamp)
            raw_floats.extend(values)

            float_text = " | ".join(f"{v:.6f}" for v in raw_floats)

            line = f"{header_text} | nf={nf} | {float_text}"
            self.monitor.append(line)

            if self.monitor_autoscroll_cb.isChecked():
                self.monitor.moveCursor(QTextCursor.End)
            return

        # NORMAL MODE — timestamp on left if available
        if timestamp is not None and self.current_data_mode == "binary_ts":
            t_display = f"{timestamp:.6f}"
        else:
            t_display = f"{t:.6f}"

        # Determine channels to show
        if self.monitor_filter_cb.isChecked():
            indices = [i for i in range(self.num_channels) if self.channel_checkboxes[i].isChecked()]
        else:
            indices = list(range(self.num_channels))

        parts = []
        for i in indices:
            val = values[i]
            name = self.channel_names[i].text()

            if self.monitor_names_cb.isChecked():
                text = f"{name}: {val:.6f}"
            else:
                text = f"{val:.6f}"

            if self.monitor_color_cb.isChecked():
                color = self.colors[i]
                text = f'<span style="color:{color}">{text}</span>'
            else:
                if self.current_theme == "dark":
                    text = f'<span style="color:white">{text}</span>'
                else:
                    text = f'<span style="color:black">{text}</span>'

            parts.append(text)

        line = f"{t_display} | " + ", ".join(parts)
        self.monitor.append(line)

        if self.monitor_autoscroll_cb.isChecked():
            self.monitor.moveCursor(QTextCursor.End)

    # --------------------------------------------------------
    #   UPDATE PLOT
    # --------------------------------------------------------
    def update_plot(self):
        if not self.time_data:
            return

        window_sec = self.window_box.value()
        latest = self.time_data[-1]
        cutoff = latest - window_sec

        # Trim old samples
        while self.time_data and self.time_data[0] < cutoff:
            self.time_data.pop(0)
            for i in range(self.num_channels):
                if self.ch_data[i]:
                    self.ch_data[i].pop(0)

        # Update curves
        for i in range(self.num_channels):
            if not self.ch_data[i]:
                self.curves[i].clear()
                continue

            if self.channel_checkboxes[i].isChecked():
                scale = self.channel_scales[i].value()
                offset = self.channel_offsets[i].value()
                y = [(v * scale) + offset for v in self.ch_data[i]]
                self.curves[i].setData(self.time_data, y)
            else:
                self.curves[i].clear()

        # Clear unused curves
        for i in range(self.num_channels, MAX_CHANNELS):
            self.curves[i].clear()

    # --------------------------------------------------------
    #   SAVE LIVE DATA
    # --------------------------------------------------------
    def save_live_data(self):
        if len(self.time_data) == 0:
            QtWidgets.QMessageBox.warning(self, "No Data", "No live data available.")
            return

        filename, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Live Data", "", "CSV Files (*.csv)"
        )
        if not filename:
            return

        headers = ["Time"] + [self.channel_names[i].text() for i in range(self.num_channels)]

        with open(filename, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(headers)

            for row_idx in range(len(self.time_data)):
                row = [self.time_data[row_idx]]
                for ch in range(self.num_channels):
                    row.append(self.ch_data[ch][row_idx])
                writer.writerow(row)

        QtWidgets.QMessageBox.information(self, "Saved", "Live data saved successfully.")

    # --------------------------------------------------------
    #   SAVE RECORDED DATA
    # --------------------------------------------------------
    def save_rec_data(self):
        if len(self.record_time_data) == 0:
            QtWidgets.QMessageBox.warning(self, "No Data", "No recorded data available.")
            return

        filename, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Recorded Data", "", "CSV Files (*.csv)"
        )
        if not filename:
            return

        headers = ["Time"] + [self.channel_names[i].text() for i in range(self.num_channels)]

        with open(filename, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(headers)

            for row_idx in range(len(self.record_time_data)):
                row = [self.record_time_data[row_idx]]
                for ch in range(self.num_channels):
                    row.append(self.record_ch_data[ch][row_idx])
                writer.writerow(row)

        QtWidgets.QMessageBox.information(self, "Saved", "Recorded data saved successfully.")

    # --------------------------------------------------------
    #   CLOSE EVENT
    # --------------------------------------------------------
    def closeEvent(self, event):
        if self.serial_thread:
            self.serial_thread.stop()
        event.accept()


# ============================================================
#   MAIN ENTRY POINT
# ============================================================
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
