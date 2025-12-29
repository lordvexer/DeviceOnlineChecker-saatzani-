import sys
import socket
import sqlite3
import pandas as pd
import subprocess
import platform
import time
import threading
import webbrowser
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QTableWidget, QTableWidgetItem,
    QLabel, QSpinBox, QMessageBox, QProgressBar,
    QHBoxLayout, QDialog, QDateTimeEdit, QHeaderView,
    QAbstractItemView, QLineEdit
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QColor

DB_NAME = "devices.db"

# --- Modern Neon Styling ---
MODERN_STYLE = """
    QWidget { background-color: #0F111A; color: #E0E0E0; font-family: 'Segoe UI'; font-size: 13px; }
    QTableWidget { 
        background-color: #161925; 
        alternate-background-color: #1C2030; 
        border: 1px solid #2D3245; 
        gridline-color: transparent; 
        outline: none;
    }
    QTableWidget::item { padding: 8px; }
    QTableWidget::item:selected { 
        color: #00F0FF; 
        background-color: #1C2030; 
        border-bottom: 2px solid #00F0FF; 
    }
    QPushButton { 
        background-color: #21263D; 
        border: 1px solid #30364D; 
        border-radius: 4px; 
        padding: 8px 15px; 
        min-width: 90px;
        font-weight: bold;
    }
    QPushButton:hover { background-color: #0078D4; border-color: #00F0FF; }
    QLineEdit, QSpinBox, QDateTimeEdit { 
        background-color: #1C2030; 
        border: 1px solid #2D3245; 
        padding: 6px; 
        border-radius: 4px; 
    }
    QHeaderView::section { 
        background-color: #0F111A; 
        padding: 10px; 
        border: none; 
        font-weight: bold; 
        color: #8B949E;
        text-transform: uppercase;
    }
    QProgressBar {
        border: 1px solid #2D3245;
        background-color: #161925;
        height: 8px;
        text-align: center;
        border-radius: 4px;
    }
    QProgressBar::chunk { background-color: #00F0FF; border-radius: 4px; }
"""

# =========================
# DIALOGS
# =========================
class AddDeviceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add New Device")
        self.setFixedWidth(350)
        self.setStyleSheet(MODERN_STYLE)
        layout = QVBoxLayout(self)
        self.name_in = QLineEdit(); self.name_in.setPlaceholderText("Device Name (e.g. Core Switch)")
        self.ip_in = QLineEdit(); self.ip_in.setPlaceholderText("IP Address (e.g. 192.168.1.1)")
        self.port_in = QSpinBox(); self.port_in.setRange(1, 65535); self.port_in.setValue(80)
        btn = QPushButton("Save Device"); btn.clicked.connect(self.accept)
        layout.addWidget(QLabel("Device Name:")); layout.addWidget(self.name_in)
        layout.addWidget(QLabel("IP Address:")); layout.addWidget(self.ip_in)
        layout.addWidget(QLabel("Service Port:")); layout.addWidget(self.port_in)
        layout.addSpacing(10); layout.addWidget(btn)

class LogWindow(QDialog):
    def __init__(self, ip, port, name):
        super().__init__()
        self.ip, self.port = ip, port
        self.setWindowTitle(f"History: {name} ({ip}:{port})")
        self.resize(850, 550); self.setStyleSheet(MODERN_STYLE)
        layout = QVBoxLayout(self)
        filter_box = QHBoxLayout()
        self.from_dt = QDateTimeEdit(datetime.now().replace(hour=0, minute=0, second=0))
        self.to_dt = QDateTimeEdit(datetime.now())
        self.from_dt.setCalendarPopup(True); self.to_dt.setCalendarPopup(True)
        btn_f = QPushButton("Filter"); btn_f.clicked.connect(self.load_logs)
        btn_e = QPushButton("Export Excel"); btn_e.clicked.connect(self.export_logs)
        filter_box.addWidget(QLabel("From:")); filter_box.addWidget(self.from_dt)
        filter_box.addWidget(QLabel("To:")); filter_box.addWidget(self.to_dt)
        filter_box.addWidget(btn_f); filter_box.addWidget(btn_e)
        layout.addLayout(filter_box)
        self.log_table = QTableWidget(0, 4)
        self.log_table.setHorizontalHeaderLabels(["Timestamp", "Ping", "Port Status", "Health"])
        self.log_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.log_table.setAlternatingRowColors(True); layout.addWidget(self.log_table)
        self.load_logs()

    def load_logs(self):
        f = self.from_dt.dateTime().toString("yyyy-MM-dd HH:mm:ss")
        t = self.to_dt.dateTime().toString("yyyy-MM-dd HH:mm:ss")
        conn = sqlite3.connect(DB_NAME)
        rows = conn.execute("SELECT timestamp, ping, port_status, overall FROM device_logs WHERE ip=? AND port=? AND timestamp BETWEEN ? AND ? ORDER BY timestamp DESC", (self.ip, self.port, f, t)).fetchall()
        conn.close()
        self.log_table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                txt = "ONLINE" if c==3 and val==1 else "OFFLINE" if c==3 else "SUCCESS" if c==1 and val==1 else "FAILED" if c==1 else "OPEN" if c==2 and val==1 else "CLOSED" if c==2 else str(val)
                item = QTableWidgetItem(txt); item.setTextAlignment(Qt.AlignCenter)
                if c == 3: item.setForeground(QColor("#00F0FF") if val == 1 else QColor("#FF4560"))
                self.log_table.setItem(r, c, item)

    def export_logs(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export", "", "Excel (*.xlsx)")
        if path:
            data = [[self.log_table.item(r, c).text() for c in range(4)] for r in range(self.log_table.rowCount())]
            pd.DataFrame(data, columns=["Time", "Ping", "Port", "Status"]).to_excel(path, index=False)

# =========================
# MAIN WINDOW
# =========================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.devices = []; self.init_db(); self.setup_ui(); self.load_from_db()
        self.worker = SerialWorker(self.devices, self.get_interval, self.get_ping_count)
        self.worker.checking_now.connect(self.mark_row_checking)
        self.worker.result_ready.connect(self.update_row)
        self.worker.tick.connect(self.update_progress); self.worker.start()

    def init_db(self):
        conn = sqlite3.connect(DB_NAME)
        conn.execute("CREATE TABLE IF NOT EXISTS devices (name TEXT, ip TEXT, port INTEGER)")
        conn.execute("CREATE TABLE IF NOT EXISTS device_logs (ip TEXT, port INTEGER, ping INTEGER, port_status INTEGER, overall INTEGER, timestamp TEXT)")
        conn.commit(); conn.close()

    def setup_ui(self):
        self.setWindowTitle("Device Network Manitor By MHZ"); self.resize(1200, 750); self.setStyleSheet(MODERN_STYLE)
        main_layout = QVBoxLayout(self); main_layout.setContentsMargins(15, 15, 15, 15)
        tools = QHBoxLayout()
        self.interval_spin = QSpinBox(); self.interval_spin.setRange(1, 3600); self.interval_spin.setValue(10)
        self.ping_spin = QSpinBox(); self.ping_spin.setRange(1, 10); self.ping_spin.setValue(1)
        btn_add = QPushButton("+ Add Manual"); btn_add.clicked.connect(self.add_manual)
        btn_excel = QPushButton("Import Excel"); btn_excel.clicked.connect(self.import_excel)
        btn_del = QPushButton("Delete Selected"); btn_del.clicked.connect(self.delete_selected)
        tools.addWidget(QLabel("Interval (s):")); tools.addWidget(self.interval_spin)
        tools.addWidget(QLabel("Pings:")); tools.addWidget(self.ping_spin)
        tools.addStretch(); tools.addWidget(btn_add); tools.addWidget(btn_excel); tools.addWidget(btn_del)
        main_layout.addLayout(tools)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["DEVICE NAME", "IP ADDRESS", "PORT", "PING", "SERVICE", "HEALTH STATUS"])
        self.table.setAlternatingRowColors(True); self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers); self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.cellDoubleClicked.connect(self.open_device_logs); self.table.viewport().installEventFilter(self)
        main_layout.addWidget(self.table)

        footer = QHBoxLayout(); self.progress = QProgressBar(); self.status_lbl = QLabel("Ready")
        footer.addWidget(self.status_lbl); footer.addWidget(self.progress); main_layout.addLayout(footer)

    def eventFilter(self, source, event):
        if event.type() == event.MouseButtonRelease and event.button() == Qt.MidButton:
            item = self.table.itemAt(event.pos())
            if item: webbrowser.open(f"http://{self.table.item(item.row(), 1).text()}")
        return super().eventFilter(source, event)

    def open_device_logs(self, row, col):
        name, ip, port = self.table.item(row, 0).text(), self.table.item(row, 1).text(), int(self.table.item(row, 2).text())
        LogWindow(ip, port, name).exec_()

    def load_from_db(self):
        conn = sqlite3.connect(DB_NAME); rows = conn.execute("SELECT name, ip, port FROM devices").fetchall(); conn.close()
        self.devices[:] = [{"name": r[0], "ip": r[1], "port": r[2]} for r in rows]; self.refresh_table_ui()

    def refresh_table_ui(self):
        self.table.setRowCount(len(self.devices))
        for r, d in enumerate(self.devices):
            it_name = QTableWidgetItem(d['name']); it_name.setTextAlignment(Qt.AlignCenter); self.table.setItem(r, 0, it_name)
            self.table.setItem(r, 1, QTableWidgetItem(d['ip'])) # IP Left aligned
            for c in range(2, 6):
                it = QTableWidgetItem(str(d['port']) if c==2 else "-"); it.setTextAlignment(Qt.AlignCenter); self.table.setItem(r, c, it)

    def mark_row_checking(self, ip, port):
        for r in range(self.table.rowCount()):
            if self.table.item(r, 1).text() == ip and self.table.item(r, 2).text() == str(port):
                it = QTableWidgetItem("Checking..."); it.setTextAlignment(Qt.AlignCenter); it.setForeground(QColor("#FFA500"))
                self.table.setItem(r, 5, it); break

    def update_row(self, res):
        for r in range(self.table.rowCount()):
            if self.table.item(r, 1).text() == res['ip'] and self.table.item(r, 2).text() == str(res['port']):
                data = {3: "SUCCESS" if res['ping'] else "FAILED", 4: "OPEN" if res['port_ok'] else "CLOSED", 5: "ONLINE" if res['overall'] else "OFFLINE"}
                for c, txt in data.items():
                    it = QTableWidgetItem(txt); it.setTextAlignment(Qt.AlignCenter)
                    if c == 5: it.setForeground(QColor("#00F0FF") if res['overall'] else QColor("#FF4560"))
                    self.table.setItem(r, c, it)

    def add_manual(self):
        dlg = AddDeviceDialog(self)
        if dlg.exec_():
            n, i, p = dlg.name_in.text().strip(), dlg.ip_in.text().strip(), dlg.port_in.value()
            if n and i:
                conn = sqlite3.connect(DB_NAME); conn.execute("INSERT INTO devices VALUES (?,?,?)", (n, i, p)); conn.commit(); conn.close(); self.load_from_db()

    def import_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import", "", "Excel (*.xlsx)")
        if path:
            df = pd.read_excel(path)
            conn = sqlite3.connect(DB_NAME)
            for _, r in df.iterrows(): conn.execute("INSERT INTO devices VALUES (?,?,?)", (str(r['name']), str(r['ip']), int(r['port'])))
            conn.commit(); conn.close(); self.load_from_db()

    def delete_selected(self):
        rows = sorted(list(set(i.row() for i in self.table.selectedIndexes())), reverse=True)
        if rows and QMessageBox.question(self, "Confirm", "Delete?", QMessageBox.Yes|QMessageBox.No) == QMessageBox.Yes:
            conn = sqlite3.connect(DB_NAME)
            for r in rows: conn.execute("DELETE FROM devices WHERE ip=? AND port=?", (self.table.item(r, 1).text(), int(self.table.item(r, 2).text())))
            conn.commit(); conn.close(); self.load_from_db()

    def update_progress(self, rem, total): self.progress.setValue(int((rem/total)*100)); self.status_lbl.setText(f"Scanning in {rem}s")
    def get_interval(self): return self.interval_spin.value()
    def get_ping_count(self): return self.ping_spin.value()

class SerialWorker(QThread):
    result_ready = pyqtSignal(dict); tick = pyqtSignal(int, int); checking_now = pyqtSignal(str, int)
    def __init__(self, dev, i_fn, p_fn): super().__init__(); self.devices = dev; self.i_fn = i_fn; self.p_fn = p_fn; self.running = True
    def run(self):
        while self.running:
            for d in list(self.devices):
                if not self.running: break
                ip, port = d['ip'], d['port']; self.checking_now.emit(ip, port)
                p_ok = subprocess.run(["ping", "-n" if platform.system().lower()=="windows" else "-c", str(self.p_fn()), ip], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode == 0
                s_ok = False
                try:
                    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s: s.settimeout(2); s_ok = (s.connect_ex((ip, port)) == 0)
                except: pass
                ov = 1 if (p_ok and s_ok) else 0
                conn = sqlite3.connect(DB_NAME); conn.execute("INSERT INTO device_logs VALUES (?,?,?,?,?,?)", (ip, port, int(p_ok), int(s_ok), ov, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))); conn.commit(); conn.close()
                self.result_ready.emit({"ip":ip, "port":port, "ping":p_ok, "port_ok":s_ok, "overall":ov})
            w = self.i_fn()
            for i in range(w, 0, -1):
                if not self.running: break
                self.tick.emit(i, w); self.msleep(1000)

if __name__ == "__main__":
    app = QApplication(sys.argv); w = MainWindow(); w.show(); sys.exit(app.exec_())
