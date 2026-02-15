import sys
import socket
import sqlite3
import pandas as pd
import subprocess
import platform
import webbrowser
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QTableWidget, QTableWidgetItem,
    QLabel, QSpinBox, QMessageBox, QProgressBar,
    QHBoxLayout, QDialog, QDateTimeEdit, QHeaderView,
    QAbstractItemView, QLineEdit, QMenu, QFormLayout,
    QTableView, QCheckBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QTimer
from PyQt5.QtGui import QColor

# تلاش برای ایمپورت pyodbc برای اس‌کیو‌ال سرور
try:
    import pyodbc
    HAS_ODBC = True
except ImportError:
    HAS_ODBC = False

DB_NAME = "devices.db"

# --- استایل نهایی و حرفه‌ای ---
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
        self.name_in = QLineEdit(); self.name_in.setPlaceholderText("Device Name")
        self.ip_in = QLineEdit(); self.ip_in.setPlaceholderText("IP Address")
        self.port_in = QSpinBox(); self.port_in.setRange(1, 65535); self.port_in.setValue(80)
        btn = QPushButton("Save Device"); btn.clicked.connect(self.accept)
        layout.addWidget(QLabel("Device Label / Name:")); layout.addWidget(self.name_in)
        layout.addWidget(QLabel("IP Address:")); layout.addWidget(self.ip_in)
        layout.addWidget(QLabel("Monitoring Port:")); layout.addWidget(self.port_in)
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
        rows = conn.execute(
            "SELECT timestamp, ping, port_status, overall "
            "FROM device_logs WHERE ip=? AND port=? AND timestamp BETWEEN ? AND ? "
            "ORDER BY timestamp DESC",
            (self.ip, self.port, f, t)
        ).fetchall()
        conn.close()
        self.log_table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                txt = (
                    "ONLINE" if c == 3 and val == 1 else
                    "OFFLINE" if c == 3 else
                    "SUCCESS" if c == 1 and val == 1 else
                    "FAILED" if c == 1 else
                    "OPEN" if c == 2 and val == 1 else
                    "CLOSED" if c == 2 else str(val)
                )
                item = QTableWidgetItem(txt); item.setTextAlignment(Qt.AlignCenter)
                if c == 3:
                    item.setForeground(QColor("#00F0FF") if val == 1 else QColor("#FF4560"))
                self.log_table.setItem(r, c, item)

    def export_logs(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export", "", "Excel (*.xlsx)")
        if path:
            data = [[self.log_table.item(r, c).text() for c in range(4)]
                    for r in range(self.log_table.rowCount())]
            pd.DataFrame(data, columns=["Time", "Ping", "Port", "Status"]).to_excel(path, index=False)


class SyncProfileDialog(QDialog):
    """
    مدیریت لیست سرورهای SQL / دیتابیس / کوئری.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("SQL Sync Profiles")
        self.resize(800, 400)
        self.setStyleSheet(MODERN_STYLE)

        self.conn = sqlite3.connect(DB_NAME)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS sync_profiles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                server TEXT,
                database TEXT,
                username TEXT,
                password TEXT,
                query TEXT,
                active INTEGER
            )
        """)
        self.conn.commit()

        layout = QVBoxLayout(self)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels([
            "TITLE", "SERVER", "DATABASE", "USERNAME", "QUERY (name, ip, port)", "ACTIVE", "ID"
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.setColumnHidden(6, True)  # ID
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        layout.addWidget(self.table)

        # Sample hint
        hint = QLabel(
            "Sample Query:\n"
            "SELECT Dev_Place AS name, Dev_IP AS ip, Dev_Port AS port\n"
            "FROM BioAtlasORG.dbo.TB_Device"
        )
        hint.setStyleSheet("color:#8B949E; font-size:11px;")
        layout.addWidget(hint)

        btns = QHBoxLayout()
        btn_add = QPushButton("Add Profile")
        btn_edit = QPushButton("Edit Selected")
        btn_del = QPushButton("Delete Selected")
        btn_close = QPushButton("Close")

        btn_add.clicked.connect(self.add_profile)
        btn_edit.clicked.connect(self.edit_profile)
        btn_del.clicked.connect(self.delete_profile)
        btn_close.clicked.connect(self.accept)

        btns.addWidget(btn_add)
        btns.addWidget(btn_edit)
        btns.addWidget(btn_del)
        btns.addStretch()
        btns.addWidget(btn_close)

        layout.addLayout(btns)

        self.load_profiles()

    def load_profiles(self):
        self.table.setRowCount(0)
        rows = self.conn.execute(
            "SELECT id, title, server, database, username, password, query, active "
            "FROM sync_profiles"
        ).fetchall()
        for row in rows:
            r = self.table.rowCount()
            self.table.insertRow(r)
            id_, title, server, db, user, pwd, query, active = row
            self.table.setItem(r, 0, QTableWidgetItem(title or ""))
            self.table.setItem(r, 1, QTableWidgetItem(server or ""))
            self.table.setItem(r, 2, QTableWidgetItem(db or ""))
            self.table.setItem(r, 3, QTableWidgetItem(user or ""))
            self.table.setItem(r, 4, QTableWidgetItem(query or ""))

            chk_item = QTableWidgetItem("YES" if active else "NO")
            chk_item.setTextAlignment(Qt.AlignCenter)
            chk_item.setForeground(QColor("#00F0FF") if active else QColor("#FF4560"))
            self.table.setItem(r, 5, chk_item)

            id_item = QTableWidgetItem(str(id_))
            self.table.setItem(r, 6, id_item)

    def add_profile(self):
        dlg = EditProfileForm(self)
        if dlg.exec_():
            data = dlg.get_data()
            self.conn.execute(
                "INSERT INTO sync_profiles (title, server, database, username, password, query, active) "
                "VALUES (?,?,?,?,?,?,?)",
                (data['title'], data['server'], data['database'], data['username'],
                 data['password'], data['query'], 1 if data['active'] else 0)
            )
            self.conn.commit()
            self.load_profiles()

    def edit_profile(self):
        row = self.table.currentRow()
        if row < 0:
            return
        id_ = int(self.table.item(row, 6).text())
        cur = self.conn.execute(
            "SELECT title, server, database, username, password, query, active "
            "FROM sync_profiles WHERE id=?",
            (id_,)
        )
        res = cur.fetchone()
        if not res:
            return
        title, server, db, user, pwd, query, active = res
        dlg = EditProfileForm(self, {
            "title": title or "",
            "server": server or "",
            "database": db or "",
            "username": user or "",
            "password": pwd or "",
            "query": query or "",
            "active": bool(active)
        })
        if dlg.exec_():
            data = dlg.get_data()
            self.conn.execute(
                "UPDATE sync_profiles SET title=?, server=?, database=?, username=?, password=?, query=?, active=? "
                "WHERE id=?",
                (data['title'], data['server'], data['database'], data['username'],
                 data['password'], data['query'], 1 if data['active'] else 0, id_)
            )
            self.conn.commit()
            self.load_profiles()

    def delete_profile(self):
        row = self.table.currentRow()
        if row < 0:
            return
        id_ = int(self.table.item(row, 6).text())
        if QMessageBox.question(
            self, "Confirm", "Delete this profile?", QMessageBox.Yes | QMessageBox.No
        ) == QMessageBox.Yes:
            self.conn.execute("DELETE FROM sync_profiles WHERE id=?", (id_,))
            self.conn.commit()
            self.load_profiles()

    def get_active_profiles(self):
        rows = self.conn.execute(
            "SELECT title, server, database, username, password, query "
            "FROM sync_profiles WHERE active=1"
        ).fetchall()
        profiles = []
        for title, server, db, user, pwd, query in rows:
            profiles.append({
                "title": title,
                "server": server,
                "database": db,
                "username": user,
                "password": pwd,
                "query": query
            })
        return profiles

    def closeEvent(self, event):
        self.conn.close()
        super().closeEvent(event)


class EditProfileForm(QDialog):
    def __init__(self, parent=None, data=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Sync Profile" if data else "Add Sync Profile")
        self.setFixedWidth(500)
        self.setStyleSheet(MODERN_STYLE)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.title_in = QLineEdit()
        self.server_in = QLineEdit()
        self.db_in = QLineEdit()
        self.user_in = QLineEdit()
        self.pass_in = QLineEdit(); self.pass_in.setEchoMode(QLineEdit.Password)
        self.query_in = QLineEdit()
        self.active_chk = QCheckBox("Active for Auto-Sync")

        self.title_in.setPlaceholderText("Sample: Main BioAtlas Devices")
        self.server_in.setPlaceholderText("e.g. 192.168.1.10 or SQLSERVER01")
        self.db_in.setPlaceholderText("e.g. BioAtlasORG")
        self.user_in.setPlaceholderText("e.g. sa")
        self.query_in.setPlaceholderText(
            "SELECT Dev_Place AS name, Dev_IP AS ip, Dev_Port AS port "
            "FROM BioAtlasORG.dbo.TB_Device"
        )

        if data:
            self.title_in.setText(data.get("title", ""))
            self.server_in.setText(data.get("server", ""))
            self.db_in.setText(data.get("database", ""))
            self.user_in.setText(data.get("username", ""))
            self.pass_in.setText(data.get("password", ""))
            self.query_in.setText(data.get("query", ""))
            self.active_chk.setChecked(data.get("active", True))
        else:
            self.active_chk.setChecked(True)

        form.addRow("Title:", self.title_in)
        form.addRow("Server:", self.server_in)
        form.addRow("Database:", self.db_in)
        form.addRow("Username:", self.user_in)
        form.addRow("Password:", self.pass_in)
        form.addRow("Query:", self.query_in)
        form.addRow("", self.active_chk)

        layout.addLayout(form)

        btns = QHBoxLayout()
        btn_ok = QPushButton("Save")
        btn_cancel = QPushButton("Cancel")
        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        btns.addStretch()
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)
        layout.addLayout(btns)

    def get_data(self):
        return {
            "title": self.title_in.text().strip(),
            "server": self.server_in.text().strip(),
            "database": self.db_in.text().strip(),
            "username": self.user_in.text().strip(),
            "password": self.pass_in.text().strip(),
            "query": self.query_in.text().strip(),
            "active": self.active_chk.isChecked()
        }


# =========================
# WORKER THREAD
# =========================
class SerialWorker(QThread):
    result_ready = pyqtSignal(dict); tick = pyqtSignal(int, int); checking_now = pyqtSignal(str, int)
    
    def __init__(self, dev, i_fn, p_fn): 
        super().__init__()
        self.devices = dev
        self.i_fn = i_fn
        self.p_fn = p_fn
        self.running = True

    def run(self):
        startupinfo = None
        creationflags = 0
        if platform.system().lower() == "windows":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0
            creationflags = 0x08000000  # CREATE_NO_WINDOW
            
        while self.running:
            for d in list(self.devices):
                if not self.running:
                    break
                ip, port = d['ip'], d['port']
                self.checking_now.emit(ip, port)
                
                try:
                    p_res = subprocess.run(
                        ["ping", "-n" if platform.system().lower()=="windows" else "-c",
                         str(self.p_fn()), ip],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        stdin=subprocess.DEVNULL,
                        startupinfo=startupinfo,
                        creationflags=creationflags
                    )
                    p_ok = (p_res.returncode == 0)
                except Exception:
                    p_ok = False
                
                s_ok = False
                try:
                    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                        s.settimeout(2)
                        if s.connect_ex((ip, port)) == 0:
                            s_ok = True
                except Exception:
                    s_ok = False
                
                ov = 1 if (p_ok and s_ok) else 0
                conn = sqlite3.connect(DB_NAME)
                conn.execute(
                    "INSERT INTO device_logs VALUES (?,?,?,?,?,?)",
                    (ip, port, int(p_ok), int(s_ok), ov,
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                )
                conn.commit()
                conn.close()
                self.result_ready.emit({
                    "ip": ip, "port": port,
                    "ping": p_ok, "port_ok": s_ok, "overall": ov
                })
            
            w = self.i_fn()
            for i in range(w, 0, -1):
                if not self.running:
                    break
                self.tick.emit(i, w)
                self.msleep(1000)


# =========================
# MAIN WINDOW
# =========================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.devices = []
        self.init_db()
        self.setup_ui()
        self.load_from_db()

        # تایمر برای سینک خودکار از SQL Server
        self.sql_timer = QTimer(self)
        self.sql_timer.setInterval(60_000)  # هر ۶۰ ثانیه
        self.sql_timer.timeout.connect(self.auto_sync_sql)
        # اگر نخواستی از ابتدا فعال باشد، می‌توانی بعد از تعریف پروفایل‌ها start کنی
        self.sql_timer.start()

        self.worker = SerialWorker(self.devices, self.get_interval, self.get_ping_count)
        self.worker.checking_now.connect(self.mark_row_checking)
        self.worker.result_ready.connect(self.update_row)
        self.worker.tick.connect(self.update_progress)
        self.worker.start()

    def init_db(self):
        conn = sqlite3.connect(DB_NAME)
        conn.execute("CREATE TABLE IF NOT EXISTS devices (name TEXT, ip TEXT, port INTEGER)")
        conn.execute(
            "CREATE TABLE IF NOT EXISTS device_logs ("
            "ip TEXT, port INTEGER, ping INTEGER, port_status INTEGER, "
            "overall INTEGER, timestamp TEXT)"
        )
        conn.commit()
        conn.close()

    def open_context_menu(self, pos):
        row = self.table.rowAt(pos.y())
        if row < 0:
            return
        menu = QMenu(self)
        edit_action = menu.addAction("✏️ Edit Device")
        delete_action = menu.addAction("🗑 Delete Device")
        action = menu.exec_(self.table.viewport().mapToGlobal(pos))
        if action == edit_action:
            self.edit_device(row)
        elif action == delete_action:
            self.table.selectRow(row)
            self.delete_selected()

    def setup_ui(self):
        self.setWindowTitle("Device Monitor By Rabinn.ir")
        self.resize(1200, 750)
        self.setStyleSheet(MODERN_STYLE)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)

        tools = QHBoxLayout()
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 3600)
        self.interval_spin.setValue(10)

        self.ping_spin = QSpinBox()
        self.ping_spin.setRange(1, 10)
        self.ping_spin.setValue(1)

        btn_add = QPushButton("+ Manual Add")
        btn_add.clicked.connect(self.add_manual)

        btn_excel = QPushButton("Import Excel")
        btn_excel.clicked.connect(self.import_excel)

        btn_sql_profiles = QPushButton("SQL Sync Profiles")
        btn_sql_profiles.clicked.connect(self.open_sync_profiles)
        btn_sql_profiles.setStyleSheet(
            "QPushButton { background-color: #2D3245; border-color: #FFA500; } "
            "QPushButton:hover { background-color: #FFA500; color: #000; }"
        )

        btn_del = QPushButton("Delete Selected")
        btn_del.clicked.connect(self.delete_selected)

        tools.addWidget(QLabel("Interval (s):"))
        tools.addWidget(self.interval_spin)
        tools.addWidget(QLabel("Pings:"))
        tools.addWidget(self.ping_spin)
        tools.addStretch()
        tools.addWidget(btn_add)
        tools.addWidget(btn_excel)
        tools.addWidget(btn_sql_profiles)
        tools.addWidget(btn_del)

        main_layout.addLayout(tools)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels([
            "DEVICE NAME", "IP ADDRESS", "PORT",
            "PING", "SERVICE", "HEALTH STATUS"
        ])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.open_context_menu)
        self.table.cellDoubleClicked.connect(self.open_device_logs)
        self.table.viewport().installEventFilter(self)

        main_layout.addWidget(self.table)

        footer = QHBoxLayout()
        self.progress = QProgressBar()
        self.status_lbl = QLabel("Monitoring Engine Ready")
        footer.addWidget(self.status_lbl)
        footer.addWidget(self.progress)
        main_layout.addLayout(footer)

    def eventFilter(self, source, event):
        if event.type() == event.MouseButtonRelease and event.button() == Qt.MidButton:
            item = self.table.itemAt(event.pos())
            if item:
                webbrowser.open(f"http://{self.table.item(item.row(), 1).text()}")
        return super().eventFilter(source, event)

    def open_device_logs(self, row, col):
        name = self.table.item(row, 0).text()
        ip = self.table.item(row, 1).text()
        port = int(self.table.item(row, 2).text())
        LogWindow(ip, port, name).exec_()

    def load_from_db(self):
        conn = sqlite3.connect(DB_NAME)
        rows = conn.execute("SELECT name, ip, port FROM devices").fetchall()
        conn.close()
        self.devices[:] = [{"name": r[0], "ip": r[1], "port": r[2]} for r in rows]
        self.refresh_table_ui()

    def refresh_table_ui(self):
        self.table.setRowCount(len(self.devices))
        for r, d in enumerate(self.devices):
            it_name = QTableWidgetItem(d['name'])
            it_name.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, 0, it_name)
            self.table.setItem(r, 1, QTableWidgetItem(d['ip']))
            for c in range(2, 6):
                it = QTableWidgetItem(str(d['port']) if c == 2 else "-")
                it.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(r, c, it)

    def mark_row_checking(self, ip, port):
        for r in range(self.table.rowCount()):
            if (self.table.item(r, 1).text() == ip and
                    self.table.item(r, 2).text() == str(port)):
                it = QTableWidgetItem("Checking...")
                it.setTextAlignment(Qt.AlignCenter)
                it.setForeground(QColor("#FFA500"))
                self.table.setItem(r, 5, it)
                break

    def update_row(self, res):
        for r in range(self.table.rowCount()):
            if (self.table.item(r, 1).text() == res['ip'] and
                    self.table.item(r, 2).text() == str(res['port'])):
                data = {
                    3: "SUCCESS" if res['ping'] else "FAILED",
                    4: "OPEN" if res['port_ok'] else "CLOSED",
                    5: "ONLINE" if res['overall'] else "OFFLINE"
                }
                for c, txt in data.items():
                    it = QTableWidgetItem(txt)
                    it.setTextAlignment(Qt.AlignCenter)
                    if c == 5:
                        it.setForeground(QColor("#00F0FF") if res['overall'] else QColor("#FF4560"))
                    self.table.setItem(r, c, it)

    def add_manual(self):
        dlg = AddDeviceDialog(self)
        if dlg.exec_():
            n = dlg.name_in.text().strip()
            i = dlg.ip_in.text().strip()
            p = dlg.port_in.value()
            if n and i:
                conn = sqlite3.connect(DB_NAME)
                conn.execute("INSERT INTO devices VALUES (?,?,?)", (n, i, p))
                conn.commit()
                conn.close()
                self.load_from_db()

    def import_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import", "", "Excel (*.xlsx)")
        if path:
            try:
                df = pd.read_excel(path)
                conn = sqlite3.connect(DB_NAME)
                for _, r in df.iterrows():
                    conn.execute(
                        "INSERT INTO devices VALUES (?,?,?)",
                        (str(r['name']), str(r['ip']), int(r['port']))
                    )
                conn.commit()
                conn.close()
                self.load_from_db()
                QMessageBox.information(self, "Success", "Excel imported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to import Excel:\n{str(e)}")

    def open_sync_profiles(self):
        if not HAS_ODBC:
            QMessageBox.critical(
                self, "Error",
                "Library 'pyodbc' is not installed.\nPlease run: pip install pyodbc"
            )
            return
        dlg = SyncProfileDialog(self)
        if dlg.exec_():
            # بعد از بستن دیالوگ می‌توانیم همینجا هم یک سینک دستی بزنیم اگر خواستی
            self.auto_sync_sql()

    def auto_sync_sql(self):
        if not HAS_ODBC:
            return

        dlg = SyncProfileDialog(self)
        profiles = dlg.get_active_profiles()

        if not profiles:
            return

        all_devices = []

        for prof in profiles:
            try:
                conn_str = (
                    f"DRIVER={{SQL Server}};"
                    f"SERVER={prof['server']};"
                    f"DATABASE={prof['database']};"
                    f"UID={prof['username']};"
                    f"PWD={prof['password']}"
                )
                sql_conn = pyodbc.connect(conn_str)
                cursor = sql_conn.cursor()
                cursor.execute(prof['query'])
                rows = cursor.fetchall()
                sql_conn.close()
            except Exception:
                # اگر یک پروفایل خطا داشت بقیه را خراب نکن
                continue

            for row in rows:
                # انتظار داریم کوئری name, ip, port بده یا با ایندکس 0,1,2
                try:
                    name = getattr(row, 'name', None) or row[0]
                    ip = getattr(row, 'ip', None) or row[1]
                    port = getattr(row, 'port', None) or row[2]
                except Exception:
                    continue
                if not ip:
                    continue
                all_devices.append({
                    "name": str(name) if name else "Unknown",
                    "ip": str(ip),
                    "port": int(port) if port else 80
                })

        if not all_devices:
            return

        conn = sqlite3.connect(DB_NAME)
        conn.execute("DELETE FROM devices")
        for d in all_devices:
            conn.execute(
                "INSERT INTO devices VALUES (?,?,?)",
                (d['name'], d['ip'], d['port'])
            )
        conn.commit()
        conn.close()
        self.load_from_db()

    def delete_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        if rows and QMessageBox.question(
            self, "Confirm", f"Delete {len(rows)} devices?", QMessageBox.Yes | QMessageBox.No
        ) == QMessageBox.Yes:
            conn = sqlite3.connect(DB_NAME)
            for r in rows:
                conn.execute(
                    "DELETE FROM devices WHERE ip=? AND port=?",
                    (self.table.item(r, 1).text(),
                     int(self.table.item(r, 2).text()))
                )
            conn.commit()
            conn.close()
            self.load_from_db()

    def update_progress(self, rem, total):
        self.progress.setValue(int((rem / total) * 100))
        self.status_lbl.setText(f"Scan in {rem}s")

    def get_interval(self):
        return self.interval_spin.value()

    def get_ping_count(self):
        return self.ping_spin.value()

    def edit_device(self, row):
        old_name = self.table.item(row, 0).text()
        old_ip = self.table.item(row, 1).text()
        old_port = int(self.table.item(row, 2).text())

        dlg = AddDeviceDialog(self)
        dlg.name_in.setText(old_name)
        dlg.ip_in.setText(old_ip)
        dlg.port_in.setValue(old_port)

        if dlg.exec_():
            new_name = dlg.name_in.text().strip()
            new_ip = dlg.ip_in.text().strip()
            new_port = dlg.port_in.value()
            if not new_name or not new_ip:
                return
            conn = sqlite3.connect(DB_NAME)
            conn.execute(
                "UPDATE devices SET name=?, ip=?, port=? WHERE ip=? AND port=?",
                (new_name, new_ip, new_port, old_ip, old_port)
            )
            conn.commit()
            conn.close()
            self.load_from_db()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())
