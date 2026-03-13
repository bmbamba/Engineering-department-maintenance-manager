"""
MIT License
Copyright (c) 2026 Clan

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software to use, copy, modify, and distribute it, subject to the
following conditions:

1. The above copyright notice must be included in all copies.
2. This software is provided "as is", without warranty of any kind.
   Clan is not liable for any damages arising from its use.
3. You may not sell this software without written permission from Clan.

─────────────────────────────────────────────────────────────────────────────
Equipment Maintenance Manager  v1.0.0
─────────────────────────────────────────────────────────────────────────────
Architecture:
  ConfigManager        — JSON config persistence
  DatabaseManager      — All SQLite logic (equipment + history)
  EquipmentValidator   — Input validation layer
  EmailManager         — SMTP alert sending (background thread)
  EmailSettingsDialog  — Configure SMTP + alert rules
  ReportExporter       — CSV + PDF generation
  TableController      — Table population, filtering, sorting
  DashboardWidget      — Live stat cards
  _GradientLabel       — Gradient-painted heading widget
  HistoryDialog        — Per-equipment service history
  ServiceDialog        — Mark-serviced form
  EquipmentDialog      — Add / Edit form
  SettingsDialog       — Preferences
  StatisticsDialog     — Full stats with chart tabs
  MainWindow           — Shell: menus, toolbar, statusbar, layout

Requirements: pip install PySide6 reportlab
"""

import sys
import json
import sqlite3
import csv
import os
import logging
import smtplib
import threading
from datetime import datetime, date, timedelta
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QLineEdit, QPushButton,
    QLabel, QDialog, QFormLayout, QSpinBox, QMessageBox, QFileDialog,
    QStatusBar, QToolBar, QFrame, QHeaderView, QAbstractItemView,
    QComboBox, QDateEdit, QGroupBox, QSizePolicy, QDialogButtonBox,
    QTextEdit, QTabWidget, QSplitter, QMenu, QScrollArea, QCheckBox
)
from PySide6.QtCore import Qt, QDate, QTimer, Signal, QSortFilterProxyModel, QSize
from PySide6.QtGui import (
    QAction, QColor, QFont, QKeySequence, QCursor,
    QStandardItemModel, QStandardItem
)

# ── Logging ──────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    handlers=[
        logging.FileHandler("maintenance_manager.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("MaintManager")

# ── Constants ─────────────────────────────────────────────────────────────────

APP_VERSION = "1.0.0"
CONFIG_FILE = "config.json"

DEFAULT_CONFIG = {
    "company_name": "Engineering Department",
    "default_interval_days": 30,
    "db_path": "equipment.db",
    "window_width": 1280,
    "window_height": 800,
    "due_soon_days": 7,
    "email_enabled":   False,
    "email_smtp_host": "",
    "email_smtp_port": 587,
    "email_smtp_user": "",
    "email_smtp_pass": "",
    "email_use_tls":   True,
    "email_from":      "",
    "email_recipients": "",   # comma-separated
    "email_on_overdue": True,
    "email_daily_digest": False,
    "email_digest_hour": 8,
}

CATEGORIES = ["General", "Electrical", "Mechanical", "Hydraulic", "HVAC", "Safety", "IT / Network", "Vehicle", "Other"]

# ── Palette ───────────────────────────────────────────────────────────────────

C = {
    "bg":        "#1e1e2e",
    "surface":   "#252535",
    "overlay":   "#313244",
    "border":    "#45475a",
    "muted":     "#6c7086",
    "subtle":    "#a6adc8",
    "text":      "#cdd6f4",
    "blue":      "#89b4fa",
    "blue_dim":  "#74c7ec",
    "green":     "#a6e3a1",
    "yellow":    "#f9e2af",
    "red":       "#f38ba8",
    "purple":    "#cba6f7",
    "peach":     "#fab387",
    "sky":       "#89dceb",
}

DARK_STYLESHEET = f"""
QMainWindow, QDialog, QWidget {{
    background-color: {C['bg']};
    color: {C['text']};
    font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
    font-size: 13px;
}}
/* ── Menubar ── */
QMenuBar {{
    background-color: {C['surface']};
    color: {C['text']};
    padding: 2px 4px;
    border-bottom: 1px solid {C['border']};
}}
QMenuBar::item:selected {{ background-color: {C['overlay']}; border-radius: 4px; }}
QMenu {{
    background-color: {C['surface']};
    color: {C['text']};
    border: 1px solid {C['border']};
    padding: 4px;
    border-radius: 6px;
}}
QMenu::item {{ padding: 6px 24px 6px 12px; border-radius: 4px; }}
QMenu::item:selected {{ background-color: {C['overlay']}; }}
QMenu::separator {{ height: 1px; background: {C['border']}; margin: 4px 8px; }}
/* ── Toolbar ── */
QToolBar {{
    background-color: {C['surface']};
    border-bottom: 1px solid {C['border']};
    spacing: 4px;
    padding: 4px 10px;
}}
QToolBar::separator {{ width: 1px; background: {C['border']}; margin: 4px 4px; }}
QToolButton {{
    background: transparent;
    color: {C['text']};
    border: none;
    border-radius: 6px;
    padding: 5px 10px;
    font-size: 12px;
}}
QToolButton:hover {{ background-color: {C['overlay']}; }}
QToolButton:pressed {{ background-color: {C['border']}; }}
/* ── Status bar ── */
QStatusBar {{
    background-color: {C['surface']};
    color: {C['subtle']};
    border-top: 1px solid {C['border']};
    padding: 2px 10px;
    font-size: 12px;
}}
/* ── Table ── */
QTableView, QTableWidget {{
    background-color: {C['bg']};
    alternate-background-color: {C['surface']};
    color: {C['text']};
    gridline-color: {C['overlay']};
    border: 1px solid {C['border']};
    border-radius: 8px;
    selection-background-color: {C['overlay']};
    selection-color: {C['text']};
    outline: none;
}}
QTableWidget::item {{ padding: 5px 10px; border: none; }}
QTableWidget::item:selected {{ background-color: {C['overlay']}; }}
QHeaderView::section {{
    background-color: {C['surface']};
    color: {C['blue']};
    padding: 7px 10px;
    border: none;
    border-right: 1px solid {C['overlay']};
    border-bottom: 1px solid {C['border']};
    font-weight: 600;
    font-size: 12px;
}}
QHeaderView::section:hover {{ background-color: {C['overlay']}; }}
QHeaderView::section:checked {{ background-color: {C['overlay']}; }}
/* ── Buttons ── */
QPushButton {{
    background-color: {C['blue']};
    color: {C['bg']};
    border: none;
    border-radius: 6px;
    padding: 7px 16px;
    font-weight: 600;
    font-size: 12px;
    min-height: 30px;
}}
QPushButton:hover {{ background-color: #9dc4ff; }}
QPushButton:pressed {{ background-color: {C['blue_dim']}; }}
QPushButton:disabled {{ background-color: {C['border']}; color: {C['muted']}; }}
QPushButton#dangerBtn  {{ background-color: {C['red']};    color: {C['bg']}; }}
QPushButton#dangerBtn:hover  {{ background-color: #f5a8bc; }}
QPushButton#successBtn {{ background-color: {C['green']};  color: {C['bg']}; }}
QPushButton#successBtn:hover {{ background-color: #b8f0b3; }}
QPushButton#warningBtn {{ background-color: {C['yellow']}; color: {C['bg']}; }}
QPushButton#warningBtn:hover {{ background-color: #fdf0c8; }}
QPushButton#neutralBtn {{ background-color: {C['overlay']}; color: {C['text']}; }}
QPushButton#neutralBtn:hover {{ background-color: {C['border']}; }}
QPushButton#purpleBtn  {{ background-color: {C['purple']}; color: {C['bg']}; }}
QPushButton#purpleBtn:hover  {{ background-color: #d8b8fc; }}
/* ── Inputs ── */
QLineEdit, QSpinBox, QDateEdit, QComboBox, QTextEdit {{
    background-color: {C['overlay']};
    color: {C['text']};
    border: 1px solid {C['border']};
    border-radius: 6px;
    padding: 5px 10px;
    selection-background-color: {C['blue']};
    selection-color: {C['bg']};
    min-height: 28px;
}}
QLineEdit:focus, QSpinBox:focus, QDateEdit:focus,
QComboBox:focus, QTextEdit:focus {{ border: 1px solid {C['blue']}; }}
QLineEdit[placeholderText] {{ color: {C['muted']}; }}
QSpinBox::up-button, QSpinBox::down-button,
QDateEdit::up-button, QDateEdit::down-button {{
    background-color: {C['border']}; border-radius: 3px; width: 16px;
}}
QComboBox::drop-down {{ border: none; padding-right: 6px; }}
QComboBox QAbstractItemView {{
    background-color: {C['overlay']}; color: {C['text']};
    border: 1px solid {C['border']}; selection-background-color: {C['border']};
    border-radius: 4px;
}}
/* ── Group boxes ── */
QGroupBox {{
    border: 1px solid {C['border']};
    border-radius: 8px;
    margin-top: 14px;
    padding: 10px 8px 8px 8px;
    color: {C['blue']};
    font-weight: 600;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 6px;
    left: 12px;
}}
/* ── Tab widget ── */
QTabWidget::pane {{
    border: 1px solid {C['border']};
    border-radius: 8px;
    background: {C['bg']};
}}
QTabBar::tab {{
    background: {C['surface']};
    color: {C['subtle']};
    padding: 7px 18px;
    border-radius: 6px 6px 0 0;
    margin-right: 2px;
    font-size: 12px;
}}
QTabBar::tab:selected {{ background: {C['overlay']}; color: {C['text']}; }}
QTabBar::tab:hover {{ background: {C['overlay']}; }}
/* ── Scrollbars ── */
QScrollBar:vertical {{
    background: {C['bg']}; width: 10px; border-radius: 5px;
}}
QScrollBar::handle:vertical {{
    background: {C['border']}; border-radius: 5px; min-height: 30px;
}}
QScrollBar::handle:vertical:hover {{ background: {C['muted']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{
    background: {C['bg']}; height: 10px; border-radius: 5px;
}}
QScrollBar::handle:horizontal {{
    background: {C['border']}; border-radius: 5px; min-width: 30px;
}}
QScrollBar::handle:horizontal:hover {{ background: {C['muted']}; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
/* ── Labels ── */
QLabel {{ color: {C['text']}; }}
QLabel#pageTitle   {{ font-size: 20px; font-weight: 700; color: {C['blue']}; }}
QLabel#pageSubtitle {{ font-size: 12px; color: {C['subtle']}; }}
QLabel#statValue   {{ font-size: 30px; font-weight: 700; }}
QLabel#statLabel   {{ font-size: 11px; color: {C['subtle']}; }}
QLabel#sectionTitle {{ font-size: 14px; font-weight: 600; color: {C['blue']}; }}
QLabel#historyItem  {{ font-size: 12px; padding: 4px 0; }}
/* ── Frames / cards ── */
QFrame#statCard {{
    background-color: {C['surface']};
    border: 1px solid {C['overlay']};
    border-radius: 10px;
}}
QFrame#overdueRow {{
    background-color: rgba(243,139,168,0.08);
    border-left: 3px solid {C['red']};
    border-radius: 4px;
}}
QFrame#dueSoonRow {{
    background-color: rgba(249,226,175,0.08);
    border-left: 3px solid {C['yellow']};
    border-radius: 4px;
}}
QFrame#separator {{ background-color: {C['border']}; }}
/* ── Dialogs ── */
QDialogButtonBox QPushButton {{ min-width: 80px; }}
/* ── Checkboxes ── */
QCheckBox {{ color: {C['text']}; spacing: 8px; }}
QCheckBox::indicator {{
    width: 16px; height: 16px;
    border-radius: 4px; border: 1px solid {C['border']};
    background-color: {C['overlay']};
}}
QCheckBox::indicator:checked {{ background-color: {C['blue']}; border-color: {C['blue']}; }}
"""


# ═════════════════════════════════════════════════════════════════════════════
# CONFIG MANAGER
# ═════════════════════════════════════════════════════════════════════════════

class ConfigManager:
    def __init__(self):
        self.config = DEFAULT_CONFIG.copy()
        self._load()

    def _load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.config.update(json.load(f))
            except Exception as e:
                log.warning(f"Config load failed: {e}")

    def save(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            log.error(f"Config save failed: {e}")

    def get(self, key, default=None):
        return self.config.get(key, default)

    def set(self, key, value):
        self.config[key] = value
        self.save()


# ═════════════════════════════════════════════════════════════════════════════
# EQUIPMENT VALIDATOR
# ═════════════════════════════════════════════════════════════════════════════

class EquipmentValidator:
    @staticmethod
    def validate(data: dict, existing_ids: set, edit_id: str = None) -> list[str]:
        errors = []
        if not data.get("name", "").strip():
            errors.append("Equipment name is required.")
        if not data.get("equip_id", "").strip():
            errors.append("Equipment ID is required.")
        else:
            eid = data["equip_id"].strip()
            if eid in existing_ids and eid != edit_id:
                errors.append(f"Equipment ID '{eid}' is already in use.")
        if not data.get("last_service", ""):
            errors.append("Last service date is required.")
        if data.get("interval", 0) < 1:
            errors.append("Maintenance interval must be at least 1 day.")
        return errors


# ═════════════════════════════════════════════════════════════════════════════
# DATABASE MANAGER
# ═════════════════════════════════════════════════════════════════════════════

class DatabaseManager:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self._migrate()

    def _migrate(self):
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS equipment (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                name         TEXT NOT NULL,
                equip_id     TEXT UNIQUE NOT NULL,
                location     TEXT DEFAULT '',
                category     TEXT DEFAULT 'General',
                interval     INTEGER NOT NULL DEFAULT 30,
                last_service TEXT NOT NULL,
                next_service TEXT NOT NULL,
                notes        TEXT DEFAULT '',
                created_at   TEXT DEFAULT (date('now'))
            );
            CREATE TABLE IF NOT EXISTS service_history (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                equipment_id INTEGER NOT NULL,
                serviced_on  TEXT NOT NULL,
                technician   TEXT DEFAULT '',
                notes        TEXT DEFAULT '',
                FOREIGN KEY (equipment_id) REFERENCES equipment(id) ON DELETE CASCADE
            );
        """)
        # Add 'category' column if upgrading from v2
        try:
            self.conn.execute("ALTER TABLE equipment ADD COLUMN category TEXT DEFAULT 'General'")
        except sqlite3.OperationalError:
            pass
        self.conn.commit()

    # ── Equipment CRUD ────────────────────────────────────────────────────────

    def add_equipment(self, d: dict) -> int:
        next_dt = self._calc_next(d["last_service"], d["interval"])
        cur = self.conn.execute("""
            INSERT INTO equipment (name, equip_id, location, category, interval, last_service, next_service, notes)
            VALUES (:name, :equip_id, :location, :category, :interval, :last_service, :next_service, :notes)
        """, {**d, "next_service": next_dt})
        self.conn.commit()
        log.info(f"Added equipment: {d['equip_id']}")
        return cur.lastrowid

    def update_equipment(self, row_id: int, d: dict):
        next_dt = self._calc_next(d["last_service"], d["interval"])
        self.conn.execute("""
            UPDATE equipment SET name=:name, equip_id=:equip_id, location=:location,
            category=:category, interval=:interval, last_service=:last_service,
            next_service=:next_service, notes=:notes WHERE id=:id
        """, {**d, "next_service": next_dt, "id": row_id})
        self.conn.commit()
        log.info(f"Updated equipment id={row_id}")

    def remove_equipment(self, row_id: int):
        self.conn.execute("DELETE FROM equipment WHERE id=?", (row_id,))
        self.conn.commit()
        log.info(f"Removed equipment id={row_id}")

    def mark_serviced(self, row_id: int, interval: int, technician: str = "", notes: str = ""):
        today = date.today().isoformat()
        next_dt = (date.today() + timedelta(days=interval)).isoformat()
        self.conn.execute(
            "UPDATE equipment SET last_service=?, next_service=? WHERE id=?",
            (today, next_dt, row_id)
        )
        self.conn.execute(
            "INSERT INTO service_history (equipment_id, serviced_on, technician, notes) VALUES (?,?,?,?)",
            (row_id, today, technician, notes)
        )
        self.conn.commit()
        log.info(f"Marked serviced: id={row_id}")

    def get_all(self) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM equipment ORDER BY name"
        ).fetchall()]

    def get_existing_ids(self) -> set:
        return {r[0] for r in self.conn.execute("SELECT equip_id FROM equipment").fetchall()}

    # ── History ───────────────────────────────────────────────────────────────

    def get_history(self, equipment_id: int) -> list[dict]:
        return [dict(r) for r in self.conn.execute("""
            SELECT serviced_on, technician, notes
            FROM service_history WHERE equipment_id=?
            ORDER BY serviced_on DESC
        """, (equipment_id,)).fetchall()]

    # ── Statistics ────────────────────────────────────────────────────────────

    def get_statistics(self) -> dict:
        today = date.today().isoformat()
        soon = (date.today() + timedelta(days=7)).isoformat()
        total    = self.conn.execute("SELECT COUNT(*) FROM equipment").fetchone()[0]
        overdue  = self.conn.execute("SELECT COUNT(*) FROM equipment WHERE next_service < ?", (today,)).fetchone()[0]
        due_soon = self.conn.execute(
            "SELECT COUNT(*) FROM equipment WHERE next_service BETWEEN ? AND ?",
            (today, soon)
        ).fetchone()[0]
        healthy  = max(0, total - overdue - due_soon)
        # Category breakdown
        cat_rows = self.conn.execute(
            "SELECT category, COUNT(*) as cnt FROM equipment GROUP BY category ORDER BY cnt DESC"
        ).fetchall()
        # Upcoming in next 30d
        upcoming = [dict(r) for r in self.conn.execute("""
            SELECT name, equip_id, next_service FROM equipment
            WHERE next_service BETWEEN ? AND ?
            ORDER BY next_service
        """, (today, (date.today() + timedelta(days=30)).isoformat())).fetchall()]
        return {
            "total": total, "overdue": overdue,
            "due_soon": due_soon, "healthy": healthy,
            "categories": [dict(r) for r in cat_rows],
            "upcoming": upcoming,
        }

    # ── Import / Export ───────────────────────────────────────────────────────

    def export_csv(self, filepath: str, rows: list[dict] = None) -> int:
        rows = rows or self.get_all()
        if not rows:
            return 0
        fields = ["name", "equip_id", "location", "category", "interval",
                  "last_service", "next_service", "notes"]
        with open(filepath, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)
        return len(rows)

    def import_csv(self, filepath: str) -> tuple[int, list[str]]:
        imported, errors = 0, []
        with open(filepath, "r", encoding="utf-8") as f:
            for i, row in enumerate(csv.DictReader(f), 1):
                try:
                    next_dt = self._calc_next(row.get("last_service", ""), int(row.get("interval", 30)))
                    self.conn.execute("""
                        INSERT OR REPLACE INTO equipment
                        (name, equip_id, location, category, interval, last_service, next_service, notes)
                        VALUES (?,?,?,?,?,?,?,?)
                    """, (
                        row.get("name",""), row.get("equip_id",""),
                        row.get("location",""), row.get("category","General"),
                        int(row.get("interval",30)), row.get("last_service",""),
                        next_dt, row.get("notes","")
                    ))
                    imported += 1
                except Exception as e:
                    errors.append(f"Row {i}: {e}")
        self.conn.commit()
        log.info(f"CSV import: {imported} records, {len(errors)} errors")
        return imported, errors

    def save_project(self, filepath: str):
        data = {
            "app": "EquipmentMaintenanceManagerPro",
            "version": APP_VERSION,
            "exported_at": datetime.now().isoformat(),
            "equipment": self.get_all(),
        }
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    def load_project(self, filepath: str) -> int:
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.conn.execute("DELETE FROM service_history")
        self.conn.execute("DELETE FROM equipment")
        for eq in data.get("equipment", []):
            self.conn.execute("""
                INSERT INTO equipment (name, equip_id, location, category, interval, last_service, next_service, notes)
                VALUES (?,?,?,?,?,?,?,?)
            """, (eq["name"], eq["equip_id"], eq.get("location",""),
                  eq.get("category","General"), eq["interval"],
                  eq["last_service"], eq["next_service"], eq.get("notes","")))
        self.conn.commit()
        return len(data.get("equipment", []))

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _calc_next(last_service: str, interval: int) -> str:
        try:
            dt = datetime.strptime(last_service, "%Y-%m-%d").date()
            return (dt + timedelta(days=interval)).isoformat()
        except ValueError:
            return date.today().isoformat()

    def close(self):
        if self.conn:
            self.conn.close()


# ═════════════════════════════════════════════════════════════════════════════
# REPORT EXPORTER  (CSV + PDF via reportlab)
# ═════════════════════════════════════════════════════════════════════════════

class ReportExporter:
    @staticmethod
    def export_pdf(filepath: str, rows: list[dict], company: str):
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib import colors
            from reportlab.lib.units import mm
            from reportlab.platypus import (
                SimpleDocTemplate, Table, TableStyle, Paragraph,
                Spacer, HRFlowable
            )
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_CENTER, TA_LEFT
        except ImportError:
            return False, "reportlab not installed. Run: pip install reportlab"

        try:
            doc = SimpleDocTemplate(
                filepath, pagesize=landscape(A4),
                leftMargin=15*mm, rightMargin=15*mm,
                topMargin=15*mm, bottomMargin=15*mm
            )
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                "title", parent=styles["Title"],
                fontSize=16, spaceAfter=4,
                textColor=colors.HexColor("#1a1a2e")
            )
            sub_style = ParagraphStyle(
                "sub", parent=styles["Normal"],
                fontSize=9, textColor=colors.HexColor("#555566")
            )

            today = date.today().isoformat()
            story = [
                Paragraph(f"{company} — Equipment Maintenance Report", title_style),
                Paragraph(f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')}  |  "
                          f"Total records: {len(rows)}", sub_style),
                Spacer(1, 6*mm),
            ]

            headers = ["#", "Equipment Name", "ID", "Location", "Category",
                       "Interval", "Last Service", "Next Service", "Status", "Notes"]
            table_data = [headers]

            for i, eq in enumerate(rows, 1):
                next_s = eq.get("next_service","")
                if next_s:
                    nd = datetime.strptime(next_s, "%Y-%m-%d").date()
                    delta = (nd - date.today()).days
                    status = f"Overdue {-delta}d" if delta < 0 else (
                        f"Due in {delta}d" if delta <= 7 else f"OK ({delta}d)"
                    )
                else:
                    status = "—"
                table_data.append([
                    str(i),
                    eq.get("name",""),
                    eq.get("equip_id",""),
                    eq.get("location","") or "—",
                    eq.get("category",""),
                    f"{eq.get('interval','')}d",
                    eq.get("last_service",""),
                    eq.get("next_service",""),
                    status,
                    (eq.get("notes","") or "")[:40],
                ])

            col_widths = [10*mm,52*mm,28*mm,38*mm,28*mm,18*mm,28*mm,28*mm,28*mm,48*mm]
            t = Table(table_data, colWidths=col_widths, repeatRows=1)

            style = TableStyle([
                ("BACKGROUND",  (0,0), (-1,0), colors.HexColor("#1a1a2e")),
                ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
                ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE",    (0,0), (-1,0), 9),
                ("FONTSIZE",    (0,1), (-1,-1), 8),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),
                    [colors.HexColor("#f8f8fc"), colors.white]),
                ("GRID",        (0,0), (-1,-1), 0.4, colors.HexColor("#ccccdd")),
                ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
                ("LEFTPADDING", (0,0), (-1,-1), 4),
                ("RIGHTPADDING",(0,0), (-1,-1), 4),
                ("TOPPADDING",  (0,0), (-1,-1), 4),
                ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ])

            # Highlight overdue rows red, due-soon yellow
            for i, eq in enumerate(rows, 1):
                next_s = eq.get("next_service","")
                if next_s:
                    nd = datetime.strptime(next_s, "%Y-%m-%d").date()
                    delta = (nd - date.today()).days
                    if delta < 0:
                        style.add("BACKGROUND", (0,i), (-1,i), colors.HexColor("#fde8ed"))
                        style.add("TEXTCOLOR",  (8,i), (8,i),  colors.HexColor("#cc2244"))
                    elif delta <= 7:
                        style.add("BACKGROUND", (0,i), (-1,i), colors.HexColor("#fef9e7"))
                        style.add("TEXTCOLOR",  (8,i), (8,i),  colors.HexColor("#aa6600"))

            t.setStyle(style)
            story.append(t)
            doc.build(story)
            return True, ""
        except Exception as e:
            log.error(f"PDF export failed: {e}")
            return False, str(e)


# ═════════════════════════════════════════════════════════════════════════════
# TABLE CONTROLLER
# ═════════════════════════════════════════════════════════════════════════════

class TableController:
    """Manages populating and filtering the equipment QTableWidget."""

    COLS = ["Equipment Name", "ID", "Location", "Category",
            "Interval", "Last Service", "Next Service", "Status"]
    COL_STATUS = 7

    def __init__(self, table: QTableWidget):
        self.table = table
        self._all_rows: list[dict] = []
        self._setup_table()

    def _setup_table(self):
        t = self.table
        t.setColumnCount(len(self.COLS))
        t.setHorizontalHeaderLabels(self.COLS)
        t.setAlternatingRowColors(True)
        t.setSelectionBehavior(QAbstractItemView.SelectRows)
        t.setSelectionMode(QAbstractItemView.ExtendedSelection)
        t.setEditTriggers(QAbstractItemView.NoEditTriggers)
        t.setSortingEnabled(True)
        t.verticalHeader().setDefaultSectionSize(36)
        t.verticalHeader().setVisible(True)
        t.horizontalHeader().setStretchLastSection(False)
        hh = t.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.Stretch)      # name
        hh.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(7, QHeaderView.Fixed)
        t.setColumnWidth(7, 130)

    def load(self, rows: list[dict]):
        self._all_rows = rows

    def apply_filter(self, text: str, status_index: int) -> int:
        text = text.strip().lower()
        today = date.today()

        visible = []
        for eq in self._all_rows:
            if text and not any(
                text in str(eq.get(k, "")).lower()
                for k in ("name", "equip_id", "location", "category", "notes")
            ):
                continue
            ns = eq.get("next_service", "")
            is_overdue = is_due_soon = False
            if ns:
                nd = datetime.strptime(ns, "%Y-%m-%d").date()
                delta = (nd - today).days
                is_overdue  = delta < 0
                is_due_soon = not is_overdue and delta <= 7
            if status_index == 1 and (is_overdue or is_due_soon):
                continue
            if status_index == 2 and not is_due_soon:
                continue
            if status_index == 3 and not is_overdue:
                continue
            visible.append(eq)

        self._populate(visible)
        return len(visible)

    def _populate(self, rows: list[dict]):
        t = self.table
        t.setSortingEnabled(False)
        t.setRowCount(0)
        today = date.today()

        for eq in rows:
            r = t.rowCount()
            t.insertRow(r)

            ns = eq.get("next_service", "")
            if ns:
                nd = datetime.strptime(ns, "%Y-%m-%d").date()
                delta = (nd - today).days
                if delta < 0:
                    status_text  = f"⚠  Overdue {-delta}d"
                    status_color = QColor(C["red"])
                    row_bg       = QColor(243, 139, 168, 25)
                elif delta <= 7:
                    status_text  = f"⏱  Due in {delta}d"
                    status_color = QColor(C["yellow"])
                    row_bg       = QColor(249, 226, 175, 20)
                else:
                    status_text  = f"✓  OK  ({delta}d)"
                    status_color = QColor(C["green"])
                    row_bg       = None
            else:
                status_text  = "—"
                status_color = QColor(C["muted"])
                row_bg       = None

            cells = [
                eq.get("name",""),
                eq.get("equip_id",""),
                eq.get("location","") or "—",
                eq.get("category","General"),
                str(eq.get("interval","")),
                eq.get("last_service","") or "—",
                eq.get("next_service","") or "—",
                status_text,
            ]

            for col, val in enumerate(cells):
                item = QTableWidgetItem(val)
                item.setTextAlignment(
                    Qt.AlignCenter if col in (4, 5, 6, 7) else Qt.AlignVCenter | Qt.AlignLeft
                )
                if col == self.COL_STATUS:
                    item.setForeground(status_color)
                if row_bg:
                    item.setBackground(row_bg)
                t.setItem(r, col, item)

            t.item(r, 0).setData(Qt.UserRole, eq["id"])

        t.setSortingEnabled(True)

    def selected_ids(self) -> list[int]:
        seen, ids = set(), []
        for item in self.table.selectedItems():
            row = item.row()
            if row not in seen:
                seen.add(row)
                d = self.table.item(row, 0).data(Qt.UserRole)
                if d is not None:
                    ids.append(d)
        return ids

    def selected_equipment(self) -> list[dict]:
        ids = set(self.selected_ids())
        return [eq for eq in self._all_rows if eq["id"] in ids]

    def first_selected(self) -> dict | None:
        sel = self.selected_equipment()
        return sel[0] if sel else None


# ═════════════════════════════════════════════════════════════════════════════
# DASHBOARD WIDGET
# ═════════════════════════════════════════════════════════════════════════════

class DashboardWidget(QWidget):
    def __init__(self):
        super().__init__()
        self._cards: dict[str, QLabel] = {}
        self._build()

    def _build(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        specs = [
            ("total",    "Total Equipment", C["blue"],   "🔧"),
            ("healthy",  "✓ Healthy",        C["green"],  "✅"),
            ("due_soon", "⏱ Due Soon (7d)",  C["yellow"], "🔔"),
            ("overdue",  "⚠ Overdue",        C["red"],    "🚨"),
        ]
        for key, label, color, icon in specs:
            card = QFrame()
            card.setObjectName("statCard")
            card.setMinimumHeight(88)
            card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            vl = QVBoxLayout(card)
            vl.setAlignment(Qt.AlignCenter)
            vl.setSpacing(3)
            vl.setContentsMargins(12, 10, 12, 10)

            val_lbl = QLabel("0")
            val_lbl.setObjectName("statValue")
            val_lbl.setStyleSheet(f"color: {color};")
            val_lbl.setAlignment(Qt.AlignCenter)

            lbl = QLabel(label)
            lbl.setObjectName("statLabel")
            lbl.setAlignment(Qt.AlignCenter)

            vl.addWidget(val_lbl)
            vl.addWidget(lbl)
            layout.addWidget(card)
            self._cards[key] = val_lbl

    def update(self, stats: dict):
        for key, lbl in self._cards.items():
            lbl.setText(str(stats.get(key, 0)))


# ═════════════════════════════════════════════════════════════════════════════
# DIALOGS
# ═════════════════════════════════════════════════════════════════════════════

class ServiceDialog(QDialog):
    """Confirm mark-as-serviced and capture technician + notes."""

    def __init__(self, parent, eq: dict):
        super().__init__(parent)
        self.setWindowTitle("Mark as Serviced")
        self.setMinimumWidth(380)
        self.setModal(True)
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(20,20,20,20)

        layout.addWidget(self._title(f"Confirm Service — {eq['name']}"))
        layout.addWidget(self._sep())

        form = QFormLayout()
        form.setSpacing(10)
        form.setLabelAlignment(Qt.AlignRight)

        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        form.addRow("Service Date", self.date_edit)

        self.tech_edit = QLineEdit()
        self.tech_edit.setPlaceholderText("Technician name (optional)")
        form.addRow("Technician", self.tech_edit)

        self.notes_edit = QLineEdit()
        self.notes_edit.setPlaceholderText("Optional service notes...")
        form.addRow("Notes", self.notes_edit)

        layout.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Mark Serviced")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_data(self):
        return {
            "date": self.date_edit.date().toString("yyyy-MM-dd"),
            "technician": self.tech_edit.text().strip(),
            "notes": self.notes_edit.text().strip(),
        }

    @staticmethod
    def _title(text):
        lbl = QLabel(text)
        lbl.setObjectName("sectionTitle")
        return lbl

    @staticmethod
    def _sep():
        f = QFrame(); f.setObjectName("separator"); f.setFixedHeight(1)
        return f


class HistoryDialog(QDialog):
    """Shows full service history for one piece of equipment."""

    def __init__(self, parent, db: DatabaseManager, eq: dict):
        super().__init__(parent)
        self.setWindowTitle(f"Service History — {eq['name']} ({eq['equip_id']})")
        self.setMinimumSize(520, 400)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20,20,20,20)
        layout.setSpacing(12)

        title = QLabel(f"Service History")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        layout.addWidget(self._sep())

        records = db.get_history(eq["id"])

        if not records:
            lbl = QLabel("No service history recorded yet.")
            lbl.setStyleSheet(f"color: {C['subtle']}; padding: 20px;")
            lbl.setAlignment(Qt.AlignCenter)
            layout.addWidget(lbl)
        else:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            inner = QWidget()
            inner_layout = QVBoxLayout(inner)
            inner_layout.setSpacing(4)
            inner_layout.setContentsMargins(0,0,0,0)

            for rec in records:
                row = QFrame()
                row.setStyleSheet(
                    f"background:{C['surface']}; border:1px solid {C['overlay']};"
                    f"border-radius:6px; padding:2px;"
                )
                rl = QHBoxLayout(row)
                rl.setContentsMargins(12,8,12,8)

                date_lbl = QLabel(f"📅  {rec['serviced_on']}")
                date_lbl.setStyleSheet(f"color:{C['blue']}; font-weight:600; min-width:120px;")
                rl.addWidget(date_lbl)

                tech = rec.get("technician","") or "—"
                tech_lbl = QLabel(f"👤 {tech}")
                tech_lbl.setStyleSheet(f"color:{C['subtle']}; min-width:140px;")
                rl.addWidget(tech_lbl)

                notes = rec.get("notes","") or "—"
                notes_lbl = QLabel(notes)
                notes_lbl.setStyleSheet(f"color:{C['text']};")
                notes_lbl.setWordWrap(True)
                rl.addWidget(notes_lbl, 1)

                inner_layout.addWidget(row)

            inner_layout.addStretch()
            scroll.setWidget(inner)
            layout.addWidget(scroll)

        info_lbl = QLabel(
            f"Equipment: {eq['name']}  |  ID: {eq['equip_id']}  |  "
            f"Interval: {eq['interval']} days  |  Next service: {eq.get('next_service','—')}"
        )
        info_lbl.setStyleSheet(f"color:{C['subtle']}; font-size:11px;")
        layout.addWidget(info_lbl)

        close_btn = QPushButton("Close")
        close_btn.setObjectName("neutralBtn")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)

    @staticmethod
    def _sep():
        f = QFrame(); f.setObjectName("separator"); f.setFixedHeight(1)
        return f


class EquipmentDialog(QDialog):
    """Add or Edit equipment."""

    def __init__(self, parent, config: ConfigManager,
                 equipment: dict = None, existing_ids: set = None):
        super().__init__(parent)
        self.config = config
        self.equipment = equipment
        self.existing_ids = existing_ids or set()
        self.is_edit = equipment is not None
        self.setWindowTitle("Edit Equipment" if self.is_edit else "Add Equipment")
        self.setMinimumWidth(440)
        self.setModal(True)
        self._build()
        if self.is_edit:
            self._populate()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(22, 20, 22, 20)

        title = QLabel("Edit Equipment" if self.is_edit else "Register New Equipment")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)

        sep = QFrame(); sep.setObjectName("separator"); sep.setFixedHeight(1)
        layout.addWidget(sep)

        form = QFormLayout()
        form.setSpacing(11)
        form.setLabelAlignment(Qt.AlignRight)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("e.g. Hydraulic Press Alpha")
        form.addRow("Equipment Name *", self.name_edit)

        self.id_edit = QLineEdit()
        self.id_edit.setPlaceholderText("e.g. HPA-001")
        form.addRow("Equipment ID *", self.id_edit)

        self.loc_edit = QLineEdit()
        self.loc_edit.setPlaceholderText("e.g. Building 3, Bay 2")
        form.addRow("Location", self.loc_edit)

        self.cat_combo = QComboBox()
        self.cat_combo.addItems(CATEGORIES)
        form.addRow("Category", self.cat_combo)

        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 3650)
        self.interval_spin.setValue(self.config.get("default_interval_days", 30))
        self.interval_spin.setSuffix(" days")
        form.addRow("Maintenance Interval *", self.interval_spin)

        self.last_service_edit = QDateEdit(QDate.currentDate())
        self.last_service_edit.setCalendarPopup(True)
        self.last_service_edit.setDisplayFormat("yyyy-MM-dd")
        form.addRow("Last Service Date *", self.last_service_edit)

        self.notes_edit = QTextEdit()
        self.notes_edit.setPlaceholderText("Optional notes about this equipment...")
        self.notes_edit.setFixedHeight(70)
        form.addRow("Notes", self.notes_edit)

        layout.addLayout(form)

        self.error_lbl = QLabel("")
        self.error_lbl.setStyleSheet(f"color: {C['red']}; font-size: 12px;")
        self.error_lbl.setWordWrap(True)
        layout.addWidget(self.error_lbl)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Save" if self.is_edit else "Add Equipment")
        btns.accepted.connect(self._validate_and_accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _populate(self):
        eq = self.equipment
        self.name_edit.setText(eq.get("name",""))
        self.id_edit.setText(eq.get("equip_id",""))
        self.loc_edit.setText(eq.get("location",""))
        idx = self.cat_combo.findText(eq.get("category","General"))
        if idx >= 0:
            self.cat_combo.setCurrentIndex(idx)
        self.interval_spin.setValue(eq.get("interval", 30))
        if eq.get("last_service"):
            self.last_service_edit.setDate(QDate.fromString(eq["last_service"],"yyyy-MM-dd"))
        self.notes_edit.setPlainText(eq.get("notes",""))

    def _validate_and_accept(self):
        data = self.get_data()
        edit_id = self.equipment["equip_id"] if self.is_edit else None
        errors = EquipmentValidator.validate(data, self.existing_ids, edit_id)
        if errors:
            self.error_lbl.setText("• " + "\n• ".join(errors))
            return
        self.accept()

    def get_data(self) -> dict:
        return {
            "name":         self.name_edit.text().strip(),
            "equip_id":     self.id_edit.text().strip(),
            "location":     self.loc_edit.text().strip(),
            "category":     self.cat_combo.currentText(),
            "interval":     self.interval_spin.value(),
            "last_service": self.last_service_edit.date().toString("yyyy-MM-dd"),
            "notes":        self.notes_edit.toPlainText().strip(),
        }


class SettingsDialog(QDialog):
    def __init__(self, parent, config: ConfigManager):
        super().__init__(parent)
        self.config = config
        self.setWindowTitle("Preferences")
        self.setMinimumWidth(420)
        self.setModal(True)
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(22, 20, 22, 20)

        title = QLabel("⚙   Application Preferences")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        sep = QFrame(); sep.setObjectName("separator"); sep.setFixedHeight(1)
        layout.addWidget(sep)

        tabs = QTabWidget()

        # ── General tab ──
        gen_tab = QWidget()
        gen_form = QFormLayout(gen_tab)
        gen_form.setSpacing(12)
        gen_form.setContentsMargins(14, 14, 14, 14)
        gen_form.setLabelAlignment(Qt.AlignRight)

        self.company_edit = QLineEdit(self.config.get("company_name",""))
        gen_form.addRow("Company / Dept. Name", self.company_edit)

        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 3650)
        self.interval_spin.setValue(self.config.get("default_interval_days", 30))
        self.interval_spin.setSuffix(" days")
        gen_form.addRow("Default Interval", self.interval_spin)

        self.due_soon_spin = QSpinBox()
        self.due_soon_spin.setRange(1, 90)
        self.due_soon_spin.setValue(self.config.get("due_soon_days", 7))
        self.due_soon_spin.setSuffix(" days")
        gen_form.addRow("'Due Soon' Warning Window", self.due_soon_spin)

        tabs.addTab(gen_tab, "⚙  General")

        # ── Database tab ──
        db_tab = QWidget()
        db_layout = QVBoxLayout(db_tab)
        db_layout.setContentsMargins(14,14,14,14)
        db_layout.setSpacing(10)

        db_path_layout = QHBoxLayout()
        self.db_path_edit = QLineEdit(self.config.get("db_path","equipment.db"))
        self.db_path_edit.setReadOnly(True)
        db_path_layout.addWidget(self.db_path_edit)
        browse_btn = QPushButton("Browse…")
        browse_btn.setObjectName("neutralBtn")
        browse_btn.clicked.connect(self._browse_db)
        db_path_layout.addWidget(browse_btn)

        db_layout.addWidget(QLabel("Database File Path:"))
        db_layout.addLayout(db_path_layout)
        db_layout.addWidget(QLabel(
            "⚠  Changing the DB path requires a restart to take effect."
        ))
        db_layout.addStretch()

        tabs.addTab(db_tab, "🗄  Database")
        layout.addWidget(tabs)

        btns = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _browse_db(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Select Database Location", self.db_path_edit.text(),
            "SQLite Database (*.db);;All Files (*)"
        )
        if path:
            self.db_path_edit.setText(path)

    def _save(self):
        self.config.set("company_name", self.company_edit.text().strip() or "Engineering Department")
        self.config.set("default_interval_days", self.interval_spin.value())
        self.config.set("due_soon_days", self.due_soon_spin.value())
        self.config.set("db_path", self.db_path_edit.text())
        self.accept()


class StatisticsDialog(QDialog):
    """Full statistics panel with bar chart (pure Qt, no matplotlib dep)."""

    def __init__(self, parent, db: DatabaseManager):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Statistics Dashboard")
        self.setMinimumSize(640, 520)
        self.setModal(True)
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20,20,20,20)
        layout.setSpacing(14)

        title = QLabel("📊  Fleet Statistics")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        sep = QFrame(); sep.setObjectName("separator"); sep.setFixedHeight(1)
        layout.addWidget(sep)

        stats = self.db.get_statistics()

        # ── Top stat cards ──
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(10)
        for key, label, color in [
            ("total",    "Total Equipment", C["blue"]),
            ("healthy",  "✓ Healthy",        C["green"]),
            ("due_soon", "⏱ Due Soon",       C["yellow"]),
            ("overdue",  "⚠ Overdue",        C["red"]),
        ]:
            card = QFrame()
            card.setObjectName("statCard")
            vl = QVBoxLayout(card)
            vl.setAlignment(Qt.AlignCenter)
            vl.setContentsMargins(10,10,10,10)
            val = QLabel(str(stats.get(key, 0)))
            val.setObjectName("statValue")
            val.setStyleSheet(f"color:{color}; font-size:28px; font-weight:700;")
            val.setAlignment(Qt.AlignCenter)
            lbl = QLabel(label)
            lbl.setObjectName("statLabel")
            lbl.setAlignment(Qt.AlignCenter)
            vl.addWidget(val)
            vl.addWidget(lbl)
            cards_layout.addWidget(card)
        layout.addLayout(cards_layout)

        # ── Tabs: Category breakdown + Upcoming ──
        tabs = QTabWidget()

        # Category tab
        cat_tab = QWidget()
        cat_layout = QVBoxLayout(cat_tab)
        cat_layout.setContentsMargins(12,12,12,12)
        cat_layout.setSpacing(8)

        cats = stats.get("categories", [])
        if cats:
            max_cnt = max(c["cnt"] for c in cats) or 1
            for cat in cats:
                row = QHBoxLayout()
                row.setSpacing(10)
                name_lbl = QLabel(cat["category"] or "Unknown")
                name_lbl.setFixedWidth(130)
                row.addWidget(name_lbl)

                bar_container = QFrame()
                bar_container.setFixedHeight(22)
                bar_container.setStyleSheet(
                    f"background:{C['overlay']}; border-radius:4px;"
                )
                bar_container.setMinimumWidth(200)

                bar_fill = QFrame(bar_container)
                w = max(4, int(300 * cat["cnt"] / max_cnt))
                bar_fill.setFixedSize(w, 22)
                bar_fill.setStyleSheet(
                    f"background:{C['blue']}; border-radius:4px;"
                )

                row.addWidget(bar_container, 1)
                count_lbl = QLabel(str(cat["cnt"]))
                count_lbl.setFixedWidth(30)
                count_lbl.setStyleSheet(f"color:{C['blue']}; font-weight:600;")
                row.addWidget(count_lbl)
                cat_layout.addLayout(row)
        else:
            cat_layout.addWidget(QLabel("No data."))
        cat_layout.addStretch()
        tabs.addTab(cat_tab, "📂  By Category")

        # Upcoming tab
        up_tab = QWidget()
        up_layout = QVBoxLayout(up_tab)
        up_layout.setContentsMargins(12,12,12,12)
        up_layout.setSpacing(6)

        upcoming = stats.get("upcoming", [])
        if upcoming:
            for eq in upcoming[:20]:
                nd = datetime.strptime(eq["next_service"], "%Y-%m-%d").date()
                delta = (nd - date.today()).days
                color = C["red"] if delta < 0 else (C["yellow"] if delta <= 7 else C["green"])
                row = QFrame()
                row.setStyleSheet(
                    f"background:{C['surface']}; border:1px solid {C['overlay']}; border-radius:5px;"
                )
                rl = QHBoxLayout(row)
                rl.setContentsMargins(10,6,10,6)
                rl.addWidget(QLabel(eq["name"]))
                rl.addStretch()
                due_lbl = QLabel(eq["next_service"])
                due_lbl.setStyleSheet(f"color:{color}; font-weight:600;")
                rl.addWidget(due_lbl)
                up_layout.addWidget(row)
        else:
            up_layout.addWidget(QLabel("No scheduled maintenance in next 30 days."))
        up_layout.addStretch()
        tabs.addTab(up_tab, "📅  Upcoming (30d)")

        # Overdue tab
        ov_tab = QWidget()
        ov_layout = QVBoxLayout(ov_tab)
        ov_layout.setContentsMargins(12,12,12,12)
        ov_layout.setSpacing(6)

        all_eq = self.db.get_all()
        overdue_eq = [e for e in all_eq if e.get("next_service","") < date.today().isoformat()]
        if overdue_eq:
            for eq in overdue_eq:
                nd = datetime.strptime(eq["next_service"], "%Y-%m-%d").date()
                days_ago = (date.today() - nd).days
                row = QFrame()
                row.setStyleSheet(
                    f"background:rgba(243,139,168,0.1); border-left:3px solid {C['red']};"
                    f"border-radius:5px;"
                )
                rl = QHBoxLayout(row)
                rl.setContentsMargins(12,7,12,7)
                rl.addWidget(QLabel(f"{eq['name']}"))
                rl.addWidget(QLabel(f"({eq['equip_id']})"))
                rl.addWidget(QLabel(eq.get("location","—")))
                rl.addStretch()
                ol = QLabel(f"{days_ago} days overdue")
                ol.setStyleSheet(f"color:{C['red']}; font-weight:600;")
                rl.addWidget(ol)
                ov_layout.addWidget(row)
        else:
            ok_lbl = QLabel("✅  No overdue equipment!")
            ok_lbl.setStyleSheet(f"color:{C['green']}; font-size:14px;")
            ok_lbl.setAlignment(Qt.AlignCenter)
            ov_layout.addWidget(ok_lbl)
        ov_layout.addStretch()
        title_suffix = f" ({len(overdue_eq)})" if overdue_eq else ""
        tabs.addTab(ov_tab, f"⚠  Overdue{title_suffix}")

        layout.addWidget(tabs, 1)

        close_btn = QPushButton("Close")
        close_btn.setObjectName("neutralBtn")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)


# ═════════════════════════════════════════════════════════════════════════════
# EMAIL MANAGER
# ═════════════════════════════════════════════════════════════════════════════

class EmailManager:
    """Send HTML alert emails via SMTP in a background thread."""

    def __init__(self, config: ConfigManager):
        self.config = config

    def _cfg(self, key, default=None):
        return self.config.get(key, default)

    def is_configured(self) -> bool:
        return bool(
            self._cfg("email_enabled") and
            self._cfg("email_smtp_host") and
            self._cfg("email_recipients")
        )

    def _recipients(self) -> list[str]:
        raw = self._cfg("email_recipients", "")
        return [r.strip() for r in raw.split(",") if r.strip()]

    def send(self, subject: str, html_body: str,
             recipients: list[str] = None, on_done=None):
        """Fire-and-forget in a daemon thread."""
        if not self.is_configured():
            return
        rcpt = recipients or self._recipients()
        if not rcpt:
            return

        cfg = self.config.config.copy()

        def _worker():
            try:
                msg = MIMEMultipart("alternative")
                msg["Subject"] = subject
                msg["From"]    = cfg.get("email_from") or cfg.get("email_smtp_user", "")
                msg["To"]      = ", ".join(rcpt)
                msg.attach(MIMEText(html_body, "html", "utf-8"))

                host = cfg.get("email_smtp_host", "")
                port = int(cfg.get("email_smtp_port", 587))

                if cfg.get("email_use_tls", True):
                    srv = smtplib.SMTP(host, port, timeout=15)
                    srv.starttls()
                else:
                    srv = smtplib.SMTP_SSL(host, port, timeout=15)

                u = cfg.get("email_smtp_user", "")
                p = cfg.get("email_smtp_pass", "")
                if u and p:
                    srv.login(u, p)

                srv.sendmail(msg["From"], rcpt, msg.as_string())
                srv.quit()
                log.info(f"Email sent: {subject}")
                if on_done:
                    on_done(True, "")
            except Exception as e:
                log.error(f"Email failed: {e}")
                if on_done:
                    on_done(False, str(e))

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def send_overdue_alert(self, overdue_items: list[dict], company: str):
        if not self._cfg("email_on_overdue", True):
            return
        rows_html = "".join(
            f"<tr>"
            f"<td style='padding:7px 14px;border-bottom:1px solid #eee'>{e['name']}</td>"
            f"<td style='padding:7px 14px;border-bottom:1px solid #eee'>{e['equip_id']}</td>"
            f"<td style='padding:7px 14px;border-bottom:1px solid #eee'>"
            f"{e.get('location','—')}</td>"
            f"<td style='padding:7px 14px;border-bottom:1px solid #eee;color:#c0392b;font-weight:600'>"
            f"{(date.today() - datetime.strptime(e['next_service'],'%Y-%m-%d').date()).days}d overdue"
            f"</td></tr>"
            for e in overdue_items
        )
        body = f"""
        <div style="font-family:Arial,sans-serif;max-width:660px;margin:auto">
          <div style="background:#1e1e2e;color:white;padding:22px 28px;border-radius:8px 8px 0 0">
            <h2 style="margin:0;font-size:18px">⚠ Overdue Maintenance Alert</h2>
            <p style="margin:6px 0 0;opacity:.65;font-size:13px">{company}</p>
          </div>
          <div style="background:#fff;padding:22px 28px;
                      border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px">
            <p style="margin:0 0 14px;color:#333">
              <strong>{len(overdue_items)}</strong> item(s) are overdue as of
              <strong>{date.today().strftime('%d %B %Y')}</strong>:
            </p>
            <table style="width:100%;border-collapse:collapse;font-size:13px">
              <thead>
                <tr style="background:#f4f4f8">
                  <th style="padding:8px 14px;text-align:left">Equipment</th>
                  <th style="padding:8px 14px;text-align:left">ID</th>
                  <th style="padding:8px 14px;text-align:left">Location</th>
                  <th style="padding:8px 14px;text-align:left">Status</th>
                </tr>
              </thead>
              <tbody>{rows_html}</tbody>
            </table>
            <p style="margin:20px 0 0;color:#888;font-size:11px">
              Sent by Equipment Maintenance Manager Pro v{APP_VERSION}
            </p>
          </div>
        </div>"""
        self.send(f"[{company}] ⚠ {len(overdue_items)} Overdue Equipment Items", body)

    def send_daily_digest(self, stats: dict, company: str):
        if not self._cfg("email_daily_digest", False):
            return
        body = f"""
        <div style="font-family:Arial,sans-serif;max-width:520px;margin:auto">
          <div style="background:#1e1e2e;color:white;padding:22px 28px;border-radius:8px 8px 0 0">
            <h2 style="margin:0;font-size:18px">📋 Daily Maintenance Digest</h2>
            <p style="margin:6px 0 0;opacity:.65;font-size:13px">
              {company} — {date.today().strftime('%d %B %Y')}
            </p>
          </div>
          <div style="background:#fff;padding:22px 28px;
                      border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px">
            <table style="width:100%;border-collapse:collapse;font-size:15px">
              <tr><td style="padding:10px 0;font-size:30px;font-weight:700;color:#89b4fa;width:60px">{stats['total']}</td>
                  <td style="color:#555">Total Equipment</td></tr>
              <tr><td style="padding:10px 0;font-size:30px;font-weight:700;color:#a6e3a1">{stats['healthy']}</td>
                  <td style="color:#555">Healthy</td></tr>
              <tr><td style="padding:10px 0;font-size:30px;font-weight:700;color:#f9e2af">{stats['due_soon']}</td>
                  <td style="color:#555">Due Soon</td></tr>
              <tr><td style="padding:10px 0;font-size:30px;font-weight:700;color:#f38ba8">{stats['overdue']}</td>
                  <td style="color:#555">Overdue</td></tr>
            </table>
            <p style="margin:20px 0 0;color:#888;font-size:11px">
              Sent by Equipment Maintenance Manager Pro v{APP_VERSION}
            </p>
          </div>
        </div>"""
        self.send(f"[{company}] Daily Maintenance Digest — {date.today().strftime('%d %b %Y')}", body)


# ═════════════════════════════════════════════════════════════════════════════
# EMAIL SETTINGS DIALOG
# ═════════════════════════════════════════════════════════════════════════════

class EmailSettingsDialog(QDialog):
    def __init__(self, parent, config: ConfigManager, email_mgr: EmailManager):
        super().__init__(parent)
        self.config    = config
        self.email_mgr = email_mgr
        self.setWindowTitle("Email Alert Settings")
        self.setMinimumWidth(480)
        self.setModal(True)
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(14)

        title = QLabel("📧  Email Alert Settings")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        sep = QFrame(); sep.setObjectName("separator"); sep.setFixedHeight(1)
        layout.addWidget(sep)

        tabs = QTabWidget()

        # ── SMTP tab ──────────────────────────────────────────────────────────
        smtp_w = QWidget()
        f = QFormLayout(smtp_w)
        f.setSpacing(10); f.setContentsMargins(14, 14, 14, 14)
        f.setLabelAlignment(Qt.AlignRight)

        self.en_cb    = QCheckBox("Enable email alerts")
        self.en_cb.setChecked(bool(self.config.get("email_enabled", False)))

        self.host_e   = QLineEdit(self.config.get("email_smtp_host", ""))
        self.host_e.setPlaceholderText("e.g. smtp.gmail.com")

        self.port_s   = QSpinBox()
        self.port_s.setRange(1, 65535)
        self.port_s.setValue(self.config.get("email_smtp_port", 587))

        self.user_e   = QLineEdit(self.config.get("email_smtp_user", ""))
        self.user_e.setPlaceholderText("your@email.com")

        self.pass_e   = QLineEdit(self.config.get("email_smtp_pass", ""))
        self.pass_e.setEchoMode(QLineEdit.Password)

        self.from_e   = QLineEdit(self.config.get("email_from", ""))
        self.from_e.setPlaceholderText("noreply@company.com (leave blank to use username)")

        self.tls_cb   = QCheckBox("Use STARTTLS")
        self.tls_cb.setChecked(bool(self.config.get("email_use_tls", True)))

        test_btn = QPushButton("📤  Send Test Email")
        test_btn.setObjectName("neutralBtn")
        test_btn.clicked.connect(self._send_test)

        f.addRow("",              self.en_cb)
        f.addRow("SMTP Host",     self.host_e)
        f.addRow("SMTP Port",     self.port_s)
        f.addRow("Username",      self.user_e)
        f.addRow("Password",      self.pass_e)
        f.addRow("From Address",  self.from_e)
        f.addRow("",              self.tls_cb)
        f.addRow("",              test_btn)
        tabs.addTab(smtp_w, "🔌  SMTP")

        # ── Recipients tab ────────────────────────────────────────────────────
        rcpt_w = QWidget()
        rl = QVBoxLayout(rcpt_w)
        rl.setContentsMargins(14, 14, 14, 14); rl.setSpacing(8)
        rl.addWidget(QLabel("Recipients (comma-separated):"))
        self.rcpt_e = QTextEdit()
        self.rcpt_e.setPlainText(self.config.get("email_recipients", ""))
        self.rcpt_e.setPlaceholderText("manager@company.com, engineer@company.com")
        self.rcpt_e.setFixedHeight(100)
        rl.addWidget(self.rcpt_e)
        rl.addStretch()
        tabs.addTab(rcpt_w, "👥  Recipients")

        # ── Alert rules tab ───────────────────────────────────────────────────
        rules_w = QWidget()
        rulesl = QVBoxLayout(rules_w)
        rulesl.setContentsMargins(14, 14, 14, 14); rulesl.setSpacing(12)

        self.ov_cb = QCheckBox("Send alert when equipment is overdue at startup")
        self.ov_cb.setChecked(bool(self.config.get("email_on_overdue", True)))

        self.dg_cb = QCheckBox("Send daily digest email")
        self.dg_cb.setChecked(bool(self.config.get("email_daily_digest", False)))

        dh_row = QHBoxLayout()
        dh_row.addWidget(QLabel("  Daily digest send hour:"))
        self.dh_s = QSpinBox()
        self.dh_s.setRange(0, 23)
        self.dh_s.setValue(self.config.get("email_digest_hour", 8))
        self.dh_s.setSuffix(":00")
        dh_row.addWidget(self.dh_s); dh_row.addStretch()

        rulesl.addWidget(self.ov_cb)
        rulesl.addWidget(self.dg_cb)
        rulesl.addLayout(dh_row)
        rulesl.addStretch()
        tabs.addTab(rules_w, "🔔  Alert Rules")

        layout.addWidget(tabs, 1)

        btns = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _save(self):
        self.config.set("email_enabled",      self.en_cb.isChecked())
        self.config.set("email_smtp_host",    self.host_e.text().strip())
        self.config.set("email_smtp_port",    self.port_s.value())
        self.config.set("email_smtp_user",    self.user_e.text().strip())
        self.config.set("email_smtp_pass",    self.pass_e.text())
        self.config.set("email_from",         self.from_e.text().strip())
        self.config.set("email_use_tls",      self.tls_cb.isChecked())
        self.config.set("email_recipients",   self.rcpt_e.toPlainText().strip())
        self.config.set("email_on_overdue",   self.ov_cb.isChecked())
        self.config.set("email_daily_digest", self.dg_cb.isChecked())
        self.config.set("email_digest_hour",  self.dh_s.value())
        self.accept()

    def _send_test(self):
        # Temporarily apply current form values so the test uses them
        self._save()
        company = self.config.get("company_name", "Company")
        body = (
            f"<div style='font-family:Arial,sans-serif;padding:20px'>"
            f"<h3>✅ Test email from {company}</h3>"
            f"<p>Equipment Maintenance Manager Pro v{APP_VERSION} is configured correctly.</p>"
            f"</div>"
        )
        def _done(ok, err):
            if ok:
                QMessageBox.information(self, "Test Email", "✅ Test email sent successfully!")
            else:
                QMessageBox.warning(self, "Test Email Failed", f"❌ Could not send:\n\n{err}")

        self.email_mgr.send(f"[Test] Maintenance Manager — {company}", body, on_done=_done)


# ═════════════════════════════════════════════════════════════════════════════
# GRADIENT LABEL  (paints text with a left→right colour gradient)
# ═════════════════════════════════════════════════════════════════════════════

class _GradientLabel(QLabel):
    """QLabel whose text is painted with a horizontal linear gradient."""
    def __init__(self, text: str, color_start: str, color_end: str, parent=None):
        super().__init__(text, parent)
        self._c1 = QColor(color_start)
        self._c2 = QColor(color_end)
        font = self.font()
        font.setPointSize(16)
        font.setWeight(QFont.Bold)
        self.setFont(font)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

    def setText(self, text: str):
        super().setText(text)
        self.update()

    def paintEvent(self, event):
        from PySide6.QtGui import QPainter, QLinearGradient, QBrush
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        grad = QLinearGradient(0, 0, self.width(), 0)
        grad.setColorAt(0.0, self._c1)
        grad.setColorAt(1.0, self._c2)
        painter.setFont(self.font())
        painter.setPen(Qt.NoPen)
        # draw text via path so we can fill with gradient
        from PySide6.QtGui import QPainterPath
        path = QPainterPath()
        fm = self.fontMetrics()
        # vertically centre
        y = (self.height() + fm.ascent() - fm.descent()) / 2
        path.addText(0, y, self.font(), self.text())
        painter.fillPath(path, QBrush(grad))
        painter.end()

    def sizeHint(self):
        fm = self.fontMetrics()
        return fm.boundingRect(self.text()).size() + \
               __import__('PySide6.QtCore', fromlist=['QSize']).QSize(4, 4)


# ═════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ═════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager()
        self.db = DatabaseManager(self.config.get("db_path", "equipment.db"))
        self.email_mgr = EmailManager(self.config)
        self._project_file: str | None = None
        self._unsaved = False

        self.setWindowTitle(self._win_title())
        self.resize(
            self.config.get("window_width", 1280),
            self.config.get("window_height", 800)
        )

        self._build_menu()
        self._build_toolbar()
        self._build_central()
        self._build_statusbar()
        self._refresh()

        # Startup overdue alert
        QTimer.singleShot(600, self._startup_alert)

        # Daily digest — check every minute
        self._digest_timer = QTimer(self)
        self._digest_timer.timeout.connect(self._check_digest)
        self._digest_timer.start(60_000)

    # ── Window title ──────────────────────────────────────────────────────────

    def _win_title(self) -> str:
        company = self.config.get("company_name","Engineering Department")
        mod     = " *" if self._unsaved else ""
        project = f"  ·  {os.path.basename(self._project_file)}" if self._project_file else ""
        return f"{company}{project}{mod}  ·  Maintenance Manager"

    def _mark_unsaved(self):
        self._unsaved = True
        self.setWindowTitle(self._win_title())

    def _mark_saved(self):
        self._unsaved = False
        self.setWindowTitle(self._win_title())

    # ── Menu ──────────────────────────────────────────────────────────────────

    def _build_menu(self):
        mb = self.menuBar()

        # File
        fm = mb.addMenu("&File")
        self._add_action(fm, "🗋  New Project",      self._new_project,   "Ctrl+N")
        self._add_action(fm, "📂  Open Project…",    self._open_project,  "Ctrl+O")
        self._add_action(fm, "💾  Save Project",      self._save_project,  "Ctrl+S")
        self._add_action(fm, "💾  Save Project As…", self._save_project_as,"Ctrl+Shift+S")
        fm.addSeparator()
        self._add_action(fm, "📤  Export CSV…",       self._export_csv,    "Ctrl+E")
        self._add_action(fm, "📤  Export PDF Report…",self._export_pdf,    "Ctrl+Shift+E")
        self._add_action(fm, "📥  Import CSV…",       self._import_csv,    "Ctrl+I")
        fm.addSeparator()
        self._add_action(fm, "✕  Exit",               self.close,          "Alt+F4")

        # Equipment
        em = mb.addMenu("&Equipment")
        self._add_action(em, "➕  Add Equipment",     self._add_equipment,  "Ctrl+Return")
        self._add_action(em, "✏  Edit Equipment",    self._edit_equipment, "F2")
        self._add_action(em, "🗑  Remove Equipment", self._remove_equipment,"Delete")
        em.addSeparator()
        self._add_action(em, "✔  Mark Serviced",     self._mark_serviced,  "Ctrl+M")
        self._add_action(em, "📋  View History",     self._view_history,   "Ctrl+H")
        em.addSeparator()
        self._add_action(em, "⚠  Select All Overdue",self._select_overdue, "")

        # View
        vm = mb.addMenu("&View")
        self._add_action(vm, "📊  Statistics",        self._show_statistics,"Ctrl+Shift+T")
        vm.addSeparator()
        self.filter_overdue_action = QAction("⚠  Show Overdue Only", self)
        self.filter_overdue_action.setCheckable(True)
        self.filter_overdue_action.triggered.connect(
            lambda c: self.filter_combo.setCurrentIndex(3 if c else 0)
        )
        vm.addAction(self.filter_overdue_action)

        # Settings
        sm = mb.addMenu("&Settings")
        self._add_action(sm, "⚙  Preferences…",       self._open_settings,   "Ctrl+,")
        self._add_action(sm, "📧  Email Alerts…",      self._open_email_settings, "Ctrl+Shift+E")

        # Help
        hm = mb.addMenu("&Help")
        self._add_action(hm, "ℹ  About",             self._show_about,     "")

    def _add_action(self, menu, label, slot, shortcut=""):
        a = QAction(label, self)
        if shortcut:
            a.setShortcut(QKeySequence(shortcut))
        a.triggered.connect(slot)
        menu.addAction(a)
        return a

    # ── Toolbar ───────────────────────────────────────────────────────────────

    def _build_toolbar(self):
        tb = QToolBar("Main")
        tb.setMovable(False)
        tb.setIconSize(QSize(16, 16))
        self.addToolBar(tb)

        for label, slot, tip in [
            ("➕  Add",           self._add_equipment,   "Add new equipment (Ctrl+Enter)"),
            ("✏  Edit",          self._edit_equipment,  "Edit selected (F2)"),
            ("🗑  Remove",       self._remove_equipment,"Remove selected (Delete)"),
        ]:
            a = QAction(label, self)
            a.setToolTip(tip)
            a.triggered.connect(slot)
            tb.addAction(a)

        tb.addSeparator()

        for label, slot, tip in [
            ("✔  Mark Serviced", self._mark_serviced,   "Mark selected as serviced (Ctrl+M)"),
            ("📋  History",      self._view_history,    "View service history (Ctrl+H)"),
        ]:
            a = QAction(label, self)
            a.setToolTip(tip)
            a.triggered.connect(slot)
            tb.addAction(a)

        tb.addSeparator()

        for label, slot in [
            ("📊  Stats",  self._show_statistics),
            ("📤  Export", self._export_pdf),
            ("📧  Email",  self._open_email_settings),
            ("⚙  Settings",self._open_settings),
        ]:
            a = QAction(label, self)
            a.triggered.connect(slot)
            tb.addAction(a)

    # ── Central ───────────────────────────────────────────────────────────────

    def _build_central(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(16, 12, 16, 10)
        layout.setSpacing(10)

        # Header
        hdr = QHBoxLayout()
        hdr.setSpacing(0)

        # App name: coloured "Maintenance Manager" + muted "Pro"
        name_widget = QWidget()
        name_widget.setStyleSheet("background:transparent;")
        name_row = QHBoxLayout(name_widget)
        name_row.setContentsMargins(0,0,0,0)
        name_row.setSpacing(6)

        self.page_title = _GradientLabel(
            self.config.get("company_name","Engineering Department"),
            C["blue"], C["purple"]
        )

        name_row.addWidget(self.page_title)
        name_row.addStretch()
        hdr.addWidget(name_widget, 1)

        date_lbl = QLabel(f"📅  {date.today().strftime('%A, %d %B %Y')}")
        date_lbl.setObjectName("pageSubtitle")
        hdr.addWidget(date_lbl)
        layout.addLayout(hdr)

        # Dashboard
        self.dashboard = DashboardWidget()
        layout.addWidget(self.dashboard)

        # Filter bar
        fb = QHBoxLayout()
        fb.setSpacing(10)
        fb.addWidget(QLabel("🔍"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search name, ID, location, category, notes…")
        self.search_edit.textChanged.connect(self._apply_filter)
        self.search_edit.setMinimumWidth(300)
        fb.addWidget(self.search_edit, 1)

        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["All Equipment","Healthy","Due Soon (7 days)","Overdue"])
        self.filter_combo.currentIndexChanged.connect(self._apply_filter)
        self.filter_combo.setMinimumWidth(185)
        fb.addWidget(self.filter_combo)

        fb.addStretch()
        self.count_lbl = QLabel("0 items")
        self.count_lbl.setObjectName("pageSubtitle")
        fb.addWidget(self.count_lbl)
        layout.addLayout(fb)

        # Table
        self.table = QTableWidget()
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._context_menu)
        self.table.doubleClicked.connect(self._edit_equipment)
        layout.addWidget(self.table, 1)

        self.table_ctrl = TableController(self.table)

        # Bottom buttons
        bb = QHBoxLayout()
        bb.setSpacing(10)
        for label, slot, obj_name in [
            ("➕  Add Equipment",  self._add_equipment,    ""),
            ("✏  Edit",           self._edit_equipment,   "neutralBtn"),
            ("🗑  Remove",        self._remove_equipment, "dangerBtn"),
        ]:
            btn = QPushButton(label)
            if obj_name:
                btn.setObjectName(obj_name)
            btn.clicked.connect(slot)
            bb.addWidget(btn)

        bb.addStretch()

        for label, slot, obj_name in [
            ("📋  History",       self._view_history,     "purpleBtn"),
            ("✔  Mark Serviced",  self._mark_serviced,    "successBtn"),
            ("📊  Statistics",    self._show_statistics,  "warningBtn"),
        ]:
            btn = QPushButton(label)
            btn.setObjectName(obj_name)
            btn.clicked.connect(slot)
            bb.addWidget(btn)

        layout.addLayout(bb)

    # ── Status bar ────────────────────────────────────────────────────────────

    def _build_statusbar(self):
        sb = QStatusBar()
        self.setStatusBar(sb)
        self._status_lbl = QLabel("Ready")
        sb.addWidget(self._status_lbl, 1)
        self._db_lbl = QLabel(f"DB: {self.config.get('db_path','equipment.db')}")
        sb.addPermanentWidget(self._db_lbl)
        ver_lbl = QLabel(f"  v{APP_VERSION}  ")
        sb.addPermanentWidget(ver_lbl)

    def _set_status(self, msg: str, ms: int = 4500):
        self._status_lbl.setText(msg)
        QTimer.singleShot(ms, lambda: self._status_lbl.setText("Ready"))

    # ── Context menu ──────────────────────────────────────────────────────────

    def _context_menu(self, pos):
        eq = self.table_ctrl.first_selected()
        if not eq:
            return
        menu = QMenu(self)
        menu.addAction("✏  Edit",          self._edit_equipment)
        menu.addAction("✔  Mark Serviced", self._mark_serviced)
        menu.addAction("📋  View History", self._view_history)
        menu.addSeparator()
        menu.addAction("🗑  Remove",       self._remove_equipment)
        menu.exec(QCursor.pos())

    # ── Refresh ───────────────────────────────────────────────────────────────

    def _refresh(self):
        rows = self.db.get_all()
        self.table_ctrl.load(rows)
        self._apply_filter()
        stats = self.db.get_statistics()
        self.dashboard.update(stats)

    def _apply_filter(self):
        text = self.search_edit.text()
        idx  = self.filter_combo.currentIndex()
        count = self.table_ctrl.apply_filter(text, idx)
        self.count_lbl.setText(f"{count} item{'s' if count != 1 else ''}")

    # ── Equipment actions ─────────────────────────────────────────────────────

    def _add_equipment(self):
        dlg = EquipmentDialog(self, self.config, existing_ids=self.db.get_existing_ids())
        if dlg.exec() == QDialog.Accepted:
            self.db.add_equipment(dlg.get_data())
            self._refresh()
            self._mark_unsaved()
            self._set_status(f"Added: {dlg.get_data()['name']}")

    def _edit_equipment(self):
        eq = self.table_ctrl.first_selected()
        if not eq:
            QMessageBox.information(self, "No Selection", "Select an item to edit.")
            return
        dlg = EquipmentDialog(self, self.config, equipment=eq,
                              existing_ids=self.db.get_existing_ids())
        if dlg.exec() == QDialog.Accepted:
            self.db.update_equipment(eq["id"], dlg.get_data())
            self._refresh()
            self._mark_unsaved()
            self._set_status(f"Updated: {dlg.get_data()['name']}")

    def _remove_equipment(self):
        selected = self.table_ctrl.selected_equipment()
        if not selected:
            QMessageBox.information(self, "No Selection", "Select item(s) to remove.")
            return
        names = "\n".join(f"  • {e['name']} ({e['equip_id']})" for e in selected[:10])
        more  = f"\n  … and {len(selected)-10} more" if len(selected) > 10 else ""
        reply = QMessageBox.question(
            self, "Confirm Remove",
            f"Remove {len(selected)} item(s)?\n\n{names}{more}\n\nThis cannot be undone.",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            for eq in selected:
                self.db.remove_equipment(eq["id"])
            self._refresh()
            self._mark_unsaved()
            self._set_status(f"Removed {len(selected)} item(s).")

    def _mark_serviced(self):
        selected = self.table_ctrl.selected_equipment()
        if not selected:
            QMessageBox.information(self, "No Selection", "Select item(s) to mark as serviced.")
            return

        if len(selected) == 1:
            dlg = ServiceDialog(self, selected[0])
            if dlg.exec() != QDialog.Accepted:
                return
            d = dlg.get_data()
            self.db.mark_serviced(selected[0]["id"], selected[0]["interval"],
                                  d["technician"], d["notes"])
            self._set_status(f"Marked serviced: {selected[0]['name']}")
        else:
            reply = QMessageBox.question(
                self, "Bulk Service",
                f"Mark {len(selected)} items as serviced today?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                for eq in selected:
                    self.db.mark_serviced(eq["id"], eq["interval"])
                self._set_status(f"Marked {len(selected)} items as serviced.")

        self._refresh()
        self._mark_unsaved()

    def _view_history(self):
        eq = self.table_ctrl.first_selected()
        if not eq:
            QMessageBox.information(self, "No Selection", "Select an item to view history.")
            return
        HistoryDialog(self, self.db, eq).exec()

    def _select_overdue(self):
        self.filter_combo.setCurrentIndex(3)
        self.table.selectAll()

    # ── File actions ──────────────────────────────────────────────────────────

    def _check_unsaved(self) -> bool:
        """Returns True if it's safe to proceed (saved or user discarded)."""
        if not self._unsaved:
            return True
        reply = QMessageBox.question(
            self, "Unsaved Changes",
            "You have unsaved changes. Save before continuing?",
            QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
        )
        if reply == QMessageBox.Save:
            self._save_project()
            return not self._unsaved  # false if save was cancelled
        return reply == QMessageBox.Discard

    def _new_project(self):
        if not self._check_unsaved():
            return
        self.db.conn.execute("DELETE FROM service_history")
        self.db.conn.execute("DELETE FROM equipment")
        self.db.conn.commit()
        self._project_file = None
        self._mark_saved()
        self._refresh()
        self._set_status("New project created.")

    def _open_project(self):
        if not self._check_unsaved():
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Project", "",
            "Maintenance Projects (*.mmp);;All Files (*)"
        )
        if not path:
            return
        try:
            count = self.db.load_project(path)
            self._project_file = path
            self._mark_saved()
            self._refresh()
            self._set_status(f"Loaded {count} records from {os.path.basename(path)}")
        except Exception as e:
            log.error(f"Open project failed: {e}")
            QMessageBox.critical(self, "Error", f"Failed to open project:\n{e}")

    def _save_project(self):
        if self._project_file:
            self._do_save(self._project_file)
        else:
            self._save_project_as()

    def _save_project_as(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Project As", "",
            "Maintenance Projects (*.mmp);;All Files (*)"
        )
        if path:
            if not path.endswith(".mmp"):
                path += ".mmp"
            self._do_save(path)

    def _do_save(self, path):
        try:
            self.db.save_project(path)
            self._project_file = path
            self._mark_saved()
            self._set_status(f"Saved: {os.path.basename(path)}")
        except Exception as e:
            log.error(f"Save failed: {e}")
            QMessageBox.critical(self, "Save Error", f"Failed to save:\n{e}")

    def _export_csv(self):
        rows = self.table_ctrl.selected_equipment() or None
        desc = f"filtered ({len(rows)} items)" if rows else "all"
        path, _ = QFileDialog.getSaveFileName(
            self, f"Export CSV ({desc})", "maintenance_export.csv",
            "CSV Files (*.csv)"
        )
        if not path:
            return
        count = self.db.export_csv(path, rows)
        if count:
            self._set_status(f"Exported {count} records to {os.path.basename(path)}")
        else:
            QMessageBox.information(self, "Empty", "No records to export.")

    def _export_pdf(self):
        rows = self.db.get_all()
        if not rows:
            QMessageBox.information(self, "Empty", "No equipment to export.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export PDF Report", "maintenance_report.pdf",
            "PDF Files (*.pdf)"
        )
        if not path:
            return
        ok, err = ReportExporter.export_pdf(
            path, rows, self.config.get("company_name","Company")
        )
        if ok:
            self._set_status(f"PDF exported: {os.path.basename(path)}")
            reply = QMessageBox.question(
                self, "PDF Exported",
                f"Report saved to:\n{path}\n\nOpen now?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                import subprocess, platform
                if platform.system() == "Windows":
                    os.startfile(path)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
        else:
            QMessageBox.warning(
                self, "PDF Export Failed",
                f"Could not generate PDF.\n\n{err}\n\n"
                "Tip: Install reportlab with:  pip install reportlab"
            )

    def _import_csv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Import CSV", "", "CSV Files (*.csv)"
        )
        if not path:
            return
        try:
            count, errors = self.db.import_csv(path)
            self._refresh()
            self._mark_unsaved()
            msg = f"Imported {count} records."
            if errors:
                msg += f"\n\n{len(errors)} rows had errors:\n" + "\n".join(errors[:5])
                if len(errors) > 5:
                    msg += f"\n… and {len(errors)-5} more (see log)"
            QMessageBox.information(self, "Import Complete", msg)
        except Exception as e:
            log.error(f"Import failed: {e}")
            QMessageBox.critical(self, "Import Error", f"Failed to import:\n{e}")

    # ── Settings & info ───────────────────────────────────────────────────────

    def _open_settings(self):
        dlg = SettingsDialog(self, self.config)
        if dlg.exec() == QDialog.Accepted:
            self.page_title.setText(self.config.get("company_name","Engineering Department"))
            self.setWindowTitle(self._win_title())
            self._set_status("Settings saved.")

    def _show_statistics(self):
        StatisticsDialog(self, self.db).exec()

    def _startup_alert(self):
        stats = self.db.get_statistics()
        if stats["overdue"] > 0:
            self._set_status(
                f"⚠  {stats['overdue']} item(s) overdue — click Statistics for details.",
                ms=8000
            )
            # Send overdue alert email if configured
            if self.email_mgr.is_configured():
                all_eq = self.db.get_all()
                today  = date.today().isoformat()
                overdue_eq = [e for e in all_eq if e.get("next_service","") < today]
                company    = self.config.get("company_name","Company")
                self.email_mgr.send_overdue_alert(overdue_eq, company)

    def _check_digest(self):
        """Called every minute — fires digest email at the configured hour:00."""
        if not self.email_mgr.is_configured():
            return
        now = datetime.now()
        if now.hour == self.config.get("email_digest_hour", 8) and now.minute == 0:
            stats   = self.db.get_statistics()
            company = self.config.get("company_name","Company")
            self.email_mgr.send_daily_digest(stats, company)

    def _open_email_settings(self):
        dlg = EmailSettingsDialog(self, self.config, self.email_mgr)
        if dlg.exec() == QDialog.Accepted:
            self._set_status("Email alert settings saved.")

    def _show_about(self):
        QMessageBox.about(
            self, "About",
            f"<h3>Equipment Maintenance Manager Pro</h3>"
            f"<p><b>Version {APP_VERSION}</b></p>"
            f"<p>Professional equipment maintenance tracking for teams.</p>"
            f"<hr>"
            f"<p><b>Features:</b> maintenance scheduling, service history,<br>"
            f"bulk operations, PDF reporting, CSV import/export,<br>"
            f"project files, category breakdown, overdue alerting.</p>"
            f"<p>Built with Python &amp; PySide6.</p>"
        )

    # ── Close ─────────────────────────────────────────────────────────────────

    def closeEvent(self, event):
        if not self._check_unsaved():
            event.ignore()
            return
        self.config.set("window_width", self.width())
        self.config.set("window_height", self.height())
        self.db.close()
        log.info("Application closed.")
        event.accept()


# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Maintenance Manager")
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName("Clan")
    app.setStyleSheet(DARK_STYLESHEET)

    # ── App icon — load from ico file (works both dev and PyInstaller exe) ────
    from PySide6.QtGui import QIcon
    import sys as _sys
    if getattr(_sys, "frozen", False):
        # Running as PyInstaller exe — icon is next to the exe
        base = os.path.dirname(_sys.executable)
    else:
        # Running as plain .py script
        base = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base, "icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    # ─────────────────────────────────────────────────────────────────────────
    # ─────────────────────────────────────────────────────────────────────────

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
