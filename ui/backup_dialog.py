"""Backup Manager Dialog for managing backup files"""

import os
import json

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QMessageBox,
    QAbstractItemView,
    QMenu,
    QTextEdit,
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction

from config import Config
from utils.helpers import parse_backup_filename, parse_resolution_string
from ui.preview_widget import IconPreviewWidget
from icon_manager import BackupComparator


class BackupManagerWindow(QDialog):
    restore_requested = pyqtSignal(str)
    list_changed_signal = pyqtSignal()

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle(self.tr("Select, Restore, or Delete Backup"))
        self.setFixedSize(1150, 650)
        self.layout = QVBoxLayout(self)
        self.layout.setSpacing(10)
        self.layout.addWidget(
            QLabel(self.tr("Select a backup to restore or right-click to delete."))
        )
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(
            self.tr("Search by tag, resolution, or date...")
        )
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self.filter_backups)
        self.layout.addWidget(self.search_input)
        h_split = QHBoxLayout()
        left_panel = QVBoxLayout()
        header_text = (
            f"{self.tr('TAG/DESCRIPTION'):<55} "
            f"| {self.tr('RESOLUTION'):<14} "
            f"| {self.tr('ICONS'):<6} "
            f"| {self.tr('TIMESTAMP')}"
        )
        header_label = QLabel(header_text)
        header_label.setStyleSheet(
            "font-family: 'Consolas', monospace; font-size: 11px; font-weight: bold; margin-bottom: 2px;"
        )
        left_panel.addWidget(header_label)
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.list_widget.setStyleSheet(
            "font-family: 'Consolas', monospace; font-size: 11px;"
        )
        left_panel.addWidget(self.list_widget)
        h_split.addLayout(left_panel, 6)
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel(self.tr("Layout Preview:")))
        self.preview_widget = IconPreviewWidget()
        right_panel.addWidget(self.preview_widget)
        self.info_label = QLabel(self.tr("Select a backup to see details."))
        self.info_label.setWordWrap(True)
        self.info_label.setStyleSheet(
            "color: #ccc; background-color: #2b2b2b; border-radius: 4px; padding: 10px; font-family: 'Segoe UI'; font-size: 12px;"
        )
        right_panel.addWidget(self.info_label)
        right_panel.addStretch()
        h_split.addLayout(right_panel, 2)
        self.layout.addLayout(h_split)
        button_layout = QHBoxLayout()
        self.btn_restore = QPushButton(self.tr("Restore Selected Layout"))
        self.btn_restore.clicked.connect(self.restore_selected)
        self.btn_restore.setEnabled(False)
        self.btn_close = QPushButton(self.tr("Close"))
        self.btn_close.clicked.connect(self.reject)
        button_layout.addWidget(self.btn_restore)
        button_layout.addStretch(1)
        button_layout.addWidget(self.btn_close)
        self.layout.addLayout(button_layout)
        self.list_widget.itemSelectionChanged.connect(self.on_selection_changed)
        self.list_widget.itemDoubleClicked.connect(self.restore_selected)
        self.list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_widget.customContextMenuRequested.connect(self.show_context_menu)
        self.load_backups()

    def load_backups(self):
        self.list_widget.clear()
        backups = self.manager.get_all_backup_filenames()
        if not backups:
            self.list_widget.addItem(QListWidgetItem(self.tr("No backups found.")))
            return
        for filename in backups:
            readable_date, resolution, _ = parse_backup_filename(filename)
            description = ""
            icon_count = "N/A"
            filepath = os.path.join(Config.BACKUP_DIR, filename)
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    description = data.get("description", "").strip()
                    icon_count = data.get("icon_count", "N/A")
            except Exception:
                pass
            description_display = f"{f'[{description[:52]}]':<56}"
            resolution_display = f"| {resolution:<15}"
            icon_count_display = f"| {icon_count:>6}"
            item_text = f"{description_display}{resolution_display}{icon_count_display} | {readable_date}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, filename)
            self.list_widget.addItem(item)

    def on_selection_changed(self):
        items = self.list_widget.selectedItems()
        if not items or items[0].data(Qt.ItemDataRole.UserRole) is None:
            self.preview_widget.update_preview({}, (1920, 1080))
            self.info_label.setText(self.tr("Select a backup to see details."))
            self.btn_restore.setEnabled(False)
            return
        filename = items[0].data(Qt.ItemDataRole.UserRole)
        filepath = os.path.join(Config.BACKUP_DIR, filename)
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                icons = data.get("icons", {})
                res_str = parse_backup_filename(filename)[1]
                res_tuple = parse_resolution_string(res_str)
                self.preview_widget.update_preview(icons, res_tuple)
                desc = data.get("description", self.tr("None"))
                ts = data.get("timestamp", "N/A")
                count = len(icons)
                info = (
                    f"<b>{self.tr('File')}:</b> {filename}<br>"
                    f"<b>{self.tr('Icons')}:</b> {count}<br>"
                    f"<b>{self.tr('Resolution')}:</b> {res_str}<br>"
                    f"<b>{self.tr('Description')}:</b> {desc}<br>"
                    f"<b>{self.tr('Timestamp')}:</b> {ts}"
                )
                self.info_label.setText(info)
                self.btn_restore.setEnabled(True)
        except Exception as e:
            self.info_label.setText(f"{self.tr('Error')}: {str(e)}")
            self.btn_restore.setEnabled(False)

    def filter_backups(self, query):
        query = query.lower()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setHidden(query not in item.text().lower())

    def get_selected_filename(self):
        selected = self.list_widget.selectedItems()
        if not selected:
            return None
        return selected[0].data(Qt.ItemDataRole.UserRole)

    def show_context_menu(self, pos):
        item = self.list_widget.itemAt(pos)
        if not item or not item.data(Qt.ItemDataRole.UserRole):
            return

        menu = QMenu(self)

        restore_action = QAction(self.tr("🔄 Restore Selected"), self)
        restore_action.triggered.connect(self.restore_selected)

        delete_action = QAction(self.tr("🗑️ Delete Selected"), self)
        delete_action.triggered.connect(self.delete_selected)

        compare_action = QAction(self.tr("📊 Compare with Latest"), self)
        compare_action.triggered.connect(self.compare_with_latest)

        menu.addAction(restore_action)
        menu.addAction(compare_action)
        menu.addSeparator()
        menu.addAction(delete_action)

        menu.exec(self.list_widget.mapToGlobal(pos))

    def restore_selected(self):
        fn = self.get_selected_filename()
        if not fn:
            return

        filepath = os.path.join(Config.BACKUP_DIR, fn)
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)

            readable_date, resolution, _ = parse_backup_filename(fn)
            description = data.get("description", "N/A")
            icon_count = data.get("icon_count", "N/A")

            reply = QMessageBox.question(
                self,
                self.tr("Confirm Restore"),
                self.tr(
                    "Restore icon positions from the selected backup file:\n\n"
                    "File: %1\n"
                    "Resolution: %2\n"
                    "Icons: %3\n"
                    "Tag: %4\n"
                    "Timestamp: %5\n\n"
                    "Are you sure you want to proceed?"
                )
                .replace("%1", fn)
                .replace("%2", resolution)
                .replace("%3", str(icon_count))
                .replace("%4", description)
                .replace("%5", readable_date),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.restore_requested.emit(fn)
                self.accept()

        except Exception as e:
            QMessageBox.critical(
                self,
                self.tr("Error"),
                self.tr("Failed to load backup file:\n%1").replace("%1", str(e)),
            )

    def delete_selected(self):
        fn = self.get_selected_filename()
        if not fn:
            return

        reply = QMessageBox.question(
            self,
            self.tr("Confirm Delete"),
            self.tr("Are you sure you want to delete this backup?\n\n%1").replace(
                "%1", fn
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.manager.delete_backup(fn):
                self.load_backups()
                self.list_changed_signal.emit()
                QMessageBox.information(
                    self, self.tr("Success"), self.tr("Backup deleted successfully.")
                )
            else:
                QMessageBox.critical(
                    self, self.tr("Error"), self.tr("Failed to delete backup file.")
                )

    def compare_with_latest(self):
        selected_filename = self.get_selected_filename()
        if not selected_filename:
            QMessageBox.warning(
                self,
                self.tr("No Selection"),
                self.tr("Please select a backup to compare."),
            )
            return

        latest_filename = self.manager.get_latest_backup_filename()
        if not latest_filename:
            QMessageBox.warning(
                self, self.tr("Error"), self.tr("No latest backup found")
            )
            return

        if selected_filename == latest_filename:
            QMessageBox.information(
                self,
                self.tr("Same Backup"),
                self.tr("You selected the latest backup. Nothing to compare."),
            )
            return

        file1 = os.path.join(Config.BACKUP_DIR, selected_filename)
        file2 = os.path.join(Config.BACKUP_DIR, latest_filename)

        report = BackupComparator.compare(file1, file2)

        if report:
            dialog = QDialog(self)
            dialog.setWindowTitle(self.tr("Comparison Results"))
            dialog.resize(650, 550)

            dialog.setStyleSheet("""
                QDialog {
                    background-color: #1e1e1e;
                }
                QLabel {
                    color: #ffffff;
                    background-color: #2d2d30;
                    border: 1px solid #3f3f46;
                    border-radius: 4px;
                    padding: 12px;
                    font-family: 'Segoe UI';
                    font-size: 11px;
                }
                QTextEdit {
                    color: #d4d4d4;
                    background-color: #1e1e1e;
                    border: 1px solid #3f3f46;
                    border-radius: 4px;
                    font-family: 'Consolas', monospace;
                    font-size: 11px;
                    selection-background-color: #264f78;
                    selection-color: #ffffff;
                }
                QPushButton {
                    color: #ffffff;
                    background-color: #0e639c;
                    border: 1px solid #1177bb;
                    border-radius: 4px;
                    padding: 8px 16px;
                    font-family: 'Segoe UI';
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1177bb;
                    border: 1px solid #1c8dd9;
                }
                QPushButton:pressed {
                    background-color: #0d5a8f;
                }
            """)

            layout = QVBoxLayout(dialog)
            layout.setSpacing(10)
            layout.setContentsMargins(15, 15, 15, 15)

            header = QLabel(
                f"<b style='color: #4ec9b0;'>{self.tr('Comparing Backups:')}</b><br>"
                f"<span style='color: #9cdcfe;'>📄 {selected_filename}</span><br>"
                f"<span style='color: #4fc1ff;'>📄 {latest_filename} ({self.tr('latest')})</span>"
            )
            layout.addWidget(header)

            text_area = QTextEdit()
            text_area.setReadOnly(True)

            html_report = self._colorize_comparison_report(report)
            text_area.setHtml(html_report)

            layout.addWidget(text_area)

            btn_close = QPushButton(self.tr("✓ Close"))
            btn_close.clicked.connect(dialog.accept)
            btn_close.setMinimumHeight(35)
            layout.addWidget(btn_close)

            dialog.exec()
        else:
            QMessageBox.critical(
                self, self.tr("Error"), self.tr("Failed to compare backups")
            )

    def _colorize_comparison_report(self, report: str) -> str:
        lines = report.split("\n")
        html_lines = []

        for line in lines:
            if line.startswith("==="):
                html_lines.append(
                    f"<p style='color: #4ec9b0; font-weight: bold; font-size: 12pt;'>{line}</p>"
                )
            elif line.startswith("---"):
                html_lines.append(
                    f"<p style='color: #dcdcaa; font-weight: bold; margin-top: 10px;'>{line}</p>"
                )
            elif "Icon(s) Added:" in line or "  + " in line:
                html_lines.append(f"<p style='color: #4ec9b0;'>{line}</p>")
            elif "Icon(s) Removed:" in line or "  - " in line:
                html_lines.append(f"<p style='color: #f48771;'>{line}</p>")
            elif "Icon(s) Moved:" in line or "  ↔" in line:
                html_lines.append(f"<p style='color: #dcdcaa;'>{line}</p>")
            elif "Icon(s) Unchanged:" in line:
                html_lines.append(f"<p style='color: #808080;'>{line}</p>")
            elif "✓" in line:
                html_lines.append(
                    f"<p style='color: #89d185; font-weight: bold;'>{line}</p>"
                )
            else:
                html_lines.append(f"<p style='color: #d4d4d4;'>{line}</p>")

        return "".join(html_lines)
