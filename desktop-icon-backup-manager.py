import sys
import os
import json
import struct
import ctypes
import win32gui
import win32con
import win32api
import win32process
import argparse
import random

from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple, Optional, Callable, List
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QPushButton, QTextEdit, QLabel, QMessageBox,
                             QHBoxLayout, QProgressBar, QDialog, QListWidget,
                             QListWidgetItem, QAbstractItemView, QMenu, QLineEdit,
                             QSystemTrayIcon, QToolTip)
from PyQt6.QtCore import (Qt, QThread, pyqtSignal, QTimer, QSettings,
                          QSize, QStandardPaths, QCoreApplication, QRect, QPoint,
                          QTranslator, QLocale, QUrl, QEvent)
from PyQt6.QtGui import (QAction, QKeySequence, QGuiApplication, QIcon, QPainter,
                         QColor, QPen, QDesktopServices)

def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

try:
    v_path = resource_path("version.txt")

    if os.path.exists(v_path):
        with open(v_path, "r", encoding="utf-8") as f:
            VERSION = f.read().strip()
    else:
        VERSION = "0.0.0"
except Exception as e:
    VERSION = "0.0.0"

# --- WIN32 CONSTANTS ---
class Win32Constants:
    LVM_GETITEMCOUNT = 0x1004
    LVM_GETITEMTEXTW = 0x1073
    LVM_GETITEMPOSITION = 0x1010
    LVM_SETITEMPOSITION = 0x100F
    MEM_COMMIT = 0x1000
    MEM_RELEASE = 0x8000
    PAGE_READWRITE = 0x04
    LVS_AUTOARRANGE = 0x0010
# -----------------------

REMOTE_BUFFER_SIZE = 4096
TEXT_BUFFER_OFFSET = 256
TEXT_BUFFER_SIZE = 2048
class LVITEMW(ctypes.Structure):
    _fields_ = [
        ("mask", ctypes.c_uint),
        ("iItem", ctypes.c_int),
        ("iSubItem", ctypes.c_int),
        ("state", ctypes.c_uint),
        ("stateMask", ctypes.c_uint),
        ("pszText", ctypes.c_void_p),
        ("cchTextMax", ctypes.c_int),
        ("iImage", ctypes.c_int),
        ("lParam", ctypes.c_void_p)
    ]
BACKUP_DIR = "icon_backups"

# --- HELPER FUNCTIONS ---
def get_display_metadata() -> Dict:
    app = QApplication.instance()
    if not app:
        return {"monitor_count": 0, "screens": [], "primary_resolution": "UnknownResolution"}

    screens = QGuiApplication.screens()
    metadata = {
        "monitor_count": len(screens),
        "screens": [
            {
                "id": i,
                "name": s.name(),
                "width": s.geometry().width(),
                "height": s.geometry().height(),
                "pixel_density": s.devicePixelRatio(),
            } for i, s in enumerate(screens)
        ]
    }
    if screens:
        primary_screen = screens[0]
        metadata["primary_resolution"] = f"{primary_screen.geometry().width()}x{primary_screen.geometry().height()}"
    else:
        metadata["primary_resolution"] = "UnknownResolution"
    return metadata

def parse_backup_filename(filename: str) -> Tuple[str, str, str]:
    try:
        parts = filename.replace('.json', '').split('_')
        resolution = "N/A"
        timestamp_part = filename.replace('.json', '')

        if len(parts) >= 3 and 'x' in parts[0] and len(parts[1]) == 8 and len(parts[2]) == 6:
            resolution = parts[0]
            timestamp_part = f"{parts[1]}_{parts[2]}"
        elif len(parts) >= 2 and len(parts[0]) == 8 and len(parts[1]) == 6:
            timestamp_part = f"{parts[0]}_{parts[1]}"
        else:
            timestamp_part = filename.replace('.json', '')
            try:
                datetime.strptime(timestamp_part, "%Y%m%d_%H%M%S")
            except ValueError:
                return filename.replace('.json', ''), "N/A", filename.replace('.json', '')

        dt_object = datetime.strptime(timestamp_part, "%Y%m%d_%H%M%S")
        readable_date = dt_object.strftime("%Y/%m/%d %H:%M:%S")
        return readable_date, resolution, timestamp_part

    except Exception:
        return filename.replace('.json', ''), "N/A", filename.replace('.json', '')

def parse_resolution_string(resolution_str: str) -> Optional[Tuple[int, int]]:
    try:
        if 'x' in resolution_str:
            width, height = map(int, resolution_str.split('x'))
            return width, height
        return None
    except Exception:
        return None

def get_readable_date(filename: str) -> str:
    return parse_backup_filename(filename)[0]

def get_resolution_from_filename(filename: str) -> str:
    return parse_backup_filename(filename)[1]

# --- WIDGET: VISUAL PREVIEW ---
class IconPreviewWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(450, 250)
        self.icons = {}
        self.screen_res = (1920, 1080)
        self.setMouseTracking(True)
        self.setStyleSheet("""
            QWidget {
                background-color: #1a1a1a;
                border: 2px solid #333;
                border-radius: 4px;
            }
            QToolTip {
                color: #000000;
                background-color: #ffffdc;
                border: 1px solid #000000;
                font-family: 'Segoe UI';
                font-size: 12px;
            }
        """)

    def update_preview(self, icons: Dict, res_tuple: Tuple[int, int]):
        self.icons = icons
        self.screen_res = res_tuple if res_tuple else (1920, 1080)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        if not self.icons:
            painter.setPen(QColor("#666"))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, self.tr("No Preview Available"))
            return
        scale_x = self.width() / self.screen_res[0]
        scale_y = self.height() / self.screen_res[1]
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor("#0078d7"))
        for pos in self.icons.values():
            px = int(pos[0] * scale_x)
            py = int(pos[1] * scale_y)
            px = max(5, min(px, self.width() - 5))
            py = max(5, min(py, self.height() - 5))
            painter.drawEllipse(px - 4, py - 4, 8, 8)

    def mouseMoveEvent(self, event):
        if not self.icons:
            return
        scale_x = self.width() / self.screen_res[0]
        scale_y = self.height() / self.screen_res[1]
        found_icon = None
        for name, pos in self.icons.items():
            ix = int(pos[0] * scale_x)
            iy = int(pos[1] * scale_y)
            dx = event.position().x() - ix
            dy = event.position().y() - iy
            if (dx*dx + dy*dy) < 144:
                found_icon = name
                break
        if found_icon:
            QToolTip.showText(event.globalPosition().toPoint(), found_icon, self)
        else:
            QToolTip.hideText()

# --- BACKEND: Icon management logic ---
class DesktopIconManager:
    def __init__(self):
        self.hwnd_listview = self._get_desktop_listview_hwnd()
        self._ensure_backup_directory()

    def _ensure_backup_directory(self) -> None:
        Path(BACKUP_DIR).mkdir(exist_ok=True)

    def _get_desktop_listview_hwnd(self) -> int:
        hwnd_progman = win32gui.FindWindow("Progman", None)
        hwnd_shell = win32gui.FindWindowEx(hwnd_progman, 0, "SHELLDLL_DefView", None)
        hwnd_listview = win32gui.FindWindowEx(hwnd_shell, 0, "SysListView32", None)

        if not hwnd_listview:
            def enum_windows_callback(hwnd, lParam):
                hwnd_shell = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)
                if hwnd_shell:
                    hwnd_listview_found = win32gui.FindWindowEx(hwnd_shell, 0, "SysListView32", None)
                    if hwnd_listview_found:
                        lParam.append(hwnd_listview_found)
                return True
            hwnds = []
            win32gui.EnumWindows(enum_windows_callback, hwnds)
            if hwnds:
                hwnd_listview = hwnds[0]

        if not hwnd_listview:
            raise Exception(QCoreApplication.translate("DesktopIconManager", "Unable to find desktop ListView control. Make sure desktop icons are visible."))
        return hwnd_listview

    def _list_backup_files(self) -> List[str]:
        if not os.path.exists(BACKUP_DIR):
            return []
        backup_files = [f for f in os.listdir(BACKUP_DIR) if f.endswith('.json')]
        backup_files.sort(key=lambda f: parse_backup_filename(f)[2], reverse=True)
        return backup_files

    def get_latest_backup_filename(self) -> Optional[str]:
        backup_files = self._list_backup_files()
        return backup_files[0] if backup_files else None

    def get_all_backup_filenames(self) -> List[str]:
        return self._list_backup_files()

    def delete_backup(self, filename: str) -> bool:
        filepath = os.path.join(BACKUP_DIR, filename)
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
                return True
            except Exception as e:
                print(f"Error deleting file {filepath}: {e}")
                return False
        return False

    def delete_all_backups(self, log_callback: Callable[[str], None]) -> bool:
        backup_files = self._list_backup_files()
        if not backup_files:
            log_callback(QCoreApplication.translate("DesktopIconManager", "No backup files found to delete."))
            return True

        deleted_count = 0
        failed_count = 0
        for filename in backup_files:
            if self.delete_backup(filename):
                deleted_count += 1
            else:
                failed_count += 1

        if deleted_count > 0:
            log_callback(QCoreApplication.translate("DesktopIconManager", "✓ Successfully deleted %n backup file(s).", None, deleted_count))
        if failed_count > 0:
            log_callback(QCoreApplication.translate("DesktopIconManager", "✗ Failed to delete %n backup file(s).", None, failed_count))
            return False
        return True

    def cleanup_old_backups(self, max_count: int, log_callback: Callable[[str], None]) -> None:
        if max_count <= 0:
            log_callback(QCoreApplication.translate("DesktopIconManager", "Automatic cleanup skipped: max_count is disabled (0)."))
            return

        backup_files = self._list_backup_files()
        current_count = len(backup_files)

        if current_count <= max_count:
            log_callback(QCoreApplication.translate("DesktopIconManager", "Cleanup skipped: Current count (%n) is within the limit (%1).", None, current_count).replace("%1", str(max_count)))
            return

        files_to_delete = backup_files[max_count:]
        deleted_count = 0

        log_callback(QCoreApplication.translate("DesktopIconManager", "Cleanup needed: Current count (%1) exceeds limit (%2). Deleting %n oldest file(s).", None, len(files_to_delete)).replace("%1", str(current_count)).replace("%2", str(max_count)))

        for filename in files_to_delete:
            if self.delete_backup(filename):
                deleted_count += 1
                log_callback(QCoreApplication.translate("DesktopIconManager", "  Deleted oldest backup: %1").replace("%1", str(filename)))
            else:
                log_callback(QCoreApplication.translate("DesktopIconManager", "  Failed to delete: %1").replace("%1", str(filename)))

        log_callback(QCoreApplication.translate("DesktopIconManager", "Cleanup complete. Total deleted: %n file(s).", None, deleted_count))

    def _get_latest_backup_path(self) -> Optional[str]:
        latest_file = self.get_latest_backup_filename()
        if latest_file:
            return os.path.join(BACKUP_DIR, latest_file)
        return None

    def save(self, log_callback: Callable[[str], None],
                 progress_callback: Optional[Callable[[int], None]] = None,
                 description: Optional[str] = None,
                 max_backup_count: int = 0) -> bool:

            display_metadata = get_display_metadata()
            resolution = display_metadata.get("primary_resolution", "UnknownResolution")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            filename = f"{resolution}_{timestamp}.json"
            filepath = os.path.join(BACKUP_DIR, filename)

            icons = {}  # Dizionario che conterrà i dati
            pid = win32process.GetWindowThreadProcessId(self.hwnd_listview)[1]
            process_handle = None
            remote_memory = None

            try:
                process_handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)
                remote_memory = win32process.VirtualAllocEx(
                    process_handle, 0, REMOTE_BUFFER_SIZE, Win32Constants.MEM_COMMIT, Win32Constants.PAGE_READWRITE
                )

                count = win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMCOUNT, 0, 0)
                log_callback(QCoreApplication.translate("DesktopIconManager", "Monitor Resolution: %1").replace("%1", str(resolution)))
                log_callback(QCoreApplication.translate("DesktopIconManager", "Found %1 icons. Starting scan...").replace("%1", str(count)))

                for i in range(count):
                    if progress_callback:
                        progress_callback(int((i / count) * 100))

                    win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMPOSITION, i, remote_memory)
                    point_data = win32process.ReadProcessMemory(process_handle, remote_memory, 8)
                    x, y = struct.unpack('ii', point_data)

                    text_buffer_remote = remote_memory + TEXT_BUFFER_OFFSET
                    lvitem = LVITEMW()
                    lvitem.mask = 0x0001
                    lvitem.iItem = i
                    lvitem.iSubItem = 0
                    lvitem.pszText = text_buffer_remote
                    lvitem.cchTextMax = 512

                    win32process.WriteProcessMemory(process_handle, remote_memory, bytes(lvitem))
                    win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMTEXTW, i, remote_memory)
                    text_raw = win32process.ReadProcessMemory(process_handle, text_buffer_remote, 512 * 2)
                    full_text = text_raw.decode("utf-16-le")
                    icon_name = full_text.split('\0', 1)[0]

                    if icon_name:
                        icons[icon_name] = (x, y)

                profile_data = {
                    "timestamp": datetime.now().isoformat(),
                    "icon_count": len(icons),
                    "description": description if description else "",
                    "display_metadata": display_metadata,
                    "icons": icons
                }

                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(profile_data, f, indent=4, ensure_ascii=False)

                log_callback(QCoreApplication.translate("DesktopIconManager", "✓ Saved %1 icons to backup file '%2'").replace("%1", str(len(icons))).replace("%2", str(filename)))

                if description:
                    log_callback(QCoreApplication.translate("DesktopIconManager", "  (Description: %1)").replace("%1", str(description)))

                self.cleanup_old_backups(max_backup_count, log_callback)

                if progress_callback:
                    progress_callback(100)
                return True

            except Exception as e:
                log_callback(QCoreApplication.translate("DesktopIconManager", "✗ Error saving: %1").replace("%1", str(str(e))))
                return False
            finally:
                if remote_memory and process_handle:
                    win32process.VirtualFreeEx(process_handle, remote_memory, 0, Win32Constants.MEM_RELEASE)
                if process_handle:
                    win32api.CloseHandle(process_handle)

    def restore(self, log_callback: Callable[[str], None],
                    filename: Optional[str] = None,
                    progress_callback: Optional[Callable[[int], None]] = None,
                    enable_scaling: bool = False) -> Tuple[bool, Optional[Dict]]:

            if filename:
                filepath = os.path.join(BACKUP_DIR, filename)
            else:
                filepath = self._get_latest_backup_path()

            if not filepath or not os.path.exists(filepath):
                log_callback(QCoreApplication.translate("DesktopIconManager", "✗ Error: Backup file not found."))
                return False, None

            filename = Path(filepath).name
            readable_date, resolution_saved, _ = parse_backup_filename(filename)

            log_callback(QCoreApplication.translate("DesktopIconManager", "Attempting to restore from backup: '%1'").replace("%1", str(filename)))
            log_callback(QCoreApplication.translate("DesktopIconManager", "Saved Resolution (from filename): %1").replace("%1", str(resolution_saved)))

            saved_metadata = None
            description = "N/A"

            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    profile_data = json.load(f)

                if isinstance(profile_data, dict) and "icons" in profile_data:
                    saved_icons = profile_data["icons"]
                    saved_metadata = profile_data.get("display_metadata")
                    description = profile_data.get("description", "N/A")
                    log_callback(QCoreApplication.translate("DesktopIconManager", "Restoring layout (saved: %1)").replace("%1", str(readable_date)))
                    log_callback(QCoreApplication.translate("DesktopIconManager", "  Description: %1").replace("%1", str(description)))
                else:
                    saved_icons = profile_data
                    log_callback(QCoreApplication.translate("DesktopIconManager", "Restoring layout (Old format, no timestamp and metadata)"))

            except json.JSONDecodeError as e:
                log_callback(QCoreApplication.translate("DesktopIconManager", "✗ Error: Invalid backup file format: %1").replace("%1", str(str(e))))
                return False, None

            scaling_active = False
            scale_x, scale_y = 1.0, 1.0
            current_metadata = get_display_metadata()
            current_res_str = current_metadata.get("primary_resolution", "UnknownResolution")
            current_res = parse_resolution_string(current_res_str)
            saved_res = parse_resolution_string(resolution_saved)

            if enable_scaling and current_res and saved_res and current_res != saved_res:
                scale_x = current_res[0] / saved_res[0]
                scale_y = current_res[1] / saved_res[1]
                scaling_active = True
                log_callback(QCoreApplication.translate("DesktopIconManager", "✓ Adaptive Scaling enabled: X=%1, Y=%2").replace("%1", f"{scale_x:.3f}").replace("%2", f"{scale_y:.3f}"))

            win32gui.SendMessage(self.hwnd_listview, win32con.WM_SETREDRAW, 0, 0)

            pid = win32process.GetWindowThreadProcessId(self.hwnd_listview)[1]
            process_handle = None
            remote_memory = None

            try:
                process_handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)
                remote_memory = win32process.VirtualAllocEx(process_handle, 0, REMOTE_BUFFER_SIZE, Win32Constants.MEM_COMMIT, Win32Constants.PAGE_READWRITE)

                count = win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMCOUNT, 0, 0)
                current_map = {}

                text_buffer_remote = remote_memory + TEXT_BUFFER_OFFSET
                for i in range(count):
                    if progress_callback:
                        progress_callback(int((i / (count * 2)) * 100))

                    lvitem = LVITEMW()
                    lvitem.mask = 0x0001
                    lvitem.iItem = i
                    lvitem.pszText = text_buffer_remote
                    lvitem.cchTextMax = 512

                    win32process.WriteProcessMemory(process_handle, remote_memory, bytes(lvitem))
                    win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMTEXTW, i, remote_memory)
                    text_raw = win32process.ReadProcessMemory(process_handle, text_buffer_remote, 512 * 2)
                    icon_name = text_raw.decode("utf-16-le").split('\0', 1)[0]
                    if icon_name:
                        current_map[icon_name] = i

                moved_count = 0
                skipped_count = 0
                total_saved = len(saved_icons)

                for idx, (name, pos) in enumerate(saved_icons.items()):
                    if progress_callback:
                        progress_callback(50 + int((idx / total_saved) * 50))

                    if name in current_map:
                        icon_idx = current_map[name]
                        x_saved, y_saved = pos

                        x_new = int(x_saved * scale_x) if scaling_active else x_saved
                        y_new = int(y_saved * scale_y) if scaling_active else y_saved

                        lparam = (y_new << 16) | (x_new & 0xFFFF)
                        win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_SETITEMPOSITION, icon_idx, lparam)
                        moved_count += 1
                    else:
                        skipped_count += 1

                log_callback(QCoreApplication.translate("DesktopIconManager", "✓ Restored %1 icons").replace("%1", str(moved_count)))
                if skipped_count > 0:
                    log_callback(QCoreApplication.translate("DesktopIconManager", "⚠ Skipped %1 icons (not found on desktop)").replace("%1", str(skipped_count)))

                if progress_callback:
                    progress_callback(100)
                return True, saved_metadata

            except Exception as e:
                log_callback(QCoreApplication.translate("DesktopIconManager", "✗ Error restoring: %1").replace("%1", str(e)))
                return False, saved_metadata
            finally:
                win32gui.SendMessage(self.hwnd_listview, win32con.WM_SETREDRAW, 1, 0)
                win32gui.InvalidateRect(self.hwnd_listview, None, True)
                if remote_memory and process_handle:
                    win32process.VirtualFreeEx(process_handle, remote_memory, 0, Win32Constants.MEM_RELEASE)
                if process_handle:
                    win32api.CloseHandle(process_handle)

    def scramble_icons(self, log_callback: Callable[[str], None],
                       progress_callback: Optional[Callable[[int], None]] = None) -> bool:
        try:
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
            margin = 100

            win32gui.SendMessage(self.hwnd_listview, win32con.WM_SETREDRAW, 0, 0)
            log_callback(QCoreApplication.translate("DesktopIconManager", "Redrawing disabled for scrambling..."))

            count = win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMCOUNT, 0, 0)
            log_callback(QCoreApplication.translate("DesktopIconManager", "Found %1 icons. Starting random positioning...").replace("%1", str(count)))

            for i in range(count):
                if progress_callback:
                    progress_callback(int((i / count) * 100))

                rand_x = random.randint(margin, screen_width - margin)
                rand_y = random.randint(margin, screen_height - margin)

                lparam = (rand_y << 16) | (rand_x & 0xFFFF)
                win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_SETITEMPOSITION, i, lparam)

            log_callback(QCoreApplication.translate("DesktopIconManager", "✓ Scrambled positions for %1 icons.").replace("%1", str(count)))

            if progress_callback:
                progress_callback(100)
            return True

        except Exception as e:
            log_callback(QCoreApplication.translate("DesktopIconManager", "✗ Error scrambling icons: %1").replace("%1", str(str(e))))
            return False
        finally:
            win32gui.SendMessage(self.hwnd_listview, win32con.WM_SETREDRAW, 1, 0)
            win32gui.InvalidateRect(self.hwnd_listview, None, True)
            win32api.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, "IconMetrics")

# --- BACKEND: Simple Backup Comparison ---
class BackupComparator:
    """Compare two backup files"""

    @staticmethod
    def compare(file1_path: str, file2_path: str) -> Optional[str]:
        try:
            # Load first backup
            with open(file1_path, 'r', encoding='utf-8') as f:
                data1 = json.load(f)
            icons1 = data1.get('icons', data1) if isinstance(data1, dict) else data1

            # Load second backup
            with open(file2_path, 'r', encoding='utf-8') as f:
                data2 = json.load(f)
            icons2 = data2.get('icons', data2) if isinstance(data2, dict) else data2

            # Calculate differences
            names1 = set(icons1.keys())
            names2 = set(icons2.keys())

            added = names2 - names1
            removed = names1 - names2
            moved = []

            for name in names1 & names2:
                pos1 = icons1[name]
                pos2 = icons2[name]
                if pos1 != pos2:
                    moved.append(name)

            # Build report
            num_added = len(added)
            num_removed = len(removed)
            num_moved = len(moved)
            num_unchanged = len(names1 & names2) - len(moved)
            report = QCoreApplication.translate("BackupComparator", "=== COMPARISON RESULTS ===") + "\n\n"
            report += QCoreApplication.translate("BackupComparator", "Icon(s) Added: %1", None, num_added).replace("%1", str(num_added)) + "\n"
            report += QCoreApplication.translate("BackupComparator", "Icon(s) Removed: %1", None, num_removed).replace("%1", str(num_removed)) + "\n"
            report += QCoreApplication.translate("BackupComparator", "Icon(s) Moved: %1", None, num_moved).replace("%1", str(num_moved)) + "\n"
            report += QCoreApplication.translate("BackupComparator", "Icon(s) Unchanged: %1", None, num_unchanged).replace("%1", str(num_unchanged)) + "\n\n"

            if added:
                report += QCoreApplication.translate("BackupComparator", "--- ADDED ICONS ---") + "\n"
                for name in sorted(added):
                    report += f"  + {name}\n"
                report += "\n"

            if removed:
                report += QCoreApplication.translate("BackupComparator", "--- REMOVED ICONS ---") + "\n"
                for name in sorted(removed):
                    report += f"  - {name}\n"
                report += "\n"

            if moved:
                report += QCoreApplication.translate("BackupComparator", "--- MOVED ICONS ---") + "\n"
                for name in sorted(moved):
                    report += f"  ↔ {name}\n"

            if not added and not removed and not moved:
                report += QCoreApplication.translate("BackupComparator", "✓ No differences - backups are identical!") + "\n"

            return report

        except Exception as e:
            return None

# --- THREADING: Worker ---
class IconWorker(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, object)

    def __init__(self, mode: str, filename: Optional[str] = None,
                 description: Optional[str] = None,
                 max_backup_count: int = 0,
                 enable_scaling: bool = False):
        super().__init__()
        self.mode = mode
        self.filename = filename
        self.description = description
        self.max_backup_count = max_backup_count
        self.enable_scaling = enable_scaling
        self.manager = DesktopIconManager()

    def run(self):
        success = False
        metadata = None
        try:
            if self.mode == 'save':
                success = self.manager.save(
                    self.log_signal.emit,
                    self.progress_signal.emit,
                    self.description,
                    self.max_backup_count
                )
            elif self.mode == 'restore':
                success, metadata = self.manager.restore(
                    self.log_signal.emit,
                    self.filename,
                    self.progress_signal.emit,
                    self.enable_scaling
                )
            elif self.mode == 'scramble':
                self.log_signal.emit(self.tr("Performing mandatory quick backup before scrambling..."))
                save_success = self.manager.save(
                    lambda msg: self.log_signal.emit(self.tr("  [Pre-Scramble Backup] %1").replace("%1", str(msg))),
                    lambda val: self.progress_signal.emit(int(val * 0.5)),
                    description=self.tr("Backup before Scramble"),
                    max_backup_count=0
                )
                if save_success:
                    self.log_signal.emit(self.tr("Pre-scramble backup completed successfully. Starting scramble..."))
                    success = self.manager.scramble_icons(
                        self.log_signal.emit,
                        lambda val: self.progress_signal.emit(50 + int(val * 0.5))
                    )
                else:
                    self.log_signal.emit(self.tr("✗ Pre-scramble backup failed. Aborting scramble operation."))
                    success = False

        except Exception as e:
            self.log_signal.emit(self.tr("✗ CRITICAL ERROR: %1").replace("%1", str(str(e))))
            success = False
        finally:
            self.finished_signal.emit(success, metadata)

# --- FRONTEND: Backup Manager Window ---
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
        self.layout.addWidget(QLabel(self.tr("Select a backup to restore or right-click to delete.")))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(self.tr("Search by tag, resolution, or date..."))
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
        header_label.setStyleSheet("font-family: 'Consolas', monospace; font-size: 11px; font-weight: bold; margin-bottom: 2px;")
        left_panel.addWidget(header_label)
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.list_widget.setStyleSheet("font-family: 'Consolas', monospace; font-size: 11px;")
        left_panel.addWidget(self.list_widget)
        h_split.addLayout(left_panel, 6)
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel(self.tr("Layout Preview:")))
        self.preview_widget = IconPreviewWidget()
        right_panel.addWidget(self.preview_widget)
        self.info_label = QLabel(self.tr("Select a backup to see details."))
        self.info_label.setWordWrap(True)
        self.info_label.setStyleSheet("color: #ccc; background-color: #2b2b2b; border-radius: 4px; padding: 10px; font-family: 'Segoe UI'; font-size: 12px;")
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
            filepath = os.path.join(BACKUP_DIR, filename)
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    description = data.get("description", "").strip()
                    icon_count = data.get("icon_count", "N/A")
            except Exception: pass
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
        filepath = os.path.join(BACKUP_DIR, filename)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                icons = data.get("icons", {})
                res_str = parse_backup_filename(filename)[1]
                res_tuple = parse_resolution_string(res_str)
                self.preview_widget.update_preview(icons, res_tuple)
                desc = data.get('description', self.tr('None'))
                ts = data.get('timestamp', 'N/A')
                count = len(icons)
                info = (f"<b>{self.tr('File')}:</b> {filename}<br>"
                        f"<b>{self.tr('Icons')}:</b> {count}<br>"
                        f"<b>{self.tr('Resolution')}:</b> {res_str}<br>"
                        f"<b>{self.tr('Description')}:</b> {desc}<br>"
                        f"<b>{self.tr('Timestamp')}:</b> {ts}")
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

    def show_context_menu(self, pos):
        item = self.list_widget.itemAt(pos)
        if item and item.data(Qt.ItemDataRole.UserRole):
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
            menu.addAction(restore_action)
            menu.addAction(delete_action)
            menu.exec(self.list_widget.mapToGlobal(pos))

    def get_selected_filename(self):
        selected = self.list_widget.selectedItems()
        return selected[0].data(Qt.ItemDataRole.UserRole) if selected else None

    def restore_selected(self):
        fn = self.get_selected_filename()
        if fn:
            self.restore_requested.emit(fn)
            self.accept()

    def delete_selected(self):
        fn = self.get_selected_filename()
        if fn and self.manager.delete_backup(fn):
            self.load_backups()
            self.list_changed_signal.emit()

    def compare_with_latest(self):
        """Compare selected backup with the latest one"""
        selected_filename = self.get_selected_filename()
        if not selected_filename:
            return

        # Get latest backup
        latest_filename = self.manager.get_latest_backup_filename()
        if not latest_filename:
            QMessageBox.warning(self, self.tr("Error"), self.tr("No latest backup found"))
            return

        if selected_filename == latest_filename:
            QMessageBox.information(
                self,
                self.tr("Same Backup"),
                self.tr("You selected the latest backup. Nothing to compare.")
            )
            return

        # Perform comparison
        file1 = os.path.join(BACKUP_DIR, selected_filename)
        file2 = os.path.join(BACKUP_DIR, latest_filename)

        report = BackupComparator.compare(file1, file2)

        if report:
            # Show results in a simple dialog
            dialog = QDialog(self)
            dialog.setWindowTitle(self.tr("Comparison Results"))
            dialog.resize(600, 500)

            layout = QVBoxLayout(dialog)

            # Header
            header = QLabel(
                self.tr("Comparing:\n  📁 %1\n  📁 %2 (Latest)")
                .replace("%1", selected_filename)
                .replace("%2", latest_filename)
            )
            header.setStyleSheet("font-weight: bold; padding: 10px; background-color: #f0f0f0;")
            layout.addWidget(header)

            # Results
            text_area = QTextEdit()
            text_area.setReadOnly(True)
            text_area.setPlainText(report)
            text_area.setStyleSheet("font-family: 'Consolas', monospace; font-size: 11px;")
            layout.addWidget(text_area)

            # Close button
            btn_close = QPushButton(self.tr("Close"))
            btn_close.clicked.connect(dialog.accept)
            layout.addWidget(btn_close)

            dialog.exec()
        else:
            QMessageBox.critical(
                self,
                self.tr("Error"),
                self.tr("Failed to compare backups")
            )

# --- FRONTEND: Main Window ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.manager = DesktopIconManager()
        self.current_resolution = get_display_metadata().get("primary_resolution", self.tr("Unknown"))

        app_path = Path(os.path.abspath(sys.argv[0])).parent
        settings_file_path = app_path / "settings.ini"
        self.settings = QSettings(str(settings_file_path), QSettings.Format.IniFormat)

        self.worker = None
        self.tray_icon = None

        self.create_tray_icon()

        self.DEFAULT_GEOMETRY = QRect(100, 100, 800, 650)

        self.setup_ui()
        self.setup_shortcuts()
        self.load_settings()

        if self.settings.value("auto_restore_on_startup", False, type=bool):
            QTimer.singleShot(1000, self.start_restore_latest)

    def create_tray_icon(self):
        icon = QIcon(resource_path("icon.ico"))
        self.tray_icon = QSystemTrayIcon(icon, self)

        tray_menu = QMenu()

        self.action_tray_save = QAction(self.tr("Quick Save"), self)
        self.action_tray_save.triggered.connect(lambda: self.start_save(description=self.tr("Quick Save (Tray)")))
        tray_menu.addAction(self.action_tray_save)

        self.action_tray_restore = QAction(self.tr("Restore Latest"), self)
        self.action_tray_restore.triggered.connect(self.start_restore_latest)
        tray_menu.addAction(self.action_tray_restore)

        tray_menu.addSeparator()

        self.action_tray_show = QAction(self.tr("Show Window"), self)
        self.action_tray_show.triggered.connect(self.show_window)
        tray_menu.addAction(self.action_tray_show)

        self.action_tray_exit = QAction(self.tr("Exit"), self)
        self.action_tray_exit.triggered.connect(self.exit_application)
        tray_menu.addAction(self.action_tray_exit)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

    def tray_icon_activated(self, reason: QSystemTrayIcon.ActivationReason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show_window()

    def show_window(self):
        self.showNormal()
        self.activateWindow()
        self.raise_()

    def exit_application(self):
        self.close()

    def setup_ui(self):
        self.setWindowTitle(self.tr("Desktop Icon Backup Manager by mapi68"))
        self.setWindowIcon(QIcon(resource_path("icon.ico")))


        menu_bar = self.menuBar()


        file_menu = menu_bar.addMenu(self.tr("&File"))

        self.action_scramble_icons = QAction(self.tr("Scramble Desktop Icons (Random)"), self)
        self.action_scramble_icons.setToolTip(self.tr("Randomizes the position of all desktop icons after creating a mandatory backup."))
        self.action_scramble_icons.triggered.connect(self.start_scramble)
        file_menu.addAction(self.action_scramble_icons)
        file_menu.addSeparator()

        self.action_remove_all = QAction(self.tr("Remove All Backups..."), self)
        self.action_remove_all.triggered.connect(self.confirm_and_delete_all_backups)
        file_menu.addAction(self.action_remove_all)

        file_menu.addSeparator()
        action_exit = QAction(self.tr("E&xit"), self)
        action_exit.setShortcut("Ctrl+Q")
        action_exit.triggered.connect(self.exit_application)
        file_menu.addAction(action_exit)


        settings_menu = menu_bar.addMenu(self.tr("&Settings"))

        self.action_start_minimized = QAction(self.tr("Start Minimized to Tray"), self, checkable=True)
        self.action_start_minimized.triggered.connect(lambda checked: self.settings.setValue("start_minimized", checked))
        settings_menu.addAction(self.action_start_minimized)

        settings_menu.addSeparator()

        self.action_auto_save = QAction(self.tr("Auto-Save on Exit"), self, checkable=True)
        self.action_auto_save.triggered.connect(lambda checked: self.settings.setValue("auto_save_on_exit", checked))
        settings_menu.addAction(self.action_auto_save)

        self.action_auto_restore = QAction(self.tr("Auto-Restore on Startup"), self, checkable=True)
        self.action_auto_restore.triggered.connect(lambda checked: self.settings.setValue("auto_restore_on_startup", checked))
        settings_menu.addAction(self.action_auto_restore)

        settings_menu.addSeparator()

        self.action_adaptive_scaling = QAction(self.tr("Enable Adaptive Scaling on Restore"), self, checkable=True)
        self.action_adaptive_scaling.triggered.connect(lambda checked: self.settings.setValue("adaptive_scaling_enabled", checked))
        settings_menu.addAction(self.action_adaptive_scaling)

        settings_menu.addSeparator()

        self.action_close_to_tray = QAction(self.tr("Minimize to Tray on Close ('X' button)"), self, checkable=True)
        self.action_close_to_tray.triggered.connect(lambda checked: self.settings.setValue("close_to_tray", checked))
        settings_menu.addAction(self.action_close_to_tray)

        settings_menu.addSeparator()

        self.cleanup_group = QMenu(self.tr("Automatic Backup Cleanup Limit"), self)
        settings_menu.addMenu(self.cleanup_group)
        self.cleanup_actions = {}

        limits = {
            self.tr("Disabled (Keep All)"): 0,
            self.tr("Keep Last 5"): 5,
            self.tr("Keep Last 10"): 10,
            self.tr("Keep Last 25"): 25,
            self.tr("Keep Last 50"): 50
        }

        for text, limit in limits.items():
            action = QAction(text, self, checkable=True)
            action.triggered.connect(lambda checked, l=limit: self._set_cleanup_limit(l))
            self.cleanup_group.addAction(action)
            self.cleanup_actions[limit] = action

        help_menu = menu_bar.addMenu(self.tr("&Help"))

        action_manual = QAction(self.tr("Online User Manual"), self)
        action_manual.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://mapi68.github.io/desktop-icon-backup-manager/manual.pdf")))
        help_menu.addAction(action_manual)

        help_menu.addSeparator()

        action_about = QAction(self.tr("&About"), self)
        action_about.triggered.connect(self.show_about_dialog)
        help_menu.addAction(action_about)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(10)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.save_tag_input = QLineEdit()
        self.save_tag_input.setPlaceholderText(self.tr("Optional: Enter a descriptive tag/description..."))

        tag_input_row = QHBoxLayout()
        tag_input_row.addWidget(QLabel(self.tr("Save Tag:")))
        tag_input_row.addWidget(self.save_tag_input, 1)
        layout.addLayout(tag_input_row)

        action_buttons_row = QHBoxLayout()
        action_buttons_row.setSpacing(10)

        self.btn_save_latest = QPushButton(self.tr("💾 SAVE QUICK BACKUP"))
        self.btn_save_latest.setMinimumHeight(50)
        self.btn_save_latest.setToolTip(self.tr("Save current desktop icon positions to a new file, using the tag above."))
        self.btn_save_latest.clicked.connect(self.quick_save_with_tag)
        self.btn_save_latest.setObjectName("saveButton")

        self.btn_restore_latest = QPushButton(self.tr("↺ RESTORE LATEST"))
        self.btn_restore_latest.setMinimumHeight(50)
        self.btn_restore_latest.setToolTip(self.tr("Restore icon positions from the LATEST backup file found."))
        self.btn_restore_latest.clicked.connect(self.start_restore_latest)
        self.btn_restore_latest.setObjectName("restoreButton")

        self.btn_restore_select = QPushButton(self.tr("↺ BACKUP MANAGER"))
        self.btn_restore_select.setMinimumHeight(50)
        self.btn_restore_select.setToolTip(self.tr("Opens a window to select a specific backup file to restore or delete."))
        self.btn_restore_select.clicked.connect(self.open_backup_manager)
        self.btn_restore_select.setObjectName("backupManagerButton")

        action_buttons_row.addWidget(self.btn_save_latest, 1)
        action_buttons_row.addWidget(self.btn_restore_latest, 1)
        action_buttons_row.addWidget(self.btn_restore_select, 1)

        layout.addLayout(action_buttons_row)

        layout.addWidget(QLabel(self.tr("Activity Log:")))

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMinimumHeight(300)
        self.log_area.setMaximumHeight(600)
        layout.addWidget(self.log_area)

        log_button_layout = QHBoxLayout()

        self.status_label = QLabel(self.tr("Current Resolution: %1").replace("%1", self.current_resolution))
        log_button_layout.addWidget(self.status_label)

        log_button_layout.addStretch(1)

        self.btn_clear_log = QPushButton(self.tr("Clear Log"))
        self.btn_clear_log.clicked.connect(self.log_area.clear)
        self.btn_clear_log.setMaximumWidth(150)
        self.btn_clear_log.setObjectName("clearLogButton")
        log_button_layout.addWidget(self.btn_clear_log)

        layout.addLayout(log_button_layout)

        self.setStyleSheet("""
            QPushButton[objectName="saveButton"],
            QPushButton[objectName="restoreButton"],
            QPushButton[objectName="backupManagerButton"],
            QPushButton[objectName="clearLogButton"]
            { color: white; font-weight: bold; border-radius: 6px; padding: 8px; font-size: 13px; }
            QPushButton[objectName="saveButton"]:hover,
            QPushButton[objectName="restoreButton"]:hover,
            QPushButton[objectName="backupManagerButton"]:hover,
            QPushButton[objectName="clearLogButton"]:hover
            { opacity: 0.8; }
            QPushButton:disabled { background-color: #cccccc; color: #666666; }
            QPushButton#saveButton { background-color: #00A65A; }
            QPushButton#backupManagerButton { background-color: #0078D7; }
            QPushButton#restoreButton { background-color: #CC0000; }
            QPushButton#clearLogButton { background-color: #6c757d; }
            QTextEdit { border: 1px solid #ddd; border-radius: 4px; padding: 5px; font-family: 'Consolas', monospace; font-size: 11px; }
            QProgressBar { border: 1px solid #ddd; border-radius: 4px; text-align: center; height: 20px; }
            QProgressBar::chunk { background-color: #0078D7; border-radius: 3px; }
        """)

    def setup_shortcuts(self):
        save_shortcut = QAction(self.tr("Save"), self)
        save_shortcut.setShortcut(QKeySequence("Ctrl+S"))
        save_shortcut.triggered.connect(lambda: self.start_save(description=self.tr("Quick Backup (Shortcut)")))
        self.addAction(save_shortcut)

    def load_settings(self):
        self.action_start_minimized.setChecked(self.settings.value("start_minimized", False, type=bool))
        self.action_auto_save.setChecked(self.settings.value("auto_save_on_exit", False, type=bool))
        self.action_auto_restore.setChecked(self.settings.value("auto_restore_on_startup", False, type=bool))
        self.action_adaptive_scaling.setChecked(self.settings.value("adaptive_scaling_enabled", False, type=bool))
        self.action_close_to_tray.setChecked(self.settings.value("close_to_tray", False, type=bool))

        current_limit = self.settings.value("cleanup_limit", 0, type=int)
        self._update_cleanup_menu_check(current_limit)

        geometry = self.settings.value("geometry", self.DEFAULT_GEOMETRY, type=QRect)
        self.setGeometry(geometry)

    def _set_cleanup_limit(self, limit: int):
        self.settings.setValue("cleanup_limit", limit)
        self._update_cleanup_menu_check(limit)
        self.log(self.tr("Automatic cleanup limit set to: %n backup(s) (0 = Disabled).", None, limit))

    def _update_cleanup_menu_check(self, current_limit: int):
        for limit, action in self.cleanup_actions.items():
            action.setChecked(limit == current_limit)

    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.append(f"[{timestamp}] {message}")

        if not self.isVisible() and ("✗" in message or "CRITICAL ERROR" in message):
             self.tray_icon.showMessage(self.tr("Desktop Icon Manager"), message, QSystemTrayIcon.MessageIcon.Warning, 5000)

    def toggle_buttons(self, enabled: bool):
        self.btn_save_latest.setEnabled(enabled)
        self.btn_restore_latest.setEnabled(enabled)
        self.btn_restore_select.setEnabled(enabled)
        self.action_remove_all.setEnabled(enabled)
        self.btn_clear_log.setEnabled(enabled)
        self.action_scramble_icons.setEnabled(enabled)

        if self.tray_icon:
            self.action_tray_save.setEnabled(enabled)
            self.action_tray_restore.setEnabled(enabled)

    def show_progress(self, show: bool):
        self.progress_bar.setVisible(show)
        if show:
            self.progress_bar.setValue(0)

    def update_progress(self, value: int):
        self.progress_bar.setValue(value)

    def open_backup_manager(self):
        manager_window = BackupManagerWindow(self.manager, self)
        manager_window.restore_requested.connect(self.start_restore_specific)
        manager_window.list_changed_signal.connect(lambda: self.log(self.tr("Backup list updated (item deleted).")))
        manager_window.exec()

    def quick_save_with_tag(self):
        tag = self.save_tag_input.text().strip()
        description = tag if tag else self.tr("Quick Backup")
        self.start_save(description=description)

    def start_restore_specific(self, filename: str):
        self._start_restore(filename)

    def show_about_dialog(self):
        QMessageBox.about(
            self,
            self.tr("About Desktop Icon Backup Manager"),
            self.tr("<h2>Desktop Icon Backup Manager</h2>"
                    "<p>A simple yet powerful tool to save and restore Windows desktop icon positions.</p>"
                    "<h3>Key Features:</h3>"
                    "<ul>"
                    "<li>**Quick Save:** Save icons with an optional descriptive tag.</li>"
                    "<li>**Backup Management:** Select, restore, or delete specific backups.</li>"
                    "<li>**Visual Preview:** See a mini-map of your layout.</li>"
                    "<li>**Adaptive Scaling:** Automatic adjustment for different resolutions.</li>"
                    "<li>**Automatic Cleanup:** Set a limit on backups to keep.</li>"
                    "<li>**Random Scramble:** Randomize icon positions after backup.</li>"
                    "<li>**Tray Integration:** Quick access via system tray.</li>"
                    "</ul>"
                    "<p><b>Version:</b> %1</p>"
                    "<p>Developed by: <b>mapi68</b></p>").replace("%1", VERSION)
        )

    def confirm_and_delete_all_backups(self):
        backup_count = len(self.manager.get_all_backup_filenames())

        if backup_count == 0:
            self.log(self.tr("No backup files found to delete."))
            QMessageBox.information(self, self.tr("No Backups Found"), self.tr("There are no backup files to delete."))
            return

        reply = QMessageBox.warning(
            self, self.tr("WARNING: Delete All Backups"),
            self.tr("Are you absolutely sure you want to permanently delete ALL %n desktop icon backup file(s)?\n\nThis action cannot be undone!", None, backup_count),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.log(self.tr("Starting deletion of all backup files..."))
            self.toggle_buttons(False)
            success = self.manager.delete_all_backups(self.log)
            self.toggle_buttons(True)

            if success:
                QMessageBox.information(self, self.tr("Success"), self.tr("All backup files have been successfully deleted."))
            else:
                QMessageBox.critical(self, self.tr("Error"), self.tr("Some files could not be deleted. Check the Activity Log for details."))

    def start_save(self, description: Optional[str] = None):
        cleanup_limit = self.settings.value("cleanup_limit", 0, type=int)

        self.log(self.tr("Starting new timestamped backup..."))
        if description:
            self.log(self.tr("  (Tag: %1)").replace("%1", str(description)))

        self.toggle_buttons(False)
        self.show_progress(True)
        self.statusBar().showMessage(self.tr("Saving..."))

        self.worker = IconWorker('save', description=description, max_backup_count=cleanup_limit)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_operation_finished)
        self.worker.start()

    def start_restore_latest(self):
        latest_backup_file = self.manager.get_latest_backup_filename()


        if not latest_backup_file:
            QMessageBox.warning(self, self.tr("Error"), self.tr("No backup files found to restore!"))
            self.log(self.tr("✗ Restore failed: No backup files found."))
            return

        formatted_date = get_readable_date(latest_backup_file)
        resolution = get_resolution_from_filename(latest_backup_file)

        description = "N/A"
        icon_count = "N/A"
        filepath = os.path.join(BACKUP_DIR, latest_backup_file)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                description = data.get("description", "N/A")
                icon_count = data.get("icon_count", "N/A")
        except Exception:
            description = "N/A (Old Format)"
            icon_count = "N/A"

        reply = QMessageBox.question(
            self, self.tr("Confirm Restore"),
            self.tr("Restore icon positions from the LATEST backup file:\n\nFile: %1\nResolution: %2\nIcons: %3\nTag: %4\nTimestamp: %5\n\nAre you sure you want to proceed?")
            .replace("%1", str(latest_backup_file)).replace("%2", str(resolution)).replace("%3", str(icon_count)).replace("%4", str(description)).replace("%5", str(formatted_date)),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self._start_restore(latest_backup_file)

    def _start_restore(self, filename: Optional[str] = None):
        enable_scaling = self.settings.value("adaptive_scaling_enabled", False, type=bool)

        self.log(self.tr("Starting restore from backup '%1'...").replace("%1", str(filename if filename else self.tr('latest'))))
        self.toggle_buttons(False)
        self.show_progress(True)
        self.statusBar().showMessage(self.tr("Restoring..."))

        self.worker = IconWorker('restore', filename, enable_scaling=enable_scaling)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_operation_finished)
        self.worker.start()

    def start_scramble(self):
        reply = QMessageBox.question(
            self, self.tr("Confirm Scramble"),
            self.tr("Are you sure you want to randomize the positions of ALL desktop icons?\n\n**A mandatory backup will be created first**.\n\nDo you want to proceed?"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.log(self.tr("Starting desktop icon scrambling (randomization)..."))
            self.toggle_buttons(False)
            self.show_progress(True)
            self.statusBar().showMessage(self.tr("Scrambling icons..."))

            self.worker = IconWorker('scramble')
            self.worker.log_signal.connect(self.log)
            self.worker.progress_signal.connect(self.update_progress)
            self.worker.finished_signal.connect(self.on_operation_finished)
            self.worker.start()

    def on_operation_finished(self, success: bool, saved_metadata: Optional[Dict]):
        mode = self.worker.mode if self.worker else "unknown"


        if mode == 'restore' and success:
            self._check_display_metadata(saved_metadata)

        if mode == 'save' and success:
            self.log(self.tr("Forcing desktop refresh..."))
            try:
                win32gui.SendMessage(self.manager.hwnd_listview, win32con.WM_SETREDRAW, 1, 0)
                win32gui.InvalidateRect(self.manager.hwnd_listview, None, True)
                win32api.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, "IconMetrics")
                self.log(self.tr("Desktop refresh signal sent successfully."))
            except Exception as e:
                self.log(self.tr("Warning: Failed to send desktop refresh signals: %1").replace("%1", str(str(e))))

        self.toggle_buttons(True)
        self.show_progress(False)

        if success:
            self.statusBar().showMessage(self.tr("Operation completed successfully"), 3000)
            if mode != 'save':
                QMessageBox.information(self, self.tr("Success"), self.tr("Operation completed successfully! (%1)").replace("%1", str(mode.capitalize())))
            if not self.isVisible():
                self.tray_icon.showMessage(self.tr("Desktop Icon Manager"), self.tr("%1 successful!").replace("%1", str(mode.capitalize())), QSystemTrayIcon.MessageIcon.Information, 2000)
        else:
            self.statusBar().showMessage(self.tr("Operation failed"), 3000)
            QMessageBox.warning(self, self.tr("Error"), self.tr("Operation failed (%1). Check the log for details.").replace("%1", str(mode.capitalize())))

        self.worker = None

    def _check_display_metadata(self, saved_metadata: Dict):
        current_metadata = get_display_metadata()
        saved_count = saved_metadata.get("monitor_count")
        current_count = current_metadata.get("monitor_count")

        if saved_count is None or current_count is None:
             self.log(self.tr("⚠ Warning: Display metadata missing or incomplete."))
             return

        if saved_count != current_count:
            self.log(self.tr("⚠ Warning: Saved (%n monitor(s)) vs Current (%1 monitor(s)).", None, saved_count).replace("%1", str(current_count)))
            QMessageBox.warning(
                self, self.tr("Monitor Mismatch Warning"),
                self.tr("The layout was saved with %1 monitor(s), but you currently have %2 monitor(s) connected.\n\nIcon positions have been restored, but they may be inaccurate.")
                .replace("%1", str(saved_count)).replace("%2", str(current_count))
            )
            return

        saved_screens = saved_metadata.get("screens", [])
        current_screens = current_metadata.get("screens", [])

        mismatch_found = False
        if len(saved_screens) == len(current_screens):
            for s_screen, c_screen in zip(saved_screens, current_screens):
                if s_screen.get('width') != c_screen.get('width') or s_screen.get('height') != c_screen.get('height'):
                    mismatch_found = True
                    break

        if mismatch_found:
            self.log(self.tr("⚠ Warning: Screen resolutions do not match the saved layout."))
            QMessageBox.warning(
                self, self.tr("Resolution Mismatch Warning"),
                self.tr("The screen resolutions for one or more monitors do not match the saved layout.\n\nIcon positions have been restored, but they may be inaccurate.")
            )

    def _run_final_cleanup(self):
        if self.isVisible():
            self.settings.setValue("geometry", self.geometry())

        if self.action_auto_save.isChecked():
            if self.isVisible(): self.log(self.tr("Auto-Save on Exit enabled. Performing silent backup..."))
            cleanup_limit = self.settings.value("cleanup_limit", 0, type=int)
            self.manager.save(
                lambda msg: print(f"Auto-Save Log: {msg}"),
                description=self.tr("Auto-Save on Exit"),
                max_backup_count=cleanup_limit
            )

    def closeEvent(self, event):
        close_to_tray = self.action_close_to_tray.isChecked()
        is_pyinstaller = getattr(sys, 'frozen', False)

        if close_to_tray and self.isVisible():
            self.settings.setValue("geometry", self.geometry())
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                self.tr("Desktop Icon Manager"),
                self.tr("Application minimized to system tray. Click or double-click to restore."),
                QSystemTrayIcon.MessageIcon.Information,
                2000
            )
            return

        self._run_final_cleanup()
        event.accept()

        if is_pyinstaller:
            try:
                hwnd_console = win32gui.GetConsoleWindow()
                if hwnd_console:
                    win32gui.PostMessage(hwnd_console, win32con.WM_CLOSE, 0, 0)
            except Exception: pass
        QApplication.quit()

if __name__ == "__main__":
    if QApplication.instance():
        app = QApplication.instance()
    else:
        app = QApplication(sys.argv)

    translator = QTranslator()
    if translator.load(QLocale.system(), "", "", resource_path("i18n")):
        app.installTranslator(translator)

    app_path = Path(os.path.abspath(sys.argv[0])).parent
    settings_file_path = app_path / "settings.ini"
    settings = QSettings(str(settings_file_path), QSettings.Format.IniFormat)

    parser = argparse.ArgumentParser(
        description=QCoreApplication.translate("CLI", "Desktop Icon Backup Manager CLI")
    )
    parser.add_argument('--backup', action='store_true',
                        help=QCoreApplication.translate("CLI", "Perform a backup"))
    parser.add_argument('--restore', type=str, metavar='FILENAME',
                        help=QCoreApplication.translate("CLI", "Restore a specific backup or latest"))
    parser.add_argument('--silent', action='store_true',
                        help=QCoreApplication.translate("CLI", "Run without showing the GUI"))

    args, unknown = parser.parse_known_args()

    if args.silent or args.backup or args.restore:
        manager = DesktopIconManager()
        app.setQuitOnLastWindowClosed(True)

        def silent_log(msg):
            prefix = QCoreApplication.translate("CLI", "[SILENT]")
            print(f"{prefix} {msg}")

        if args.backup:
            cleanup_limit = settings.value("cleanup_limit", 0, type=int)
            print(QCoreApplication.translate("CLI", "Starting silent backup..."))

            success = manager.save(
                silent_log,
                description=QCoreApplication.translate("CLI", "Silent CLI Backup"),
                max_backup_count=cleanup_limit
            )
            sys.exit(0 if success else 1)

        elif args.restore:
            enable_scaling = settings.value("adaptive_scaling_enabled", False, type=bool)

            filename = None
            if args.restore.lower() == 'latest':
                filename = manager.get_latest_backup_filename()
                if not filename:
                    print(QCoreApplication.translate("CLI", "Error: No backup files found for latest restore."))
                    sys.exit(1)
            else:
                filename = args.restore

            msg_restore = QCoreApplication.translate("CLI", "Starting silent restore from: %1").replace("%1", filename)
            print(msg_restore)

            success, _ = manager.restore(silent_log, filename=filename, enable_scaling=enable_scaling)
            sys.exit(0 if success else 1)

        if args.silent:
            sys.exit(0)

    app.setQuitOnLastWindowClosed(False)
    app.setStyle("Fusion")

    try:
        window = MainWindow()

        if settings.value("start_minimized", False, type=bool):
            window.hide()
            window.tray_icon.showMessage(
                window.tr("Desktop Icon Manager"),
                window.tr("Application started minimized to system tray."),
                QSystemTrayIcon.MessageIcon.Information,
                2000
            )
        else:
            window.show()

        sys.exit(app.exec())
    except Exception as e:
        error_title = QCoreApplication.translate("Main", "Critical Error")
        error_msg = QCoreApplication.translate("Main", "Failed to start application:\n%1").replace("%1", str(e))
        QMessageBox.critical(None, error_title, error_msg)
        sys.exit(1)