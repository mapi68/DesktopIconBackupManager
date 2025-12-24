import sys
import os
import json
import struct
import ctypes
import win32gui
import win32con
import win32api
import win32process
from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple, Optional, Callable, List
import random

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QPushButton, QTextEdit, QLabel, QMessageBox,
                             QHBoxLayout, QProgressBar, QDialog, QListWidget,
                             QListWidgetItem, QAbstractItemView, QMenu, QLineEdit,
                             QSystemTrayIcon)
from PyQt6.QtCore import (Qt, QThread, pyqtSignal, QTimer, QSettings,
                          QSize, QStandardPaths, QCoreApplication, QRect, QPoint)
from PyQt6.QtGui import (QAction, QKeySequence, QGuiApplication, QIcon, QPainter,
                         QColor, QPen)

# --- AUTOMATIC VERSIONING ---
now = datetime.now()
VERSION = f"0.{now.year % 10}.{now.month}.{now.day}"
# ---------------------------

# --- PYINSTALLER RESOURCE PATH FIX ---
def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
# ------------------------------------

# --- WIN32 CONSTANTS ---
class Win32Constants:
    LVM_GETITEMCOUNT = 0x1004
    LVM_GETITEMTEXT = 0x102D
    LVM_GETITEMPOSITION = 0x1010
    LVM_SETITEMPOSITION = 0x100F
    MEM_COMMIT = 0x1000
    MEM_RELEASE = 0x8000
    PAGE_READWRITE = 0x04
    LVS_AUTOARRANGE = 0x0010
# -----------------------

REMOTE_BUFFER_SIZE = 4096
TEXT_BUFFER_OFFSET = 128
TEXT_BUFFER_SIZE = 1024
LVITEM_FORMAT = 'IIIIIQI'
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

        # New format: Resolution_Date_Time (e.g., 1920x1080_20240101_103000)
        if len(parts) >= 3 and 'x' in parts[0] and len(parts[1]) == 8 and len(parts[2]) == 6:
            resolution = parts[0]
            timestamp_part = f"{parts[1]}_{parts[2]}"
        # Old format: Date_Time (e.g., 20240101_103000)
        elif len(parts) >= 2 and len(parts[0]) == 8 and len(parts[1]) == 6:
            timestamp_part = f"{parts[0]}_{parts[1]}"
        # Fallback
        else:
            timestamp_part = filename.replace('.json', '')
            try:
                datetime.strptime(timestamp_part, "%Y%m%d_%H%M%S")
            except ValueError:
                return filename.replace('.json', ''), "N/A", filename.replace('.json', '')

        dt_object = datetime.strptime(timestamp_part, "%Y%m%d_%H%M%S")
        readable_date = dt_object.strftime("%Y/%m/%d at %H:%M:%S")
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
    """A widget that renders a mini-map of the saved icon layout."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(220, 124)  # 16:9 ratio approximately
        self.icons = {}
        self.screen_res = (1920, 1080)
        self.setStyleSheet("background-color: #1a1a1a; border: 2px solid #333; border-radius: 4px;")

    def update_preview(self, icons: Dict, res_tuple: Tuple[int, int]):
        """Update the preview with new icon positions and resolution."""
        self.icons = icons
        self.screen_res = res_tuple if res_tuple else (1920, 1080)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        if not self.icons:
            painter.setPen(QColor("#666"))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "No Preview Available")
            return

        # Calculate scaling factors
        scale_x = self.width() / self.screen_res[0]
        scale_y = self.height() / self.screen_res[1]

        # Draw icons as small dots
        painter.setPen(QPen(QColor("#0078d7"), 3))
        for pos in self.icons.values():
            x = int(pos[0] * scale_x)
            y = int(pos[1] * scale_y)
            # Basic clipping within preview bounds
            x = max(2, min(x, self.width() - 2))
            y = max(2, min(y, self.height() - 2))
            painter.drawPoint(x, y)


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
            raise Exception("Unable to find desktop ListView control. Make sure desktop icons are visible.")
        return hwnd_listview

    def _list_backup_files(self) -> List[str]:
        if not os.path.exists(BACKUP_DIR):
            return []
        backup_files = [f for f in os.listdir(BACKUP_DIR) if f.endswith('.json')]
        # Sort based on timestamp embedded in the filename
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
            log_callback("No backup files found to delete.")
            return True

        deleted_count = 0
        failed_count = 0
        for filename in backup_files:
            if self.delete_backup(filename):
                deleted_count += 1
            else:
                failed_count += 1

        if deleted_count > 0:
            log_callback(f"✓ Successfully deleted {deleted_count} backup files.")
        if failed_count > 0:
            log_callback(f"✗ Failed to delete {failed_count} backup files.")
            return False
        return True

    def cleanup_old_backups(self, max_count: int, log_callback: Callable[[str], None]) -> None:
        if max_count <= 0:
            log_callback("Automatic cleanup skipped: max_count is disabled (0).")
            return

        backup_files = self._list_backup_files()
        current_count = len(backup_files)

        if current_count <= max_count:
            log_callback(f"Cleanup skipped: Current count ({current_count}) is within the limit ({max_count}).")
            return

        files_to_delete = backup_files[max_count:]
        deleted_count = 0

        log_callback(f"Cleanup needed: Current count ({current_count}) exceeds limit ({max_count}). Deleting {len(files_to_delete)} oldest files.")

        for filename in files_to_delete:
            if self.delete_backup(filename):
                deleted_count += 1
                log_callback(f"  Deleted oldest backup: {filename}")
            else:
                log_callback(f"  Failed to delete: {filename}")

        log_callback(f"Cleanup complete. Total deleted: {deleted_count}.")

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

        # NEW FILENAME FORMAT: [Resolution]_[Timestamp].json
        filename = f"{resolution}_{timestamp}.json"
        filepath = os.path.join(BACKUP_DIR, filename)

        icons = {}
        pid = win32process.GetWindowThreadProcessId(self.hwnd_listview)[1]
        process_handle = None
        remote_memory = None

        try:
            process_handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)
            remote_memory = win32process.VirtualAllocEx(
                process_handle, 0, REMOTE_BUFFER_SIZE, Win32Constants.MEM_COMMIT, Win32Constants.PAGE_READWRITE
            )

            count = win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMCOUNT, 0, 0)
            log_callback(f"Monitor Resolution: {resolution}")
            log_callback(f"Found {count} icons. Starting scan...")

            for i in range(count):
                if progress_callback:
                    progress_callback(int((i / count) * 100))

                # Get position
                win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMPOSITION, i, remote_memory)
                point_data = win32process.ReadProcessMemory(process_handle, remote_memory, 8)
                x, y = struct.unpack('ii', point_data)

                # Get name
                text_buffer_remote = remote_memory + TEXT_BUFFER_OFFSET
                lvitem_data = struct.pack(
                    LVITEM_FORMAT, 0x0001, i, 0, 0, 0, text_buffer_remote, 512
                )
                win32process.WriteProcessMemory(process_handle, remote_memory, lvitem_data)
                win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMTEXT, i, remote_memory)
                text_raw = win32process.ReadProcessMemory(process_handle, text_buffer_remote, TEXT_BUFFER_SIZE)
                icon_name = text_raw.decode('utf-16', errors='ignore').split('\x00')[0]

                if icon_name:
                    icons[icon_name] = (x, y)

            # Save with metadata
            profile_data = {
                "timestamp": datetime.now().isoformat(),
                "icon_count": len(icons),
                "description": description if description else "",
                "display_metadata": display_metadata,
                "icons": icons
            }

            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(profile_data, f, indent=4)

            log_callback(f"✓ Saved {len(icons)} icons to backup file '{filename}'")
            if description:
                log_callback(f"  (Description: {description})")

            self.cleanup_old_backups(max_backup_count, log_callback)

            if progress_callback:
                progress_callback(100)
            return True

        except Exception as e:
            log_callback(f"✗ Error saving: {str(e)}")
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
            log_callback(f"✗ Error: Backup file not found.")
            return False, None

        filename = Path(filepath).name
        readable_date, resolution_saved, _ = parse_backup_filename(filename)
        log_callback(f"Attempting to restore from backup: '{filename}'")
        log_callback(f"Saved Resolution (from filename): {resolution_saved}")

        saved_metadata = None
        description = "N/A"

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                profile_data = json.load(f)

            if "icons" in profile_data:
                saved_icons = profile_data["icons"]
                saved_metadata = profile_data.get("display_metadata")
                description = profile_data.get("description", "N/A")
                log_callback(f"Restoring layout (saved: {readable_date})")
                log_callback(f"  Description: {description}")
            else:
                saved_icons = profile_data
                log_callback("Restoring layout (Old format, no timestamp/metadata)")

        except json.JSONDecodeError as e:
            log_callback(f"✗ Error: Invalid backup file format: {str(e)}")
            return False, None

        # --- ADAPTIVE SCALING LOGIC ---
        scaling_active = False
        scale_x = 1.0
        scale_y = 1.0

        current_metadata = get_display_metadata()
        current_res_str = current_metadata.get("primary_resolution", "UnknownResolution")
        current_res = parse_resolution_string(current_res_str)
        saved_res = parse_resolution_string(resolution_saved)

        if enable_scaling and current_res and saved_res and current_res != saved_res:
            saved_width, saved_height = saved_res
            current_width, current_height = current_res
            scale_x = current_width / saved_width
            scale_y = current_height / saved_height
            scaling_active = True
            log_callback(f"✓ Adaptive Scaling enabled: Saved {saved_width}x{saved_height} -> Current {current_width}x{current_height}")
            log_callback(f"  **[SCALING APPLIED]** Scaling factors: X={scale_x:.3f}, Y={scale_y:.3f}")
        else:
            if enable_scaling:
                log_callback("Adaptive Scaling enabled, but resolutions match or are invalid. Scaling skipped.")
            else:
                log_callback("Adaptive Scaling is disabled. Restoring raw coordinates.")
        # --- END SCALING LOGIC ---

        # Actual restoration logic
        win32gui.SendMessage(self.hwnd_listview, win32con.WM_SETREDRAW, 0, 0)

        pid = win32process.GetWindowThreadProcessId(self.hwnd_listview)[1]
        process_handle = None
        remote_memory = None

        try:
            process_handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)
            remote_memory = win32process.VirtualAllocEx(
                process_handle, 0, REMOTE_BUFFER_SIZE, Win32Constants.MEM_COMMIT, Win32Constants.PAGE_READWRITE
            )

            count = win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMCOUNT, 0, 0)
            log_callback(f"Restoring positions for {len(saved_icons)} saved icons...")

            # Map current indices
            current_map = {}
            for i in range(count):
                if progress_callback:
                    progress_callback(int((i / (count * 2)) * 100))

                text_buffer_remote = remote_memory + TEXT_BUFFER_OFFSET
                lvitem_data = struct.pack(
                    LVITEM_FORMAT, 0x0001, i, 0, 0, 0, text_buffer_remote, 512
                )
                win32process.WriteProcessMemory(process_handle, remote_memory, lvitem_data)
                win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMTEXT, i, remote_memory)
                text_raw = win32process.ReadProcessMemory(process_handle, text_buffer_remote, TEXT_BUFFER_SIZE)
                icon_name = text_raw.decode('utf-16', errors='ignore').split('\x00')[0]
                if icon_name:
                    current_map[icon_name] = i

            # Restore positions
            moved_count = 0
            skipped_count = 0
            for idx, (name, pos) in enumerate(saved_icons.items()):
                if progress_callback:
                    progress_callback(int((50 + (idx / len(saved_icons)) * 50)))

                if name in current_map:
                    icon_idx = current_map[name]
                    x_saved, y_saved = pos

                    # Apply scaling
                    if scaling_active:
                        x_new = int(x_saved * scale_x)
                        y_new = int(y_saved * scale_y)
                    else:
                        x_new = x_saved
                        y_new = y_saved

                    lparam = (y_new << 16) | (x_new & 0xFFFF)
                    win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_SETITEMPOSITION, icon_idx, lparam)
                    moved_count += 1
                else:
                    skipped_count += 1

            log_callback(f"✓ Restored {moved_count} icons")
            if skipped_count > 0:
                log_callback(f"⚠ Skipped {skipped_count} icons (not found on desktop)")

            if progress_callback:
                progress_callback(100)
            return True, saved_metadata

        except Exception as e:
            log_callback(f"✗ Error restoring: {str(e)}")
            return False, saved_metadata
        finally:
            if remote_memory and process_handle:
                win32process.VirtualFreeEx(process_handle, remote_memory, 0, Win32Constants.MEM_RELEASE)
            if process_handle:
                win32api.CloseHandle(process_handle)

    # --- NEW FUNCTION: SCRAMBLE ICONS ---
    def scramble_icons(self, log_callback: Callable[[str], None],
                       progress_callback: Optional[Callable[[int], None]] = None) -> bool:
        try:
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
            margin = 100

            win32gui.SendMessage(self.hwnd_listview, win32con.WM_SETREDRAW, 0, 0)
            log_callback("Redrawing disabled for scrambling...")

            count = win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_GETITEMCOUNT, 0, 0)
            log_callback(f"Found {count} icons. Starting random positioning...")

            for i in range(count):
                if progress_callback:
                    progress_callback(int((i / count) * 100))

                rand_x = random.randint(margin, screen_width - margin)
                rand_y = random.randint(margin, screen_height - margin)

                lparam = (rand_y << 16) | (rand_x & 0xFFFF)
                win32gui.SendMessage(self.hwnd_listview, Win32Constants.LVM_SETITEMPOSITION, i, lparam)

            log_callback(f"✓ Scrambled positions for {count} icons.")

            if progress_callback:
                progress_callback(100)
            return True

        except Exception as e:
            log_callback(f"✗ Error scrambling icons: {str(e)}")
            return False
        finally:
            # Re-enable redraw and force a refresh
            win32gui.SendMessage(self.hwnd_listview, win32con.WM_SETREDRAW, 1, 0)
            win32gui.InvalidateRect(self.hwnd_listview, None, True)
            win32api.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, "IconMetrics")
    # --- END SCRAMBLE ICONS ---


# --- THREADING: Worker to avoid blocking the GUI ---
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
            # --- SCRAMBLE MODE ---
            elif self.mode == 'scramble':
                self.log_signal.emit("Performing mandatory quick backup before scrambling...")
                # 1. Mandatory Save (50% progress)
                save_success = self.manager.save(
                    lambda msg: self.log_signal.emit(f"  [Pre-Scramble Backup] {msg}"),
                    lambda val: self.progress_signal.emit(int(val * 0.5)), # 0-50%
                    description=f"Auto-Backup before Scramble (Random)",
                    max_backup_count=0
                )
                if save_success:
                    self.log_signal.emit("Pre-scramble backup completed successfully. Starting scramble...")
                    # 2. Scramble Icons (50%-100% progress)
                    success = self.manager.scramble_icons(
                        self.log_signal.emit,
                        lambda val: self.progress_signal.emit(50 + int(val * 0.5))
                    )
                else:
                    self.log_signal.emit("✗ Pre-scramble backup failed. Aborting scramble operation.")
                    success = False
            # --- END SCRAMBLE MODE ---

        except Exception as e:
            self.log_signal.emit(f"✗ CRITICAL ERROR: {str(e)}")
            success = False
        finally:
            self.finished_signal.emit(success, metadata)

# --- FRONTEND: Backup Manager Window ---
class BackupManagerWindow(QDialog):
    restore_requested = pyqtSignal(str)
    list_changed_signal = pyqtSignal()

    def __init__(self, manager: DesktopIconManager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Select, Restore, or Delete Backup")
        self.setFixedSize(800, 450)

        self.layout = QVBoxLayout(self)
        self.layout.setSpacing(10)

        self.layout.addWidget(QLabel("Select a backup to restore or right-click to delete."))

        # --- HORIZONTAL SPLIT: List and Preview Panel ---
        h_split = QHBoxLayout()

        # Left side: List of backups
        left_panel = QVBoxLayout()

        # List header
        header_text = (
            f"{'TAG/DESCRIPTION':<30} "
            f"| {'RESOLUTION':<10} "
            f"| {'ICONS':<5} "
            f"| TIMESTAMP"
        )
        header_label = QLabel(header_text)
        header_label.setStyleSheet("font-family: 'Consolas', monospace; font-size: 11px; font-weight: bold; margin-bottom: 2px;")
        left_panel.addWidget(header_label)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.list_widget.setStyleSheet("font-family: 'Consolas', monospace; font-size: 11px;")
        left_panel.addWidget(self.list_widget)

        h_split.addLayout(left_panel, 3)

        # Right side: Preview and Info Panel
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel("Layout Preview:"))

        self.preview_widget = IconPreviewWidget()
        right_panel.addWidget(self.preview_widget)

        self.info_label = QLabel("Select a backup to see details.")
        self.info_label.setWordWrap(True)
        self.info_label.setStyleSheet("color: #555; font-size: 11px; padding: 10px;")
        right_panel.addWidget(self.info_label)
        right_panel.addStretch()

        h_split.addLayout(right_panel, 1)
        self.layout.addLayout(h_split)
        # --- END HORIZONTAL SPLIT ---

        # Buttons
        button_layout = QHBoxLayout()

        self.btn_restore = QPushButton("Restore Selected Layout")
        self.btn_restore.clicked.connect(self.restore_selected)
        self.btn_restore.setEnabled(False)
        self.btn_restore.setObjectName("restoreButton")

        self.btn_close = QPushButton("Close")
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
            QListWidgetItem("No backups found.", self.list_widget)
            self.btn_restore.setEnabled(False)
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
            except Exception:
                pass

            # Format the item text: [Tag] [Res] [Count] Timestamp
            description_display = f"{f'[{description[:28]}]':<31}"
            resolution_display = f"| {resolution:<10}"
            icon_count_display = f"| {icon_count:>5}"

            item_text = (
                f"{description_display}"
                f"{resolution_display}"
                f"{icon_count_display}"
                f" | {readable_date}"
            )

            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, filename)
            self.list_widget.addItem(item)

        self.update_button_states()

    def on_selection_changed(self):
        """Handle selection change to update preview and info."""
        items = self.list_widget.selectedItems()
        if not items:
            self.preview_widget.update_preview({}, (1920, 1080))
            self.info_label.setText("Select a backup to see details.")
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

                # Update preview widget
                self.preview_widget.update_preview(icons, res_tuple)

                # Update info label
                description = data.get('description', 'None')
                timestamp = data.get('timestamp', 'N/A')
                info_text = (
                    f"File: {filename}\n"
                    f"Icons: {len(icons)}\n"
                    f"Resolution: {res_str}\n"
                    f"Description: {description}\n"
                    f"Timestamp: {timestamp}"
                )
                self.info_label.setText(info_text)
                self.btn_restore.setEnabled(True)
        except Exception as e:
            self.preview_widget.update_preview({}, (1920, 1080))
            self.info_label.setText(f"Error loading backup:\n{str(e)}")
            self.btn_restore.setEnabled(False)

    def update_button_states(self):
        has_selection = bool(self.list_widget.selectedItems())
        self.btn_restore.setEnabled(has_selection)

    def show_context_menu(self, pos: QPoint):
        selected_item = self.list_widget.itemAt(pos)
        if selected_item:
            context_menu = QMenu(self)

            action_restore = QAction("Restore Selected", self)
            action_restore.triggered.connect(self.restore_selected)
            context_menu.addAction(action_restore)

            context_menu.addSeparator()

            action_delete = QAction("Delete Selected", self)
            action_delete.triggered.connect(self.delete_selected)
            context_menu.addAction(action_delete)

            context_menu.exec(self.list_widget.mapToGlobal(pos))

    def get_selected_filename(self) -> Optional[str]:
        selected = self.list_widget.selectedItems()
        if selected:
            return selected[0].data(Qt.ItemDataRole.UserRole)
        return None

    def restore_selected(self):
        filename = self.get_selected_filename()
        if filename:
            self.restore_requested.emit(filename)
            self.accept()

    def delete_selected(self):
        filename = self.get_selected_filename()
        if not filename: return

        reply = QMessageBox.question(
            self, "Confirm Deletion",
            f"Are you sure you want to permanently delete the backup:\n{filename}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.manager.delete_backup(filename):
                self.load_backups()
                self.list_changed_signal.emit()
            else:
                QMessageBox.critical(self, "Error", f"Failed to delete backup '{filename}'.")


# --- FRONTEND: PyQt6 GUI ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.manager = DesktopIconManager()
        self.current_resolution = get_display_metadata().get("primary_resolution", "Unknown")

        # --- QSettings for INI file ---
        app_path = Path(os.path.abspath(sys.argv[0])).parent
        settings_file_path = app_path / "settings.ini"
        self.settings = QSettings(str(settings_file_path), QSettings.Format.IniFormat)

        self.worker = None

        # --- System Tray Initialization ---
        self.tray_icon = None
        self.create_tray_icon()
        # ---------------------------------

        self.DEFAULT_GEOMETRY = QRect(100, 100, 650, 450)

        self.setup_ui()
        self.setup_shortcuts()
        self.load_settings()

        # Check for auto-restore on startup
        if self.settings.value("auto_restore_on_startup", False, type=bool):
            QTimer.singleShot(1000, self.start_restore_latest)

    # --- System Tray Icon Methods ---
    def create_tray_icon(self):
        icon = QIcon(resource_path("icon.ico"))
        self.tray_icon = QSystemTrayIcon(icon, self)

        tray_menu = QMenu()

        self.action_tray_save = QAction("Quick Save", self)
        self.action_tray_save.triggered.connect(lambda: self.start_save(description="Quick Save (Tray)"))
        tray_menu.addAction(self.action_tray_save)

        self.action_tray_restore = QAction("Restore Latest", self)
        self.action_tray_restore.triggered.connect(self.start_restore_latest)
        tray_menu.addAction(self.action_tray_restore)

        tray_menu.addSeparator()

        self.action_tray_show = QAction("Show Window", self)
        self.action_tray_show.triggered.connect(self.show_window)
        tray_menu.addAction(self.action_tray_show)

        self.action_tray_exit = QAction("Exit", self)
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
    # --- END: System Tray Icon Methods ---

    def setup_ui(self):
        self.setWindowTitle("Desktop Icon Backup Manager by mapi68")
        self.setWindowIcon(QIcon(resource_path("icon.ico")))

        menu_bar = self.menuBar()

        # File Menu
        file_menu = menu_bar.addMenu("&File")

        self.action_scramble_icons = QAction("Scramble Desktop Icons (Random)", self)
        self.action_scramble_icons.setToolTip("Randomizes the position of all desktop icons after creating a mandatory backup.")
        self.action_scramble_icons.triggered.connect(self.start_scramble)
        file_menu.addAction(self.action_scramble_icons)
        file_menu.addSeparator()

        self.action_remove_all = QAction("Remove All Backups...", self)
        self.action_remove_all.triggered.connect(self.confirm_and_delete_all_backups)
        file_menu.addAction(self.action_remove_all)

        file_menu.addSeparator()
        action_exit = QAction("E&xit", self)
        action_exit.setShortcut("Ctrl+Q")
        action_exit.triggered.connect(self.exit_application)
        file_menu.addAction(action_exit)

        # Settings Menu
        settings_menu = menu_bar.addMenu("&Settings")

        self.action_start_minimized = QAction("Start Minimized to Tray", self, checkable=True)
        self.action_start_minimized.triggered.connect(lambda checked: self.settings.setValue("start_minimized", checked))
        settings_menu.addAction(self.action_start_minimized)

        settings_menu.addSeparator()

        self.action_auto_save = QAction("Auto-Save on Exit", self, checkable=True)
        self.action_auto_save.triggered.connect(lambda checked: self.settings.setValue("auto_save_on_exit", checked))
        settings_menu.addAction(self.action_auto_save)

        self.action_auto_restore = QAction("Auto-Restore on Startup", self, checkable=True)
        self.action_auto_restore.triggered.connect(lambda checked: self.settings.setValue("auto_restore_on_startup", checked))
        settings_menu.addAction(self.action_auto_restore)

        settings_menu.addSeparator()

        self.action_adaptive_scaling = QAction("Enable Adaptive Scaling on Restore", self, checkable=True)
        self.action_adaptive_scaling.triggered.connect(lambda checked: self.settings.setValue("adaptive_scaling_enabled", checked))
        settings_menu.addAction(self.action_adaptive_scaling)

        settings_menu.addSeparator()

        self.action_close_to_tray = QAction("Minimize to Tray on Close ('X' button)", self, checkable=True)
        self.action_close_to_tray.triggered.connect(lambda checked: self.settings.setValue("close_to_tray", checked))
        settings_menu.addAction(self.action_close_to_tray)

        settings_menu.addSeparator()

        # Automatic Cleanup Settings
        self.cleanup_group = QMenu("Automatic Backup Cleanup Limit", self)
        settings_menu.addMenu(self.cleanup_group)
        self.cleanup_actions = {}

        limits = {
            "Disabled (Keep All)": 0,
            "Keep Last 5": 5,
            "Keep Last 10": 10,
            "Keep Last 25": 25,
            "Keep Last 50": 50
        }

        for text, limit in limits.items():
            action = QAction(text, self, checkable=True)
            action.triggered.connect(lambda checked, l=limit: self._set_cleanup_limit(l))
            self.cleanup_group.addAction(action)
            self.cleanup_actions[limit] = action

        # Help Menu
        help_menu = menu_bar.addMenu("&Help")
        action_about = QAction("&About", self)
        action_about.triggered.connect(self.show_about_dialog)
        help_menu.addAction(action_about)


        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(10)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # --- LOGICAL BUTTONS AND TAG INPUT LAYOUT ---

        # 1. Save tag input
        self.save_tag_input = QLineEdit()
        self.save_tag_input.setPlaceholderText("Optional: Enter a descriptive tag/description...")

        tag_input_row = QHBoxLayout()
        tag_input_row.addWidget(QLabel("Save Tag:"))
        tag_input_row.addWidget(self.save_tag_input, 1)
        layout.addLayout(tag_input_row)

        # 2. Main Action Buttons Row
        action_buttons_row = QHBoxLayout()
        action_buttons_row.setSpacing(10)

        self.btn_save_latest = QPushButton("💾 SAVE QUICK BACKUP")
        self.btn_save_latest.setMinimumHeight(50)
        self.btn_save_latest.setToolTip("Save current desktop icon positions to a new file, using the tag above.")
        self.btn_save_latest.clicked.connect(self.quick_save_with_tag)
        self.btn_save_latest.setObjectName("saveButton")

        self.btn_restore_latest = QPushButton("↺ RESTORE LATEST")
        self.btn_restore_latest.setMinimumHeight(50)
        self.btn_restore_latest.setToolTip("Restore icon positions from the LATEST backup file found.")
        self.btn_restore_latest.clicked.connect(self.start_restore_latest)
        self.btn_restore_latest.setObjectName("restoreButton")

        self.btn_restore_select = QPushButton("↺ BACKUP MANAGER")
        self.btn_restore_select.setMinimumHeight(50)
        self.btn_restore_select.setToolTip("Opens a window to select a specific backup file to restore or delete.")
        self.btn_restore_select.clicked.connect(self.open_backup_manager)
        self.btn_restore_select.setObjectName("backupManagerButton")

        action_buttons_row.addWidget(self.btn_save_latest, 1)
        action_buttons_row.addWidget(self.btn_restore_latest, 1)
        action_buttons_row.addWidget(self.btn_restore_select, 1)

        layout.addLayout(action_buttons_row)

        # Log area
        layout.addWidget(QLabel("Activity Log:"))

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        layout.addWidget(self.log_area)

        # Clear Log Button
        self.btn_clear_log = QPushButton("Clear Log")
        self.btn_clear_log.clicked.connect(self.log_area.clear)
        self.btn_clear_log.setMaximumWidth(150)
        self.btn_clear_log.setObjectName("clearLogButton")

        log_button_layout = QHBoxLayout()
        log_button_layout.addStretch(1)
        log_button_layout.addWidget(self.btn_clear_log)
        layout.addLayout(log_button_layout)


        # Status bar
        self.status_label = QLabel(f"Current Resolution: {self.current_resolution}")
        self.statusBar().addWidget(self.status_label)
        self.statusBar().showMessage("Ready")

        # --- STYLES ---
        self.setStyleSheet("""
            /* Base Style */
            QPushButton[objectName="saveButton"],
            QPushButton[objectName="restoreButton"],
            QPushButton[objectName="backupManagerButton"],
            QPushButton[objectName="clearLogButton"]
            {
                color: white;
                font-weight: bold;
                border-radius: 6px;
                padding: 8px;
                font-size: 13px;
            }

            /* Hover/Pressed Effects */
            QPushButton[objectName="saveButton"]:hover,
            QPushButton[objectName="restoreButton"]:hover,
            QPushButton[objectName="backupManagerButton"]:hover,
            QPushButton[objectName="clearLogButton"]:hover
            { opacity: 0.8; }

            QPushButton[objectName="saveButton"]:pressed,
            QPushButton[objectName="restoreButton"]:pressed,
            QPushButton[objectName="backupManagerButton"]:pressed,
            QPushButton[objectName="clearLogButton"]:pressed
            { opacity: 0.6; }

            QPushButton:disabled { background-color: #cccccc; color: #666666; }

            QPushButton#saveButton { background-color: #00A65A; }
            QPushButton#saveButton:hover { background-color: #008748; }
            QPushButton#saveButton:pressed { background-color: #006836; }

            QPushButton#backupManagerButton { background-color: #0078D7; }
            QPushButton#backupManagerButton:hover { background-color: #0063AD; }
            QPushButton#backupManagerButton:pressed { background-color: #004D84; }

            QPushButton#restoreButton { background-color: #CC0000; }
            QPushButton#restoreButton:hover { background-color: #A30000; }
            QPushButton#restoreButton:pressed { background-color: #7A0000; }

            QPushButton#clearLogButton { background-color: #6c757d; }
            QPushButton#clearLogButton:hover { background-color: #5a6268; }
            QPushButton#clearLogButton:pressed { background-color: #495057; }

            QTextEdit { border: 1px solid #ddd; border-radius: 4px; padding: 5px; font-family: 'Consolas', monospace; font-size: 11px; }
            QToolTip { background-color: #2b2b2b; color: white; border: 1px solid #555; padding: 5px; border-radius: 3px; font-size: 12px; }
            QProgressBar { border: 1px solid #ddd; border-radius: 4px; text-align: center; height: 20px; }
            QProgressBar::chunk { background-color: #0078D7; border-radius: 3px; }
            QLabel { color: #333; }
        """)

    def setup_shortcuts(self):
        save_shortcut = QAction("Save", self)
        save_shortcut.setShortcut(QKeySequence("Ctrl+S"))
        save_shortcut.triggered.connect(lambda: self.start_save(description="Quick Backup (Shortcut)"))
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

    # --- Cleanup Menu Logic ---
    def _set_cleanup_limit(self, limit: int):
        self.settings.setValue("cleanup_limit", limit)
        self._update_cleanup_menu_check(limit)
        self.log(f"Automatic cleanup limit set to: {limit} backups (0 = Disabled).")

    def _update_cleanup_menu_check(self, current_limit: int):
        for limit, action in self.cleanup_actions.items():
            action.setChecked(limit == current_limit)
    # --------------------------------

    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.append(f"[{timestamp}] {message}")

        if not self.isVisible() and ("✗" in message or "CRITICAL ERROR" in message):
             self.tray_icon.showMessage("Desktop Icon Manager", message, QSystemTrayIcon.MessageIcon.Warning, 5000)

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

    # --- Backup Manager Logic ---
    def open_backup_manager(self):
        manager_window = BackupManagerWindow(self.manager, self)
        manager_window.restore_requested.connect(self.start_restore_specific)
        manager_window.list_changed_signal.connect(lambda: self.log("Backup list updated (item deleted)."))
        manager_window.exec()

    def quick_save_with_tag(self):
        tag = self.save_tag_input.text().strip()
        description = tag if tag else "Quick Backup"
        self.start_save(description=description)

    def start_restore_specific(self, filename: str):
        self._start_restore(filename)

    def show_about_dialog(self):
        QMessageBox.about(
            self,
            "About Desktop Icon Backup Manager",
            "<h2>Desktop Icon Backup Manager</h2>"
            "<p>A simple yet powerful tool to save and restore Windows desktop icon positions.</p>"
            "<h3>Key Features:</h3>"
            "<ul>"
            "<li>**Quick Save:** Save icons with an optional descriptive tag.</li>"
            "<li>**Backup Management:** Select, restore, or delete specific backups in a dedicated window.</li>"
            "<li>**Visual Preview:** See a mini-map of your icon layout before restoring.</li>"
            "<li>**Adaptive Scaling:** Option to automatically scale icon positions for different monitor resolutions.</li>"
            "<li>**Automatic Cleanup:** Set a maximum limit on the number of backups to keep.</li>"
            "<li>**Random Scramble:** Randomize icon positions after an automatic backup.</li>"
            "<li>**Tray Icon Integration:** Quick access to save and restore, with options to start/close minimized.</li>"
            "</ul>"
            f"<p><b>Version:</b> {VERSION}</p>"
            "<p>Developed by: <b>mapi68</b></p>"
        )

    def confirm_and_delete_all_backups(self):
        backup_count = len(self.manager.get_all_backup_filenames())

        if backup_count == 0:
            self.log("No backup files found to delete.")
            QMessageBox.information(self, "No Backups Found", "There are no backup files to delete.")
            return

        reply = QMessageBox.warning(
            self, "WARNING: Delete All Backups",
            f"Are you absolutely sure you want to permanently delete ALL {backup_count} desktop icon backup files?\n\n"
            f"This action cannot be undone!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.log("Starting deletion of all backup files...")
            self.toggle_buttons(False)
            success = self.manager.delete_all_backups(self.log)
            self.toggle_buttons(True)

            if success:
                QMessageBox.information(self, "Success", "All backup files have been successfully deleted.")
            else:
                QMessageBox.critical(self, "Error", "Some files could not be deleted. Check the Activity Log for details.")

    # --- Save/Restore Logic ---
    def start_save(self, description: Optional[str] = None):
        cleanup_limit = self.settings.value("cleanup_limit", 0, type=int)

        self.log(f"Starting new timestamped backup...")
        if description:
            self.log(f"  (Tag: {description})")

        self.toggle_buttons(False)
        self.show_progress(True)
        self.statusBar().showMessage("Saving...")

        self.worker = IconWorker('save', description=description, max_backup_count=cleanup_limit)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_operation_finished)
        self.worker.start()

    def start_restore_latest(self):
        latest_backup_file = self.manager.get_latest_backup_filename()

        if not latest_backup_file:
            QMessageBox.warning(self, "Error", "No backup files found to restore!")
            self.log("✗ Restore failed: No backup files found.")
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
            self, "Confirm Restore",
            f"Restore icon positions from the LATEST backup file:\n\n"
            f"File: {latest_backup_file}\n"
            f"Resolution: {resolution}\n"
            f"Icons: {icon_count}\n"
            f"Tag: {description}\n"
            f"Timestamp: {formatted_date}\n\n"
            f"Are you sure you want to proceed?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self._start_restore(latest_backup_file)

    def _start_restore(self, filename: Optional[str] = None):
        enable_scaling = self.settings.value("adaptive_scaling_enabled", False, type=bool)

        self.log(f"Starting restore from backup '{filename if filename else 'latest'}'...")
        self.toggle_buttons(False)
        self.show_progress(True)
        self.statusBar().showMessage("Restoring...")

        self.worker = IconWorker('restore', filename, enable_scaling=enable_scaling)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_operation_finished)
        self.worker.start()

    # --- START SCRAMBLE FUNCTION ---
    def start_scramble(self):
        reply = QMessageBox.question(
            self, "Confirm Scramble",
            "Are you sure you want to randomize the positions of ALL desktop icons?\n\n"
            "**A mandatory backup will be created first** (tag: Auto-Backup before Scramble).\n\n"
            "Do you want to proceed?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.log("Starting desktop icon scrambling (randomization)...")
            self.toggle_buttons(False)
            self.show_progress(True)
            self.statusBar().showMessage("Scrambling icons...")

            self.worker = IconWorker('scramble')
            self.worker.log_signal.connect(self.log)
            self.worker.progress_signal.connect(self.update_progress)
            self.worker.finished_signal.connect(self.on_operation_finished)
            self.worker.start()
    # --- END START SCRAMBLE FUNCTION ---

    def on_operation_finished(self, success: bool, saved_metadata: Optional[Dict]):
        mode = self.worker.mode if self.worker else "unknown"

        if mode == 'restore' and success:
            self._check_display_metadata(saved_metadata)

        # For 'save' mode, force desktop refresh
        if mode == 'save' and success:
            self.log("Forcing desktop refresh...")
            try:
                win32gui.SendMessage(self.manager.hwnd_listview, win32con.WM_SETREDRAW, 1, 0)
                win32gui.InvalidateRect(self.manager.hwnd_listview, None, True)
                win32api.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, "IconMetrics")
                self.log("Desktop refresh signal sent successfully.")
            except Exception as e:
                self.log(f"Warning: Failed to send desktop refresh signals: {e}")

        self.toggle_buttons(True)
        self.show_progress(False)

        if success:
            self.statusBar().showMessage("Operation completed successfully", 3000)
            if mode != 'save':
                QMessageBox.information(self, "Success", f"Operation completed successfully! ({mode.capitalize()})")
            if not self.isVisible():
                self.tray_icon.showMessage("Desktop Icon Manager", f"{mode.capitalize()} successful!", QSystemTrayIcon.MessageIcon.Information, 2000)
        else:
            self.statusBar().showMessage("Operation failed", 3000)
            QMessageBox.warning(self, "Error", f"Operation failed ({mode.capitalize()}). Check the log for details.")

        self.worker = None

    def _check_display_metadata(self, saved_metadata: Dict):
        current_metadata = get_display_metadata()

        saved_count = saved_metadata.get("monitor_count")
        current_count = current_metadata.get("monitor_count")

        if saved_count is None or current_count is None:
             self.log("⚠ Warning: Display metadata missing or incomplete.")
             return

        # Check 1: Monitor Count
        if saved_count != current_count:
            self.log(f"⚠ Warning: Saved ({saved_count} monitors) vs Current ({current_count} monitors). Icon positions may be inaccurate.")
            QMessageBox.warning(
                self, "Monitor Mismatch Warning",
                f"The layout was saved with {saved_count} monitor(s), but you currently have {current_count} monitor(s) connected.\n\n"
                f"Icon positions have been restored, but they may be inaccurate or off-screen due to the configuration change."
            )
            return

        # Check 2: Individual Screen Resolutions (basic check)
        saved_screens = saved_metadata.get("screens", [])
        current_screens = current_metadata.get("screens", [])

        mismatch_found = False
        if len(saved_screens) == len(current_screens):
            for s_screen, c_screen in zip(saved_screens, current_screens):
                if s_screen.get('width') != c_screen.get('width') or s_screen.get('height') != c_screen.get('height'):
                    mismatch_found = True
                    break

        if mismatch_found:
            self.log("⚠ Warning: Screen resolutions do not match the saved layout. Positions may be inaccurate.")
            QMessageBox.warning(
                self, "Resolution Mismatch Warning",
                f"The screen resolutions for one or more monitors do not match the saved layout.\n\n"
                f"Icon positions have been restored, but they may be inaccurate."
            )

    # --- Application Management Cleanup ---
    def _run_final_cleanup(self):
        # Save geometry
        if self.isVisible():
            self.settings.setValue("geometry", self.geometry())

        # Auto-save
        if self.action_auto_save.isChecked():
            if self.isVisible():
                 self.log("Auto-Save on Exit enabled. Performing silent backup...")
            else:
                 print("Auto-Save on Exit enabled. Performing silent backup...")

            cleanup_limit = self.settings.value("cleanup_limit", 0, type=int)

            self.manager.save(
                lambda msg: print(f"Auto-Save Log: {msg}"),
                description="Auto-Save on Exit",
                max_backup_count=cleanup_limit
            )

            if self.isVisible():
                self.log("Auto-Save complete.")
            else:
                print("Auto-Save complete.")

    def closeEvent(self, event):
        close_to_tray = self.action_close_to_tray.isChecked()
        is_pyinstaller = getattr(sys, 'frozen', False)

        # 1. MINIMIZE TO TRAY PATH
        if close_to_tray and self.isVisible():
            self.settings.setValue("geometry", self.geometry())
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "Desktop Icon Manager",
                "Application minimized to system tray. Click or double-click to restore.",
                QSystemTrayIcon.MessageIcon.Information,
                2000
            )
            return

        # 2. FULL EXIT PATH
        self._run_final_cleanup()

        event.accept()

        if is_pyinstaller:
            try:
                hwnd_console = win32gui.GetConsoleWindow()
                if hwnd_console:
                    win32gui.PostMessage(hwnd_console, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass

        QApplication.quit()

if __name__ == "__main__":
    if QApplication.instance():
        app = QApplication.instance()
    else:
        app = QApplication(sys.argv)

    app.setQuitOnLastWindowClosed(False)
    app.setStyle("Fusion")

    try:
        window = MainWindow()

        if window.settings.value("start_minimized", False, type=bool):
            window.hide()
            window.tray_icon.showMessage(
                "Desktop Icon Manager",
                "Application started minimized to system tray.",
                QSystemTrayIcon.MessageIcon.Information,
                2000
            )
        else:
            window.show()

        sys.exit(app.exec())
    except Exception as e:
        QMessageBox.critical(None, "Critical Error", f"Failed to start application:\n{str(e)}")
        sys.exit(1)
