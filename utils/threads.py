"""Worker threads for asynchronous operations"""

from typing import Optional

import win32con
import win32gui

from PyQt6.QtCore import QThread, pyqtSignal

from icon_manager import DesktopIconManager


class SaveThread(QThread):
    """Thread for asynchronous backup save operation"""

    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool)

    def __init__(self, manager, description: str, max_backup_count: int):
        super().__init__()
        self.manager = manager
        self.description = description
        self.max_backup_count = max_backup_count

    def run(self):
        """Execute save operation"""
        success = self.manager.save(
            self.log_signal.emit, self.description, self.max_backup_count
        )
        self.finished_signal.emit(success)


class RestoreThread(QThread):
    """Thread for asynchronous backup restore operation"""

    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, object)

    def __init__(self, manager, filename: str, enable_scaling: bool):
        super().__init__()
        self.manager = manager
        self.filename = filename
        self.enable_scaling = enable_scaling

    def run(self):
        """Execute restore operation"""
        success, metadata = self.manager.restore(
            self.log_signal.emit, self.filename, self.enable_scaling
        )
        self.finished_signal.emit(success, metadata)


class IconWorker(QThread):
    """Worker thread for icon operations (save/restore/scramble)"""

    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, object)

    def __init__(
        self,
        mode: str,
        filename: Optional[str] = None,
        description: Optional[str] = None,
        max_backup_count: int = 0,
        enable_scaling: bool = False,
    ):
        super().__init__()
        self.mode = mode
        self.filename = filename
        self.description = description
        self.max_backup_count = max_backup_count
        self.enable_scaling = enable_scaling
        self.manager = DesktopIconManager()

    def run(self):
        """Execute the requested operation"""
        success = False
        metadata = None
        try:
            if self.mode == "save":
                success = self.manager.save(
                    self.log_signal.emit,
                    self.progress_signal.emit,
                    self.description,
                    self.max_backup_count,
                )

                if success:
                    self.log_signal.emit(self.tr("Forcing desktop refresh..."))
                    try:
                        win32gui.SendMessage(
                            self.manager.hwnd_listview, win32con.WM_SETREDRAW, 1, 0
                        )
                        win32gui.InvalidateRect(self.manager.hwnd_listview, None, True)
                        win32gui.SendMessageTimeout(
                            win32con.HWND_BROADCAST,
                            win32con.WM_SETTINGCHANGE,
                            0,
                            "IconMetrics",
                            win32con.SMTO_ABORTIFHUNG,
                            5000,
                        )
                        self.log_signal.emit(
                            self.tr("Desktop refresh signal sent successfully.")
                        )
                    except Exception as e:
                        self.log_signal.emit(
                            self.tr(
                                "Warning: Failed to send desktop refresh signals: %1"
                            ).replace("%1", str(e))
                        )
            elif self.mode == "restore":
                success, metadata = self.manager.restore(
                    self.log_signal.emit,
                    self.filename,
                    self.progress_signal.emit,
                    self.enable_scaling,
                )
            elif self.mode == "scramble":
                self.log_signal.emit(
                    self.tr("Performing mandatory quick backup before scrambling...")
                )
                save_success = self.manager.save(
                    lambda msg: self.log_signal.emit(
                        self.tr("  [Pre-Scramble Backup] %1").replace("%1", str(msg))
                    ),
                    lambda val: self.progress_signal.emit(int(val * 0.5)),
                    description=self.tr("Backup before Scramble"),
                    max_backup_count=0,
                )
                if save_success:
                    self.log_signal.emit(
                        self.tr(
                            "Pre-scramble backup completed successfully. Starting scramble..."
                        )
                    )
                    success = self.manager.scramble_icons(
                        self.log_signal.emit,
                        lambda val: self.progress_signal.emit(50 + int(val * 0.5)),
                    )
                else:
                    self.log_signal.emit(
                        self.tr(
                            "✗ Pre-scramble backup failed. Aborting scramble operation."
                        )
                    )
                    success = False

        except Exception as e:
            self.log_signal.emit(
                self.tr("✗ CRITICAL ERROR: %1").replace("%1", str(str(e)))
            )
            success = False
        finally:
            self.finished_signal.emit(success, metadata)
