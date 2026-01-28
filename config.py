"""Configuration and constants for Desktop Icon Backup Manager"""

import os
import sys
import ctypes


class Config:
    """Centralized configuration for the application"""

    # Backup settings
    BACKUP_DIR = "icon_backups"
    MAX_BACKUPS_DEFAULT = 10
    VERSION_FILE = "version.txt"

    # Win32 memory settings
    REMOTE_BUFFER_SIZE = 4096
    TEXT_BUFFER_OFFSET = 256
    TEXT_BUFFER_SIZE = 2048

    # UI dimensions
    PREVIEW_WIDTH = 450
    PREVIEW_HEIGHT = 250
    DIALOG_MIN_WIDTH = 400
    DIALOG_MIN_HEIGHT = 350
    SHORTCUTS_DIALOG_MIN_WIDTH = 400
    SHORTCUTS_DIALOG_MIN_HEIGHT = 350

    # Icon preview
    ICON_DOT_SIZE = 8
    ICON_DOT_MARGIN = 5

    # Colors
    COLOR_BACKGROUND = "#1a1a1a"
    COLOR_BORDER = "#333"
    COLOR_TEXT_DIM = "#666"
    COLOR_ICON_DOT = "#0078d7"
    COLOR_TOOLTIP_BG = "#ffffdc"
    COLOR_TOOLTIP_TEXT = "#000000"

    # Tray notification duration (ms)
    TRAY_NOTIFICATION_DURATION = 2000

    # Application version (will be overridden from file)
    VERSION = "0.0.0"


class Win32Constants:
    """Win32 API constants"""

    LVM_GETITEMCOUNT = 0x1004
    LVM_GETITEMTEXTW = 0x1073
    LVM_GETITEMPOSITION = 0x1010
    LVM_SETITEMPOSITION = 0x100F
    MEM_COMMIT = 0x1000
    MEM_RELEASE = 0x8000
    PAGE_READWRITE = 0x04
    LVS_AUTOARRANGE = 0x0010


class LVITEMW(ctypes.Structure):
    """Win32 LVITEMW structure"""

    _fields_ = [
        ("mask", ctypes.c_uint),
        ("iItem", ctypes.c_int),
        ("iSubItem", ctypes.c_int),
        ("state", ctypes.c_uint),
        ("stateMask", ctypes.c_uint),
        ("pszText", ctypes.c_void_p),
        ("cchTextMax", ctypes.c_int),
        ("iImage", ctypes.c_int),
        ("lParam", ctypes.c_void_p),
    ]


def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# Load version from file
try:
    v_path = resource_path(Config.VERSION_FILE)
    if os.path.exists(v_path):
        with open(v_path, "r", encoding="utf-8") as f:
            Config.VERSION = f.read().strip()
except Exception:
    Config.VERSION = "0.0.0"
