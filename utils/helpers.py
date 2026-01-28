"""Helper functions for Desktop Icon Backup Manager"""

import argparse
from typing import Dict, Tuple, Optional
from datetime import datetime

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QGuiApplication
from PyQt6.QtCore import QCoreApplication


def get_display_metadata() -> Dict:
    """Get metadata about connected displays"""
    app = QApplication.instance()
    if not app:
        return {
            "monitor_count": 0,
            "screens": [],
            "primary_resolution": "UnknownResolution",
        }

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
            }
            for i, s in enumerate(screens)
        ],
    }

    if screens:
        primary_screen = screens[0]
        metadata["primary_resolution"] = (
            f"{primary_screen.geometry().width()}x{primary_screen.geometry().height()}"
        )
    else:
        metadata["primary_resolution"] = "UnknownResolution"
    return metadata


def parse_backup_filename(filename: str) -> Tuple[str, str, str]:
    """
    Parse backup filename to extract date, resolution, and timestamp
    Returns: (readable_date, resolution, timestamp_part)
    """
    try:
        clean_name = filename.replace(".json", "")
        parts = clean_name.split("_")
        resolution = "N/A"
        timestamp_part = clean_name

        if (
            len(parts) >= 3
            and "x" in parts[0]
            and len(parts[1]) == 8
            and len(parts[2]) == 6
        ):
            resolution = parts[0]
            timestamp_part = f"{parts[1]}_{parts[2]}"

        elif len(parts) >= 2 and len(parts[0]) == 8 and len(parts[1]) == 6:
            timestamp_part = f"{parts[0]}_{parts[1]}"

        else:
            try:
                datetime.strptime(timestamp_part, "%Y%m%d_%H%M%S")
            except ValueError:
                return clean_name, "N/A", clean_name

        dt_object = datetime.strptime(timestamp_part, "%Y%m%d_%H%M%S")
        readable_date = dt_object.strftime("%Y/%m/%d %H:%M:%S")
        return readable_date, resolution, timestamp_part

    except Exception:
        return filename.replace(".json", ""), "N/A", filename.replace(".json", "")


def parse_resolution_string(resolution_str: str) -> Optional[Tuple[int, int]]:
    """Parse resolution string like '1920x1080' to tuple (1920, 1080)"""
    try:
        if "x" in resolution_str:
            width, height = map(int, resolution_str.split("x"))
            return width, height
        return None
    except Exception:
        return None


def get_readable_date(filename: str) -> str:
    """Extract readable date from backup filename"""
    return parse_backup_filename(filename)[0]


def get_resolution_from_filename(filename: str) -> str:
    """Extract resolution from backup filename"""
    return parse_backup_filename(filename)[1]


def setup_cli_parser():
    """Setup and parse command line arguments"""
    parser = argparse.ArgumentParser(
        description=QCoreApplication.translate("CLI", "Desktop Icon Backup Manager CLI")
    )
    parser.add_argument(
        "--backup",
        action="store_true",
        help=QCoreApplication.translate("CLI", "Perform a backup"),
    )
    parser.add_argument(
        "--restore",
        type=str,
        metavar="FILENAME",
        help=QCoreApplication.translate("CLI", "Restore a specific backup or latest"),
    )
    parser.add_argument(
        "--silent",
        action="store_true",
        help=QCoreApplication.translate("CLI", "Run without showing the GUI"),
    )

    args, unknown = parser.parse_known_args()
    return args, unknown
