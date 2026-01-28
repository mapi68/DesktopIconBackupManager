"""Desktop Icon Backup Manager - Main Entry Point"""

import sys
import os
import argparse
from pathlib import Path

from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import QSettings, QTranslator, QLocale, QCoreApplication
from PyQt6.QtGui import QIcon

from config import Config, resource_path
from icon_manager import DesktopIconManager
from utils.helpers import setup_cli_parser
from ui.main_window import MainWindow

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
                max_backup_count=cleanup_limit,
            )
            sys.exit(0 if success else 1)

        elif args.restore:
            enable_scaling = settings.value(
                "adaptive_scaling_enabled", False, type=bool
            )

            filename = None
            if args.restore.lower() == "latest":
                filename = manager.get_latest_backup_filename()
                if not filename:
                    print(
                        QCoreApplication.translate(
                            "CLI", "Error: No backup files found for latest restore."
                        )
                    )
                    sys.exit(1)
            else:
                filename = args.restore

            msg_restore = QCoreApplication.translate(
                "CLI", "Starting silent restore from: %1"
            ).replace("%1", filename)
            print(msg_restore)

            success, _ = manager.restore(
                silent_log, filename=filename, enable_scaling=enable_scaling
            )
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
                Config.TRAY_NOTIFICATION_DURATION,
            )
        else:
            window.show()

        sys.exit(app.exec())
    except Exception as e:
        error_title = QCoreApplication.translate("Main", "Critical Error")
        error_msg = QCoreApplication.translate(
            "Main", "Failed to start application:\n%1"
        ).replace("%1", str(e))
        QMessageBox.critical(None, error_title, error_msg)
        sys.exit(1)
