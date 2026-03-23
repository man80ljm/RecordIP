import ctypes
import os
import sys
from pathlib import Path

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication

from ui.main_window import MainWindow


def _icon_path() -> Path:
    """返回图标路径，兼容开发环境与 PyInstaller onefile 打包环境。"""
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / "ip.ico"
    return Path(__file__).resolve().parent.parent / "ip.ico"


def _set_windows_app_id() -> None:
    """设置 Windows AppUserModelID，改善任务栏/任务管理器图标识别。"""
    if os.name != "nt":
        return

    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "RecordIP.IPMonitorPyQt6.1"
        )
    except Exception:
        # 失败时不影响主流程。
        pass


def main() -> None:
    """程序入口。"""
    _set_windows_app_id()
    app = QApplication(sys.argv)
    app.setApplicationName("ip_monitor_pyqt6")
    icon_p = _icon_path()
    if icon_p.exists():
        app.setWindowIcon(QIcon(str(icon_p)))

    project_root = Path(__file__).resolve().parent
    window = MainWindow(project_root=project_root)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
