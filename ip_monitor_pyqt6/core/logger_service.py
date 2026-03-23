from datetime import datetime
from typing import Callable


class LoggerService:
    """简单日志服务，用于向界面输出统一格式日志。"""

    def __init__(self):
        self._listeners: list[Callable[[str], None]] = []

    def add_listener(self, callback: Callable[[str], None]) -> None:
        """注册日志监听回调。"""
        self._listeners.append(callback)

    def _emit(self, level: str, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] [{level}] {message}"
        for callback in self._listeners:
            callback(line)

    def info(self, message: str) -> None:
        self._emit("INFO", message)

    def warning(self, message: str) -> None:
        self._emit("WARN", message)

    def error(self, message: str) -> None:
        self._emit("ERROR", message)
