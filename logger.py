import os
from enum import IntEnum
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()


class LogLevel(IntEnum):
    DEBUG = 1
    INFO = 2
    WARNING = 3
    ERROR = 4


def _get_default_level() -> LogLevel:
    level_name = os.getenv("LOG_LEVEL", "INFO").upper()
    return LogLevel[level_name] if level_name in LogLevel.__members__ else LogLevel.INFO

DEFAULT_LEVEL = _get_default_level()


class Logger:
    def __init__(self, name: str, level: LogLevel = DEFAULT_LEVEL):
        self.name = name
        self.level = level

    def _log(self, level: LogLevel, message: str) -> None:
        if level >= self.level:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}] [{level.name}] [{self.name}] {message}")

    def debug(self, message: str) -> None:
        self._log(LogLevel.DEBUG, message)

    def info(self, message: str) -> None:
        self._log(LogLevel.INFO, message)

    def warning(self, message: str) -> None:
        self._log(LogLevel.WARNING, message)

    def error(self, message: str) -> None:
        self._log(LogLevel.ERROR, message)

    def critical(self, message: str) -> None:
        self._log(LogLevel.CRITICAL, message)

    def set_level(self, level: LogLevel) -> None:
        self.level = level
