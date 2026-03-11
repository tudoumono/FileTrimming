"""
logging_utils.py - リアルタイム表示/即時書き込み向けのログ設定
"""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


class ImmediateFileHandler(logging.StreamHandler):
    """各レコードを書いた直後に flush + fsync するハンドラ。"""

    def __init__(self, path: Path):
        self._path = path
        self._path.parent.mkdir(parents=True, exist_ok=True)
        stream = self._path.open("a", encoding="utf-8", buffering=1)
        super().__init__(stream)

    @property
    def path(self) -> Path:
        return self._path

    def flush(self) -> None:
        super().flush()
        if self.stream is None or self.stream.closed:
            return
        try:
            os.fsync(self.stream.fileno())
        except OSError:
            pass


def configure_logging(level: str, log_file: Path | str | None) -> Path | None:
    """標準出力とログファイルの両方へリアルタイム出力する。"""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(line_buffering=True, write_through=True)

    formatter = logging.Formatter(LOG_FORMAT, datefmt=LOG_DATE_FORMAT)
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    log_path = Path(log_file) if log_file else None
    if log_path is not None:
        handlers.append(ImmediateFileHandler(log_path))

    root_logger = logging.getLogger()
    for handler in list(root_logger.handlers):
        root_logger.removeHandler(handler)
        handler.close()

    root_logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    for handler in handlers:
        handler.setFormatter(formatter)
        root_logger.addHandler(handler)

    return log_path
