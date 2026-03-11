"""utils/file_utils.py - ファイル操作ユーティリティ"""

from pathlib import Path
import logging

logger = logging.getLogger(__name__)

FILE_TYPE_MAP: dict[str, str] = {
    ".xlsx": "excel", ".xls": "excel_legacy", ".xlsm": "excel",
    ".docx": "word", ".doc": "word_legacy",
    ".pptx": "pptx", ".ppt": "pptx_legacy",
    ".pdf": "pdf",
    ".rtf": "rtf", ".trf": "trf",
    ".txt": "text", ".csv": "text", ".tsv": "text", ".md": "text", ".log": "text",
    ".bgl": "bagles", ".bag": "bagles", ".def": "bagles",
}

LEGACY_CONVERSIONS: dict[str, str] = {
    "excel_legacy": "excel", "word_legacy": "word",
    "pptx_legacy": "pptx", "rtf": "word", "trf": "word",
}


def detect_file_type(filepath: Path) -> str:
    return FILE_TYPE_MAP.get(filepath.suffix.lower(), "unknown")


def read_text_file(filepath: Path, fallback_encodings: list[str] | None = None) -> str:
    encodings = fallback_encodings or ["utf-8", "shift_jis", "cp932", "euc-jp"]
    for enc in encodings:
        try:
            return filepath.read_text(encoding=enc)
        except (UnicodeDecodeError, LookupError):
            continue
    logger.warning("テキスト読み込み: 全エンコーディング失敗 %s", filepath)
    return filepath.read_text(encoding="utf-8", errors="replace")


def classify_files(directory: Path) -> dict[str, list[Path]]:
    classified: dict[str, list[Path]] = {}
    for f in sorted(directory.rglob("*")):
        if f.is_file() and not f.name.startswith("."):
            ftype = detect_file_type(f)
            classified.setdefault(ftype, []).append(f)
    return classified


def should_process(target: Path, conflict_mode: str) -> bool:
    if not target.exists():
        return True
    if conflict_mode == "overwrite":
        return True
    logger.debug("スキップ: %s (既存)", target.name)
    return False
