"""
utils/file_classifier.py - ファイル種別判定（拡張子 + magic bytes）
"""
import zipfile
from pathlib import Path

EXTENSION_MAP: dict[str, str] = {
    ".xlsx": "excel", ".xls": "excel_legacy", ".xlsm": "excel", ".xlsb": "excel",
    ".docx": "word", ".doc": "word_legacy",
    ".pptx": "pptx", ".ppt": "pptx_legacy",
    ".pdf": "pdf", ".rtf": "rtf", ".trf": "trf",
    ".txt": "text", ".csv": "text", ".tsv": "text", ".md": "text", ".log": "text",
}

LEGACY_TYPES: dict[str, str] = {
    "excel_legacy": "excel", "word_legacy": "word",
    "pptx_legacy": "pptx", "rtf": "word", "trf": "word",
}

def classify_file(filepath: Path) -> str:
    ext = filepath.suffix.lower()
    ftype = EXTENSION_MAP.get(ext, "unknown")
    if ftype == "unknown":
        ftype = _detect_by_magic(filepath)
    return ftype

def classify_directory(directory: Path) -> dict[str, list[Path]]:
    classified: dict[str, list[Path]] = {}
    for f in sorted(directory.rglob("*")):
        if f.is_file() and not f.name.startswith("."):
            classified.setdefault(classify_file(f), []).append(f)
    return classified

def is_legacy_format(file_type: str) -> bool:
    return file_type in LEGACY_TYPES

def get_normalized_type(file_type: str) -> str:
    return LEGACY_TYPES.get(file_type, file_type)

def _detect_by_magic(filepath: Path) -> str:
    try:
        header = filepath.read_bytes()[:16]
    except OSError:
        return "unknown"
    if header[:4] == b"PK\x03\x04":
        return _detect_ooxml(filepath)
    if header[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "word_legacy"
    if header[:5] == b"%PDF-":
        return "pdf"
    if header[:5] == b"{\\rtf":
        return "rtf"
    return "unknown"

def _detect_ooxml(filepath: Path) -> str:
    try:
        with zipfile.ZipFile(filepath) as z:
            names = z.namelist()
            if any("word/" in n for n in names): return "word"
            if any("xl/" in n for n in names): return "excel"
            if any("ppt/" in n for n in names): return "pptx"
    except zipfile.BadZipFile:
        pass
    return "unknown"
