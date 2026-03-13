"""拡張子 → 抽出器マッピング

拡張子に応じて適切な抽出関数を返す。
Phase 1 では Word 系のみ対応。Excel / BAGLES / テキスト系は将来追加。
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Callable
    from pathlib import Path

    from src.config import PipelineConfig
    from src.models.metadata import ExtractedFileRecord, StepResult

# 抽出関数の型: (docx_path, source_path, source_ext, config) -> (record, result)
ExtractorFunc = "Callable[[Path, str, str, PipelineConfig], tuple[ExtractedFileRecord, StepResult]]"

# 拡張子 → 抽出関数のレジストリ
_REGISTRY: dict[str, str] = {
    # Word 系 (Phase 1) — 正規化後は全て .docx
    ".docx": "src.extractors.word:extract_docx",
    # Excel 系 (将来)
    # ".xlsx": "src.extractors.excel:extract_xlsx",
    # BAGLES 系 (将来)
    # ".bik": "src.extractors.bagles:extract_bagles",
    # テキスト系 (将来)
    # ".txt": "src.extractors.text:extract_text",
}


def get_extractor(ext: str) -> "Callable[[Path, str, str, PipelineConfig], tuple[ExtractedFileRecord, StepResult]] | None":
    """拡張子に対応する抽出関数を返す。未対応なら None。"""
    module_func = _REGISTRY.get(ext.lower())
    if module_func is None:
        return None

    module_path, func_name = module_func.split(":")
    import importlib
    mod = importlib.import_module(module_path)
    return getattr(mod, func_name)


def supported_extensions() -> set[str]:
    """現在対応している拡張子一覧"""
    return set(_REGISTRY.keys())
