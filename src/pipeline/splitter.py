"""15MB 超ファイルの分割

Dify UI のファイルサイズ上限 (15MB) を超えるファイルを
見出し境界で分割する。
"""

from __future__ import annotations

import re
from logging import getLogger
from pathlib import Path

from src.config import PipelineConfig
from src.models.metadata import ProcessStatus, StepResult

logger = getLogger(__name__)

# 見出しパターン（# で始まる行）
_HEADING_RE = re.compile(r"^(#{1,6})\s+", re.MULTILINE)


def split_if_needed(
    md_path: Path,
    config: PipelineConfig,
) -> list[StepResult]:
    """ファイルサイズが上限を超えている場合に分割する。

    分割は見出し境界で行い、元ファイルは保持する。
    分割されたファイルには _part01, _part02, ... のサフィックスを付与。

    Returns:
        分割結果のリスト（分割不要なら空リスト）
    """
    if not md_path.exists():
        return []

    file_size = md_path.stat().st_size
    if file_size <= config.max_file_size_bytes:
        return []

    logger.info("15MB 超: %s (%d bytes) — 分割開始", md_path.name, file_size)

    text = md_path.read_text(encoding="utf-8")
    sections = _split_by_headings(text)

    if len(sections) <= 1:
        logger.warning("見出しなしのため分割不可: %s", md_path.name)
        return [StepResult(
            file_path=str(md_path), step="split",
            status=ProcessStatus.WARNING,
            message="no headings found, cannot split",
        )]

    # セクションを上限以内のパートにまとめる
    parts = _pack_sections(sections, config.max_file_size_bytes)

    results: list[StepResult] = []
    stem = md_path.stem
    parent = md_path.parent

    for i, part_text in enumerate(parts, start=1):
        part_name = f"{stem}_part{i:02d}.md"
        part_path = parent / part_name
        part_path.write_text(part_text, encoding="utf-8")
        size_kb = len(part_text.encode("utf-8")) / 1024
        logger.info("  分割パート: %s (%.1fKB)", part_name, size_kb)
        results.append(StepResult(
            file_path=str(part_path), step="split",
            status=ProcessStatus.SUCCESS,
            message=f"part {i}/{len(parts)}, size={size_kb:.1f}KB",
        ))

    return results


def _split_by_headings(text: str) -> list[str]:
    """テキストを見出し境界でセクションに分割する。"""
    positions = [m.start() for m in _HEADING_RE.finditer(text)]

    if not positions:
        return [text]

    # 最初の見出しの前にテキストがあれば最初のセクションに含める
    sections: list[str] = []
    if positions[0] > 0:
        sections.append(text[:positions[0]])

    for i, pos in enumerate(positions):
        end = positions[i + 1] if i + 1 < len(positions) else len(text)
        sections.append(text[pos:end])

    return sections


def _pack_sections(
    sections: list[str],
    max_bytes: int,
) -> list[str]:
    """セクションを上限以内のパートにまとめる。"""
    parts: list[str] = []
    current: list[str] = []
    current_size = 0

    for section in sections:
        section_size = len(section.encode("utf-8"))

        if current and current_size + section_size > max_bytes:
            parts.append("".join(current))
            current = []
            current_size = 0

        current.append(section)
        current_size += section_size

    if current:
        parts.append("".join(current))

    return parts
