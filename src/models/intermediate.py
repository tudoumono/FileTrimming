"""中間表現データモデル

パイプライン Step2（構造抽出）の出力として使用する。
全ファイル形式で共通の中間表現を定義し、Step3（Markdown 変換）への入力とする。

設計方針:
  - dataclass + asdict() で JSON シリアライズ可能
  - LLM なしでも完全な中間表現が作れる（LLM は品質向上ブースター）
  - 要素の出現順序を保持（元文書の構造再現に必要）
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any


class ElementType(str, Enum):
    """文書要素の種別"""
    HEADING = "heading"
    PARAGRAPH = "paragraph"
    TABLE = "table"
    SHAPE = "shape"           # 浮動図形・テキストボックス・ワークフロー図
    IMAGE = "image"           # インライン画像（テキスト説明に変換）
    PAGE_BREAK = "page_break"


class Confidence(str, Enum):
    """変換品質の信頼度"""
    HIGH = "high"
    MEDIUM = "medium"
    LOW = "low"               # LOW_CONFIDENCE マーカー付与対象


@dataclass
class HeadingElement:
    """見出し要素"""
    level: int                # 1-6 (Markdown の # レベルに対応)
    text: str
    detection_method: str     # "style", "font_size", "heuristic" 等


@dataclass
class ParagraphElement:
    """段落要素"""
    text: str
    is_list_item: bool = False
    list_level: int = 0       # インデントレベル (0 = トップ)


@dataclass
class CellData:
    """表のセルデータ"""
    text: str
    row: int
    col: int
    rowspan: int = 1
    colspan: int = 1
    is_header: bool = False


@dataclass
class TableElement:
    """表要素"""
    rows: list[list[CellData]]
    caption: str = ""         # 表の直前段落から取得した見出し的テキスト
    has_merged_cells: bool = False
    confidence: Confidence = Confidence.HIGH
    fallback_reason: str = ""  # TABLE_FALLBACK 時の理由


@dataclass
class ShapeElement:
    """図形要素（浮動図形・テキストボックス・フロー図）"""
    shape_type: str           # "text_box", "flowchart", "group", "picture" 等
    texts: list[str]          # 図形内のテキスト群
    description: str = ""     # LLM 生成または機械的な説明文
    confidence: Confidence = Confidence.HIGH
    fallback_reason: str = ""


@dataclass
class ImageElement:
    """画像要素"""
    alt_text: str = ""
    description: str = ""     # LLM 生成の説明文
    original_filename: str = ""


@dataclass
class DocumentElement:
    """文書内の1要素（出現順序を保持するためのラッパー）"""
    type: ElementType
    content: HeadingElement | ParagraphElement | TableElement | ShapeElement | ImageElement | None
    source_index: int = 0     # 元文書内での出現位置（デバッグ用）

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"type": self.type.value, "source_index": self.source_index}
        if self.content is not None:
            d["content"] = asdict(self.content)
        return d


@dataclass
class IntermediateDocument:
    """1ファイル分の中間表現

    Step2 の出力、Step3 の入力として使う。
    JSON シリアライズ / デシリアライズをサポートする。
    """
    elements: list[DocumentElement] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "elements": [e.to_dict() for e in self.elements],
        }

    def add_heading(self, level: int, text: str, detection_method: str = "unknown",
                    source_index: int = 0) -> None:
        self.elements.append(DocumentElement(
            type=ElementType.HEADING,
            content=HeadingElement(level=level, text=text, detection_method=detection_method),
            source_index=source_index,
        ))

    def add_paragraph(self, text: str, is_list_item: bool = False, list_level: int = 0,
                      source_index: int = 0) -> None:
        if not text.strip():
            return  # 空段落はスキップ
        self.elements.append(DocumentElement(
            type=ElementType.PARAGRAPH,
            content=ParagraphElement(text=text, is_list_item=is_list_item, list_level=list_level),
            source_index=source_index,
        ))

    def add_table(self, rows: list[list[CellData]], caption: str = "",
                  has_merged_cells: bool = False,
                  confidence: Confidence = Confidence.HIGH,
                  fallback_reason: str = "",
                  source_index: int = 0) -> None:
        self.elements.append(DocumentElement(
            type=ElementType.TABLE,
            content=TableElement(
                rows=rows, caption=caption, has_merged_cells=has_merged_cells,
                confidence=confidence, fallback_reason=fallback_reason,
            ),
            source_index=source_index,
        ))

    def add_shape(self, shape_type: str, texts: list[str], description: str = "",
                  confidence: Confidence = Confidence.HIGH,
                  fallback_reason: str = "",
                  source_index: int = 0) -> None:
        self.elements.append(DocumentElement(
            type=ElementType.SHAPE,
            content=ShapeElement(
                shape_type=shape_type, texts=texts, description=description,
                confidence=confidence, fallback_reason=fallback_reason,
            ),
            source_index=source_index,
        ))
