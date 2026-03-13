"""Step2: Word 構造抽出器

.docx ファイルから段落・表・図形を抽出し、中間表現 (IntermediateDocument) を構築する。

設計方針:
  - python-docx で確定的に抽出（LLM は使わない）
  - Oasys/Win スタイル対応: 太字が取れないため、フォントサイズ差と
    文字数ヒューリスティクスで疑似見出しを検出
  - 要素の出現順序を文書内の XML 順で保持
  - 変更履歴テーブル検出: 1行目に「ページ」「種別」「年月」「記事」のうち3個以上
"""

from __future__ import annotations

import re
import time
from logging import getLogger
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from src.config import PipelineConfig
from src.models.intermediate import (
    CellData,
    Confidence,
    DocumentElement,
    ElementType,
    IntermediateDocument,
    ShapeElement,
)
from src.models.metadata import (
    ExtractedFileRecord,
    FileMetadata,
    ProcessStatus,
    StepResult,
)

logger = getLogger(__name__)


# ---------------------------------------------------------------------------
# 疑似見出し検出
# ---------------------------------------------------------------------------

def _get_font_size_pt(para: Paragraph) -> float | None:
    """段落の実効フォントサイズ (pt) を取得する。

    Run レベル → 段落スタイル → デフォルトの順で探索。
    """
    # Run レベルで最頻値を取る
    sizes: list[float] = []
    for run in para.runs:
        if run.font.size is not None:
            sizes.append(run.font.size.pt)
    if sizes:
        return max(set(sizes), key=sizes.count)

    # スタイルチェーン
    style = para.style
    while style is not None:
        if style.font and style.font.size is not None:
            return style.font.size.pt
        style = style.base_style
    return None


def _detect_heading(
    para: Paragraph,
    config: PipelineConfig,
) -> tuple[int, str] | None:
    """疑似見出しを検出する。

    Returns:
        (level, detection_method) or None
    """
    text = para.text.strip()
    if not text or len(text) > config.heading_max_chars:
        return None

    # 1. Word 見出しスタイル（Heading 1, Heading 2 等）
    style_name = para.style.name or ""
    if style_name.startswith("Heading"):
        try:
            level = int(style_name.split()[-1])
            return (level, "style")
        except (ValueError, IndexError):
            pass

    # 2. outlineLevel
    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        outline = pPr.find(qn("w:outlineLvl"))
        if outline is not None:
            val = outline.get(qn("w:val"))
            if val is not None:
                level = int(val) + 1  # outlineLevel は 0-indexed
                return (min(level, 6), "outline_level")

    # 3. フォントサイズ差（Oasys/Win 対応）
    size = _get_font_size_pt(para)
    if size is not None and size >= config.heading_font_size_min_pt:
        # サイズに応じてレベルを推定
        if size >= 16.0:
            level = 1
        elif size >= 14.0:
            level = 2
        elif size >= 12.0:
            level = 3
        else:
            level = 4
        return (level, f"font_size:{size}pt")

    # 4. 短文 + 行末句点なし（見出しらしいパターン）
    if len(text) <= 30 and not text.endswith(("。", ".", "、", ",")):
        # ただし数字のみ、空白のみは除外
        if re.search(r"[\u3040-\u9fff\uff01-\uff5ea-zA-Z]", text):
            return (3, "heuristic:short_no_period")

    return None


# ---------------------------------------------------------------------------
# 表の抽出
# ---------------------------------------------------------------------------

def _extract_table(table: Table) -> tuple[list[list[CellData]], bool]:
    """表からセルデータを抽出する。

    Returns:
        (rows, has_merged_cells)
    """
    has_merged = False
    rows: list[list[CellData]] = []

    for r_idx, row in enumerate(table.rows):
        row_data: list[CellData] = []
        for c_idx, cell in enumerate(row.cells):
            # 結合セルの検出（gridSpan / vMerge）
            tc = cell._tc
            rowspan = 1
            colspan = 1

            grid_span = tc.find(qn("w:tcPr"))
            if grid_span is not None:
                gs = grid_span.find(qn("w:gridSpan"))
                if gs is not None:
                    colspan = int(gs.get(qn("w:val"), "1"))
                vm = grid_span.find(qn("w:vMerge"))
                if vm is not None:
                    val = vm.get(qn("w:val"), "")
                    if val != "restart":
                        # 継続セル（上のセルに結合されている）→ スキップ
                        continue

            if colspan > 1 or rowspan > 1:
                has_merged = True

            cell_text = cell.text.strip()
            row_data.append(CellData(
                text=cell_text,
                row=r_idx,
                col=c_idx,
                rowspan=rowspan,
                colspan=colspan,
                is_header=(r_idx == 0),  # 1行目をヘッダーとみなす
            ))

        if row_data:
            rows.append(row_data)

    return rows, has_merged


def _is_change_history_table(
    rows: list[list[CellData]],
    config: PipelineConfig,
) -> bool:
    """変更履歴テーブルかどうかを判定する。

    1行目のセルテキストから全角スペースを除去した上で
    キーワードマッチングを行う。
    """
    if not rows or not rows[0]:
        return False

    matched = 0
    for cell in rows[0]:
        normalized = re.sub(r"[\s\u3000]+", "", cell.text)
        for kw in config.change_history_keywords:
            if kw in normalized:
                matched += 1
                break
    return matched >= config.change_history_min_match


# ---------------------------------------------------------------------------
# 図形の抽出
# ---------------------------------------------------------------------------

def _extract_shapes(doc: Document) -> list[ShapeElement]:
    """文書内の浮動図形・テキストボックスを抽出する。"""
    shapes: list[ShapeElement] = []

    # InlineShapes
    for ishape in doc.inline_shapes:
        shape_type = str(ishape.type) if ishape.type else "unknown"
        shapes.append(ShapeElement(
            shape_type=f"inline:{shape_type}",
            texts=[],
            confidence=Confidence.MEDIUM,
        ))

    # 浮動図形 (w:drawing, mc:AlternateContent 内の図形)
    body = doc.element.body
    for drawing in body.iter(qn("w:drawing")):
        texts = [t.text for t in drawing.iter(qn("a:t")) if t.text]
        shapes.append(ShapeElement(
            shape_type="floating",
            texts=texts,
            confidence=Confidence.MEDIUM if texts else Confidence.LOW,
            fallback_reason="" if texts else "no_text_content",
        ))

    # VML 図形 (w:pict)
    for pict in body.iter(qn("w:pict")):
        texts = []
        # VML textbox 内のテキスト
        for t_elem in pict.iter():
            if t_elem.tag.endswith("}t") or t_elem.tag == "t":
                if t_elem.text:
                    texts.append(t_elem.text)
        shapes.append(ShapeElement(
            shape_type="vml",
            texts=texts,
            confidence=Confidence.MEDIUM if texts else Confidence.LOW,
            fallback_reason="" if texts else "no_text_content",
        ))

    return shapes


# ---------------------------------------------------------------------------
# 文書要素の順序付き抽出
# ---------------------------------------------------------------------------

def _build_element_order(doc: Document) -> list[tuple[str, int]]:
    """文書本文の XML から要素の出現順序を構築する。

    Returns:
        [(element_type, index), ...] — "paragraph" or "table"
    """
    order: list[tuple[str, int]] = []
    para_idx = 0
    table_idx = 0

    for child in doc.element.body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            order.append(("paragraph", para_idx))
            para_idx += 1
        elif tag == "tbl":
            order.append(("table", table_idx))
            table_idx += 1
    return order


# ---------------------------------------------------------------------------
# メインの抽出関数
# ---------------------------------------------------------------------------

def extract_docx(
    docx_path: Path,
    source_path: str,
    source_ext: str,
    config: PipelineConfig,
) -> tuple[ExtractedFileRecord, StepResult]:
    """1つの .docx ファイルから中間表現を抽出する。

    Args:
        docx_path: 正規化済み .docx のパス
        source_path: 元ファイルの相対パス (追跡用)
        source_ext: 元の拡張子
        config: パイプライン設定

    Returns:
        (ExtractedFileRecord, StepResult)
    """
    t0 = time.perf_counter()

    try:
        doc = Document(str(docx_path))
    except Exception as e:
        elapsed = time.perf_counter() - t0
        result = StepResult(
            file_path=source_path, step="extract",
            status=ProcessStatus.ERROR, message=str(e),
            duration_sec=round(elapsed, 2),
        )
        meta = FileMetadata(source_path=source_path, source_ext=source_ext)
        record = ExtractedFileRecord(metadata=meta, document={})
        return record, result

    intermediate = IntermediateDocument()

    # --- 表を先に抽出（テーブルオブジェクト → データ変換）---
    table_data_list: list[tuple[list[list[CellData]], bool, bool]] = []
    for table in doc.tables:
        rows, has_merged = _extract_table(table)
        is_ch = _is_change_history_table(rows, config)
        table_data_list.append((rows, has_merged, is_ch))

    # --- 要素を出現順に中間表現へ追加 ---
    element_order = _build_element_order(doc)
    paragraphs = doc.paragraphs
    source_idx = 0
    table_counter = 0

    for elem_type, idx in element_order:
        if elem_type == "paragraph" and idx < len(paragraphs):
            para = paragraphs[idx]
            text = para.text.strip()

            # 疑似見出し検出
            heading_info = _detect_heading(para, config)
            if heading_info:
                level, method = heading_info
                intermediate.add_heading(level, text, method, source_index=source_idx)
            elif text:
                intermediate.add_paragraph(text, source_index=source_idx)

        elif elem_type == "table" and table_counter < len(table_data_list):
            rows, has_merged, is_ch = table_data_list[table_counter]
            table_counter += 1

            # 表の直前段落をキャプション候補として取得
            caption = ""
            if intermediate.elements:
                last = intermediate.elements[-1]
                if last.type.value == "paragraph" and last.content is not None:
                    candidate = last.content.text  # type: ignore[union-attr]
                    if len(candidate) <= 60:
                        caption = candidate

            confidence = Confidence.HIGH
            fallback_reason = ""
            if has_merged:
                confidence = Confidence.MEDIUM
            if is_ch:
                fallback_reason = "change_history_table"

            intermediate.add_table(
                rows=rows,
                caption=caption,
                has_merged_cells=has_merged,
                confidence=confidence,
                fallback_reason=fallback_reason,
                source_index=source_idx,
            )

        source_idx += 1

    # --- 図形の抽出 ---
    shapes = _extract_shapes(doc)
    for shape in shapes:
        intermediate.elements.append(DocumentElement(
            type=ElementType.SHAPE,
            content=shape,
            source_index=source_idx,
        ))
        source_idx += 1

    # --- 変更履歴テーブル数から doc_role 推定 ---
    total_tables = len(table_data_list)
    ch_count = sum(1 for _, _, is_ch in table_data_list if is_ch)
    if total_tables == 0:
        doc_role = "unknown"
    elif ch_count > 0 and ch_count == total_tables:
        doc_role = "change_history"
    elif ch_count > 0:
        doc_role = "mixed"
    else:
        doc_role = "spec_body"

    # --- メタデータ構築 ---
    meta = FileMetadata(
        source_path=source_path,
        source_ext=source_ext,
        source_size_bytes=docx_path.stat().st_size,
        normalized_from=source_ext if source_ext != ".docx" else "",
        doc_role_guess=doc_role,
    )

    record = ExtractedFileRecord(
        metadata=meta,
        document=intermediate.to_dict(),
    )

    elapsed = time.perf_counter() - t0
    warnings: list[str] = []
    if ch_count > 0:
        warnings.append(f"change_history_tables={ch_count}")
    if shapes:
        warnings.append(f"shapes={len(shapes)}")

    status = ProcessStatus.SUCCESS
    msg = f"elements={len(intermediate.elements)}, tables={total_tables}, doc_role={doc_role}"
    if warnings:
        status = ProcessStatus.WARNING
        msg += ", " + ", ".join(warnings)

    result = StepResult(
        file_path=source_path, step="extract",
        status=status, message=msg,
        duration_sec=round(elapsed, 2),
    )
    logger.info("抽出完了: %s (%s)", source_path, msg)

    return record, result
