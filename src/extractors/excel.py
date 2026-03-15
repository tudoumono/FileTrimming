"""Step2: Excel 構造抽出器

.xlsx ファイルからシート・表・コメントを抽出し、中間表現 (IntermediateDocument) を構築する。

設計方針:
  - openpyxl で確定的に抽出（LLM は使わない）
  - シート単位でセクション化（シート名 → 見出し）
  - 結合セル対応: openpyxl の merged_cells_ranges から検出
  - コメント: セルテキストに「※注: ...」として付記
  - 非表示シート: スキップし、ログに記録
  - 画像: 存在のみ IMAGE 要素として記録（テキスト抽出不可）
  - 数式: openpyxl は data_only=False で数式文字列、data_only=True でキャッシュ値。
    ここでは data_only=True で値優先、取れない場合は数式文字列をフォールバック
  - 空行・空列の自動トリミング: データ範囲外の空セルは除外
"""

from __future__ import annotations

import time
from logging import getLogger
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.worksheet.worksheet import Worksheet

from src.config import PipelineConfig
from src.models.intermediate import (
    CellData,
    Confidence,
    DocumentElement,
    ElementType,
    ImageElement,
    IntermediateDocument,
)
from src.models.metadata import (
    ExtractedFileRecord,
    FileMetadata,
    ProcessStatus,
    StepResult,
)

logger = getLogger(__name__)


# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------

def _get_data_bounds(ws: Worksheet) -> tuple[int, int, int, int] | None:
    """シートのデータ範囲 (min_row, min_col, max_row, max_col) を返す。

    完全に空のシートなら None を返す。
    openpyxl の ws.max_row/max_col は書式だけのセルも含むため、
    実データのある範囲を走査して特定する。
    """
    min_r = min_c = float("inf")
    max_r = max_c = 0

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                # 結合セルの一部（内容は親セルが持つ）→ 範囲には含める
                if cell.row < min_r:
                    min_r = cell.row
                if cell.row > max_r:
                    max_r = cell.row
                if cell.column < min_c:
                    min_c = cell.column
                if cell.column > max_c:
                    max_c = cell.column
                continue
            if cell.value is not None:
                if cell.row < min_r:
                    min_r = cell.row
                if cell.row > max_r:
                    max_r = cell.row
                if cell.column < min_c:
                    min_c = cell.column
                if cell.column > max_c:
                    max_c = cell.column

    if max_r == 0:
        return None
    return int(min_r), int(min_c), int(max_r), int(max_c)


def _build_merge_map(ws: Worksheet) -> dict[tuple[int, int], tuple[int, int, int, int]]:
    """結合セルのマップを構築する。

    Returns:
        {(row, col): (min_row, min_col, max_row, max_col)} — 結合範囲の左上セルの座標
    """
    merge_map: dict[tuple[int, int], tuple[int, int, int, int]] = {}
    for merge_range in ws.merged_cells.ranges:
        bounds = (
            merge_range.min_row,
            merge_range.min_col,
            merge_range.max_row,
            merge_range.max_col,
        )
        for row in range(merge_range.min_row, merge_range.max_row + 1):
            for col in range(merge_range.min_col, merge_range.max_col + 1):
                merge_map[(row, col)] = bounds
    return merge_map


def _cell_text(cell: Cell, ws: Worksheet) -> str:
    """セルのテキスト表現を取得する。

    - None → ""
    - 数値・日付はそのまま str()
    - コメントがあれば「※注: ...」を付記
    """
    if isinstance(cell, MergedCell):
        return ""

    val = cell.value
    if val is None:
        text = ""
    else:
        text = str(val).strip()

    # コメント付記
    if cell.comment and cell.comment.text:
        comment_text = cell.comment.text.strip()
        if comment_text:
            text = f"{text} ※注: {comment_text}" if text else f"※注: {comment_text}"

    return text


# ---------------------------------------------------------------------------
# シート → 表データ変換
# ---------------------------------------------------------------------------

def _extract_sheet_table(
    ws: Worksheet,
    bounds: tuple[int, int, int, int],
    merge_map: dict[tuple[int, int], tuple[int, int, int, int]],
) -> tuple[list[list[CellData]], bool]:
    """シートのデータ範囲を CellData の2次元リストに変換する。

    Returns:
        (rows, has_merged_cells)
    """
    min_r, min_c, max_r, max_c = bounds
    has_merged = bool(ws.merged_cells.ranges)
    rows: list[list[CellData]] = []

    # 結合セルの左上のみを出力するために、既に出力済みの結合範囲を追跡
    seen_merges: set[tuple[int, int, int, int]] = set()

    for r in range(min_r, max_r + 1):
        row_data: list[CellData] = []
        for c in range(min_c, max_c + 1):
            cell = ws.cell(row=r, column=c)

            # 結合セルの場合
            if (r, c) in merge_map:
                merge_bounds = merge_map[(r, c)]
                m_min_r, m_min_c, m_max_r, m_max_c = merge_bounds

                if (r, c) != (m_min_r, m_min_c):
                    # 結合範囲の左上以外 → スキップ
                    continue

                if merge_bounds in seen_merges:
                    continue
                seen_merges.add(merge_bounds)

                # 左上セルからテキスト取得
                top_left = ws.cell(row=m_min_r, column=m_min_c)
                text = _cell_text(top_left, ws)
                rowspan = m_max_r - m_min_r + 1
                colspan = m_max_c - m_min_c + 1

                row_data.append(CellData(
                    text=text,
                    row=r - min_r,
                    col=c - min_c,
                    rowspan=rowspan,
                    colspan=colspan,
                    is_header=(r == min_r),
                ))
            else:
                text = _cell_text(cell, ws)
                row_data.append(CellData(
                    text=text,
                    row=r - min_r,
                    col=c - min_c,
                    rowspan=1,
                    colspan=1,
                    is_header=(r == min_r),
                ))

        if row_data:
            rows.append(row_data)

    return rows, has_merged


# ---------------------------------------------------------------------------
# 画像検出
# ---------------------------------------------------------------------------

def _count_images(ws: Worksheet) -> int:
    """シート内の画像数を返す。"""
    try:
        return len(ws._images)
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# メイン抽出関数
# ---------------------------------------------------------------------------

def extract_xlsx(
    xlsx_path: Path,
    source_path: str,
    source_ext: str,
    config: PipelineConfig,
) -> tuple[ExtractedFileRecord, StepResult]:
    """1つの .xlsx ファイルから中間表現を抽出する。

    Args:
        xlsx_path: .xlsx ファイルパス（正規化済み、または元ファイル）
        source_path: 元ファイルの相対パス（追跡用）
        source_ext: 元の拡張子
        config: パイプライン設定

    Returns:
        (ExtractedFileRecord, StepResult)
    """
    t0 = time.perf_counter()

    try:
        # data_only=True でキャッシュ値を読む（数式ではなく計算結果）
        wb = load_workbook(str(xlsx_path), data_only=True, read_only=False)
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
    source_idx = 0
    total_sheets = 0
    hidden_sheets = 0
    total_tables = 0
    total_images = 0
    total_merged = 0
    warnings: list[str] = []

    for ws in wb.worksheets:
        # 非表示シートのスキップ
        if ws.sheet_state != "visible":
            hidden_sheets += 1
            logger.debug("非表示シートをスキップ: %s / %s", source_path, ws.title)
            continue

        total_sheets += 1

        # シート名を見出しとして追加
        intermediate.add_heading(
            level=2,
            text=ws.title,
            detection_method="sheet_name",
            source_index=source_idx,
        )
        source_idx += 1

        # 画像の検出
        img_count = _count_images(ws)
        if img_count > 0:
            total_images += img_count
            for i in range(img_count):
                intermediate.elements.append(DocumentElement(
                    type=ElementType.IMAGE,
                    content=ImageElement(
                        alt_text="",
                        description=f"画像 ({ws.title} 内)",
                    ),
                    source_index=source_idx,
                ))
                source_idx += 1

        # データ範囲の特定
        bounds = _get_data_bounds(ws)
        if bounds is None:
            intermediate.add_paragraph(
                "(空のシート)",
                source_index=source_idx,
            )
            source_idx += 1
            continue

        # 結合セルマップの構築
        merge_map = _build_merge_map(ws)

        # シートの表データ抽出
        rows, has_merged = _extract_sheet_table(ws, bounds, merge_map)
        if has_merged:
            total_merged += 1

        if rows:
            total_tables += 1

            # 大きすぎるシートの警告
            num_rows = len(rows)
            num_cols = max(len(r) for r in rows) if rows else 0
            if num_rows > config.excel_large_sheet_rows or num_cols > config.excel_large_sheet_cols:
                warnings.append(
                    f"large_sheet:{ws.title}({num_rows}r×{num_cols}c)"
                )

            confidence = Confidence.HIGH
            fallback_reason = ""
            if has_merged:
                confidence = Confidence.MEDIUM

            intermediate.add_table(
                rows=rows,
                caption=ws.title,
                has_merged_cells=has_merged,
                confidence=confidence,
                fallback_reason=fallback_reason,
                source_index=source_idx,
            )
            source_idx += 1

    wb.close()

    # --- メタデータ構築 ---
    meta = FileMetadata(
        source_path=source_path,
        source_ext=source_ext,
        source_size_bytes=xlsx_path.stat().st_size,
        normalized_from=source_ext if source_ext != ".xlsx" else "",
        doc_role_guess="data_sheet",
    )

    record = ExtractedFileRecord(
        metadata=meta,
        document=intermediate.to_dict(),
    )

    elapsed = time.perf_counter() - t0

    if hidden_sheets > 0:
        warnings.append(f"hidden_sheets={hidden_sheets}")
    if total_images > 0:
        warnings.append(f"images={total_images}")
    if total_merged > 0:
        warnings.append(f"merged_cell_sheets={total_merged}")

    status = ProcessStatus.SUCCESS
    msg = (
        f"sheets={total_sheets}, tables={total_tables}, "
        f"elements={len(intermediate.elements)}"
    )
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
