"""Step3: 中間表現 → 半構造化 Markdown 変換

設計方針 (Task.md §6 決定事項):
  - 見出し階層は ## / ### で保持
  - 表は項目ラベル付き半構造化テキストに変換（Markdown テーブルではない）
  - 説明文はそのまま残す
  - 図形はテキスト説明に変換（復元困難時はフォールバック）
  - 品質マーカー (LOW_CONFIDENCE 等) は Markdown に埋め込まない
    （Dify がテキストとして扱うためノイズになる。品質情報は中間 JSON に記録済み）
  - YAML front matter は付けない（Dify が認識しないため）
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json
import time
from logging import getLogger
from pathlib import Path
from typing import Any

from src.llm.base import LLMBackend, ReconstructionUnit, TableInterpretationResult
from src.models.metadata import ProcessStatus, StepResult

logger = getLogger(__name__)

TableCell = dict[str, Any]
TableRow = list[TableCell]
TableRows = list[TableRow]
ExpandedRow = list[tuple[str, int]]

_LLM_MAX_CANDIDATE_ROWS = 200
_LLM_MAX_CANDIDATE_TEXT_CELLS = 2000
_CHECKBOX_PREFIXES = ("□", "■", "☑", "☐", "✓")
_FORM_GRID_ROW_KINDS = {
    "parallel_labels",
    "field_pairs",
    "check_item",
    "section_header",
    "banner",
    "text",
}
_DATA_TABLE_ROW_KINDS = {
    "data_record",
    "parallel_labels",
    "field_pairs",
    "check_item",
    "section_header",
    "banner",
    "text",
    "skip",
}
_ROW_RENDER_KINDS = _FORM_GRID_ROW_KINDS | _DATA_TABLE_ROW_KINDS


@dataclass(frozen=True)
class _PreparedTable:
    """描画前に正規化した表データ。

    Step3 の初期整理として、行の rowspan 展開と列数推定を
    ひとまとまりの準備処理として扱う。
    """
    rows: TableRows
    total_cols: int


@dataclass(frozen=True)
class _TableAnalysis:
    """LLMなし解釈に渡す表の分析結果。"""
    rows: TableRows
    total_cols: int
    caption: str
    has_merged_cells: bool
    extract_confidence: str
    extract_fallback_reason: str
    source_row_start: int | None
    source_col_start: int | None
    source_row_end: int | None
    source_col_end: int | None


@dataclass(frozen=True)
class _TableProfile:
    """LLM 適用可否のための表プロファイル。"""
    non_empty_row_count: int
    text_cell_count: int
    max_text_cells_per_row: int
    merged_text_cell_count: int
    max_text_colspan: int


@dataclass(frozen=True)
class _TableInterpretation:
    """LLMなしモードで得た表の解釈結果。"""
    render_kind: str
    labels: list[str]
    data_start: int
    header_found: bool
    active_cols: list[int]
    row_role_overrides: dict[int, str] = field(default_factory=dict)
    summary_labels: list[str] = field(default_factory=list)
    markdown_lines: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class _RenderContext:
    """表ごとの再構成コンテキスト。"""
    source_path: str
    source_ext: str
    doc_role_guess: str
    current_sheet_name: str
    heading_context: tuple[str, ...]
    table_index: int
    previous_table_context: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class _TableObservation:
    """LLM 解釈の観察用レコード。"""
    unit: dict[str, Any]
    fallback_result: dict[str, Any]
    llm_result: dict[str, Any] | None
    applied_result: dict[str, Any]
    used_for_rendering: bool
    selection_reason: str = ""
    error: str = ""


def _render_heading(content: dict[str, Any]) -> str:
    level = min(content.get("level", 3), 6)
    text = content.get("text", "")
    return f"{'#' * level} {text}"


def _render_paragraph(content: dict[str, Any]) -> str:
    text = content.get("text", "")
    if content.get("is_list_item"):
        indent = "  " * content.get("list_level", 0)
        return f"{indent}- {text}"
    return text


def _fill_rowspan(rows: TableRows) -> TableRows:
    """rowspan でカバーされている列の値を後続行に展開する。

    Excel の縦結合セル（rowspan > 1）は、元の行にのみセルが存在し、
    結合先の行にはセルが含まれない。この関数はそれを補完し、
    各行が全列分のセルを持つようにする。

    例: 「入力系」rowspan=4 → 行1〜4 全てに「入力系」が col=0 として出現
    """
    # rowspan フィールドを持つセルがなければ何もしない
    has_rowspan = any(
        cell.get("rowspan", 1) > 1
        for row in rows
        for cell in row
    )
    if not has_rowspan:
        return rows

    # アクティブな rowspan を追跡: {col: (cell_data, remaining_rows)}
    active_spans: dict[int, tuple[TableCell, int]] = {}
    result: TableRows = []

    for row in rows:
        # 元の行が完全に空かどうか判定（スペーサー行検出）
        row_originally_empty = all(not cell.get("text", "") for cell in row)

        # 現在の行に存在する列を記録
        cols_in_row: set[int] = set()
        for cell in row:
            col = cell.get("col", -1)
            cs = cell.get("colspan", 1)
            for c in range(col, col + cs):
                cols_in_row.add(c)

        # アクティブな rowspan から、この行にないセルを補完
        # ただし元々空のスペーサー行には伝播しない
        new_row = list(row)
        for col, (span_cell, remaining) in list(active_spans.items()):
            if remaining > 0:
                if not row_originally_empty and col not in cols_in_row:
                    new_row.append({
                        "text": span_cell.get("text", ""),
                        "col": col,
                        "colspan": span_cell.get("colspan", 1),
                        "rowspan": 1,
                        "is_header": span_cell.get("is_header", False),
                    })
                active_spans[col] = (span_cell, remaining - 1)
            elif remaining <= 0:
                del active_spans[col]

        # 現在の行の rowspan > 1 のセルを追跡開始
        for cell in row:
            rs = cell.get("rowspan", 1)
            if rs > 1:
                col = cell.get("col", 0)
                active_spans[col] = (cell, rs - 1)

        # col でソートして行の順序を保持
        new_row.sort(key=lambda c: c.get("col", 0))
        result.append(new_row)

    return result


def _expand_row_to_positions(row: TableRow) -> ExpandedRow:
    """行のセルを列位置に展開する。

    セルの `col` フィールドがある場合はそれを使って正しい列位置に配置する。
    rowspan で上の行がカバーしている列がある場合、行のセルは途中の列から
    始まるため、col フィールドなしでは位置がずれる。

    Returns:
        [(text, colspan), ...] — 各列位置のテキストと元の colspan
    """
    # col フィールドがあるか確認
    has_col = any("col" in cell for cell in row)

    if has_col:
        # col フィールドを使って正しい列位置に配置
        # まず必要な配列サイズを計算
        max_pos = 0
        for cell in row:
            col = cell.get("col", 0)
            cs = cell.get("colspan", 1)
            end = col + cs
            if end > max_pos:
                max_pos = end

        positions: ExpandedRow = [("", 0)] * max_pos
        for cell in row:
            text = cell.get("text", "")
            col = cell.get("col", 0)
            cs = cell.get("colspan", 1)
            if col < max_pos:
                positions[col] = (text, cs)
                for offset in range(1, cs):
                    if col + offset < max_pos:
                        positions[col + offset] = (text, 0)
        return positions
    else:
        # col フィールドなし: 従来のシーケンシャル展開
        positions = []
        for cell in row:
            text = cell.get("text", "")
            cs = cell.get("colspan", 1)
            positions.append((text, cs))
            for _ in range(cs - 1):
                positions.append((text, 0))
        return positions


def _is_empty_row(row: TableRow) -> bool:
    """行が完全に空（全セルのテキストが空）かどうか判定する。"""
    return all(not cell.get("text", "") for cell in row)


def _text_cells(row: TableRow) -> TableRow:
    """空文字でないセルだけを返す。"""
    return [cell for cell in row if cell.get("text", "")]


def _prepare_table(rows: TableRows) -> _PreparedTable:
    """表描画の前処理をまとめて行う。"""
    normalized_rows = _fill_rowspan(rows)
    return _PreparedTable(
        rows=normalized_rows,
        total_cols=_estimate_total_cols(normalized_rows),
    )


def _analyze_table(content: dict[str, Any]) -> _TableAnalysis:
    """レンダリング前の表準備と基本情報抽出を行う。"""
    prepared = _prepare_table(content.get("rows", []))
    return _TableAnalysis(
        rows=prepared.rows,
        total_cols=prepared.total_cols,
        caption=content.get("caption", ""),
        has_merged_cells=bool(content.get("has_merged_cells", False)),
        extract_confidence=str(content.get("confidence", "")),
        extract_fallback_reason=str(content.get("fallback_reason", "")),
        source_row_start=content.get("source_row_start"),
        source_col_start=content.get("source_col_start"),
        source_row_end=content.get("source_row_end"),
        source_col_end=content.get("source_col_end"),
    )


def _build_table_profile(analysis: _TableAnalysis) -> _TableProfile:
    """表の形状を簡易に集計する。"""
    non_empty_rows = [row for row in analysis.rows if not _is_empty_row(row)]
    text_rows = [_text_cells(row) for row in non_empty_rows]
    text_cell_count = sum(len(row) for row in text_rows)
    merged_text_cell_count = sum(
        1 for row in text_rows for cell in row
        if cell.get("colspan", 1) > 1
    )
    max_text_colspan = max(
        (cell.get("colspan", 1) for row in text_rows for cell in row),
        default=0,
    )
    max_text_cells_per_row = max((len(row) for row in text_rows), default=0)
    return _TableProfile(
        non_empty_row_count=len(non_empty_rows),
        text_cell_count=text_cell_count,
        max_text_cells_per_row=max_text_cells_per_row,
        merged_text_cell_count=merged_text_cell_count,
        max_text_colspan=max_text_colspan,
    )


def _build_reconstruction_unit(
    analysis: _TableAnalysis, render_context: _RenderContext,
) -> ReconstructionUnit:
    """現在の表を LLM 共通契約の入力へ変換する。"""
    table_name = f"table_{render_context.table_index:04d}"
    sheet_name = render_context.current_sheet_name or "document"
    heading_context = [h for h in render_context.heading_context if h]
    profile = _build_table_profile(analysis)
    hints: dict[str, Any] = {}
    if render_context.doc_role_guess:
        hints["doc_role_guess"] = render_context.doc_role_guess
    if analysis.has_merged_cells:
        hints["has_merged_cells"] = True
    if analysis.extract_confidence:
        hints["extract_confidence"] = analysis.extract_confidence
    if analysis.extract_fallback_reason:
        hints["extract_fallback_reason"] = analysis.extract_fallback_reason

    return ReconstructionUnit(
        schema_version="1.0",
        unit_id=f"{render_context.source_path}::{sheet_name}::{table_name}",
        source_path=render_context.source_path,
        source_ext=render_context.source_ext,
        sheet_name=sheet_name,
        table_caption=analysis.caption,
        rows=analysis.rows,
        context={
            "nearby_headings": heading_context,
            "table_shape": {
                "row_count": len(analysis.rows),
                "total_cols": analysis.total_cols,
            },
            "table_profile": {
                "non_empty_row_count": profile.non_empty_row_count,
                "text_cell_count": profile.text_cell_count,
                "max_text_cells_per_row": profile.max_text_cells_per_row,
                "merged_text_cell_count": profile.merged_text_cell_count,
                "max_text_colspan": profile.max_text_colspan,
            },
            "source_bounds": {
                "row_start": analysis.source_row_start,
                "col_start": analysis.source_col_start,
                "row_end": analysis.source_row_end,
                "col_end": analysis.source_col_end,
            },
            "previous_table": render_context.previous_table_context,
        },
        hints=hints,
    )


def _extract_following_table_context(content: dict[str, Any]) -> dict[str, Any]:
    """後続テーブルが参照できる最小限の前後文脈を作る。"""
    analysis = _analyze_table(content)
    labels, _, header_found = _build_column_labels(analysis.rows, analysis.total_cols)
    column_labels_by_col: list[dict[str, Any]] = []
    if labels and analysis.source_col_start is not None:
        for idx, label in enumerate(labels):
            if label:
                column_labels_by_col.append({
                    "col": analysis.source_col_start + idx,
                    "label": label,
                })

    return {
        "caption": analysis.caption,
        "header_found": header_found,
        "column_labels_by_col": column_labels_by_col,
        "source_bounds": {
            "row_start": analysis.source_row_start,
            "col_start": analysis.source_col_start,
            "row_end": analysis.source_row_end,
            "col_end": analysis.source_col_end,
        },
    }


def _interpretation_to_result(
    unit_id: str, interpretation: _TableInterpretation,
) -> TableInterpretationResult:
    """内部解釈結果を LLM 共通契約へ写像する。"""
    table_type = {
        "form_grid": "form",
        "key_value": "key_value",
        "data_table": "data_table",
    }.get(interpretation.render_kind, "unknown")
    render_plan: dict[str, Any] = {}
    if interpretation.row_role_overrides:
        max_idx = max(interpretation.row_role_overrides)
        row_roles = [""] * (max_idx + 1)
        for idx, role in interpretation.row_role_overrides.items():
            if 0 <= idx < len(row_roles):
                row_roles[idx] = role
        render_plan["row_roles"] = row_roles
    if interpretation.summary_labels:
        render_plan["summary_labels"] = interpretation.summary_labels
    if interpretation.markdown_lines:
        render_plan["markdown_lines"] = interpretation.markdown_lines

    return TableInterpretationResult(
        schema_version="1.0",
        unit_id=unit_id,
        table_type=table_type,
        render_strategy=interpretation.render_kind,
        header_rows=[0] if interpretation.header_found else [],
        data_start_row=interpretation.data_start,
        column_labels=interpretation.labels,
        active_columns=interpretation.active_cols,
        render_plan=render_plan,
    )


def _result_to_interpretation(
    result: TableInterpretationResult, analysis: _TableAnalysis,
) -> _TableInterpretation:
    """LLM 共通契約の結果を内部描画用の解釈へ変換する。"""
    render_kind = result.render_strategy or result.table_type or "data_table"
    if render_kind == "form":
        render_kind = "form_grid"
    if render_kind not in {"form_grid", "key_value", "data_table"}:
        logger.warning(
            "未知の render_strategy を受信したため data_table へフォールバック: unit_id=%s, render_strategy=%s",
            result.unit_id,
            render_kind,
        )
        render_kind = "data_table"

    active_cols = [
        col for col in result.active_columns
        if isinstance(col, int) and 0 <= col < analysis.total_cols
    ]
    data_start = result.data_start_row
    if not isinstance(data_start, int) or data_start < 0 or data_start > len(analysis.rows):
        logger.warning(
            "不正な data_start_row を受信したため 0 へ補正: unit_id=%s, data_start_row=%s",
            result.unit_id,
            data_start,
        )
        data_start = 0

    labels = [label for label in result.column_labels if label]
    header_found = bool(result.header_rows)
    row_role_overrides: dict[int, str] = {}
    raw_row_roles = result.render_plan.get("row_roles")
    if isinstance(raw_row_roles, list):
        for idx, role in enumerate(raw_row_roles):
            if isinstance(role, str) and role in _ROW_RENDER_KINDS and idx < len(analysis.rows):
                row_role_overrides[idx] = role
    summary_labels: list[str] = []
    raw_summary_labels = result.render_plan.get("summary_labels")
    if isinstance(raw_summary_labels, list):
        summary_labels = [label for label in raw_summary_labels if isinstance(label, str) and label]
    markdown_lines: list[str] = []
    raw_markdown_lines = result.render_plan.get("markdown_lines")
    if isinstance(raw_markdown_lines, list):
        markdown_lines = [
            line.rstrip()
            for line in raw_markdown_lines
            if isinstance(line, str)
        ]
    return _TableInterpretation(
        render_kind=render_kind,
        labels=labels,
        data_start=data_start,
        header_found=header_found,
        active_cols=active_cols,
        row_role_overrides=row_role_overrides,
        summary_labels=summary_labels,
        markdown_lines=markdown_lines,
    )


def _sanitize_form_grid_row_role_overrides(
    analysis: _TableAnalysis,
    row_role_overrides: dict[int, str],
) -> dict[int, str]:
    """deterministic 判定を壊しにくい形で form_grid の上書きを制限する。"""
    sanitized: dict[int, str] = {}
    for idx, role in row_role_overrides.items():
        if idx < 0 or idx >= len(analysis.rows):
            continue
        text_cells = _text_cells(analysis.rows[idx])
        baseline = _classify_form_grid_row(analysis.rows[idx], analysis.total_cols)

        # 明確に判定できる行は基本的に deterministic 判定を優先する。
        if baseline in {"parallel_labels", "check_item", "banner"}:
            continue
        if role == "parallel_labels" and len(text_cells) < 3:
            continue
        if baseline == "field_pairs" and role in {"section_header", "banner"}:
            continue

        if role != baseline:
            sanitized[idx] = role
    return sanitized


def _sanitize_data_table_row_role_overrides(
    analysis: _TableAnalysis,
    row_role_overrides: dict[int, str],
    data_start: int,
) -> dict[int, str]:
    """data_table の行上書きを安全側で制限する。"""
    sanitized: dict[int, str] = {}
    for idx, role in row_role_overrides.items():
        if idx < 0 or idx >= len(analysis.rows):
            continue
        if role not in _DATA_TABLE_ROW_KINDS:
            continue

        row = analysis.rows[idx]
        if _is_empty_row(row):
            continue

        if idx < data_start:
            if role in {"section_header", "banner", "field_pairs", "check_item", "parallel_labels", "text"}:
                sanitized[idx] = role
            continue

        # ヘッダー行相当を強制的に壊さない
        if idx == data_start - 1 and role != "data_record":
            continue

        if role in {"section_header", "banner", "field_pairs", "check_item", "parallel_labels", "text", "skip", "data_record"}:
            sanitized[idx] = role

    return sanitized


def _get_llm_confidence(result: TableInterpretationResult) -> str:
    confidence = result.self_assessment.get("confidence", "")
    if confidence in {"high", "medium", "low"}:
        return confidence
    return ""


def _is_llm_table_too_large(profile: _TableProfile) -> bool:
    return (
        profile.non_empty_row_count > _LLM_MAX_CANDIDATE_ROWS
        or profile.text_cell_count > _LLM_MAX_CANDIDATE_TEXT_CELLS
    )


def _is_summary_header_only_table(analysis: _TableAnalysis) -> bool:
    return len(analysis.rows) == 1 and _is_summary_header_only_row(
        analysis.rows[0],
        analysis.total_cols,
    )


def _should_request_llm_interpretation(
    analysis: _TableAnalysis,
    fallback: _TableInterpretation,
    profile: _TableProfile,
    has_previous_table_context: bool,
) -> tuple[bool, str]:
    """LLM 解釈を試す価値がある表かを事前判定する。"""
    if _is_llm_table_too_large(profile):
        return False, "fallback: table_too_large_for_llm"
    if profile.non_empty_row_count == 1 and profile.text_cell_count == 1:
        return False, "fallback: trivial_banner_table"
    if _is_summary_header_only_table(analysis):
        if has_previous_table_context:
            return True, "llm_requested: summary_header_table_candidate"
        return False, "fallback: summary_without_context"
    if fallback.render_kind not in {"data_table", "form_grid"}:
        return False, "fallback: existing_non_data_strategy"
    if not analysis.has_merged_cells:
        return False, "fallback: no_merged_cells"
    if not _looks_like_small_merged_form(analysis, profile):
        return False, "fallback: pattern_not_target"
    return True, "llm_requested: small_merged_form_candidate"


def _looks_like_small_merged_form(
    analysis: _TableAnalysis, profile: _TableProfile,
) -> bool:
    if analysis.total_cols < 4:
        return False
    if profile.non_empty_row_count == 0 or profile.non_empty_row_count > 12:
        return False
    if profile.max_text_cells_per_row == 0 or profile.max_text_cells_per_row > 4:
        return False
    if profile.text_cell_count == 0:
        return False
    merged_ratio = profile.merged_text_cell_count / profile.text_cell_count
    dominant_span = profile.max_text_colspan >= max(2, int(analysis.total_cols * 0.5))
    return merged_ratio >= 0.5 or dominant_span


def _infer_key_value_active_cols(analysis: _TableAnalysis) -> list[int]:
    """KV 型に使う実列位置を表から決定論的に推定する。"""
    for row in analysis.rows:
        if _is_empty_row(row):
            continue
        positions = _expand_row_to_positions(row)
        cols = [i for i, (text, cs) in enumerate(positions) if text and cs > 0]
        if len(cols) >= 2:
            return cols[:2]
    return []


def _sanitize_summary_labels(
    analysis: _TableAnalysis,
    summary_labels: list[str],
) -> list[str]:
    """summary header only table に対する安全な summary_labels だけ採用する。"""
    if not _is_summary_header_only_table(analysis):
        return []
    if not analysis.rows:
        return []

    value_count = max(0, len(_text_cells(analysis.rows[0])) - 1)
    cleaned = [label.strip() for label in summary_labels if label.strip()]
    if value_count == 0 or len(cleaned) != value_count:
        return []
    return cleaned


def _sanitize_markdown_lines(
    analysis: _TableAnalysis,
    markdown_lines: list[str],
) -> list[str]:
    """LLM 生成の Markdown 行を安全側で整える。"""
    if not markdown_lines:
        return []

    cleaned: list[str] = []
    max_lines = max(8, len(analysis.rows) * 6)
    for line in markdown_lines:
        if len(cleaned) >= max_lines:
            break
        text = line.rstrip()
        if text.startswith("```"):
            return []
        cleaned.append(text)

    while cleaned and not cleaned[-1]:
        cleaned.pop()

    visible_source_texts = {
        cell.get("text", "").strip()
        for row in analysis.rows
        for cell in row
        if cell.get("text", "").strip()
    }
    joined = "\n".join(cleaned)
    if not any(text in joined for text in visible_source_texts):
        return []
    if len(visible_source_texts) <= 12 and not all(
        text in joined for text in visible_source_texts
    ):
        return []

    return cleaned


def _render_text_row(row: TableRow) -> list[str]:
    texts = [cell.get("text", "") for cell in _text_cells(row) if cell.get("text", "")]
    if not texts:
        return []
    if len(texts) == 1:
        return texts
    return [" | ".join(texts)]


def _render_row_by_kind(row: TableRow, total_cols: int, row_kind: str) -> list[str]:
    if row_kind == "skip":
        return []
    if row_kind == "banner":
        banner_text = row[0].get("text", "") if row else ""
        return [f"**{banner_text}**"] if banner_text else []
    if row_kind == "section_header":
        text = _get_section_header_text(row, total_cols)
        return [f"**{text}**"] if text else []
    if row_kind == "parallel_labels":
        return _render_parallel_label_row(row)
    if row_kind in {"check_item", "field_pairs"}:
        return _render_form_field_row(row)
    if row_kind == "text":
        return _render_text_row(row)
    return []


def _derive_summary_labels_from_previous_table(
    analysis: _TableAnalysis,
    previous_table_context: dict[str, Any] | None,
) -> list[str]:
    """直前テーブルの列ラベルから summary 行の値ラベル候補を引く。"""
    if not previous_table_context or not _is_summary_header_only_table(analysis):
        return []
    if not analysis.rows or analysis.source_col_start is None:
        return []

    label_by_col = {
        item.get("col"): item.get("label", "")
        for item in previous_table_context.get("column_labels_by_col", [])
        if isinstance(item, dict) and item.get("label")
    }
    if not label_by_col:
        return []

    row = analysis.rows[0]
    labels: list[str] = []
    for cell in _text_cells(row)[1:]:
        absolute_col = analysis.source_col_start + int(cell.get("col", 0))
        label = label_by_col.get(absolute_col, "")
        if label:
            labels.append(label)
    return labels


def _select_interpretation_with_llm(
    analysis: _TableAnalysis,
    fallback: _TableInterpretation,
    result: TableInterpretationResult,
    profile: _TableProfile,
    summary_label_candidates: list[str] | None = None,
) -> tuple[_TableInterpretation, str]:
    """LLM 結果をどこまで採用するかを安全側で決める。"""
    confidence = _get_llm_confidence(result)
    if confidence == "low":
        return fallback, "fallback: low_confidence"

    is_summary_candidate = _is_summary_header_only_table(analysis)
    llm_interpretation = _result_to_interpretation(result, analysis)
    if not is_summary_candidate and not _looks_like_small_merged_form(analysis, profile):
        return fallback, "fallback: pattern_not_target"

    if not is_summary_candidate:
        markdown_lines = _sanitize_markdown_lines(
            analysis,
            llm_interpretation.markdown_lines,
        )
        if markdown_lines:
            return _TableInterpretation(
                render_kind=fallback.render_kind,
                labels=fallback.labels,
                data_start=fallback.data_start,
                header_found=fallback.header_found,
                active_cols=fallback.active_cols,
                row_role_overrides=fallback.row_role_overrides,
                summary_labels=fallback.summary_labels,
                markdown_lines=markdown_lines,
            ), "llm_adopted: markdown_lines"

    if fallback.render_kind == "form_grid":
        row_role_overrides = _sanitize_form_grid_row_role_overrides(
            analysis,
            llm_interpretation.row_role_overrides,
        )
        if llm_interpretation.render_kind == "form_grid" and row_role_overrides:
            return _TableInterpretation(
                render_kind=fallback.render_kind,
                labels=fallback.labels,
                data_start=fallback.data_start,
                header_found=fallback.header_found,
                active_cols=fallback.active_cols,
                row_role_overrides=row_role_overrides,
                summary_labels=fallback.summary_labels,
                markdown_lines=fallback.markdown_lines,
            ), "llm_adopted: form_grid_plan"
        return fallback, "fallback: no_effective_render_plan"

    if fallback.render_kind != "data_table":
        return fallback, "fallback: existing_non_data_strategy"

    data_table_row_overrides = _sanitize_data_table_row_role_overrides(
        analysis,
        llm_interpretation.row_role_overrides,
        fallback.data_start,
    )

    summary_labels = _sanitize_summary_labels(
        analysis,
        llm_interpretation.summary_labels,
    )
    if (
        not summary_labels
        and llm_interpretation.summary_labels
        and summary_label_candidates is not None
    ):
        summary_labels = _sanitize_summary_labels(
            analysis,
            summary_label_candidates,
        )
    if (
        not summary_labels
        and is_summary_candidate
        and summary_label_candidates is not None
        and llm_interpretation.render_kind in {"data_table", "key_value"}
    ):
        summary_labels = _sanitize_summary_labels(
            analysis,
            summary_label_candidates,
        )
    if llm_interpretation.render_kind == "data_table" and summary_labels:
        return _TableInterpretation(
            render_kind=fallback.render_kind,
            labels=fallback.labels,
            data_start=fallback.data_start,
            header_found=fallback.header_found,
            active_cols=fallback.active_cols,
            row_role_overrides=data_table_row_overrides,
            summary_labels=summary_labels,
            markdown_lines=fallback.markdown_lines,
        ), "llm_adopted: summary_labels"

    if summary_labels or data_table_row_overrides:
        return _TableInterpretation(
            render_kind=fallback.render_kind,
            labels=fallback.labels,
            data_start=fallback.data_start,
            header_found=fallback.header_found,
            active_cols=fallback.active_cols,
            row_role_overrides=data_table_row_overrides,
            summary_labels=summary_labels,
            markdown_lines=fallback.markdown_lines,
        ), "llm_adopted: data_table_plan"

    if result.render_strategy == "form_grid":
        row_role_overrides = _sanitize_form_grid_row_role_overrides(
            analysis,
            llm_interpretation.row_role_overrides,
        )
        return _TableInterpretation(
            render_kind="form_grid",
            labels=[],
            data_start=0,
            header_found=False,
            active_cols=[],
            row_role_overrides=row_role_overrides,
            summary_labels=[],
            markdown_lines=markdown_lines,
        ), "llm_adopted: form_grid"

    if result.render_strategy == "key_value":
        active_cols = _infer_key_value_active_cols(analysis)
        if len(active_cols) < 2:
            return fallback, "fallback: key_value_columns_not_found"
        return _TableInterpretation(
            render_kind="key_value",
            labels=[],
            data_start=0,
            header_found=False,
            active_cols=active_cols,
            summary_labels=[],
            markdown_lines=markdown_lines,
        ), "llm_adopted: key_value"

    return fallback, "fallback: data_table_or_unsupported"


def _is_form_field_row(row: TableRow, total_cols: int) -> bool:
    """行がフォームフィールド（ラベル-値ペア）行か判定する。

    フォーム型 Excel では「項目名 [colspan=N] + 値 [colspan=M]」のように
    少数のセルが大きく結合されている行がある。これはヘッダー行ではない。

    判定基準:
      - 総列数が 3 以上（2列テーブルはラベル+値が通常の構造）
      - テキストのあるセル数が総列数の 3/4 未満
      - テキストのある全セルが colspan >= 2（セル結合によるレイアウト）
        ※端の空セル（cs=1）はレイアウトの残骸なので無視
    """
    if not row or total_cols <= 2:
        return False
    # テキストのあるセルのみで判定（空セルは無視）
    text_cells = _text_cells(row)
    if not text_cells:
        return False
    if len(text_cells) >= total_cols * 3 / 4:
        return False
    return all(cell.get("colspan", 1) >= 2 for cell in text_cells)


def _estimate_total_cols(rows: TableRows) -> int:
    """テーブル全体の列数を推定する（最大展開幅）。"""
    max_cols = 0
    for row in rows:
        positions = _expand_row_to_positions(row)
        if len(positions) > max_cols:
            max_cols = len(positions)
    return max_cols


def _render_form_field_row(row: TableRow) -> list[str]:
    """フォームフィールド行をラベル-値ペアとして出力する。

    セルの並び方に応じて自動判定:
      - 2セル: 単一のラベル-値ペア
      - 偶数セル: ラベル-値ペアの繰り返し
      - 奇数セル: 最後のセルは単独出力
      - 1セル: そのまま出力
    """
    cells = _text_cells(row)
    lines: list[str] = []

    if len(cells) == 0:
        return lines
    elif len(cells) == 1:
        lines.append(cells[0]["text"])
    elif len(cells) == 2:
        lines.append(f"{cells[0]['text']}: {cells[1]['text']}")
    else:
        # 複数ペア: 交互にラベル-値
        for j in range(0, len(cells) - 1, 2):
            label_text = cells[j].get("text", "")
            value_text = cells[j + 1].get("text", "") if j + 1 < len(cells) else ""
            if label_text and value_text:
                lines.append(f"{label_text}: {value_text}")
            elif label_text:
                lines.append(label_text)
        if len(cells) % 2 == 1:
            lines.append(cells[-1].get("text", ""))

    return lines


def _is_parallel_label_row(row: TableRow, total_cols: int) -> bool:
    """行が並列見出し行か判定する。

    例:
      設備購入稟議書 | 担当課長 | 部長 | 役員

    この種の行はフォーム型テーブルの一部だが、交互のラベル-値ペアではない。
    先頭に大きなタイトルセル、その後ろに承認欄などの短い見出しが並ぶ形を
    優先的に検出する。
    """
    if total_cols <= 4:
        return False

    cells = _text_cells(row)
    if len(cells) < 3:
        return False

    if any(cell.get("colspan", 1) < 2 for cell in cells):
        return False

    widest_colspan = max(cell.get("colspan", 1) for cell in cells)
    if widest_colspan <= total_cols / 2:
        return False

    # 交互のラベル-値ペアで出したい行よりも、短い見出しが横並びになる行を優先。
    short_cell_count = sum(1 for cell in cells if len(cell.get("text", "").strip()) <= 12)
    return short_cell_count == len(cells)


def _render_parallel_label_row(row: TableRow) -> list[str]:
    """並列見出し行を横並びのテキストとして出力する。"""
    cells = _text_cells(row)
    if not cells:
        return []
    return [" | ".join(cell.get("text", "") for cell in cells if cell.get("text", ""))]


def _is_checkbox_field_row(row: TableRow) -> bool:
    """行がチェック項目形式か判定する。"""
    cells = _text_cells(row)
    if len(cells) != 2:
        return False
    value_text = cells[1].get("text", "").strip()
    return any(value_text.startswith(prefix) for prefix in _CHECKBOX_PREFIXES)


def _is_two_cell_merged_field_row(row: TableRow, total_cols: int) -> bool:
    """2セルの結合フィールド行か判定する。

    例:
      件名 | 受注 CSV 取込レイアウト変更
      起案理由 | 長文説明...
    """
    if total_cols <= 2:
        return False

    cells = _text_cells(row)
    if len(cells) != 2:
        return False

    label_cell, value_cell = cells
    label_text = label_cell.get("text", "").strip()
    value_text = value_cell.get("text", "").strip()
    label_span = label_cell.get("colspan", 1)
    value_span = value_cell.get("colspan", 1)

    if not label_text or not value_text:
        return False
    if label_span < 1 or value_span < 2:
        return False
    if label_span + value_span < total_cols * 0.7:
        return False
    if len(label_text) > 24:
        return False
    return value_span > label_span


def _classify_form_grid_row(row: TableRow, total_cols: int) -> str:
    """フォーム型テーブル内の行種別を判定する。"""
    if _is_empty_row(row):
        return "empty"

    text_cells = _text_cells(row)
    positions = _expand_row_to_positions(row)
    if _is_banner_row(row, len(positions)):
        return "banner"
    if _is_checkbox_field_row(row):
        return "check_item"
    if len(text_cells) == 2 and _is_form_field_row(row, total_cols):
        return "field_pairs"
    if _is_section_header_row(row, total_cols):
        return "section_header"
    if _is_parallel_label_row(row, total_cols):
        return "parallel_labels"
    if _is_form_field_row(row, total_cols):
        return "field_pairs"
    return "text"


def _render_form_grid_row(
    row: TableRow,
    total_cols: int,
    row_kind_override: str = "",
) -> list[str]:
    """フォーム型テーブルの1行を行種別に応じて描画する。"""
    row_kind = (
        row_kind_override
        if row_kind_override in _FORM_GRID_ROW_KINDS else _classify_form_grid_row(row, total_cols)
    )

    if row_kind == "empty":
        return []
    if row_kind == "banner":
        banner_text = row[0].get("text", "")
        return [f"**{banner_text}**"] if banner_text else []
    if row_kind == "section_header":
        text = _get_section_header_text(row, total_cols)
        return [f"**{text}**"] if text else []
    if row_kind == "parallel_labels":
        return _render_parallel_label_row(row)
    if row_kind in {"check_item", "field_pairs"}:
        return _render_form_field_row(row)

    text_cells = _text_cells(row)
    return [cell.get("text", "") for cell in text_cells if cell.get("text", "")]


def _is_summary_header_only_row(row: TableRow, total_cols: int) -> bool:
    """1行だけの集計表に見えるか判定する。"""
    if total_cols <= 4:
        return False

    cells = _text_cells(row)
    if len(cells) < 4:
        return False

    first = cells[0]
    if first.get("colspan", 1) < 2:
        return False

    return all(cell.get("colspan", 1) == 1 for cell in cells[1:])


def _render_summary_header_only_row(
    row: TableRow,
    summary_labels: list[str] | None = None,
) -> str:
    """1行だけの集計表を重複なしで描画する。"""
    cells = _text_cells(row)
    if not cells:
        return ""

    label = cells[0].get("text", "")
    values = [cell.get("text", "") for cell in cells[1:] if cell.get("text", "")]
    if not values:
        return label
    if summary_labels and len(summary_labels) == len(values):
        lines = [label]
        for value_label, value in zip(summary_labels, values):
            lines.append(f"  {value_label}: {value}")
        return "\n".join(lines)
    return f"{label}: {' | '.join(values)}"


def _is_form_grid_table(rows: TableRows, total_cols: int) -> bool:
    """テーブル全体がフォーム型（ヘッダー行なし）か判定する。

    判定基準: 非空・非バナー・非セクション見出しの全行がフォームフィールド行であればフォーム型。
    データテーブル型なら少なくとも1行はヘッダー候補（セル数 ≈ 列数）がある。
    """
    if total_cols <= 2:
        return False

    content_rows = 0
    form_rows = 0
    for row in rows:
        if _is_empty_row(row):
            continue
        positions = _expand_row_to_positions(row)
        if _is_banner_row(row, len(positions)):
            continue
        if _is_section_header_row(row, total_cols):
            continue
        content_rows += 1
        if _is_form_field_row(row, total_cols):
            form_rows += 1

    return content_rows > 0 and form_rows == content_rows


def _looks_like_merged_field_pairs_table(rows: TableRows, total_cols: int) -> bool:
    """分離抽出された merged field row 群を deterministic に form_grid 扱いする。"""
    if total_cols <= 2:
        return False

    content_rows = 0
    field_like_rows = 0
    for row in rows:
        if _is_empty_row(row):
            continue
        positions = _expand_row_to_positions(row)
        if _is_banner_row(row, len(positions)):
            return False

        content_rows += 1
        if (
            _is_two_cell_merged_field_row(row, total_cols)
            or _is_checkbox_field_row(row)
            or _is_parallel_label_row(row, total_cols)
            or _is_form_field_row(row, total_cols)
        ):
            field_like_rows += 1
            continue
        return False

    return content_rows > 0 and content_rows == field_like_rows


def _should_skip_as_header(row: TableRow, total_cols: int) -> bool:
    """ヘッダー候補から除外すべき行か判定する。"""
    if _is_empty_row(row):
        return True
    positions = _expand_row_to_positions(row)
    if _is_banner_row(row, len(positions)):
        return True
    if _is_section_header_row(row, total_cols):
        return True
    if _is_form_field_row(row, total_cols):
        return True
    return False


def _find_header_row(
    rows: TableRows, total_cols: int,
) -> tuple[int, ExpandedRow, bool]:
    """先頭のバナー行・空行・セクション見出し・フォーム行をスキップしてヘッダー行を見つける。

    Returns:
        (header_idx, header_positions, found): ヘッダー行のインデックス、
        展開済み列位置、ヘッダーが見つかったかどうか
    """
    for i, row in enumerate(rows):
        if _should_skip_as_header(row, total_cols):
            continue
        return i, _expand_row_to_positions(row), True
    return -1, [], False


def _build_column_labels(
    rows: TableRows, total_cols: int,
) -> tuple[list[str], int, bool]:
    """ヘッダー行からカラム位置ベースのラベルを構築する。

    多段ヘッダー対応:
      - row[0] に colspan > 1 のセルがあり、row[1] がサブヘッダーに見える場合、
        「親ラベル/子ラベル」形式で結合する。
      - 先頭のバナー行（全列スパンのタイトル行）と空行はスキップする。

    Returns:
        (labels, data_start_idx, header_found): ラベルリスト、データ開始行、
        ヘッダーが見つかったかどうか
    """
    if not rows:
        return [], 0, False

    header_idx, header_positions, found = _find_header_row(rows, total_cols)
    if not found:
        return [], 0, False

    hdr_cols = len(header_positions)
    labels = [t or f"列{i+1}" for i, (t, _) in enumerate(header_positions)]

    data_start = header_idx + 1
    has_parent_colspan = any(cell.get("colspan", 1) > 1 for cell in rows[header_idx])
    has_parent_rowspan = any(cell.get("rowspan", 1) > 1 for cell in rows[header_idx])

    if (has_parent_colspan or has_parent_rowspan) and header_idx + 1 < len(rows):
        row_next = rows[header_idx + 1]
        row_next_positions = _expand_row_to_positions(row_next)

        # サブヘッダー判定:
        #   - 展開後の列数が一致
        #   - 行全体がバナーでない
        #   - colspan型、または rowspan型で親子の差分が1列以上ある
        changed_columns = sum(
            1
            for (parent, _), (child, _) in zip(header_positions, row_next_positions)
            if child and child != parent
        )
        if (
            len(row_next_positions) == hdr_cols
            and not _is_banner_row(row_next, hdr_cols)
            and changed_columns > 0
        ):
            combined: list[str] = []
            for i, ((parent, _), (child, _)) in enumerate(
                zip(header_positions, row_next_positions)
            ):
                if parent == child or not child:
                    combined.append(parent or f"列{i+1}")
                elif not parent:
                    combined.append(child)
                else:
                    combined.append(f"{parent}/{child}")
            labels = combined
            data_start = header_idx + 2

    return labels, data_start, True


def _is_banner_row(row: TableRow, total_cols: int) -> bool:
    """行が全列スパンのバナー行（セクション区切り等）か判定する。"""
    if len(row) == 1 and row[0].get("colspan", 1) >= total_cols:
        return True
    # 全セルが同一テキストの場合もバナー扱い（横結合の残骸対策）
    if len(row) > 1:
        texts = [c.get("text", "") for c in row]
        if texts and all(t == texts[0] for t in texts) and texts[0]:
            return True
    return False


def _is_section_header_row(row: TableRow, total_cols: int) -> bool:
    """行がセクション見出し行か判定する。

    1つのセルが列数の 2/3 超を占める場合、セクション区切りと見なす。
    例: 「■ 売上データ」cs=8 + 「問い合わせ履歴」cs=1（11列テーブル）
    バナー行（全列スパン）との違い: 端に小さなセルが付いている場合にも対応。

    閾値 2/3: 合計行等の通常の結合セル（cs=2 in 3列 = 67%）を除外しつつ、
    セクション見出し（cs=8 in 11列 = 73%）を検出する。
    """
    if total_cols <= 4:
        return False
    threshold = total_cols * 2 / 3
    for cell in row:
        if cell.get("text", "") and cell.get("colspan", 1) > threshold:
            return True
    return False


def _get_section_header_text(row: TableRow, total_cols: int) -> str:
    """セクション見出し行からテキストを取得する。支配的セルのテキストを返す。"""
    parts = []
    for cell in row:
        text = cell.get("text", "")
        if text:
            parts.append(text)
    return " / ".join(parts) if parts else ""


def _render_form_grid(
    rows: TableRows,
    total_cols: int,
    row_role_overrides: dict[int, str] | None = None,
) -> str:
    """フォーム型テーブルを全行ラベル-値ペアとして出力する。

    フォーム型: ヘッダー行が存在せず、全行がラベル-値ペアの結合セルで構成される。
    業務申請書、稟議書、設定シート等でよく見られるレイアウト。
    """
    lines: list[str] = []
    for idx, row in enumerate(rows):
        row_kind_override = ""
        if row_role_overrides is not None:
            row_kind_override = row_role_overrides.get(idx, "")
        field_lines = _render_form_grid_row(row, total_cols, row_kind_override=row_kind_override)
        lines.extend(field_lines)
        if field_lines:
            lines.append("")
    return "\n".join(lines)


def _render_pre_header_rows(
    rows: TableRows,
    data_start: int,
    total_cols: int,
    row_role_overrides: dict[int, str] | None = None,
) -> list[str]:
    """ヘッダーより前の行を出力する（バナー→太字、フォーム→ラベル: 値）。"""
    lines: list[str] = []
    for idx, row in enumerate(rows[:data_start]):
        if _is_empty_row(row):
            continue

        row_kind_override = ""
        if row_role_overrides is not None:
            row_kind_override = row_role_overrides.get(idx, "")
        if row_kind_override in _DATA_TABLE_ROW_KINDS:
            rendered = _render_row_by_kind(row, total_cols, row_kind_override)
            lines.extend(rendered)
            if rendered:
                lines.append("")
            continue

        positions = _expand_row_to_positions(row)
        if _is_banner_row(row, len(positions)) or _is_section_header_row(row, total_cols):
            text = _get_section_header_text(row, total_cols)
            if text:
                lines.append(f"**{text}**")
                lines.append("")
        elif _is_form_field_row(row, total_cols):
            field_lines = _render_form_field_row(row)
            lines.extend(field_lines)
            if field_lines:
                lines.append("")
    return lines


def _detect_active_columns(
    rows: TableRows, data_start: int, total_cols: int,
) -> list[int]:
    """データ行で一貫してデータが入っている列を検出する。

    「アクティブ」= データ行の 50% 以上で非空のセルがある列。
    キーバリュー型テーブル（広い表だが2列程度しか使われていない）の検出に使用。
    """
    col_fill_count = [0] * total_cols
    data_row_count = 0
    for row in rows[data_start:]:
        if _is_empty_row(row):
            continue
        positions = _expand_row_to_positions(row)
        if _is_banner_row(row, len(positions)):
            continue
        if _is_section_header_row(row, total_cols):
            continue
        data_row_count += 1
        for i, (text, cs) in enumerate(positions):
            if text and cs > 0 and i < total_cols:
                col_fill_count[i] += 1
    if data_row_count == 0:
        return list(range(total_cols))
    threshold = data_row_count * 0.5
    return [i for i, count in enumerate(col_fill_count) if count >= threshold]


def _should_render_as_key_value(
    header_found: bool, active_cols: list[int], total_cols: int,
) -> bool:
    """現在のヒューリスティクスに基づき KV 型として扱うか判定する。"""
    return header_found and len(active_cols) >= 2 and len(active_cols) <= total_cols // 2


def _looks_like_key_value_memo_table(analysis: _TableAnalysis) -> bool:
    """2列の部門メモ・補足表のような KV 型を決定論的に検出する。"""
    if analysis.total_cols < 4:
        return False

    content_rows = 0
    for row in analysis.rows:
        if _is_empty_row(row):
            continue
        text_cells = _text_cells(row)
        if len(text_cells) != 2:
            return False

        label_cell, value_cell = text_cells
        label_text = label_cell.get("text", "").strip()
        value_text = value_cell.get("text", "").strip()
        label_span = label_cell.get("colspan", 1)
        value_span = value_cell.get("colspan", 1)

        if not label_text or not value_text:
            return False
        if label_span != 1:
            return False
        if value_span < max(2, analysis.total_cols - 1):
            return False
        if len(label_text) > 20:
            return False

        content_rows += 1

    return 0 < content_rows <= 12


def _interpret_table_no_llm(analysis: _TableAnalysis) -> _TableInterpretation:
    """既存ヒューリスティクスで表の出力戦略を決める。"""
    if _looks_like_key_value_memo_table(analysis):
        active_cols = _infer_key_value_active_cols(analysis)
        return _TableInterpretation(
            render_kind="key_value",
            labels=[],
            data_start=0,
            header_found=False,
            active_cols=active_cols,
        )

    if _looks_like_merged_field_pairs_table(analysis.rows, analysis.total_cols):
        return _TableInterpretation(
            render_kind="form_grid",
            labels=[],
            data_start=0,
            header_found=False,
            active_cols=[],
        )

    if _is_form_grid_table(analysis.rows, analysis.total_cols):
        return _TableInterpretation(
            render_kind="form_grid",
            labels=[],
            data_start=0,
            header_found=False,
            active_cols=[],
        )

    labels, data_start, header_found = _build_column_labels(analysis.rows, analysis.total_cols)
    active_cols = _detect_active_columns(analysis.rows, data_start, analysis.total_cols)
    render_kind = "key_value" if _should_render_as_key_value(
        header_found, active_cols, analysis.total_cols,
    ) else "data_table"
    return _TableInterpretation(
        render_kind=render_kind,
        labels=labels,
        data_start=data_start,
        header_found=header_found,
        active_cols=active_cols,
    )


def _interpret_table(
    analysis: _TableAnalysis,
    render_context: _RenderContext | None = None,
    backend: LLMBackend | None = None,
    observation_only: bool = False,
) -> tuple[_TableInterpretation, _TableObservation | None]:
    """LLM あり/なしを吸収して表の解釈結果を得る。"""
    fallback = _interpret_table_no_llm(analysis)
    profile = _build_table_profile(analysis)
    fallback_observation_result: dict[str, Any] | None = None
    if backend is None or not backend.supports_table_interpretation():
        return fallback, None
    if render_context is None:
        logger.warning("render_context がないため LLM 解釈をスキップします。")
        return fallback, None

    unit = _build_reconstruction_unit(analysis, render_context)
    summary_label_candidates = _derive_summary_labels_from_previous_table(
        analysis,
        render_context.previous_table_context,
    )
    fallback_observation_result = _interpretation_to_result(unit.unit_id, fallback).to_dict()
    should_request_llm, skip_reason = _should_request_llm_interpretation(
        analysis,
        fallback,
        profile,
        bool(render_context.previous_table_context.get("column_labels_by_col")),
    )
    if not should_request_llm:
        observation = _TableObservation(
            unit=unit.to_dict(),
            fallback_result=fallback_observation_result,
            llm_result=None,
            applied_result=fallback_observation_result,
            used_for_rendering=False,
            selection_reason=skip_reason,
            error=f"skipped: {skip_reason}",
        )
        return fallback, observation
    try:
        result = backend.interpret_table(unit)
    except Exception as exc:
        logger.warning(
            "LLM テーブル解釈に失敗したため既存解釈へフォールバック: unit_id=%s, error=%s",
            unit.unit_id,
            exc,
        )
        observation = _TableObservation(
            unit=unit.to_dict(),
            fallback_result=fallback_observation_result,
            llm_result=None,
            applied_result=fallback_observation_result,
            used_for_rendering=False,
            selection_reason="fallback: llm_error",
            error=str(exc),
        )
        return fallback, observation

    if result.unit_id and result.unit_id != unit.unit_id:
        logger.warning(
            "LLM 解釈結果の unit_id が入力と一致しません: expected=%s, actual=%s",
            unit.unit_id,
            result.unit_id,
        )

    selected, selection_reason = _select_interpretation_with_llm(
        analysis,
        fallback,
        result,
        profile,
        summary_label_candidates=summary_label_candidates,
    )
    selected_result = _interpretation_to_result(unit.unit_id, selected).to_dict()
    if observation_only:
        applied = fallback
        used_for_rendering = False
        applied_result = selected_result
    else:
        applied = selected
        used_for_rendering = (
            selected.render_kind != fallback.render_kind
            or selected.data_start != fallback.data_start
            or selected.labels != fallback.labels
            or selected.active_cols != fallback.active_cols
            or selected.row_role_overrides != fallback.row_role_overrides
            or selected.summary_labels != fallback.summary_labels
            or selected.markdown_lines != fallback.markdown_lines
        )
        applied_result = selected_result

    observation = _TableObservation(
        unit=unit.to_dict(),
        fallback_result=fallback_observation_result,
        llm_result=result.to_dict(),
        applied_result=applied_result,
        used_for_rendering=used_for_rendering,
        selection_reason=selection_reason,
    )
    return applied, observation


def _render_key_value_table(
    rows: TableRows,
    labels: list[str],
    data_start: int,
    total_cols: int,
    active_cols: list[int],
) -> str:
    """キーバリュー型テーブルを出力する。

    広い表（6列等）だが実際にデータがあるのは2列程度のパターン。
    第1アクティブ列の値をキー、第2アクティブ列の値をバリューとして出力する。
    それ以外の列にデータがある場合はヘッダーラベル付きで追加出力する。
    """
    lines: list[str] = []

    # ヘッダーより前の行を出力
    lines.extend(_render_pre_header_rows(rows, data_start, total_cols))

    key_col = active_cols[0]
    val_col = active_cols[1]

    for row in rows[data_start:]:
        if _is_empty_row(row):
            continue

        positions = _expand_row_to_positions(row)
        text_cells = _text_cells(row)

        # 2セル前後のフォーム行は、セクション見出し判定より KV として扱う。
        key = positions[key_col][0] if key_col < len(positions) else ""
        val = positions[val_col][0] if val_col < len(positions) else ""
        if len(text_cells) <= 2 and key and val:
            lines.append(f"{key}: {val}")
            lines.append("")
            continue

        if _is_banner_row(row, len(positions)):
            banner_text = row[0].get("text", "")
            if banner_text:
                lines.append(f"**{banner_text}**")
                lines.append("")
            continue

        if _is_section_header_row(row, total_cols):
            text = _get_section_header_text(row, total_cols)
            if text:
                lines.append(f"**{text}**")
                lines.append("")
            continue

        if key and val:
            lines.append(f"{key}: {val}")
        elif key:
            lines.append(key)
        elif val:
            lines.append(val)
        else:
            continue

        # アクティブ2列以外の非空列をヘッダーラベル付きで追加
        for i, (text, cs) in enumerate(positions):
            if i in (key_col, val_col) or not text or cs == 0:
                continue
            label = labels[i] if i < len(labels) else f"列{i+1}"
            lines.append(f"  {label}: {text}")

        lines.append("")

    return "\n".join(lines)


def _render_data_table(
    rows: TableRows,
    labels: list[str],
    data_start: int,
    total_cols: int,
    summary_labels: list[str] | None = None,
    row_role_overrides: dict[int, str] | None = None,
) -> str:
    """データテーブル型を出力する。

    対応するレイアウト:
      - ヘッダー前のフォーム行（混在型: 請求書の上部にフォーム部分）
      - セクション分割テーブル（バナー行の後に新ヘッダーが出現 → ラベルを再構築）
    """
    lines: list[str] = []

    # ヘッダーより前の行を出力
    lines.extend(_render_pre_header_rows(
        rows,
        data_start,
        total_cols,
        row_role_overrides=row_role_overrides,
    ))

    # データ行を出力（セクション分割対応）
    display_row_num = 1
    i = data_start
    while i < len(rows):
        row = rows[i]

        if _is_empty_row(row):
            i += 1
            continue

        positions = _expand_row_to_positions(row)
        row_kind_override = ""
        if row_role_overrides is not None:
            row_kind_override = row_role_overrides.get(i, "")

        # セクション見出し行 → 見出し出力 + 次行のヘッダー再検出
        if row_kind_override == "section_header" or (
            not row_kind_override and _is_section_header_row(row, total_cols)
        ):
            text = _get_section_header_text(row, total_cols)
            if text:
                lines.append(f"**{text}**")
                lines.append("")

            # セクション見出し後の次の非空行がヘッダー候補か確認
            j = i + 1
            while j < len(rows) and _is_empty_row(rows[j]):
                j += 1
            if j < len(rows) and not _should_skip_as_header(rows[j], total_cols):
                # 新しいヘッダー行を検出 → ラベルを再構築
                new_positions = _expand_row_to_positions(rows[j])
                labels = [t or f"列{k+1}" for k, (t, _) in enumerate(new_positions)]
                display_row_num = 1
                i = j + 1  # ヘッダー行をスキップ
                continue

            i += 1
            continue

        # バナー行 → 太字出力のみ（ヘッダー再検出しない）
        if row_kind_override == "banner" or (
            not row_kind_override and _is_banner_row(row, len(positions))
        ):
            banner_text = row[0].get("text", "")
            if banner_text:
                lines.append(f"**{banner_text}**")
                lines.append("")
            i += 1
            continue

        if row_kind_override == "skip":
            i += 1
            continue

        # フォームフィールド行がデータ部分に混在する場合
        if row_kind_override in {"field_pairs", "check_item", "parallel_labels", "text"}:
            rendered = _render_row_by_kind(row, total_cols, row_kind_override)
            lines.extend(rendered)
            if rendered:
                lines.append("")
            i += 1
            continue

        if not row_kind_override and _is_form_field_row(row, total_cols):
            field_lines = _render_form_field_row(row)
            lines.extend(field_lines)
            if field_lines:
                lines.append("")
            i += 1
            continue

        lines.append(f"[行{display_row_num}]")

        # 行のセルを列位置に展開
        row_positions = _expand_row_to_positions(row)
        for pos_idx, (value, cs) in enumerate(row_positions):
            if cs == 0:
                continue
            label = labels[pos_idx] if pos_idx < len(labels) else f"列{pos_idx+1}"
            if value:
                lines.append(f"  {label}: {value}")

        lines.append("")
        display_row_num += 1
        i += 1

    # データ行がない場合（ヘッダーのみ）
    if data_start >= len(rows):
        if rows and _is_summary_header_only_row(rows[0], total_cols):
            summary_line = _render_summary_header_only_row(
                rows[0],
                summary_labels=summary_labels,
            )
            if summary_line:
                lines.append(summary_line)
        else:
            lines.append("  " + " | ".join(labels))
        lines.append("")

    return "\n".join(lines)


def _render_table(analysis: _TableAnalysis, interpretation: _TableInterpretation) -> str:
    """分析結果と解釈結果に基づいて表を描画する。"""
    lines: list[str] = []

    if analysis.caption:
        lines.append(f"**{analysis.caption}**")
        lines.append("")

    if interpretation.markdown_lines:
        lines.extend(interpretation.markdown_lines)
        return "\n".join(lines)

    if interpretation.render_kind == "form_grid":
        lines.append(_render_form_grid(
            analysis.rows,
            analysis.total_cols,
            row_role_overrides=interpretation.row_role_overrides,
        ))
    elif interpretation.render_kind == "key_value":
        lines.append(_render_key_value_table(
            analysis.rows,
            interpretation.labels,
            interpretation.data_start,
            analysis.total_cols,
            interpretation.active_cols,
        ))
    else:
        effective_cols = len(interpretation.labels) if interpretation.labels else analysis.total_cols
        lines.append(_render_data_table(
            analysis.rows,
            interpretation.labels,
            interpretation.data_start,
            effective_cols,
            summary_labels=interpretation.summary_labels,
            row_role_overrides=interpretation.row_role_overrides,
        ))

    return "\n".join(lines)


def _render_table_as_labeled_text(
    content: dict[str, Any],
    render_context: _RenderContext | None = None,
    backend: LLMBackend | None = None,
    observation_only: bool = False,
    observation_records: list[dict[str, Any]] | None = None,
) -> str:
    """表を項目ラベル付き半構造化テキストに変換する。

    Task.md §6 の決定事項:
    「表は Markdown テーブルではなく項目ラベル付き半構造化テキストに変換して渡す。
     行列の意味や制約・対応関係を壊さないことを優先」

    テーブル型の自動判定:
      1. フォーム型: 全行がラベル-値ペア（全セル colspan >= 2）→ 全行をペア出力
      2. データテーブル型: ヘッダー行あり → ヘッダー + ラベル付きデータ行
      3. 混在型: ヘッダー前にフォーム行、以降データ行
    """
    if not content.get("rows", []):
        return ""

    analysis = _analyze_table(content)
    interpretation, observation = _interpret_table(
        analysis,
        render_context=render_context,
        backend=backend,
        observation_only=observation_only,
    )
    if observation is not None and observation_records is not None:
        observation_records.append({
            "table_index": render_context.table_index if render_context is not None else -1,
            "caption": analysis.caption,
            "record": {
                "unit": observation.unit,
                "fallback_result": observation.fallback_result,
                "llm_result": observation.llm_result,
                "applied_result": observation.applied_result,
                "used_for_rendering": observation.used_for_rendering,
                "selection_reason": observation.selection_reason,
                "error": observation.error,
            },
            "decision": {
                "selection_reason": observation.selection_reason,
                "used_for_rendering": observation.used_for_rendering,
                "fallback_render_strategy": observation.fallback_result.get("render_strategy", ""),
                "llm_render_strategy": (
                    observation.llm_result.get("render_strategy", "")
                    if observation.llm_result is not None else ""
                ),
                "applied_render_strategy": observation.applied_result.get("render_strategy", ""),
            },
        })
    return _render_table(analysis, interpretation)


_SHAPE_TYPE_LABEL: dict[str, str] = {
    "vml_textbox": "テキストボックス",
    "vml_rect": "矩形オブジェクト",
    "vml": "図形",
    "floating": "図形",
    "workflow": "フロー図",
}


def _render_shape(content: dict[str, Any]) -> str:
    """図形をテキスト説明に変換する。

    テキストなし矩形 (vml_rect) はオーバーレイパターンで suppressed 済みだが、
    残存した場合も出力しない（ノイズになるだけのため）。
    """
    texts = content.get("texts", [])
    description = content.get("description", "")
    shape_type = content.get("shape_type", "")

    # テキストなし矩形オブジェクトはスキップ
    if shape_type == "vml_rect" and not texts and not description:
        return ""

    label = _SHAPE_TYPE_LABEL.get(shape_type, "図形")
    lines: list[str] = []

    if description:
        lines.append(description)
    elif shape_type == "workflow" and texts:
        lines.append(f"[{label}]")
        for idx, text in enumerate(texts, 1):
            lines.append(f"  {idx}. {text}")
    elif texts:
        lines.append(f"[{label}]")
        for t in texts:
            for part in t.splitlines():
                if part.strip():
                    lines.append(f"  - {part.strip()}")
    else:
        if label == "図形" and shape_type:
            lines.append(f"[図形: {shape_type}]")
        else:
            lines.append(f"[{label}]")

    return "\n".join(lines)


def transform_to_markdown(
    extracted_json: dict[str, Any],
    backend: LLMBackend | None = None,
    observation_only: bool = False,
    observation_records: list[dict[str, Any]] | None = None,
) -> str:
    """中間表現 JSON → Markdown 文字列に変換する。

    Args:
        extracted_json: ExtractedFileRecord.to_dict() の結果

    Returns:
        Markdown テキスト
    """
    document = extracted_json.get("document", {})
    elements = document.get("elements", [])
    metadata = extracted_json.get("metadata", {})
    source_path = metadata.get("source_path", "")
    source_ext = metadata.get("source_ext", "")
    doc_role_guess = metadata.get("doc_role_guess", "")

    parts: list[str] = []
    recent_headings: list[str] = []
    current_sheet_name = ""
    table_index = 0
    previous_table_context: dict[str, Any] = {}

    for elem in elements:
        elem_type = elem.get("type", "")
        content = elem.get("content", {})

        if elem_type == "heading":
            heading_text = content.get("text", "")
            if heading_text:
                recent_headings.append(heading_text)
                recent_headings = recent_headings[-3:]
            if content.get("detection_method") == "sheet_name" and heading_text:
                current_sheet_name = heading_text
                previous_table_context = {}
            parts.append(_render_heading(content))
            parts.append("")  # 見出し後の空行

        elif elem_type == "paragraph":
            parts.append(_render_paragraph(content))
            parts.append("")

        elif elem_type == "table":
            render_context = _RenderContext(
                source_path=source_path,
                source_ext=source_ext,
                doc_role_guess=doc_role_guess,
                current_sheet_name=current_sheet_name,
                heading_context=tuple(recent_headings),
                table_index=table_index,
                previous_table_context=previous_table_context,
            )
            parts.append(_render_table_as_labeled_text(
                content,
                render_context=render_context,
                backend=backend,
                observation_only=observation_only,
                observation_records=observation_records,
            ))
            previous_table_context = _extract_following_table_context(content)
            table_index += 1

        elif elem_type == "image":
            # 画像の存在を示すプレースホルダー
            desc = content.get("description", "")
            alt = content.get("alt_text", "")
            if desc:
                parts.append(f"[画像: {desc}]")
            elif alt:
                parts.append(f"[画像: {alt}]")
            else:
                parts.append("[画像]")
            parts.append("")

        elif elem_type == "shape":
            rendered = _render_shape(content)
            if rendered:
                parts.append(rendered)
                parts.append("")

        elif elem_type == "page_break":
            parts.append("---")
            parts.append("")

    # 末尾の余分な空行を整理
    text = "\n".join(parts).strip()
    return text + "\n"


def transform_file(
    json_path: Path,
    output_path: Path,
    backend: LLMBackend | None = None,
    observation_only: bool = False,
    observation_path: Path | None = None,
) -> StepResult:
    """1つの中間表現 JSON ファイルを Markdown に変換して書き出す。

    Args:
        json_path: Step2 出力の JSON ファイルパス
        output_path: 出力 Markdown ファイルパス

    Returns:
        StepResult
    """
    t0 = time.perf_counter()

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        elapsed = time.perf_counter() - t0
        return StepResult(
            file_path=str(json_path), step="transform",
            status=ProcessStatus.ERROR, message=f"JSON read error: {e}",
            duration_sec=round(elapsed, 2),
        )

    observations: list[dict[str, Any]] = []
    md_text = transform_to_markdown(
        data,
        backend=backend,
        observation_only=observation_only,
        observation_records=observations,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(md_text, encoding="utf-8")

    if observation_path is not None and observations:
        observation_payload = {
            "schema_version": "1.0",
            "source_json": str(json_path),
            "markdown_output": str(output_path),
            "observation_only": observation_only,
            "tables": observations,
        }
        observation_path.parent.mkdir(parents=True, exist_ok=True)
        observation_path.write_text(
            json.dumps(observation_payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    elapsed = time.perf_counter() - t0
    size_kb = len(md_text.encode("utf-8")) / 1024

    logger.info("変換完了: %s → %s (%.1fKB, %.1fs)", json_path.name, output_path.name, size_kb, elapsed)
    return StepResult(
        file_path=str(json_path), step="transform",
        status=ProcessStatus.SUCCESS,
        message=f"output={output_path.name}, size={size_kb:.1f}KB",
        duration_sec=round(elapsed, 2),
    )
