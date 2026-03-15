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

import json
import time
from logging import getLogger
from pathlib import Path
from typing import Any

from src.models.metadata import ProcessStatus, StepResult

logger = getLogger(__name__)


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


def _expand_row_to_positions(row: list[dict[str, Any]]) -> list[tuple[str, int]]:
    """行のセルを列位置に展開する。

    Returns:
        [(text, colspan), ...] — 各列位置のテキストと元の colspan
    """
    positions: list[tuple[str, int]] = []
    for cell in row:
        text = cell.get("text", "")
        cs = cell.get("colspan", 1)
        positions.append((text, cs))
        # colspan > 1 の場合、残りの位置は同じテキストで埋める
        for _ in range(cs - 1):
            positions.append((text, 0))  # 0 = 展開済み位置
    return positions


def _build_column_labels(rows: list[list[dict[str, Any]]]) -> tuple[list[str], int]:
    """ヘッダー行からカラム位置ベースのラベルを構築する。

    多段ヘッダー対応:
      - row[0] に colspan > 1 のセルがあり、row[1] がサブヘッダーに見える場合、
        「親ラベル/子ラベル」形式で結合する。

    Returns:
        (labels, data_start_idx): ラベルリストとデータ開始行インデックス
    """
    if not rows:
        return [], 0

    header_positions = _expand_row_to_positions(rows[0])
    total_cols = len(header_positions)
    labels = [t or f"列{i+1}" for i, (t, _) in enumerate(header_positions)]

    data_start = 1
    has_parent_colspan = any(cell.get("colspan", 1) > 1 for cell in rows[0])

    if has_parent_colspan and len(rows) > 1:
        row1 = rows[1]
        row1_positions = _expand_row_to_positions(row1)

        # サブヘッダー判定: 展開後の列数が一致し、かつ行全体がバナーでない
        if len(row1_positions) == total_cols and not _is_banner_row(row1, total_cols):
            combined: list[str] = []
            for i, ((parent, _), (child, _)) in enumerate(
                zip(header_positions, row1_positions)
            ):
                if parent == child or not child:
                    combined.append(parent or f"列{i+1}")
                elif not parent:
                    combined.append(child)
                else:
                    combined.append(f"{parent}/{child}")
            labels = combined
            data_start = 2

    return labels, data_start


def _is_banner_row(row: list[dict[str, Any]], total_cols: int) -> bool:
    """行が全列スパンのバナー行（セクション区切り等）か判定する。"""
    if len(row) == 1 and row[0].get("colspan", 1) >= total_cols:
        return True
    # 全セルが同一テキストの場合もバナー扱い（横結合の残骸対策）
    if len(row) > 1:
        texts = [c.get("text", "") for c in row]
        if texts and all(t == texts[0] for t in texts) and texts[0]:
            return True
    return False


def _render_table_as_labeled_text(content: dict[str, Any]) -> str:
    """表を項目ラベル付き半構造化テキストに変換する。

    Task.md §6 の決定事項:
    「表は Markdown テーブルではなく項目ラベル付き半構造化テキストに変換して渡す。
     行列の意味や制約・対応関係を壊さないことを優先」

    変換戦略:
      - 多段ヘッダー対応: colspan > 1 のヘッダー + サブヘッダーを「親/子」形式で結合
      - バナー行: 全列スパンの行はセクション区切りとして単独出力
      - 通常行: 「ラベル: 値」形式で出力
      - 空セル値は明示しない（省略）
    """
    rows = content.get("rows", [])
    if not rows:
        return ""

    lines: list[str] = []

    caption = content.get("caption", "")
    if caption:
        lines.append(f"**{caption}**")
        lines.append("")

    # ラベル構築（多段ヘッダー対応）
    labels, data_start = _build_column_labels(rows)
    total_cols = len(labels) if labels else len(rows[0])

    # データ行を出力
    display_row_num = 2  # 表示上の行番号（ヘッダー分をスキップ）
    for row in rows[data_start:]:
        # バナー行判定
        if _is_banner_row(row, total_cols):
            banner_text = row[0].get("text", "")
            if banner_text:
                lines.append(f"**{banner_text}**")
                lines.append("")
            display_row_num += 1
            continue

        lines.append(f"[行{display_row_num}]")

        # 行のセルを列位置に展開
        row_positions = _expand_row_to_positions(row)
        for pos_idx, (value, cs) in enumerate(row_positions):
            if cs == 0:
                continue  # colspan 展開済みの位置はスキップ
            label = labels[pos_idx] if pos_idx < len(labels) else f"列{pos_idx+1}"
            if value:
                lines.append(f"  {label}: {value}")

        lines.append("")
        display_row_num += 1

    # データ行がない場合（ヘッダーのみ）
    if data_start >= len(rows):
        lines.append("  " + " | ".join(labels))
        lines.append("")

    return "\n".join(lines)


_SHAPE_TYPE_LABEL: dict[str, str] = {
    "vml_textbox": "テキストボックス",
    "vml_rect": "矩形オブジェクト",
    "vml": "図形",
    "floating": "図形",
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


def transform_to_markdown(extracted_json: dict[str, Any]) -> str:
    """中間表現 JSON → Markdown 文字列に変換する。

    Args:
        extracted_json: ExtractedFileRecord.to_dict() の結果

    Returns:
        Markdown テキスト
    """
    document = extracted_json.get("document", {})
    elements = document.get("elements", [])

    parts: list[str] = []

    for elem in elements:
        elem_type = elem.get("type", "")
        content = elem.get("content", {})

        if elem_type == "heading":
            parts.append(_render_heading(content))
            parts.append("")  # 見出し後の空行

        elif elem_type == "paragraph":
            parts.append(_render_paragraph(content))
            parts.append("")

        elif elem_type == "table":
            parts.append(_render_table_as_labeled_text(content))

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

    md_text = transform_to_markdown(data)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(md_text, encoding="utf-8")

    elapsed = time.perf_counter() - t0
    size_kb = len(md_text.encode("utf-8")) / 1024

    logger.info("変換完了: %s → %s (%.1fKB, %.1fs)", json_path.name, output_path.name, size_kb, elapsed)
    return StepResult(
        file_path=str(json_path), step="transform",
        status=ProcessStatus.SUCCESS,
        message=f"output={output_path.name}, size={size_kb:.1f}KB",
        duration_sec=round(elapsed, 2),
    )
