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


def _render_table_as_labeled_text(content: dict[str, Any]) -> str:
    """表を項目ラベル付き半構造化テキストに変換する。

    Task.md §6 の決定事項:
    「表は Markdown テーブルではなく項目ラベル付き半構造化テキストに変換して渡す。
     行列の意味や制約・対応関係を壊さないことを優先」

    変換戦略:
      - 1行目をヘッダー（ラベル名）として使用
      - 2行目以降の各行を「ラベル: 値」形式で出力
      - 結合セルや変更履歴テーブルは注意マーカー付き
    """
    rows = content.get("rows", [])
    if not rows:
        return ""

    lines: list[str] = []

    caption = content.get("caption", "")
    if caption:
        lines.append(f"**{caption}**")
        lines.append("")

    # ヘッダー行からラベルを取得
    header_row = rows[0]
    labels = [cell.get("text", f"列{i+1}") or f"列{i+1}" for i, cell in enumerate(header_row)]

    # データ行を「ラベル: 値」形式で出力
    for row_idx, row in enumerate(rows[1:], start=2):
        lines.append(f"[行{row_idx}]")
        for col_idx, cell in enumerate(row):
            label = labels[col_idx] if col_idx < len(labels) else f"列{col_idx+1}"
            value = cell.get("text", "")
            if value:
                lines.append(f"  {label}: {value}")
        lines.append("")

    # データ行がない場合（ヘッダーのみ）
    if len(rows) <= 1:
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
