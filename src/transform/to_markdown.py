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


def _fill_rowspan(rows: list[list[dict[str, Any]]]) -> list[list[dict[str, Any]]]:
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
    active_spans: dict[int, tuple[dict[str, Any], int]] = {}
    result: list[list[dict[str, Any]]] = []

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


def _expand_row_to_positions(row: list[dict[str, Any]]) -> list[tuple[str, int]]:
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

        positions: list[tuple[str, int]] = [("", 0)] * max_pos
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


def _is_empty_row(row: list[dict[str, Any]]) -> bool:
    """行が完全に空（全セルのテキストが空）かどうか判定する。"""
    return all(not cell.get("text", "") for cell in row)


def _is_form_field_row(row: list[dict[str, Any]], total_cols: int) -> bool:
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
    text_cells = [c for c in row if c.get("text", "")]
    if not text_cells:
        return False
    if len(text_cells) >= total_cols * 3 / 4:
        return False
    return all(cell.get("colspan", 1) >= 2 for cell in text_cells)


def _estimate_total_cols(rows: list[list[dict[str, Any]]]) -> int:
    """テーブル全体の列数を推定する（最大展開幅）。"""
    max_cols = 0
    for row in rows:
        positions = _expand_row_to_positions(row)
        if len(positions) > max_cols:
            max_cols = len(positions)
    return max_cols


def _render_form_field_row(row: list[dict[str, Any]]) -> list[str]:
    """フォームフィールド行をラベル-値ペアとして出力する。

    セルの並び方に応じて自動判定:
      - 2セル: 単一のラベル-値ペア
      - 偶数セル: ラベル-値ペアの繰り返し
      - 奇数セル: 最後のセルは単独出力
      - 1セル: そのまま出力
    """
    cells = [c for c in row if c.get("text", "")]
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


def _is_form_grid_table(rows: list[list[dict[str, Any]]], total_cols: int) -> bool:
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


def _should_skip_as_header(row: list[dict[str, Any]], total_cols: int) -> bool:
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
    rows: list[list[dict[str, Any]]], total_cols: int,
) -> tuple[int, list[tuple[str, int]], bool]:
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
    rows: list[list[dict[str, Any]]], total_cols: int,
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

    if has_parent_colspan and header_idx + 1 < len(rows):
        row_next = rows[header_idx + 1]
        row_next_positions = _expand_row_to_positions(row_next)

        # サブヘッダー判定: 展開後の列数が一致し、かつ行全体がバナーでない
        if len(row_next_positions) == hdr_cols and not _is_banner_row(row_next, hdr_cols):
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


def _is_section_header_row(row: list[dict[str, Any]], total_cols: int) -> bool:
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


def _get_section_header_text(row: list[dict[str, Any]], total_cols: int) -> str:
    """セクション見出し行からテキストを取得する。支配的セルのテキストを返す。"""
    parts = []
    for cell in row:
        text = cell.get("text", "")
        if text:
            parts.append(text)
    return " / ".join(parts) if parts else ""


def _render_form_grid(rows: list[list[dict[str, Any]]], total_cols: int) -> str:
    """フォーム型テーブルを全行ラベル-値ペアとして出力する。

    フォーム型: ヘッダー行が存在せず、全行がラベル-値ペアの結合セルで構成される。
    業務申請書、稟議書、設定シート等でよく見られるレイアウト。
    """
    lines: list[str] = []
    for row in rows:
        if _is_empty_row(row):
            continue
        positions = _expand_row_to_positions(row)
        if _is_banner_row(row, len(positions)):
            banner_text = row[0].get("text", "")
            if banner_text:
                lines.append(f"**{banner_text}**")
                lines.append("")
            continue
        field_lines = _render_form_field_row(row)
        lines.extend(field_lines)
        if field_lines:
            lines.append("")
    return "\n".join(lines)


def _render_pre_header_rows(
    rows: list[list[dict[str, Any]]], data_start: int, total_cols: int,
) -> list[str]:
    """ヘッダーより前の行を出力する（バナー→太字、フォーム→ラベル: 値）。"""
    lines: list[str] = []
    for row in rows[:data_start]:
        if _is_empty_row(row):
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
    rows: list[list[dict[str, Any]]], data_start: int, total_cols: int,
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


def _render_key_value_table(
    rows: list[list[dict[str, Any]]],
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

        # キーとバリューを取得
        key = positions[key_col][0] if key_col < len(positions) else ""
        val = positions[val_col][0] if val_col < len(positions) else ""

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
    rows: list[list[dict[str, Any]]],
    labels: list[str],
    data_start: int,
    total_cols: int,
) -> str:
    """データテーブル型を出力する。

    対応するレイアウト:
      - ヘッダー前のフォーム行（混在型: 請求書の上部にフォーム部分）
      - セクション分割テーブル（バナー行の後に新ヘッダーが出現 → ラベルを再構築）
    """
    lines: list[str] = []

    # ヘッダーより前の行を出力
    lines.extend(_render_pre_header_rows(rows, data_start, total_cols))

    # データ行を出力（セクション分割対応）
    display_row_num = 1
    i = data_start
    while i < len(rows):
        row = rows[i]

        if _is_empty_row(row):
            i += 1
            continue

        positions = _expand_row_to_positions(row)

        # セクション見出し行 → 見出し出力 + 次行のヘッダー再検出
        if _is_section_header_row(row, total_cols):
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
        if _is_banner_row(row, len(positions)):
            banner_text = row[0].get("text", "")
            if banner_text:
                lines.append(f"**{banner_text}**")
                lines.append("")
            i += 1
            continue

        # フォームフィールド行がデータ部分に混在する場合
        if _is_form_field_row(row, total_cols):
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
        lines.append("  " + " | ".join(labels))
        lines.append("")

    return "\n".join(lines)


def _render_table_as_labeled_text(content: dict[str, Any]) -> str:
    """表を項目ラベル付き半構造化テキストに変換する。

    Task.md §6 の決定事項:
    「表は Markdown テーブルではなく項目ラベル付き半構造化テキストに変換して渡す。
     行列の意味や制約・対応関係を壊さないことを優先」

    テーブル型の自動判定:
      1. フォーム型: 全行がラベル-値ペア（全セル colspan >= 2）→ 全行をペア出力
      2. データテーブル型: ヘッダー行あり → ヘッダー + ラベル付きデータ行
      3. 混在型: ヘッダー前にフォーム行、以降データ行
    """
    rows = content.get("rows", [])
    if not rows:
        return ""

    # rowspan で縦結合されたセルの値を後続行に展開
    rows = _fill_rowspan(rows)

    lines: list[str] = []

    caption = content.get("caption", "")
    if caption:
        lines.append(f"**{caption}**")
        lines.append("")

    total_cols = _estimate_total_cols(rows)

    # テーブル型判定: フォーム型 vs キーバリュー型 vs データテーブル型
    if _is_form_grid_table(rows, total_cols):
        # フォーム型: 全行をラベル-値ペアとして出力
        lines.append(_render_form_grid(rows, total_cols))
    else:
        # データテーブル型（混在型含む）
        labels, data_start, header_found = _build_column_labels(rows, total_cols)
        effective_cols = len(labels) if labels else total_cols

        # キーバリュー型判定: 広い表だが実データは2列程度のみ
        active_cols = _detect_active_columns(rows, data_start, total_cols)
        if header_found and len(active_cols) >= 2 and len(active_cols) <= total_cols // 2:
            lines.append(_render_key_value_table(
                rows, labels, data_start, total_cols, active_cols,
            ))
        else:
            lines.append(_render_data_table(rows, labels, data_start, effective_cols))

    return "\n".join(lines)


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
