"""テーブル解釈用のプロンプト生成と JSON パース。"""

from __future__ import annotations

import json
from typing import Any

from src.llm.base import ReconstructionUnit, TableInterpretationResult

TABLE_INTERPRETATION_PROMPT_VERSION = "table_interpretation.v2"

TABLE_INTERPRETATION_SYSTEM_PROMPT = f"""
prompt_version: {TABLE_INTERPRETATION_PROMPT_VERSION}

あなたは Excel / Word から抽出した表を再構成するアシスタントです。
与えられた JSON を読み、表の出力戦略を判定してください。

## テーブルタイプの定義
- "form": 帳票・申請書。ラベルと値がセル結合で並ぶレイアウト。ヘッダー行がない。
  例: 稟議書、申込書、設定シート、チェックリスト、各種届出
- "key_value": キーと値のペアが縦に並ぶ構造。列数は2が典型だが、広い表で2列のみ使用のケースもある。
  例: 仕様一覧、設定パラメータ表、属性リスト、プロファイル情報
- "data_table": 列ヘッダーがあり、各行が同構造のレコード。最も一般的。
  例: 売上表、社員名簿、ログ一覧、在庫表、スケジュール
- "unknown": 上記に当てはまらない、または構造が不明瞭な表。

## レンダリング戦略の選択基準
- "form_grid": form タイプに対応。各行をラベル: 値ペアとして出力。
- "key_value": key_value タイプに対応。第1列をキー、第2列を値として出力。
- "data_table": data_table タイプに対応。ヘッダー + データ行の Markdown テーブルとして出力。

## row_roles の定義
各行に1つ割り当てる。入力 rows と同じ長さの配列で返すこと。
- "data_record": 通常のデータ行（data_table のメイン行）
- "field_pairs": ラベルと値のペアが横に並ぶフォーム行（例: "件名 | 受注CSV取込"）
- "parallel_labels": 複数の見出しが横に並ぶ行（例: "担当 | 課長 | 部長"）。3つ以上のテキストセルが必要。
- "check_item": チェックボックス付きの項目行（□, ■, ☑ で始まる）
- "section_header": セクション区切りの見出し行（太字で出力される）
- "banner": 表の全幅に渡るタイトル行（単一セルが全列を占める）
- "text": 自由記述テキスト行（そのまま出力）
- "skip": 不要な行（空行、装飾のみの行）

## confidence の判断基準
- "high": テーブル構造が明確で、タイプ・戦略に迷いがない
- "medium": おおよそ判断できるが、一部の行の役割が曖昧
- "low": 構造が不規則で判断に自信がない

必ず JSON オブジェクトだけを返してください。説明文や Markdown は不要です。
""".strip()


_FEW_SHOT_EXAMPLES = """
### 例1: data_table（ヘッダー付きデータ表）
入力 rows (3行, 3列):
  行0: [{"text":"名前","colspan":1}, {"text":"部署","colspan":1}, {"text":"入社日","colspan":1}]
  行1: [{"text":"山田","colspan":1}, {"text":"営業部","colspan":1}, {"text":"2020-04","colspan":1}]
  行2: [{"text":"鈴木","colspan":1}, {"text":"開発部","colspan":1}, {"text":"2021-10","colspan":1}]
出力:
  {"table_type":"data_table", "render_strategy":"data_table",
   "header_rows":[0], "data_start_row":1,
   "column_labels":["名前","部署","入社日"], "active_columns":[0,1,2],
   "render_plan":{"row_roles":["data_record","data_record","data_record"]},
   "self_assessment":{"confidence":"high"}}

### 例2: form（結合セルによる帳票）
入力 rows (3行, 6列):
  行0: [{"text":"設備購入稟議書","colspan":6}]
  行1: [{"text":"件名","colspan":1}, {"text":"サーバ増設","colspan":5}]
  行2: [{"text":"申請者","colspan":1}, {"text":"山田太郎","colspan":2}, {"text":"申請日","colspan":1}, {"text":"2026-03-01","colspan":2}]
出力:
  {"table_type":"form", "render_strategy":"form_grid",
   "header_rows":[], "data_start_row":0,
   "render_plan":{"row_roles":["banner","field_pairs","field_pairs"]},
   "self_assessment":{"confidence":"high"}}

### 例3: key_value（広い表で2列のみ使用）
入力 rows (3行, 6列):
  行0: [{"text":"項目","colspan":1}, {"text":"","colspan":3}, {"text":"値","colspan":1}, {"text":"","colspan":1}]
  行1: [{"text":"担当部署","colspan":1}, {"text":"","colspan":3}, {"text":"営業本部","colspan":1}, {"text":"","colspan":1}]
  行2: [{"text":"納期","colspan":1}, {"text":"","colspan":3}, {"text":"2026年4月","colspan":1}, {"text":"","colspan":1}]
出力:
  {"table_type":"key_value", "render_strategy":"key_value",
   "header_rows":[0], "data_start_row":1,
   "column_labels":["項目","値"], "active_columns":[0,4],
   "render_plan":{"row_roles":["data_record","data_record","data_record"]},
   "self_assessment":{"confidence":"high"}}
""".strip()


class TableInterpretationParseError(ValueError):
    """LLM 応答を TableInterpretationResult へ変換できない場合の例外。"""


def build_table_interpretation_prompt(unit: ReconstructionUnit) -> str:
    """ReconstructionUnit から解釈用プロンプトを作る。"""
    return (
        "次の表データを読み、出力戦略を JSON で返してください。\n\n"
        f"## 判定例\n{_FEW_SHOT_EXAMPLES}\n\n"
        "## 返却 JSON スキーマ\n"
        "{\n"
        '  "schema_version": "1.0",\n'
        '  "unit_id": "<入力の unit_id をそのまま返す>",\n'
        '  "table_type": "form | key_value | data_table | unknown",\n'
        '  "render_strategy": "form_grid | key_value | data_table",\n'
        '  "header_rows": [0],\n'
        '  "data_start_row": 1,\n'
        '  "column_labels": ["列1", "列2"],\n'
        '  "active_columns": [0, 1],\n'
        '  "render_plan": {\n'
        '    "row_roles": ["field_pairs", "field_pairs"],\n'
        '    "summary_labels": ["1/水", "2/木"],\n'
        '    "markdown_lines": ["件名: 受注 CSV 取込レイアウト変更"]\n'
        '  },\n'
        '  "notes": ["判断根拠を短く記述"],\n'
        '  "self_assessment": {"confidence": "high | medium | low"}\n'
        "}\n\n"
        "## 補足\n"
        "- row_roles は入力 rows と同じ長さの配列で返してください。\n"
        "- render_plan は必要な場合のみ指定してください。\n"
        "- data_table でも row_roles を使って、前置き行・途中のフォーム行・"
        "セクション行・不要行を調整して構いません。\n"
        "- summary_labels は、1 行だけの集計表に対して値ごとのラベルを"
        "付けたい場合だけ指定してください。\n"
        "- markdown_lines を使う場合は、セルの文字列を失わずに Markdown "
        "本文としてそのまま出力したい行だけを返してください。\n"
        "- markdown_lines はテーブル全体の本文であり、表キャプションは"
        "含めないでください。\n"
        "- context.nearby_headings にはこの表の周辺にある見出しが含まれます。"
        "表の用途を推測する手がかりにしてください。\n"
        "- context.previous_table が与えられている場合は、列ラベルや"
        "列位置との対応を参考にしてください。\n"
        "- context.rule_based_interpretation がある場合は、ルールベースの"
        "解釈結果です。参考にしつつ、より適切な判断があれば修正してください。\n"
        "- hints フィールドに has_merged_cells, doc_role_guess 等の"
        "ヒントがある場合は判断の参考にしてください。\n"
        "- notes には判断根拠を短く記述してください。\n"
        "- Markdown 本文そのものは返さないでください。\n\n"
        "## 入力 JSON\n"
        f"{json.dumps(unit.to_dict(), ensure_ascii=False, indent=2)}"
    )


def _extract_first_json_object(text: str) -> str | None:
    decoder = json.JSONDecoder()
    start = text.find("{")
    while start != -1:
        candidate = text[start:]
        try:
            parsed, end = decoder.raw_decode(candidate)
        except json.JSONDecodeError:
            start = text.find("{", start + 1)
            continue
        if isinstance(parsed, dict):
            return candidate[:end]
        start = text.find("{", start + 1)
    return None


def _extract_json_text(text: str) -> str:
    """LLM 応答から JSON 文字列部分を抽出する。"""
    stripped = text.strip()
    if stripped.startswith("{"):
        candidate = _extract_first_json_object(stripped)
        if candidate is not None:
            return candidate
        if stripped.endswith("}"):
            return stripped

    if "```json" in stripped:
        start = stripped.find("```json") + len("```json")
        end = stripped.find("```", start)
        if end != -1:
            candidate = stripped[start:end].strip()
            parsed = _extract_first_json_object(candidate)
            if parsed is not None:
                return parsed
            return candidate

    if "```" in stripped:
        start = stripped.find("```") + len("```")
        end = stripped.find("```", start)
        if end != -1:
            candidate = stripped[start:end].strip()
            parsed = _extract_first_json_object(candidate)
            if parsed is not None:
                return parsed
            if candidate.startswith("{") and candidate.endswith("}"):
                return candidate

    candidate = _extract_first_json_object(stripped)
    if candidate is not None:
        return candidate

    start = stripped.find("{")
    end = stripped.rfind("}")
    if start != -1 and end != -1 and start < end:
        return stripped[start:end + 1]
    return stripped


def _string(value: Any, default: str = "") -> str:
    return value if isinstance(value, str) else default


def _string_list(value: Any) -> list[str]:
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, str)]


def _int_list(value: Any) -> list[int]:
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, int)]


def _dict(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def parse_table_interpretation_response(
    text: str, fallback_unit_id: str,
) -> TableInterpretationResult:
    """LLM 応答文字列を共通契約へ正規化する。"""
    json_text = _extract_json_text(text)
    try:
        raw = json.loads(json_text)
    except json.JSONDecodeError as exc:
        preview = " ".join(text.strip().split())[:100]
        raise TableInterpretationParseError(
            "LLM table interpretation JSON parse error: "
            f"unit_id={fallback_unit_id}, "
            f"line={exc.lineno}, column={exc.colno}, message={exc.msg}, "
            f"response_preview={preview!r}"
        ) from exc
    if not isinstance(raw, dict):
        raise TableInterpretationParseError(
            "LLM table interpretation response must be a JSON object: "
            f"unit_id={fallback_unit_id}, root_type={type(raw).__name__}"
        )

    table_type = _string(raw.get("table_type"), "unknown")
    if table_type not in {"form", "key_value", "data_table", "unknown"}:
        table_type = "unknown"

    render_strategy = _string(raw.get("render_strategy"), "data_table")
    if render_strategy not in {"form_grid", "key_value", "data_table"}:
        render_strategy = "data_table"

    return TableInterpretationResult(
        schema_version=_string(raw.get("schema_version"), "1.0"),
        unit_id=_string(raw.get("unit_id"), fallback_unit_id) or fallback_unit_id,
        table_type=table_type,
        render_strategy=render_strategy,
        header_rows=_int_list(raw.get("header_rows")),
        data_start_row=raw.get("data_start_row", 0)
        if isinstance(raw.get("data_start_row", 0), int) else 0,
        column_labels=_string_list(raw.get("column_labels")),
        active_columns=_int_list(raw.get("active_columns")),
        render_plan=_dict(raw.get("render_plan")),
        notes=_string_list(raw.get("notes")),
        self_assessment=_dict(raw.get("self_assessment")),
    )
