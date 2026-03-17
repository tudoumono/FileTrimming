"""テーブル解釈用のプロンプト生成と JSON パース。"""

from __future__ import annotations

import json
from typing import Any

from src.llm.base import ReconstructionUnit, TableInterpretationResult

TABLE_INTERPRETATION_SYSTEM_PROMPT = """
あなたは Excel / Word から抽出した表を再構成するアシスタントです。
与えられた JSON を読み、表の出力戦略だけを判定してください。

必ず JSON オブジェクトだけを返してください。説明文や Markdown は不要です。
使用可能な値:
- table_type: "form" | "key_value" | "data_table" | "unknown"
- render_strategy: "form_grid" | "key_value" | "data_table"
- render_plan.row_roles: "data_record" | "parallel_labels" | "field_pairs" | "check_item" | "section_header" | "banner" | "text" | "skip"
- render_plan.summary_labels: 1 行だけの集計表の値に対応するラベル一覧
- render_plan.markdown_lines: テーブル本文として採用したい Markdown 行の配列
""".strip()


def build_table_interpretation_prompt(unit: ReconstructionUnit) -> str:
    """ReconstructionUnit から解釈用プロンプトを作る。"""
    return (
        "次の表データを読み、出力戦略を JSON で返してください。\n\n"
        "返却 JSON スキーマ:\n"
        "{\n"
        '  "schema_version": "1.0",\n'
        '  "unit_id": "<入力の unit_id をそのまま返す>",\n'
        '  "table_type": "form | key_value | data_table | unknown",\n'
        '  "render_strategy": "form_grid | key_value | data_table",\n'
        '  "header_rows": [0],\n'
        '  "data_start_row": 1,\n'
        '  "column_labels": ["列1", "列2"],\n'
        '  "active_columns": [0, 1],\n'
        '  "render_plan": {"row_roles": ["field_pairs", "field_pairs"], "summary_labels": ["1/水", "2/木"], "markdown_lines": ["件名: 受注 CSV 取込レイアウト変更"]},\n'
        '  "notes": ["任意の短い補足"],\n'
        '  "self_assessment": {"confidence": "high | medium | low"}\n'
        "}\n\n"
        "補足:\n"
        "- render_plan は必要な場合のみ指定してください。\n"
        "- row_roles は入力 rows の各行を Markdown でどう出すかの方針です。\n"
        "- data_table でも row_roles を使って、前置き行・途中のフォーム行・セクション行・不要行を調整して構いません。\n"
        "- summary_labels は、1 行だけの集計表に対して値ごとのラベルを付けたい場合だけ指定してください。\n"
        "- markdown_lines を使う場合は、セルの文字列を失わずに Markdown 本文としてそのまま出力したい行だけを返してください。\n"
        "- markdown_lines はテーブル全体の本文であり、表キャプションは含めないでください。\n"
        "- previous_table が与えられている場合は、その列ラベルや列位置との対応を参考にしても構いません。\n"
        "- Markdown 本文そのものは返さないでください。\n\n"
        "入力 JSON:\n"
        f"{json.dumps(unit.to_dict(), ensure_ascii=False, indent=2)}"
    )


def _extract_json_text(text: str) -> str:
    """LLM 応答から JSON 文字列部分を抽出する。"""
    stripped = text.strip()
    if stripped.startswith("{") and stripped.endswith("}"):
        return stripped

    if "```json" in stripped:
        start = stripped.find("```json") + len("```json")
        end = stripped.find("```", start)
        if end != -1:
            return stripped[start:end].strip()

    if "```" in stripped:
        start = stripped.find("```") + len("```")
        end = stripped.find("```", start)
        if end != -1:
            candidate = stripped[start:end].strip()
            if candidate.startswith("{") and candidate.endswith("}"):
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
    raw = json.loads(_extract_json_text(text))

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
