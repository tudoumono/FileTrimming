import pytest

from src.llm.table_interpretation import (
    TABLE_INTERPRETATION_PROMPT_VERSION,
    TABLE_INTERPRETATION_SYSTEM_PROMPT,
    TableInterpretationParseError,
    parse_table_interpretation_response,
)


def test_parse_table_interpretation_response_accepts_trailing_text_with_braces():
    text = (
        '{"schema_version":"1.0","unit_id":"unit-1","table_type":"key_value",'
        '"render_strategy":"key_value","data_start_row":1}\n'
        "補足メモ: } は説明文側の文字です。"
    )

    result = parse_table_interpretation_response(text, "fallback-unit")

    assert result.unit_id == "unit-1"
    assert result.render_strategy == "key_value"
    assert result.data_start_row == 1


def test_parse_table_interpretation_response_raises_preview_on_invalid_json():
    text = """```json
{"schema_version":"1.0","table_type":
```
"""

    with pytest.raises(TableInterpretationParseError) as excinfo:
        parse_table_interpretation_response(text, "unit-parse-error")

    message = str(excinfo.value)
    assert "unit_id=unit-parse-error" in message
    assert "response_preview=" in message


def test_system_prompt_contains_prompt_version():
    assert TABLE_INTERPRETATION_PROMPT_VERSION in TABLE_INTERPRETATION_SYSTEM_PROMPT
