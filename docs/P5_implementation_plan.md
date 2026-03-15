# P5: キャプション重複の解消 — 実装計画書

## 概要

表の直前の短い段落がキャプションとして表に設定されるが、
段落要素自体も残るため、同じテキストが2回出力される。

## 現状の問題

### mixed_complex.md の実例

```markdown
入力データの妥当性を検証する機能。       ← 段落として出力（L40）

**入力データの妥当性を検証する機能。**   ← 表キャプションとして出力（L42）

[行2]
  チェック項目: 必須チェック
```

### 原因

`extract_docx()` L677-684 のキャプション取得ロジック:

```python
# 表の直前段落をキャプション候補として取得
caption = ""
if intermediate.elements:
    last = intermediate.elements[-1]
    if last.type.value == "paragraph" and last.content is not None:
        candidate = last.content.text
        if len(candidate) <= 60:
            caption = candidate
```

段落をキャプションに使用した後も、`intermediate.elements` からその段落を削除していない。
結果として Markdown 出力時に段落テキストとキャプション `**text**` の両方が出力される。

---

## 変更対象ファイル

| ファイル | 変更内容 |
|---------|---------|
| `src/extractors/word.py` | キャプション使用時に段落要素を除去 |
| `tests/test_extract.py` | P5 テスト追加 |
| `tests/test_transform.py` | P5 テスト追加 |

**`to_markdown.py` の変更はなし。**

---

## 変更仕様

### 変更1: キャプション使用時の段落除去（word.py）

**変更箇所**: `extract_docx()` 内の表キャプション取得ロジック（現在 L677-684）

```python
# 変更前
caption = ""
if intermediate.elements:
    last = intermediate.elements[-1]
    if last.type.value == "paragraph" and last.content is not None:
        candidate = last.content.text
        if len(candidate) <= 60:
            caption = candidate

# 変更後
caption = ""
if intermediate.elements:
    last = intermediate.elements[-1]
    if last.type.value == "paragraph" and last.content is not None:
        candidate = last.content.text
        if len(candidate) <= 60:
            caption = candidate
            # ★ P5: キャプションに使用した段落を除去（重複出力防止）
            intermediate.elements.pop()
```

**ポイント:**
- `intermediate.elements.pop()` で末尾の段落要素を除去
- キャプションに使われた段落は表の `caption` フィールドに移動した形になる
- Markdown 変換時は `**{caption}**` として表の冒頭に1回だけ出力される

### 注意事項

1. **heading は対象外**: 条件 `last.type.value == "paragraph"` のため、見出しが表のキャプションに使われることはない（既存動作のまま）

2. **空段落は対象外**: `IntermediateDocument.add_paragraph()` が空テキストをスキップするため、空段落は elements に入らない

3. **60文字超の段落は対象外**: 長い段落はキャプションとして採用されないため、pop されない

4. **source_index への影響**: 除去した段落の source_index は失われるが、キャプションは表の一部として出力されるため問題ない

---

## テスト仕様

### test_extract.py に追加: `TestTableCaptionDedup`

```python
class TestTableCaptionDedup:
    """表キャプション重複除去テスト（P5）"""

    def test_caption_paragraph_removed(self, tmp_path, config):
        """表の直前段落がキャプションに使われた場合、段落要素が除去されること"""
        from docx import Document as DocxDocument

        doc = DocxDocument()
        doc.add_heading("セクション", level=2)
        doc.add_paragraph("機能一覧表")  # ← キャプション候補（60文字以内）
        table = doc.add_table(rows=2, cols=2)
        table.rows[0].cells[0].text = "項目"
        table.rows[0].cells[1].text = "値"
        table.rows[1].cells[0].text = "A"
        table.rows[1].cells[1].text = "1"

        path = tmp_path / "caption_dedup.docx"
        doc.save(str(path))

        record, _ = extract_docx(path, "caption_dedup.docx", ".docx", config)
        elements = record.document["elements"]

        # "機能一覧表" が段落要素として存在しないこと
        paragraphs = [e for e in elements if e["type"] == "paragraph"]
        para_texts = [p["content"]["text"] for p in paragraphs]
        assert "機能一覧表" not in para_texts

        # 表のキャプションとして設定されていること
        tables = [e for e in elements if e["type"] == "table"]
        assert tables[0]["content"]["caption"] == "機能一覧表"

    def test_long_paragraph_not_removed(self, tmp_path, config):
        """60文字超の段落はキャプションに使用されず、段落として残ること"""
        from docx import Document as DocxDocument

        doc = DocxDocument()
        long_text = "この段落は表のキャプションとして使用するには長すぎるテキストです。" * 2
        doc.add_paragraph(long_text)
        table = doc.add_table(rows=2, cols=2)
        table.rows[0].cells[0].text = "A"
        table.rows[0].cells[1].text = "B"
        table.rows[1].cells[0].text = "1"
        table.rows[1].cells[1].text = "2"

        path = tmp_path / "long_para.docx"
        doc.save(str(path))

        record, _ = extract_docx(path, "long_para.docx", ".docx", config)
        elements = record.document["elements"]

        # 長い段落は段落として残る
        paragraphs = [e for e in elements if e["type"] == "paragraph"]
        assert any(long_text in p["content"]["text"] for p in paragraphs)

        # 表のキャプションは空
        tables = [e for e in elements if e["type"] == "table"]
        assert tables[0]["content"]["caption"] == ""

    def test_heading_not_removed_as_caption(self, tmp_path, config):
        """見出しはキャプションに使用されず、見出しとして残ること"""
        from docx import Document as DocxDocument

        doc = DocxDocument()
        doc.add_heading("セクション見出し", level=2)
        table = doc.add_table(rows=2, cols=2)
        table.rows[0].cells[0].text = "A"
        table.rows[0].cells[1].text = "B"
        table.rows[1].cells[0].text = "1"
        table.rows[1].cells[1].text = "2"

        path = tmp_path / "heading_not_caption.docx"
        doc.save(str(path))

        record, _ = extract_docx(path, "heading_not_caption.docx", ".docx", config)
        elements = record.document["elements"]

        # 見出しは残っている
        headings = [e for e in elements if e["type"] == "heading"]
        assert any("セクション見出し" in h["content"]["text"] for h in headings)

        # 表のキャプションは空（見出しはキャプションに使われない）
        tables = [e for e in elements if e["type"] == "table"]
        assert tables[0]["content"]["caption"] == ""
```

### test_transform.py に追加: テスト確認

既存の `test_table_caption` がキャプションの出力を検証済み。
P5 で変更するのは抽出側（word.py）のみなので、追加のtransform テストは不要。

ただし、E2E 的な確認として:

```python
class TestCaptionNoDuplication:
    """キャプション重複なしの E2E テスト（P5）"""

    def test_caption_appears_once_in_markdown(self):
        """キャプションがMarkdownに1回だけ出力されること"""
        doc = IntermediateDocument()
        # キャプション段落は word.py 側で除去済みなので、
        # ここでは段落なし + caption 付きテーブルをテスト
        doc.add_table(
            rows=[
                [CellData(text="項目", row=0, col=0, is_header=True),
                 CellData(text="値", row=0, col=1, is_header=True)],
                [CellData(text="A", row=1, col=0),
                 CellData(text="1", row=1, col=1)],
            ],
            caption="機能一覧表",
        )

        md = transform_to_markdown(_make_record(doc))
        assert md.count("機能一覧表") == 1
        assert "**機能一覧表**" in md
```

---

## 実装の注意点

1. **`pop()` の安全性**: `intermediate.elements` が空でないことは `if intermediate.elements:` で保証済み。`last.type.value == "paragraph"` の条件を通過した後なので、末尾要素は確実に段落。

2. **既存テストへの影響**:
   - `test_basic_table` (`simple_docx`): 表の直前が見出し `1.2 制約事項` なのでキャプションは空 → 影響なし
   - `test_table_caption` (`test_transform.py`): 直接 `IntermediateDocument` を構築しているので word.py の変更は影響なし
   - `test_detect_with_fullwidth_spaces` (`change_history_docx`): 直前段落 "変更履歴" は60文字以内でキャプションになる → この段落が pop される。ただしテストは `fallback_reason == "change_history_table"` のみ検証しており、直前段落の有無はチェックしていないため影響なし

3. **`test_normal_table_not_detected`**: `simple_docx` の表の直前段落のテストだが、直前が見出しなのでキャプション設定されない → 影響なし

---

## ファイル変更サマリー

| ファイル | 変更行数（見積） | 内容 |
|---------|-------------|------|
| `src/extractors/word.py` | +2行 | キャプション使用時に `pop()` |
| `tests/test_extract.py` | +55行 | `TestTableCaptionDedup` クラス（3テスト） |
| `tests/test_transform.py` | +15行 | `TestCaptionNoDuplication` クラス（1テスト） |
