# P3: 図形フロー改善 — 実装計画書

## 概要

ワークフロー図の複数テキストボックスが個別の `[テキストボックス]` として出力され、
フロー全体の構造・順序関係が失われている。また矢印テキスト（→、↓）が見出しに誤昇格している。

## 現状の問題

### 問題1: テキストボックスが個別出力でフロー構造が失われる

**overlay_workflow.md の現状:**
```markdown
[テキストボックス]
  - 申請者
  - 申請書作成

[テキストボックス]
  - 上長
  - 内容確認

[テキストボックス]
  - 部長
  - 承認判断
```

**期待する出力:**
```markdown
[フロー図]
  1. 申請者 / 申請書作成
  2. 上長 / 内容確認
  3. 部長 / 承認判断
  4. 経理 / 処理実行
  5. 完了
```

### 問題2: 位置情報が未使用

`ShapeElement` には `left_pt`, `top_pt` があるが、出力順序に反映されていない。
エラーハンドリングフローでは正常系（top=0）とエラー系（top=80）が混在するが、
位置によるグルーピング・ソートが行われていない。

### 問題3: 矢印テキストの見出し誤昇格

```markdown
### 正常系: → → → →      ← 見出しに誤昇格
### エラー系: ↓ → → →     ← 見出しに誤昇格
### ↓ の矢印で接続          ← 見出しに誤昇格
```

これらは短文・句点なしのため `heuristic:short_no_period` に引っかかる。

---

## 変更対象ファイル

| ファイル | 変更内容 |
|---------|---------|
| `src/extractors/word.py` | 連続図形のフローグルーピング + 位置ソート |
| `src/transform/to_markdown.py` | フロー図のグループ描画 |
| `src/models/intermediate.py` | なし（既存の ShapeElement で対応可能） |
| `tests/test_extract.py` | P3 テストクラス追加 |
| `tests/test_transform.py` | P3 テストクラス追加 |

---

## 変更仕様

### 変更1: 矢印テキストの見出し誤昇格防止（word.py）

#### 1-1. 矢印判定関数の追加

**新規関数**: `_is_arrow_annotation(text: str) -> bool`

```python
_ARROW_CHARS = set("→←↑↓⇒⇐⇑⇓▶▷►▸◀◁◄◂➔➡➜➞")

def _is_arrow_annotation(text: str) -> bool:
    """矢印記号を含む注釈テキスト（フロー図の接続表現）か判定する。

    例:
      "→ → → → →"          → True
      "正常系: → → → →"     → True
      "↓ の矢印で接続"       → True
      "エラー処理について"     → False
    """
    return any(c in _ARROW_CHARS for c in text)
```

#### 1-2. _detect_heading での除外

**変更箇所**: `_detect_heading()` 関数内、ヒューリスティクス判定部分（現在 L138-142）

```python
# 4. 短文 + 行末句点なし（見出しらしいパターン）
if len(text) <= 30 and not text.endswith(("。", ".", "、", ",")):
    # ただし数字のみ、空白のみは除外
    if re.search(r"[\u3040-\u9fff\uff01-\uff5ea-zA-Z]", text):
        # ★ P3: 矢印記号を含む注釈テキストは見出しにしない
        if not _is_arrow_annotation(text):
            return (3, "heuristic:short_no_period")
```

**注意**: Word見出しスタイル（style）やフォントサイズでの見出し検出はそのまま維持。
ヒューリスティクス判定のみ矢印テキストを除外する。

---

### 変更2: 連続図形のフローグルーピング（word.py）

#### 2-1. 方針

現在 `_build_element_order()` → `_merge_overlapping_shapes()` で蓄積された図形が
テキスト段落到達時にフラッシュされる。このフラッシュ処理で、位置情報を使ったソートと
フローグルーピングを行う。

#### 2-2. フローグルーピング関数の追加

**新規関数**: `_group_shapes_as_flow(shapes: list[ShapeElement]) -> list[ShapeElement]`

```python
def _group_shapes_as_flow(shapes: list[ShapeElement]) -> list[ShapeElement]:
    """連続する図形を位置情報でソートし、テキストありの図形群をフロー図としてグルーピングする。

    条件:
      - テキストあり図形が3個以上連続
      - 位置情報（left_pt, top_pt）を持つ図形が2個以上

    上記条件を満たす場合、1つの ShapeElement にまとめる:
      - shape_type = "workflow"
      - texts = ソート済みの各図形テキスト（改行→" / " に変換）
      - description = "" (LLM 生成用に空けておく)

    条件を満たさない場合はそのまま返す。
    """
    if len(shapes) < 3:
        return shapes

    # テキストあり図形のみ抽出
    text_shapes = [s for s in shapes if s.texts]
    if len(text_shapes) < 3:
        return shapes

    # 位置情報でソート（top_pt → left_pt の順）
    has_pos = [s for s in text_shapes if s.top_pt is not None and s.left_pt is not None]
    no_pos = [s for s in text_shapes if s.top_pt is None or s.left_pt is None]

    if len(has_pos) >= 2:
        has_pos.sort(key=lambda s: (s.top_pt or 0, s.left_pt or 0))
        sorted_shapes = has_pos + no_pos
    else:
        sorted_shapes = text_shapes

    # 各図形のテキストを結合（改行 → " / "）
    flow_texts = []
    for s in sorted_shapes:
        combined = " / ".join(t.strip() for t in s.texts if t.strip())
        if combined:
            flow_texts.append(combined)

    if not flow_texts:
        return shapes

    # 1つの workflow ShapeElement にまとめる
    workflow = ShapeElement(
        shape_type="workflow",
        texts=flow_texts,
        confidence=Confidence.MEDIUM,
        fallback_reason="",
    )
    return [workflow]
```

#### 2-3. _build_element_order() への組み込み

フラッシュ処理の3箇所（テキスト段落到達時、テーブル前、末尾）で
`_merge_overlapping_shapes()` の後に `_group_shapes_as_flow()` を適用:

```python
# 変更前
merged = _merge_overlapping_shapes(pending_shapes)
for shape in merged:
    order.append(("shape_inline", shape))

# 変更後
merged = _merge_overlapping_shapes(pending_shapes)
grouped = _group_shapes_as_flow(merged)
for shape in grouped:
    order.append(("shape_inline", shape))
```

**3箇所すべてに同じ変更を適用**（L462-466, L473-477, L482-485）。

---

### 変更3: フロー図のMarkdown描画（to_markdown.py）

#### 3-1. _SHAPE_TYPE_LABEL への追加

```python
_SHAPE_TYPE_LABEL: dict[str, str] = {
    "vml_textbox": "テキストボックス",
    "vml_rect": "矩形オブジェクト",
    "vml": "図形",
    "floating": "図形",
    "workflow": "フロー図",     # ★ P3 追加
}
```

#### 3-2. _render_shape() のフロー図対応

フロー図（`shape_type == "workflow"`）では、テキストを番号付きリストで出力する:

```python
def _render_shape(content: dict[str, Any]) -> str:
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
        # ★ P3: フロー図は番号付きリストで出力
        lines.append(f"[{label}]")
        for i, t in enumerate(texts, 1):
            lines.append(f"  {i}. {t}")
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
```

---

## テスト仕様

### test_extract.py に追加: `TestFlowGrouping`

```python
class TestFlowGrouping:
    """図形フローグルーピングテスト（P3）"""

    def test_arrow_text_not_heading(self, tmp_path: Path, config: PipelineConfig):
        """矢印テキストが見出しに昇格しないこと"""
        from docx import Document as DocxDocument

        doc = DocxDocument()
        doc.add_paragraph("→ → → → →")
        doc.add_paragraph("正常系: → → → →")
        doc.add_paragraph("↓ の矢印で接続")
        doc.add_paragraph("エラー処理について")  # こちらは見出し候補

        path = tmp_path / "arrows.docx"
        doc.save(str(path))

        record, _ = extract_docx(path, "arrows.docx", ".docx", config)
        elements = record.document["elements"]

        headings = [e for e in elements if e["type"] == "heading"]
        heading_texts = [h["content"]["text"] for h in headings]

        # 矢印テキストは見出しにならない
        assert "→ → → → →" not in heading_texts
        assert "正常系: → → → →" not in heading_texts
        assert "↓ の矢印で接続" not in heading_texts

        # 通常の短文は見出し候補のまま
        assert "エラー処理について" in heading_texts
```

### test_transform.py に追加: `TestWorkflowRendering`

```python
class TestWorkflowRendering:
    """フロー図描画テスト（P3）"""

    def test_workflow_numbered_list(self):
        """フロー図が番号付きリストで出力されること"""
        doc = IntermediateDocument()
        doc.add_shape(
            shape_type="workflow",
            texts=["申請者 / 申請書作成", "上長 / 内容確認", "部長 / 承認判断"],
        )

        md = transform_to_markdown(_make_record(doc))
        assert "[フロー図]" in md
        assert "1. 申請者 / 申請書作成" in md
        assert "2. 上長 / 内容確認" in md
        assert "3. 部長 / 承認判断" in md

    def test_workflow_with_description(self):
        """description がある場合はそちらが優先されること"""
        doc = IntermediateDocument()
        doc.add_shape(
            shape_type="workflow",
            texts=["ステップA", "ステップB"],
            description="申請→承認→処理のフロー",
        )

        md = transform_to_markdown(_make_record(doc))
        assert "申請→承認→処理のフロー" in md
        assert "1." not in md  # 番号リストにはならない
```

---

## 実装の注意点

1. **グルーピング閾値**: テキストあり図形が「3個以上」でフロー判定。2個は通常の図形配置でもありうるため。

2. **位置ソートの優先順位**: `top_pt` → `left_pt` の順。これにより:
   - 横方向フロー: top が同じ → left で左から右に
   - 分岐フロー: top が異なる → 上段→下段の順に

3. **テキスト結合**: 各テキストボックス内の複数行テキスト（例: `["申請者", "申請書作成"]`）は `" / "` で結合。

4. **_merge_overlapping_shapes との連携**:
   - まず `_merge_overlapping_shapes()` でテキストなし矩形を除去
   - その後 `_group_shapes_as_flow()` で残った図形をグルーピング
   - この順序が重要（矩形除去前にグルーピングすると空テキストが混入する）

5. **既存テストへの影響**:
   - `test_shape_with_texts`: shape_type="text_box" でテキスト3つだが、ShapeElement 1個なのでグルーピング対象外 → 影響なし
   - `test_shape_no_text_keeps_placeholder`: テキストなし → 影響なし
   - `test_shape_with_description`: description あり → 影響なし

6. **矢印判定 `_is_arrow_annotation()`**:
   - ヒューリスティクス見出し検出のみに影響
   - Word見出しスタイルやフォントサイズでの検出は除外しない
   - 「正常系: → → → →」のようなラベル+矢印も矢印テキストとして判定する

---

## 変更しないもの（スコープ外）

- 分岐関係の抽出（正常系/エラー系のリンク）→ LLM description に委譲
- VML の `v:line` 矢印要素の解析 → 複雑すぎるため見送り
- 位置情報を使った図形間の接続関係推定 → 将来タスク

---

## ファイル変更サマリー

| ファイル | 変更行数（見積） | 内容 |
|---------|-------------|------|
| `src/extractors/word.py` | +50行 | `_is_arrow_annotation()`, `_group_shapes_as_flow()`, フラッシュ3箇所修正 |
| `src/transform/to_markdown.py` | +8行 | `_SHAPE_TYPE_LABEL` に "workflow" 追加、`_render_shape` にフロー分岐 |
| `tests/test_extract.py` | +30行 | `TestFlowGrouping` クラス（1テスト） |
| `tests/test_transform.py` | +30行 | `TestWorkflowRendering` クラス（2テスト） |
