# P4: 見出し検出の精度向上 — 実装計画書

## 概要

P2（図キャプション除外）、P3（矢印テキスト除外）で heuristic の主要な誤検出は解消済み。
P4 では残る2つの問題に対処する:

1. **font-size 検出のレベルが section 番号と矛盾する問題**
2. **heuristic 検出での caption パターン除外漏れ**

## 現状の問題

### 問題1: font-size vs section 番号のレベル競合

`mixed_complex.json` の実例:
```
L2  4. 機能詳細              (style)            ← 親
L2  4.1 入力チェック機能      (font_size:14.0pt) ← 子なのに L2
L2  4.2 データ出力機能        (font_size:14.0pt) ← 子なのに L2
```

font-size 検出は pt 値のみでレベルを決定するため、
section 番号の階層構造（`4.1` = 2階層 → L3）を無視している。

**期待する出力:**
```
L2  4. 機能詳細              (style)
L3  4.1 入力チェック機能      (font_size:14.0pt + section_number)
L3  4.2 データ出力機能        (font_size:14.0pt + section_number)
```

### 問題2: heuristic が caption パターンを見出しにする

P2 の `_is_figure_caption()` は画像直後の段落でのみチェックされる。
画像の直後でない場合（空段落が挟まる等）、`図: XXX` がヒューリスティクス見出しになりうる。

同様に `表N: XXX` パターンもヒューリスティクスで見出しになりうる。

---

## 変更対象ファイル

| ファイル | 変更内容 |
|---------|---------|
| `src/extractors/word.py` | section 番号検出、caption 除外 |
| `tests/test_extract.py` | P4 テストクラス追加 |

**`to_markdown.py` の変更はなし。**

---

## 変更仕様

### 変更1: section 番号によるレベル補正

#### 1-1. section 番号検出関数の追加

**新規関数**: `_detect_section_number_depth(text: str) -> int | None`

```python
_SECTION_NUMBER_RE = re.compile(
    r"^(?:第(\d+)章|(\d+(?:\.\d+)*)\.?\s)"
)

def _detect_section_number_depth(text: str) -> int | None:
    """テキストの先頭から section 番号の階層深度を検出する。

    Returns:
        階層の深さ（1始まり）。検出できなければ None。

    例:
        "第1章 概要"      → 1
        "1. 概要"          → 1
        "4.1 入力チェック"  → 2
        "2.3.1 詳細"       → 3
        "概要"             → None
    """
    m = _SECTION_NUMBER_RE.match(text.strip())
    if not m:
        return None
    if m.group(1):
        # "第N章" パターン
        return 1
    if m.group(2):
        # "N.N.N" パターン: ドット数 + 1 = 階層深度
        return m.group(2).count(".") + 1
    return None
```

#### 1-2. _detect_heading() でのレベル補正

**変更箇所**: `_detect_heading()` の font-size 検出部分（現在 L130-142）

```python
# 3. フォントサイズ差（Oasys/Win 対応）
size = _get_font_size_pt(para)
if size is not None and size >= config.heading_font_size_min_pt:
    # まず section 番号の深度を確認
    depth = _detect_section_number_depth(text)
    if depth is not None:
        # section 番号の深度に基づくレベル（depth 1 → L2, depth 2 → L3, ...）
        level = min(depth + 1, 6)
    else:
        # サイズに応じてレベルを推定（既存ロジック）
        if size >= 16.0:
            level = 1
        elif size >= 14.0:
            level = 2
        elif size >= 12.0:
            level = 3
        else:
            level = 4
    return (level, f"font_size:{size}pt")
```

**ポイント:**
- `depth + 1` とするのは、section 番号の深度1（例: `1.`）が通常 H2 に相当するため
- `depth 1` → L2（`# ` ではなく `## `）、`depth 2` → L3、...
- `第N章` は depth=1 → L2（H1 は文書タイトル用に予約）

#### 1-3. heuristic 検出でも section 番号を活用

**変更箇所**: `_detect_heading()` のヒューリスティクス部分（現在 L144-149）

```python
# 4. 短文 + 行末句点なし（見出しらしいパターン）
if len(text) <= 30 and not text.endswith(("。", ".", "、", ",")):
    if re.search(r"[\u3040-\u9fff\uff01-\uff5ea-zA-Z]", text):
        if not _is_arrow_annotation(text):
            # ★ P4: caption パターンは見出しにしない
            if _is_figure_caption(text) is not None:
                return None
            if _is_table_caption(text) is not None:
                return None
            # section 番号があればその深度でレベル決定
            depth = _detect_section_number_depth(text)
            if depth is not None:
                level = min(depth + 1, 6)
                return (level, "heuristic:section_number")
            return (3, "heuristic:short_no_period")
```

---

### 変更2: 表キャプション判定関数の追加

**新規関数**: `_is_table_caption(text: str) -> str | None`

```python
_TABLE_CAPTION_RE = re.compile(
    r"^(?:表|Table)\s*[:：]?\s*(.+)",
    re.IGNORECASE,
)

def _is_table_caption(text: str) -> str | None:
    """段落テキストが表キャプションかどうか判定する。

    Returns:
        キャプション本文。表キャプションでなければ None。
    """
    stripped = text.strip()
    if not stripped or len(stripped) > 60 or stripped.endswith("。"):
        return None

    match = _TABLE_CAPTION_RE.match(stripped)
    if not match:
        return None

    caption = match.group(1).strip()
    return caption or None
```

**判定パターン:**
- `表1 画面一覧` → `"1 画面一覧"` (caption detected)
- `表: ユーザー管理` → `"ユーザー管理"` (caption detected)
- `Table 3: API` → `"3: API"` (caption detected)
- `表現力が高い。` → `None` (句点あり、本文)
- `表面処理について` → `None` (**注意**: 実際にはマッチしてしまう）

**「表面」の誤検出対策**: 正規表現に先読みを追加:

```python
_TABLE_CAPTION_RE = re.compile(
    r"^(?:表|Table)(?=\s*\d|\s*[:：]|\s+)(?:\s*[:：]?\s*)(.+)$",
    re.IGNORECASE,
)
```

これは P2 の `_FIGURE_CAPTION_RE` と同じパターンの先読み:
- `表1` → `(?=\s*\d)` でマッチ
- `表: ` → `(?=\s*[:：])` でマッチ
- `表 ` → `(?=\s+)` でマッチ
- `表面` → いずれの先読みにもマッチしない → 除外

---

## テスト仕様

### test_extract.py に追加: `TestHeadingPrecision`

```python
class TestHeadingPrecision:
    """見出し検出精度テスト（P4）"""

    def test_section_number_overrides_font_size_level(self, tmp_path, config):
        """section 番号で font-size 見出しのレベルが補正されること"""
        from docx import Document as DocxDocument
        from docx.shared import Pt

        doc = DocxDocument()
        # 親見出し（スタイル）
        doc.add_heading("4. 機能詳細", level=2)
        # 子見出し（font-size のみ、スタイルなし）
        p1 = doc.add_paragraph()
        run1 = p1.add_run("4.1 入力チェック機能")
        run1.font.size = Pt(14)
        doc.add_paragraph("入力データの妥当性を検証する。")
        p2 = doc.add_paragraph()
        run2 = p2.add_run("4.2 データ出力機能")
        run2.font.size = Pt(14)

        path = tmp_path / "section_num.docx"
        doc.save(str(path))

        record, _ = extract_docx(path, "section_num.docx", ".docx", config)
        elements = record.document["elements"]

        headings = [e for e in elements if e["type"] == "heading"]
        # "4. 機能詳細" → L2 (style)
        h_parent = [h for h in headings if "機能詳細" in h["content"]["text"]]
        assert h_parent[0]["content"]["level"] == 2

        # "4.1 入力チェック" → L3 (font_size + section_number depth=2)
        h_child = [h for h in headings if "入力チェック" in h["content"]["text"]]
        assert h_child[0]["content"]["level"] == 3

    def test_table_caption_not_heading(self, tmp_path, config):
        """「表N: XXX」パターンがヒューリスティクス見出しにならないこと"""
        from docx import Document as DocxDocument

        doc = DocxDocument()
        doc.add_paragraph("表1: ユーザー管理テーブル")
        table = doc.add_table(rows=2, cols=2)
        table.rows[0].cells[0].text = "ID"
        table.rows[0].cells[1].text = "名前"
        table.rows[1].cells[0].text = "1"
        table.rows[1].cells[1].text = "田中"

        path = tmp_path / "table_caption.docx"
        doc.save(str(path))

        record, _ = extract_docx(path, "table_caption.docx", ".docx", config)
        elements = record.document["elements"]

        headings = [e for e in elements if e["type"] == "heading"]
        heading_texts = [h["content"]["text"] for h in headings]
        assert "表1: ユーザー管理テーブル" not in heading_texts

    def test_figure_caption_standalone_not_heading(self, tmp_path, config):
        """画像直後でない「図: XXX」もヒューリスティクス見出しにならないこと"""
        from docx import Document as DocxDocument

        doc = DocxDocument()
        doc.add_paragraph("前のセクション。")
        doc.add_paragraph("")  # 空段落
        doc.add_paragraph("図: 概念図")

        path = tmp_path / "standalone_fig_caption.docx"
        doc.save(str(path))

        record, _ = extract_docx(path, "standalone_fig_caption.docx", ".docx", config)
        elements = record.document["elements"]

        headings = [e for e in elements if e["type"] == "heading"]
        heading_texts = [h["content"]["text"] for h in headings]
        assert "図: 概念図" not in heading_texts

    def test_chapter_heading_level(self, tmp_path, config):
        """「第N章」パターンが L2 で検出されること"""
        from docx import Document as DocxDocument

        doc = DocxDocument()
        doc.add_paragraph("第1章 システム概要")
        doc.add_paragraph("本章ではシステムの概要を述べる。")

        path = tmp_path / "chapter.docx"
        doc.save(str(path))

        record, _ = extract_docx(path, "chapter.docx", ".docx", config)
        elements = record.document["elements"]

        headings = [e for e in elements if e["type"] == "heading"]
        h = [h for h in headings if "システム概要" in h["content"]["text"]]
        assert len(h) == 1
        assert h[0]["content"]["level"] == 2
```

---

## 実装の注意点

1. **`_detect_section_number_depth()` の正規表現**: `^(\d+(?:\.\d+)*)\.?\s` で `4.1 ` にマッチ。末尾の `.?\s` は `4.` (ドット+スペース) と `4.1 ` (ドットなし+スペース) の両方に対応。

2. **`第N章` は depth=1 → L2**: 日本語の仕様書では `第1章` が最上位の見出し構造だが、H1 は文書タイトル用に予約されているため L2 とする。

3. **style 見出しは section 番号補正の対象外**: Word 見出しスタイルのレベルはそのまま信頼する。補正するのは font-size 検出と heuristic 検出のみ。

4. **`_is_table_caption()` と `_is_figure_caption()` は heuristic 内で使用**: font-size 検出ではキャプションを除外しない（大きなフォントのキャプションは稀なため）。

5. **既存テストへの影響**:
   - `test_font_size_heading`: テストデータに section 番号パターンが含まれない（「機能概要」「入力条件」等）ため影響なし
   - `test_heuristic_heading`: 「エラーコード一覧」「入力チェック」は caption パターンにマッチしないため影響なし

---

## 変更しないもの（スコープ外）

- Word 見出しスタイルのレベル補正 → スタイルを信頼する方針
- 文書全体のレベル整合性チェック（L1→L3 のスキップ検出等）→ 将来タスク
- テーブルキャプション重複（P5 スコープ）

---

## ファイル変更サマリー

| ファイル | 変更行数（見積） | 内容 |
|---------|-------------|------|
| `src/extractors/word.py` | +40行 | `_detect_section_number_depth()`, `_is_table_caption()`, `_detect_heading()` 修正 |
| `tests/test_extract.py` | +70行 | `TestHeadingPrecision` クラス（4テスト） |
