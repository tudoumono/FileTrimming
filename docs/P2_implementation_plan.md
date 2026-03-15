# P2: 画像/図キャプション改善 — 実装計画書

## 概要

現在の出力では画像が裸の `[画像]` プレースホルダーのみで、キャプション情報が失われている。
「図: システム構成図」のような段落が見出し（`### 図: システム構成図`）に誤昇格し、
Dify のチャンク分割境界を不適切に増やしている。

## 現状の問題（many_images.md の実例）

```markdown
[画像]

### 図: システム構成図    ← 画像直後の「図:」が見出しに誤昇格
```

**期待する出力:**
```markdown
[画像: システム構成図]    ← キャプションが画像に統合される
```

---

## 変更対象ファイル

| ファイル | 変更内容 |
|---------|---------|
| `src/extractors/word.py` | 図キャプション検出 + 画像との統合 |
| `src/transform/to_markdown.py` | テキストなし図形の `[図形: type]` 出力修正 |
| `tests/test_extract.py` | P2 テストクラス追加 |
| `tests/test_transform.py` | P2 テストクラス追加 |
| `tests/conftest.py` | 画像+キャプション用フィクスチャ追加 |

---

## 変更仕様

### 変更1: 図キャプション検出と画像への統合（word.py）

#### 1-1. 図キャプション判定関数の追加

**新規関数**: `_is_figure_caption(text: str) -> str | None`

```python
import re

_FIGURE_CAPTION_RE = re.compile(
    r"^(?:図|Fig\.?|Figure)\s*[:：]?\s*(.+)",
    re.IGNORECASE,
)

def _is_figure_caption(text: str) -> str | None:
    """段落テキストが図キャプションかどうか判定する。

    Returns:
        キャプション本文（例: "システム構成図"）。図キャプションでなければ None。
    """
    m = _FIGURE_CAPTION_RE.match(text.strip())
    if m:
        return m.group(1).strip()
    return None
```

**判定パターン:**
- `図: システム構成図` → `"システム構成図"`
- `図：ER図` → `"ER図"`
- `図1 画面遷移図` → `"画面遷移図"`  ※ 番号付き
- `Fig. 1: Network Diagram` → `"1: Network Diagram"`
- `図のような構成で...` → `None`（長文・句点ありは除外）

**追加の除外条件:**
- テキストが60文字を超える場合は除外（本文の可能性）
- 句点（。）で終わる場合は除外

#### 1-2. 画像直後のキャプション統合（extract_docx 内）

**変更箇所**: `extract_docx()` 関数内の要素追加ループ（現在 L543 付近）

**現在の処理フロー:**
```
image → 即座に ImageElement(alt_text=alt, description="") を追加
次の paragraph → _detect_heading() で見出し判定
  → "図: XXX" が短文なので heuristic:short_no_period で見出しになる
```

**変更後の処理フロー:**
```
image → ImageElement を追加（ただし直後のキャプション統合のため保留フラグを立てる）
次の paragraph → まず _is_figure_caption() でチェック
  → キャプションなら:
    - 直前の ImageElement.description にキャプション本文を設定
    - この段落は要素として追加しない（画像に統合済み）
  → キャプションでなければ: 通常の見出し/段落判定
```

**実装方針:**

`extract_docx()` 内のメインループで `last_image_element` 変数を導入:

```python
last_image_element: ImageElement | None = None

for elem_type, idx_or_data in element_order:
    if elem_type == "image":
        alt_text = idx_or_data if isinstance(idx_or_data, str) else ""
        img = ImageElement(alt_text=alt_text, description="")
        intermediate.elements.append(DocumentElement(
            type=ElementType.IMAGE, content=img, source_index=source_idx,
        ))
        last_image_element = img
        source_idx += 1
        image_count += 1
        continue

    if elem_type == "paragraph" and idx_or_data < len(paragraphs):
        para = paragraphs[idx_or_data]
        text = para.text.strip()

        # ★ 図キャプション統合: 直前が画像なら優先チェック
        if last_image_element is not None and text:
            caption_text = _is_figure_caption(text)
            if caption_text:
                # 画像の description にキャプションを統合
                last_image_element.description = caption_text
                last_image_element = None
                source_idx += 1
                continue

        last_image_element = None  # 画像直後でなければリセット

        # 以下、既存の見出し/段落判定（変更なし）
        heading_info = _detect_heading(para, config)
        ...
```

**注意点:**
- `last_image_element` は画像の直後1段落のみ有効
- 画像→画像の連続時はリセットされない（各画像に対応するキャプションのみ）
- 段落以外（table, shape）が来たらリセット

#### 1-3. 表キャプションとの競合回避

現在、表の直前段落を `caption` として取得するロジック（L584-591）がある。
図キャプション（`図: XXX`）は表のキャプションにはならないため、競合は発生しない。
ただし、画像→キャプション段落→表の順序の場合、キャプション段落が画像に統合されるため
表のキャプションは空になる。これは正しい動作。

---

### 変更2: テキストなし図形の type 付きプレースホルダー（to_markdown.py）

#### 現在の問題

テスト `test_shape_no_text_keeps_placeholder` が期待する出力:
```
[図形: vml]
```

現在のコード出力:
```
[図形]
```

#### 修正箇所

`_render_shape()` 関数（L206-207）:

```python
# 変更前
else:
    lines.append(f"[{label}]")

# 変更後
else:
    lines.append(f"[{label}: {shape_type}]")
```

ただし `label` が既に `shape_type` と同等の場合は冗長になるため:

```python
else:
    if label != "図形" or not shape_type:
        lines.append(f"[{label}]")
    else:
        lines.append(f"[図形: {shape_type}]")
```

**テスト `test_shape_no_text_keeps_placeholder` の期待値:**
- `shape_type="vml"`, `texts=[]` → `[図形: vml]`

---

### 変更3: 画像プレースホルダーの出力改善（to_markdown.py）

#### 現在の出力

```python
# description あり → [画像: {description}]
# alt_text あり   → [画像: {alt_text}]
# なし            → [画像]
```

これは変更不要（P2 の変更1 で description が設定されるようになるため、
自動的に `[画像: システム構成図]` 形式で出力される）。

---

## テスト仕様

### test_extract.py に追加: `TestImageCaptionDetection`

```python
class TestImageCaptionDetection:
    """画像キャプション検出テスト（P2）"""

    def test_figure_caption_merged_to_image(self, tmp_path, config):
        """画像直後の「図: XXX」がImageElementのdescriptionに統合されること"""
        from docx import Document as DocxDocument
        doc = DocxDocument()
        doc.add_paragraph("以下にシステム構成図を示す。")
        # インライン画像を追加（python-docxのadd_picture）
        # ※ テスト用ダミー画像が必要 → 1x1 PNG を tmp_path に作成
        _create_dummy_image(tmp_path / "dummy.png")
        doc.add_picture(str(tmp_path / "dummy.png"))
        doc.add_paragraph("図: システム構成図")  # ← キャプション
        doc.add_paragraph("本文が続く。")

        path = tmp_path / "img_caption.docx"
        doc.save(str(path))

        record, result = extract_docx(path, "img_caption.docx", ".docx", config)
        elements = record.document["elements"]

        # 画像要素の description にキャプションが設定されていること
        images = [e for e in elements if e["type"] == "image"]
        assert len(images) == 1
        assert images[0]["content"]["description"] == "システム構成図"

        # キャプション段落が見出しとして残っていないこと
        headings = [e for e in elements if e["type"] == "heading"]
        caption_headings = [h for h in headings if "システム構成図" in h["content"]["text"]]
        assert len(caption_headings) == 0

    def test_non_caption_not_merged(self, tmp_path, config):
        """画像直後でも図キャプションでない段落は統合されないこと"""
        from docx import Document as DocxDocument
        doc = DocxDocument()
        _create_dummy_image(tmp_path / "dummy.png")
        doc.add_picture(str(tmp_path / "dummy.png"))
        doc.add_paragraph("この画像は参考資料である。")  # ← 句点あり、キャプションではない

        path = tmp_path / "img_no_caption.docx"
        doc.save(str(path))

        record, result = extract_docx(path, "img_no_caption.docx", ".docx", config)
        elements = record.document["elements"]

        images = [e for e in elements if e["type"] == "image"]
        assert len(images) == 1
        assert images[0]["content"]["description"] == ""  # 統合されない

    def test_figure_caption_with_number(self, tmp_path, config):
        """「図1 画面遷移図」形式のキャプションが統合されること"""
        from docx import Document as DocxDocument
        doc = DocxDocument()
        _create_dummy_image(tmp_path / "dummy.png")
        doc.add_picture(str(tmp_path / "dummy.png"))
        doc.add_paragraph("図1 画面遷移図")

        path = tmp_path / "img_numbered.docx"
        doc.save(str(path))

        record, result = extract_docx(path, "img_numbered.docx", ".docx", config)
        elements = record.document["elements"]

        images = [e for e in elements if e["type"] == "image"]
        assert len(images) == 1
        # "1 画面遷移図" or "画面遷移図" — 番号を含むかは正規表現次第
        assert "画面遷移図" in images[0]["content"]["description"]
```

**ヘルパー関数（conftest.py または test_extract.py 内）:**

```python
def _create_dummy_image(path: Path):
    """テスト用 1x1 PNG を作成する"""
    import struct, zlib
    # 最小限の PNG (1x1 白ピクセル)
    def _minimal_png():
        sig = b'\x89PNG\r\n\x1a\n'
        ihdr_data = struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)
        ihdr_crc = zlib.crc32(b'IHDR' + ihdr_data)
        ihdr = struct.pack('>I', 13) + b'IHDR' + ihdr_data + struct.pack('>I', ihdr_crc)
        raw = b'\x00\x00\x00\x00'
        idat_data = zlib.compress(raw)
        idat_crc = zlib.crc32(b'IDAT' + idat_data)
        idat = struct.pack('>I', len(idat_data)) + b'IDAT' + idat_data + struct.pack('>I', idat_crc)
        iend_crc = zlib.crc32(b'IEND')
        iend = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', iend_crc)
        return sig + ihdr + idat + iend
    path.write_bytes(_minimal_png())
```

**代替案（よりシンプル）:** Pillow が利用可能なら:
```python
from PIL import Image
img = Image.new("RGB", (1, 1), "white")
img.save(str(path))
```

### test_transform.py に追加: テストは変更不要

P2 の Markdown 変換側の変更は `_render_shape()` のテキストなし図形のみ。
既存テスト `test_shape_no_text_keeps_placeholder` が通るようになれば OK。

---

## 実装の注意点

1. **`_is_figure_caption()` の正規表現は保守的に**: 「図のような構成で動作する。」のような本文を誤検出しないよう、文字数上限（60文字）と句点チェックを入れる。

2. **`last_image_element` のライフサイクル**: 画像直後の1段落のみ有効。`shape_inline` や `table` が来たらリセット。

3. **`ImageElement` はミュータブル**: `dataclass` なので `description` フィールドを後から書き換え可能。中間表現の `elements` リストに追加済みの `ImageElement` インスタンスを直接書き換える（参照渡し）。

4. **テスト画像の作成**: `python-docx` の `doc.add_picture()` にはファイルパスが必要。1x1 PNG を動的生成する。Pillow が依存関係にあれば使用、なければ raw bytes で生成。

5. **既存テストへの影響**:
   - `test_shape_no_text_keeps_placeholder`: `_render_shape` 修正で PASS になる
   - 他のテストへの影響なし（画像キャプション統合は新機能）

---

## 変更しないもの（スコープ外）

- テーブルセル内の画像抽出 → P2 スコープ外（影響範囲が大きいため別タスク）
- `original_filename` の実装 → LLM 連携時に検討
- LLM による画像 description 生成 → オフライン制約のため見送り

---

## 依存関係の確認

```
pip list | grep -i pillow
```

Pillow がなければ raw PNG bytes で対応。`python-docx` の `add_picture` は Pillow を内部で使用するため、Pillow が入っていなければテスト用画像挿入に工夫が必要。

## ファイル変更サマリー

| ファイル | 変更行数（見積） | 内容 |
|---------|-------------|------|
| `src/extractors/word.py` | +30行 | `_is_figure_caption()` 追加、`extract_docx()` 内キャプション統合ロジック |
| `src/transform/to_markdown.py` | +3行 | `_render_shape()` テキストなし図形の type 表示 |
| `tests/test_extract.py` | +60行 | `TestImageCaptionDetection` クラス（3テスト） |
| `tests/conftest.py` | +15行 | `_create_dummy_image` ヘルパー（または test_extract.py 内） |
