# P6: 評価スクリプト強化 — 実装計画書

## 概要

P1〜P5 で追加した機能（結合セル、画像キャプション、フロー図、見出し精度、キャプション重複除去）
に対応する評価チェックを `tools/evaluate_results.py` に追加する。

現在の評価スクリプトは基本的なメタデータ・要素数のチェックのみで、
P1〜P5 の品質改善が正しく反映されているか検証できていない。

## 現状の問題

### evaluate_json() に不足しているチェック

1. **画像の説明文（P2）**: image 要素の `description` が設定されているか未チェック
2. **ワークフロー図形（P3）**: shape 要素の `shape_type == "workflow"` が存在するか未チェック
3. **見出しレベル整合性（P4）**: section 番号の深度と見出しレベルが矛盾していないか未チェック
4. **キャプション重複（P5）**: 表の caption と直前の段落テキストが重複していないか未チェック
5. **図/表キャプションの誤見出し化（P2/P4）**: `図:` `表:` パターンが見出しになっていないか未チェック

### evaluate_markdown() に不足しているチェック

1. **画像プレースホルダの説明文（P2）**: `[画像]` のみ（説明なし）の出力がないか未チェック
2. **フロー図出力（P3）**: `[フロー図]` の後に番号付きリストがあるか未チェック
3. **キャプション重複（P5）**: 同じテキストが段落と `**太字キャプション**` の両方に出現していないか未チェック
4. **バナー行（P1）**: `**太字**` 行が表内セクション区切りとして存在するか（存在時は情報として出力）

---

## 変更対象ファイル

| ファイル | 変更内容 |
|---------|---------|
| `tools/evaluate_results.py` | P1〜P5 対応チェック追加 |

**テストファイルの変更はなし。** 評価スクリプトはツールであり、パイプライン本体ではない。

---

## 変更仕様

### 変更1: evaluate_json() への P1〜P5 チェック追加

**変更箇所**: `evaluate_json()` 関数の末尾（L142 の `return ev` の前）

#### 1-1. 画像説明文チェック（P2）

```python
# --- P2: 画像の説明文 ---
images = [e for e in elements if e.get("type") == "image"]
ev.add("画像抽出", True, f"{len(images)}個")
if images:
    with_desc = sum(1 for img in images if img["content"].get("description"))
    bare_images = len(images) - with_desc
    ev.add("画像説明文あり", bare_images == 0,
           f"説明あり={with_desc}/{len(images)}" if bare_images == 0
           else f"説明なし={bare_images}/{len(images)} — 図キャプション未検出の可能性")
```

#### 1-2. ワークフロー図形チェック（P3）

```python
# --- P3: ワークフロー図形 ---
if shapes:
    workflows = [s for s in shapes if s["content"].get("shape_type") == "workflow"]
    ev.add("ワークフロー図形", True,
           f"{len(workflows)}個（{len(workflows)}/{len(shapes)}図形がワークフロー）")
    for i, wf in enumerate(workflows):
        texts = wf["content"].get("texts", [])
        ev.add(f"  ワークフロー{i+1}テキスト", len(texts) >= 2,
               f"{len(texts)}ステップ" if len(texts) >= 2
               else f"テキスト不足（{len(texts)}個）")
```

#### 1-3. 見出しレベル整合性チェック（P4）

```python
# --- P4: 見出しレベルと section 番号の整合性 ---
import re
_SECTION_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s")

if headings:
    level_mismatches = []
    for h in headings:
        text = h["content"].get("text", "")
        level = h["content"].get("level", 0)
        m = _SECTION_RE.match(text)
        if m:
            expected_depth = m.group(1).count(".") + 1
            expected_level = expected_depth + 1  # depth 1 → L2
            if level < expected_level:
                level_mismatches.append(
                    f"'{text[:30]}' L{level} (期待 L{expected_level})")
    ev.add("見出しレベル整合", len(level_mismatches) == 0,
           "整合" if not level_mismatches
           else f"{len(level_mismatches)}件の不整合: {'; '.join(level_mismatches[:3])}")
```

**ポイント:**
- `level < expected_level` のみチェック（style 見出しは section 番号より上位レベルになりうるため `>` は許容）
- 例: `4.1 入力チェック` が L2 → NG（L3 期待）、`4. 機能詳細` が L2 → OK

#### 1-4. キャプション重複チェック（P5）

```python
# --- P5: キャプション重複チェック ---
if tables:
    para_texts = set(
        e["content"]["text"]
        for e in elements if e.get("type") == "paragraph"
    )
    dup_captions = []
    for t in tables:
        cap = t["content"].get("caption", "")
        if cap and cap in para_texts:
            dup_captions.append(cap[:30])
    ev.add("キャプション重複なし", len(dup_captions) == 0,
           "重複なし" if not dup_captions
           else f"{len(dup_captions)}件の重複: {'; '.join(dup_captions)}")
```

#### 1-5. 図/表キャプション誤見出し化チェック（P2/P4）

```python
# --- P2/P4: 図・表キャプションが見出しになっていないか ---
if headings:
    _CAPTION_HEAD_RE = re.compile(
        r"^(?:図|表|Figure|Table)\s*[\d:：]",
        re.IGNORECASE,
    )
    caption_headings = [
        h["content"]["text"][:30]
        for h in headings
        if _CAPTION_HEAD_RE.match(h["content"].get("text", ""))
    ]
    ev.add("キャプション誤見出し化なし", len(caption_headings) == 0,
           "クリーン" if not caption_headings
           else f"{len(caption_headings)}件: {'; '.join(caption_headings)}")
```

### 変更2: evaluate_markdown() への P1〜P5 チェック追加

**変更箇所**: `evaluate_markdown()` 関数の末尾（L211 の `return ev` の前）

#### 2-1. 画像プレースホルダ説明文チェック（P2）

```python
# --- P2: 画像プレースホルダの説明文 ---
image_placeholders = [l for l in lines if l.strip().startswith("[画像")]
ev.add("画像プレースホルダ", True, f"{len(image_placeholders)}個")
if image_placeholders:
    bare = [l for l in image_placeholders if l.strip() == "[画像]"]
    ev.add("画像説明文付き", len(bare) == 0,
           f"全て説明あり" if not bare
           else f"説明なし={len(bare)}個 — [画像] のみの出力あり")
```

#### 2-2. フロー図出力チェック（P3）

```python
# --- P3: フロー図出力 ---
flow_lines = [i for i, l in enumerate(lines) if l.strip() == "[フロー図]"]
ev.add("フロー図", True, f"{len(flow_lines)}個")
if flow_lines:
    # フロー図の直後に番号付きリストがあるか
    flow_with_steps = 0
    for idx in flow_lines:
        # 次の非空行を探す
        for j in range(idx + 1, min(idx + 5, len(lines))):
            if lines[j].strip():
                if re.match(r"\s+\d+\.\s", lines[j]):
                    flow_with_steps += 1
                break
    ev.add("フロー図ステップあり", flow_with_steps == len(flow_lines),
           f"{flow_with_steps}/{len(flow_lines)}個にステップあり")
```

#### 2-3. キャプション重複チェック（P5）

```python
# --- P5: キャプション重複（段落 + 太字キャプション） ---
import re as _re
bold_captions = set()
plain_paragraphs = set()
_BOLD_RE = _re.compile(r"^\*\*(.+)\*\*$")

for l in lines:
    stripped = l.strip()
    if not stripped:
        continue
    m = _BOLD_RE.match(stripped)
    if m:
        bold_captions.add(m.group(1))
    elif not stripped.startswith(("#", "[", "  ")):
        plain_paragraphs.add(stripped)

dup_in_md = bold_captions & plain_paragraphs
ev.add("MD内キャプション重複なし", len(dup_in_md) == 0,
       "重複なし" if not dup_in_md
       else f"{len(dup_in_md)}件: {'; '.join(list(dup_in_md)[:3])}")
```

#### 2-4. バナー行情報（P1）

```python
# --- P1: バナー行（表内セクション区切り） ---
import re as _re2
banner_lines = [l for l in lines
                if l.strip().startswith("**") and l.strip().endswith("**")
                and not l.strip().startswith("**#")]
ev.add("バナー行（表内区切り）", True, f"{len(banner_lines)}個")
```

---

## 変更3: `import re` の追加

`evaluate_markdown()` と `evaluate_json()` の新チェックで `re` モジュールを使用するため、
ファイル先頭の import に `import re` を追加する。

```python
import re  # ★追加
```

---

## 実装の注意点

1. **チェックの pass/fail 判定**: P1〜P5 の改善が正しく反映されている場合に pass となるよう設計。
   既存のチェックと同様に、情報表示のみ（常に pass）のチェックと、品質判定（pass/fail）のチェックを混在させる。

2. **既存チェックとの重複回避**:
   - 既存の `図形抽出` チェック（L131-135）は shapes 全体のカウント。P3 のワークフロー図チェックは shapes ブロック内に追加し、既存の `テキスト付き図形` チェックの後に配置。
   - 既存の `図形プレースホルダ` チェック（L187-188）は `[図形` のみ。P2 の画像チェックは `[画像` を新規追加。

3. **`_SECTION_RE` のスコープ**: `evaluate_json()` 内でのみ使用するため、関数内でコンパイルする（モジュールレベルに出さない）。

4. **`evaluate_markdown()` での `re` 使用**: 関数内で `import re` は冗長なので、ファイル先頭で1回 import する。

5. **既存テストへの影響**: 評価スクリプトにはユニットテストがないため、影響なし。実行して結果を確認する形で検証。

---

## 検証方法

実装後、以下のコマンドで最新の出力に対して評価スクリプトを実行し、
P1〜P5 のチェックが全て pass することを確認する:

```bash
python tools/evaluate_results.py
```

### 期待する新規チェック結果（mixed_complex.json の例）

```
OK 画像抽出: 2個
OK 画像説明文あり: 説明あり=2/2
OK ワークフロー図形: 1個（1/1図形がワークフロー）
OK   ワークフロー1テキスト: 3ステップ
OK 見出しレベル整合: 整合
OK キャプション重複なし: 重複なし
OK キャプション誤見出し化なし: クリーン
```

### 期待する新規チェック結果（mixed_complex.md の例）

```
OK 画像プレースホルダ: 2個
OK 画像説明文付き: 全て説明あり
OK フロー図: 1個
OK フロー図ステップあり: 1/1個にステップあり
OK MD内キャプション重複なし: 重複なし
OK バナー行（表内区切り）: 0個
```

---

## ファイル変更サマリー

| ファイル | 変更行数（見積） | 内容 |
|---------|-------------|------|
| `tools/evaluate_results.py` | +80行 | P1〜P5 対応チェック追加、`import re` 追加 |
