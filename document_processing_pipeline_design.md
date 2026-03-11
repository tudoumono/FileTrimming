# ドキュメント処理パイプライン 詳細設計書

## 1. 全体アーキテクチャ

```
[入力フォルダ]
    │
    ▼
┌─────────────────────────────┐
│ Phase 1: コピー & 準備       │
│  - 作業フォルダへコピー       │
│  - ファイル種別の自動判別     │
└─────────────────────────────┘
    │
    ▼
┌─────────────────────────────┐
│ Phase 2: フォーマット正規化   │
│  - .doc → .docx              │
│  - .xls → .xlsx              │
│  - .ppt → .pptx              │
│  - .rtf → .docx（TRF対応）    │
└─────────────────────────────┘
    │
    ▼
┌─────────────────────────────┐
│ Phase 3: ファイルサイズ検査   │
│  & 物理分割                  │
│  - トークン推定              │
│  - 閾値超過時に分割          │
└─────────────────────────────┘
    │
    ▼
┌─────────────────────────────┐
│ Phase 4: Markdown変換        │
│  - Excel → xlwings解析       │
│  - Word/PPT → MarkItDown     │
│  - PDF → MarkItDown          │
│  - Bugless → カスタムパーサ   │
│  - TRF/RTF → pandoc          │
│  - テキスト → そのまま        │
└─────────────────────────────┘
    │
    ▼
┌─────────────────────────────┐
│ Phase 5: 構造化 & ノイズ除去  │
│  - ルールベース前処理         │
│  - OpenAI による構造化成形    │
│    （必要なファイルのみ）      │
└─────────────────────────────┘
    │
    ▼
┌─────────────────────────────┐
│ Phase 6: チャンキング & 出力  │
│  - セマンティックチャンキング │
│  - メタデータ付与             │
│  - RAG用インデックス出力      │
└─────────────────────────────┘
    │
    ▼
[出力フォルダ（Markdown + メタデータJSON）]
```

---

## 2. Phase 1: コピー & 準備

### 処理内容

```python
import shutil
from pathlib import Path

def phase1_copy_and_classify(input_dir: str, work_dir: str) -> dict:
    """入力フォルダを作業フォルダにコピーし、ファイル種別を分類する"""
    
    shutil.copytree(input_dir, work_dir, dirs_exist_ok=True)
    
    FILE_TYPE_MAP = {
        # Excel系
        '.xlsx': 'excel', '.xls': 'excel_legacy', '.xlsm': 'excel',
        # Word系
        '.docx': 'word', '.doc': 'word_legacy',
        # PowerPoint系
        '.pptx': 'pptx', '.ppt': 'pptx_legacy',
        # PDF
        '.pdf': 'pdf',
        # リッチテキスト（TRF含む）
        '.rtf': 'rtf', '.trf': 'trf',
        # テキスト
        '.txt': 'text', '.csv': 'text', '.tsv': 'text', '.md': 'text',
        # Bugless（富士通独自形式）
        '.bgl': 'bugless', '.bug': 'bugless',
    }
    
    classified = {}
    for f in Path(work_dir).rglob('*'):
        if f.is_file():
            ext = f.suffix.lower()
            ftype = FILE_TYPE_MAP.get(ext, 'unknown')
            classified.setdefault(ftype, []).append(str(f))
    
    return classified
```

### 設計ポイント

- **拡張子だけでなく、magic bytes（ファイルヘッダ）でも判定**することを推奨。拡張子が間違っているケースがあるため。
- Bugless ファイルの拡張子が不明な場合は、ファイルの先頭数バイトや内容のパターンマッチで判別するカスタムロジックが必要。
- `unknown` に分類されたファイルはログに記録し、手動確認フローへ回す。

---

## 3. Phase 2: フォーマット正規化

### 処理内容

古い形式のファイルを新しい形式に変換する。

```python
import subprocess

def phase2_normalize_formats(classified: dict, work_dir: str) -> dict:
    """レガシー形式を現行形式に変換"""
    
    CONVERSIONS = {
        'excel_legacy': ('xlsx', 'calc'),    # .xls → .xlsx
        'word_legacy':  ('docx', 'writer'),  # .doc → .docx
        'pptx_legacy':  ('pptx', 'impress'), # .ppt → .pptx
        'rtf':          ('docx', 'writer'),  # .rtf → .docx
    }
    
    for ftype, (target_ext, lo_filter) in CONVERSIONS.items():
        for filepath in classified.get(ftype, []):
            convert_with_libreoffice(filepath, target_ext)
    
    return reclassify(work_dir)  # 変換後に再分類


def convert_with_libreoffice(filepath: str, target_ext: str):
    """LibreOfficeによる形式変換"""
    cmd = [
        'soffice', '--headless', '--norestore',
        '--convert-to', target_ext,
        '--outdir', str(Path(filepath).parent),
        filepath
    ]
    result = subprocess.run(cmd, capture_output=True, timeout=120)
    if result.returncode != 0:
        raise ConversionError(f"変換失敗: {filepath}")
```

### 設計ポイント

| 変換元 | 変換先 | 使用ツール | 注意点 |
|--------|--------|-----------|--------|
| `.xls` | `.xlsx` | LibreOffice | マクロ付き(.xlsm)は別途対応 |
| `.doc` | `.docx` | LibreOffice | 複雑なレイアウトは崩れる場合あり |
| `.ppt` | `.pptx` | LibreOffice | アニメーション情報は一部欠落 |
| `.rtf`/`.trf` | `.docx` | LibreOffice | TRFは事前検証が必要 |

**TRFファイルについて**：富士通のTRF形式がRTF互換であればLibreOfficeで変換可能。独自バイナリの場合は、富士通のプレビューツールで一旦PDFやテキストにエクスポートしてから処理する方針が現実的。

**Buglessファイルについて**：COBOLの設計書＋ソースコードが混在する独自形式のため、以下の2段階アプローチを提案：

1. **構造解析フェーズ**：テキストとして読み込み、設計書部分とCOBOLコード部分をセクション分離
2. **Markdown化フェーズ**：設計書部分はMarkdown散文へ、COBOL部分はコードブロック（```cobol）として出力

---

## 4. Phase 3: ファイルサイズ検査 & 物理分割

### トークン推定ロジック

```python
import tiktoken

# OpenAI API の実質的な安全上限（モデルにより異なる）
TOKEN_LIMITS = {
    'gpt-4o':      128_000,   # 入力上限
    'gpt-4o-mini': 128_000,
}
# 安全マージンを考慮した処理上限（出力トークン分を差し引き）
PROCESSING_LIMIT = 80_000  # 入力に使える実質上限

def estimate_tokens(filepath: str) -> int:
    """ファイルのトークン数を推定"""
    enc = tiktoken.encoding_for_model("gpt-4o")
    
    ext = Path(filepath).suffix.lower()
    
    if ext in ('.xlsx', '.xlsm'):
        return estimate_excel_tokens(filepath, enc)
    elif ext == '.docx':
        text = extract_text_docx(filepath)
        return len(enc.encode(text))
    elif ext == '.pdf':
        text = extract_text_pdf(filepath)
        return len(enc.encode(text))
    elif ext == '.pptx':
        text = extract_text_pptx(filepath)
        return len(enc.encode(text))
    else:
        with open(filepath, 'r', errors='ignore') as f:
            return len(enc.encode(f.read()))


def estimate_excel_tokens(filepath: str, enc) -> int:
    """Excelファイルのシートごとのトークン推定"""
    import openpyxl
    wb = openpyxl.load_workbook(filepath, data_only=True)
    total = 0
    sheet_tokens = {}
    for ws in wb.worksheets:
        text = '\n'.join(
            '\t'.join(str(c.value or '') for c in row)
            for row in ws.iter_rows()
        )
        tokens = len(enc.encode(text))
        sheet_tokens[ws.title] = tokens
        total += tokens
    return total, sheet_tokens
```

### 物理分割戦略

ファイル形式ごとに最適な分割単位が異なる。

```python
def phase3_split_if_needed(filepath: str, file_type: str) -> list[str]:
    """必要に応じてファイルを物理分割し、分割後のパスリストを返す"""
    
    tokens = estimate_tokens(filepath)
    
    if tokens <= PROCESSING_LIMIT:
        return [filepath]  # 分割不要
    
    SPLIT_STRATEGIES = {
        'excel': split_excel_by_sheet,     # シート単位で分割
        'word':  split_word_by_heading,    # 見出し単位で分割
        'pptx':  split_pptx_by_slides,    # N枚ごとに分割
        'pdf':   split_pdf_by_pages,       # Nページごとに分割
        'text':  split_text_by_lines,      # N行ごとに分割
        'bugless': split_bugless_by_section, # セクション単位で分割
    }
    
    splitter = SPLIT_STRATEGIES.get(file_type, split_text_by_lines)
    return splitter(filepath, PROCESSING_LIMIT)
```

#### Excel分割の詳細

```python
def split_excel_by_sheet(filepath: str, token_limit: int) -> list[str]:
    """
    Excel分割戦略（優先順位）:
    1. シート単位で分割（最も意味的にまとまりがある）
    2. 1シートが巨大な場合 → 行範囲で分割
    """
    import openpyxl
    
    wb = openpyxl.load_workbook(filepath, data_only=True)
    parts = []
    
    for ws in wb.worksheets:
        sheet_tokens = estimate_sheet_tokens(ws)
        
        if sheet_tokens <= token_limit:
            # シート全体で1ファイル
            parts.append(save_single_sheet(filepath, ws.title))
        else:
            # シート内を行範囲で分割
            row_groups = chunk_rows(ws, token_limit)
            for i, (start_row, end_row) in enumerate(row_groups):
                parts.append(
                    save_sheet_range(filepath, ws.title, start_row, end_row, i)
                )
    
    return parts
```

#### Word/PDF分割の詳細

```python
def split_word_by_heading(filepath: str, token_limit: int) -> list[str]:
    """
    Heading1/Heading2 を境界として分割。
    見出しがない場合はN段落ごとに分割。
    """
    pass  # python-docx で見出しスタイルを検出し分割

def split_pdf_by_pages(filepath: str, token_limit: int) -> list[str]:
    """
    ページ単位で分割。
    各チャンクが token_limit を超えないようにページ数を調整。
    """
    pass  # PyMuPDF (fitz) でページ単位分割
```

---

## 5. Phase 4: Markdown変換（★設計の核心）

### 提案: ファイル形式別の最適ツール選定

ここが設計上最も重要な判断ポイント。結論として **ハイブリッドアプローチ** を推奨する。

| ファイル形式 | 推奨ツール | 理由 |
|-------------|-----------|------|
| **Excel** | **xlwings（推奨）** | セル書式・結合・色・罫線などの構造情報を保持でき、Markdownテーブルの精度が高い |
| **Word** | **MarkItDown** | 見出し構造・リスト・テーブルをそのままMarkdown化できる |
| **PowerPoint** | **MarkItDown** | スライド構造を維持してMarkdown化できる |
| **PDF** | **MarkItDown + PyMuPDF** | テキスト抽出はMarkItDown、レイアウト解析が必要ならPyMuPDFを併用 |
| **Bugless** | **カスタムパーサ** | 独自形式のため専用ロジックが必要 |
| **TRF/RTF** | **pandoc** | `pandoc -f rtf -t markdown` で高品質変換可能 |
| **テキスト** | **直接読み込み** | そのままMarkdownとして扱える |

### なぜExcelにxlwingsを推奨するか

MarkItDown は Excel を処理できるが、以下の情報が欠落する：

1. **セル結合情報** — 結合セルはMarkdownテーブルで表現困難だが、xlwingsなら検出・注記できる
2. **セルの背景色・フォント色** — 日本企業のExcelでは「黄色＝要確認」「赤字＝変更箇所」等の意味を持つことが多い
3. **条件付き書式・データバリデーション** — 業務ロジックが埋め込まれている場合がある
4. **複数テーブルが1シートに存在するレイアウト** — BAGLESのようなExcel設計書に多いパターン
5. **セルコメント・メモ** — レビュー情報や補足説明

```python
# xlwings による高精度Excel→Markdown変換

import xlwings as xw

def excel_to_markdown_xlwings(filepath: str) -> str:
    """
    xlwingsによるExcel→Markdown変換
    
    ※ xlwings は Windows/macOS の Excel COM/AppleScript 連携が必要。
    ※ Linux サーバ上では xlwings PRO の xlwings.pro.reports や
       openpyxl へのフォールバックを検討。
    """
    
    app = xw.App(visible=False)
    try:
        wb = app.books.open(filepath)
        md_parts = []
        
        for sheet in wb.sheets:
            md_parts.append(f"# シート: {sheet.name}\n")
            
            used_range = sheet.used_range
            if used_range is None:
                md_parts.append("（空のシート）\n")
                continue
            
            # === ステップ1: テーブル領域の検出 ===
            tables = detect_table_regions(sheet)
            
            for table_idx, table_region in enumerate(tables):
                md_parts.append(
                    convert_region_to_markdown(sheet, table_region, table_idx)
                )
            
            # === ステップ2: コメント・メモの抽出 ===
            comments = extract_comments(sheet)
            if comments:
                md_parts.append("\n## コメント・メモ\n")
                for c in comments:
                    md_parts.append(f"- **{c['cell']}**: {c['text']}\n")
        
        wb.close()
        return '\n'.join(md_parts)
    
    finally:
        app.quit()


def detect_table_regions(sheet) -> list:
    """
    シート内の独立したテーブル領域を検出する。
    
    検出ロジック:
    1. 使用範囲を取得
    2. 空行・空列をスキャンして、データブロックの境界を特定
    3. 各ブロックをテーブル領域として返す
    
    これにより、1シートに複数テーブルが配置されている
    BAGLES形式のExcelに対応できる。
    """
    pass


def convert_region_to_markdown(sheet, region, idx: int) -> str:
    """
    検出したテーブル領域をMarkdownテーブルに変換。
    
    処理内容:
    - 結合セルの検出 → 結合範囲を注記として追記
    - ヘッダ行の推定（罫線・色・太字から判断）
    - 数値セルのフォーマット保持
    - 背景色が意味を持つ場合の注記追加
    """
    
    values = sheet.range(region).value
    
    # 結合セル検出
    merged = detect_merged_cells(sheet, region)
    
    # ヘッダ行推定
    header_row = estimate_header_row(sheet, region)
    
    # Markdownテーブル生成
    md = format_as_markdown_table(values, header_row, merged)
    
    return md
```

### xlwings の制約と代替案

**重要な制約**: xlwings はExcelアプリケーション（COM/AppleScript）を必要とするため、Linux サーバ上では直接動作しない。

| 実行環境 | 推奨アプローチ |
|---------|--------------|
| **Windows サーバ** | xlwings（フル機能利用可能） |
| **macOS** | xlwings（AppleScript経由） |
| **Linux サーバ** | openpyxl + カスタムロジック（推奨フォールバック） |
| **Docker/Lambda** | openpyxl + MarkItDown の組み合わせ |

```python
# Linux環境向けフォールバック: openpyxl による同等処理

import openpyxl

def excel_to_markdown_openpyxl(filepath: str) -> str:
    """
    openpyxlによるExcel→Markdown変換（Linux互換）
    
    xlwings の代替として、以下の情報を取得可能:
    - セル値・数式
    - セル結合情報（merged_cells）
    - フォント・塗りつぶし色
    - コメント
    - 条件付き書式（一部）
    
    取得できない情報:
    - 計算済み数式の結果（data_only=True で一部対応）
    - グラフ・画像の詳細
    - VBAマクロの内容
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    md_parts = []
    
    for ws in wb.worksheets:
        md_parts.append(f"# シート: {ws.title}\n")
        
        # 結合セル情報の取得
        merged_ranges = list(ws.merged_cells.ranges)
        
        # データ領域の検出
        if ws.max_row is None or ws.max_column is None:
            md_parts.append("（空のシート）\n")
            continue
        
        # テーブル領域をスキャンして分離
        table_regions = scan_table_regions(ws)
        
        for region in table_regions:
            md_parts.append(
                region_to_markdown(ws, region, merged_ranges)
            )
        
        # コメント抽出
        for cell in ws._cells.values():
            if hasattr(cell, 'comment') and cell.comment:
                md_parts.append(
                    f"- **{cell.coordinate}**: {cell.comment.text}\n"
                )
    
    return '\n'.join(md_parts)
```

### MarkItDown の利用箇所

```python
from markitdown import MarkItDown

def convert_with_markitdown(filepath: str) -> str:
    """
    MarkItDown によるMarkdown変換
    対象: Word, PowerPoint, PDF, HTML
    """
    mid = MarkItDown()
    result = mid.convert(filepath)
    return result.text_content
```

### Bugless ファイル用カスタムパーサ

```python
import re

def parse_bugless_file(filepath: str) -> str:
    """
    Bugless（バグレス）ファイルのパース
    
    想定構造:
    - ヘッダ部: プログラム名、作成者、作成日等のメタ情報
    - 設計書部: 日本語による処理概要・仕様記述
    - ソース部: COBOLソースコード
    - 各セクションは特定の区切り文字やキーワードで分離
    
    ※ 実際のBugless形式の仕様を確認の上、パターンを調整する必要がある
    """
    
    with open(filepath, 'r', encoding='shift_jis', errors='replace') as f:
        content = f.read()
    
    md_parts = []
    
    # メタ情報の抽出（例: ヘッダ行パターン）
    meta = extract_bugless_metadata(content)
    if meta:
        md_parts.append("## メタ情報\n")
        for k, v in meta.items():
            md_parts.append(f"- **{k}**: {v}")
    
    # 設計書セクションの抽出
    design_sections = extract_design_sections(content)
    for section in design_sections:
        md_parts.append(f"\n## {section['title']}\n")
        md_parts.append(section['body'])
    
    # COBOLコードセクションの抽出
    cobol_sections = extract_cobol_sections(content)
    for code in cobol_sections:
        md_parts.append(f"\n## COBOL: {code['name']}\n")
        md_parts.append(f"```cobol\n{code['source']}\n```\n")
    
    return '\n'.join(md_parts)
```

---

## 6. Phase 5: 構造化 & ノイズ除去

### 3つのアプローチの比較

ここで「いつLLMを使うか」について3つの選択肢を比較する。

#### アプローチA: Markdown化後にバッチでLLM処理（現在の案）

```
[各形式] → [Markdown変換] → [全Markdownに対してOpenAI一括処理]
```

**メリット**: パイプラインがシンプル、各ステップの責務が明確  
**デメリット**: Markdown変換で情報が欠落した場合、LLMで補完できない

#### アプローチB: 変換中にリアルタイムでLLM処理

```
[各形式] → [セルやページ単位でOpenAI呼び出し] → [構造化Markdown直接出力]
```

**メリット**: 元ファイルの構造情報を活かしてLLMが判断できる  
**デメリット**: API呼び出し回数が爆発的に増加、コスト大

#### アプローチC: ★ ハイブリッド（推奨）

```
[各形式] → [ルールベースMarkdown変換]
               │
               ├── 品質OKのもの → そのまま出力（LLM不使用）
               │
               └── 品質NGのもの → OpenAIで構造化補正
```

### 推奨: アプローチC の詳細設計

```python
import openai
from enum import Enum

class MarkdownQuality(Enum):
    HIGH = "high"       # LLM処理不要
    MEDIUM = "medium"   # 軽量なLLM処理で補正
    LOW = "low"         # フルLLM処理が必要


def assess_markdown_quality(md_content: str, source_type: str) -> MarkdownQuality:
    """
    Markdown変換結果の品質を自動判定
    
    判定基準:
    - テーブルの整合性（列数が揃っているか）
    - 見出し構造の存在
    - 文字化け・ガベージの有無
    - 情報密度（テキスト量 vs ノイズ量）
    """
    
    score = 100
    
    # テーブル崩れチェック
    if has_broken_tables(md_content):
        score -= 30
    
    # 文字化け検出（制御文字、異常なUnicodeパターン）
    garbage_ratio = detect_garbage_ratio(md_content)
    score -= int(garbage_ratio * 50)
    
    # 構造の存在チェック
    if not has_heading_structure(md_content):
        score -= 10
    
    # ノイズ比率チェック（罫線記号、装飾文字など）
    noise_ratio = detect_noise_ratio(md_content)
    score -= int(noise_ratio * 30)
    
    if score >= 70:
        return MarkdownQuality.HIGH
    elif score >= 40:
        return MarkdownQuality.MEDIUM
    else:
        return MarkdownQuality.LOW


def phase5_structure_and_denoise(
    md_content: str,
    quality: MarkdownQuality,
    source_metadata: dict
) -> str:
    """品質に応じた構造化処理"""
    
    if quality == MarkdownQuality.HIGH:
        # ルールベースの軽微な補正のみ
        return rule_based_cleanup(md_content)
    
    elif quality == MarkdownQuality.MEDIUM:
        # 部分的なLLM補正
        return partial_llm_cleanup(md_content, source_metadata)
    
    else:  # LOW
        # フルLLM構造化
        return full_llm_restructure(md_content, source_metadata)


def rule_based_cleanup(md_content: str) -> str:
    """
    LLMを使わないルールベースのクリーンアップ
    
    処理内容:
    - 連続する空行の圧縮（3行以上→2行）
    - 不要な装飾文字の除去（━━━, ═══ 等）
    - 全角スペースの正規化
    - 制御文字の除去
    - Markdownテーブルの列揃え
    """
    import re
    
    # 連続空行の圧縮
    md_content = re.sub(r'\n{3,}', '\n\n', md_content)
    
    # 装飾罫線の除去
    md_content = re.sub(r'^[━═─┌┐└┘├┤┬┴┼│]+$', '', md_content, flags=re.MULTILINE)
    
    # 制御文字除去（改行・タブは保持）
    md_content = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', md_content)
    
    return md_content.strip()


def partial_llm_cleanup(md_content: str, metadata: dict) -> str:
    """
    品質MEDIUM向け: 問題箇所のみLLMで補正
    """
    
    # 問題セクションを特定
    sections = split_into_sections(md_content)
    cleaned_sections = []
    
    for section in sections:
        if needs_llm_help(section):
            cleaned = call_openai_for_cleanup(section, metadata)
            cleaned_sections.append(cleaned)
        else:
            cleaned_sections.append(rule_based_cleanup(section))
    
    return '\n\n'.join(cleaned_sections)


def full_llm_restructure(md_content: str, metadata: dict) -> str:
    """
    品質LOW向け: LLMによるフル構造化
    """
    
    client = openai.OpenAI()
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",  # コスト最適化のためminiを使用
        messages=[
            {
                "role": "system",
                "content": """あなたはドキュメント構造化の専門家です。
以下のMarkdownテキストを整形してください。

ルール:
1. 元の情報を一切削除しない（ノイズのみ除去）
2. 適切な見出し階層（#, ##, ###）を付与
3. テーブルが崩れている場合は再構成
4. 箇条書きが適切な箇所はリスト化
5. コードブロックが含まれる場合は適切な言語タグを付与
6. メタ情報（作成日、作成者等）は冒頭にYAML Front Matterとして整理
7. 出力はMarkdown形式のみ（説明文は不要）"""
            },
            {
                "role": "user",
                "content": f"ファイル種別: {metadata.get('source_type', '不明')}\n"
                          f"ファイル名: {metadata.get('filename', '不明')}\n\n"
                          f"---\n\n{md_content}"
            }
        ],
        max_tokens=16000,
        temperature=0.1  # 創造性を最小限に
    )
    
    return response.choices[0].message.content
```

### LLMをExcel処理中に呼ぶ場合（アプローチBの部分適用）

特定のケースではxlwings処理中にLLMを呼ぶことが有効:

```python
def smart_excel_conversion(sheet, region) -> str:
    """
    複雑なExcelレイアウト（BAGLES等）では、
    テーブル領域の検出結果をLLMに渡して、
    人間可読な構造化を支援させる。
    
    ただし、これはシート全体ではなく
    「テーブル領域の境界判定」に限定して使用する。
    """
    
    raw_values = sheet.range(region).value
    cell_colors = get_cell_colors(sheet, region)
    merged_info = get_merged_info(sheet, region)
    
    # セル情報をコンパクトな表現にして LLM に送る
    compact_repr = format_compact_representation(
        raw_values, cell_colors, merged_info
    )
    
    # LLMにテーブル構造の解釈を依頼
    client = openai.OpenAI()
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "system",
                "content": """あなたはExcel構造解析の専門家です。
以下のセルデータを分析し、テーブルの論理構造を判定してください。

出力形式（JSON）:
{
  "tables": [
    {
      "header_rows": [0, 1],
      "data_start_row": 2,
      "key_columns": [0, 1],
      "description": "テーブルの説明"
    }
  ],
  "non_table_areas": [
    {"rows": [10, 15], "type": "free_text", "description": "備考欄"}
  ]
}"""
            },
            {"role": "user", "content": compact_repr}
        ],
        response_format={"type": "json_object"},
        temperature=0
    )
    
    structure = json.loads(response.choices[0].message.content)
    
    # 判定結果に基づいてMarkdown生成（ルールベース）
    return generate_markdown_from_structure(raw_values, structure)
```

**この手法を使うべきケース**:
- BAGLES定義書のように、1シートに複数の異なるテーブルが配置されている
- ヘッダ行が複数段で結合セルが多用されている
- テーブル間にフリーテキストの説明が挟まっている

**使うべきでないケース（コスト面で非効率）**:
- 単純な一覧表（ヘッダ+データ行の典型的なテーブル）
- テキスト主体の文書
- 1シート1テーブルの整形されたExcel

---

## 7. Phase 6: チャンキング & 出力

### すべてをチャンク化しない設計

```python
from dataclasses import dataclass
from typing import Optional

@dataclass
class ProcessedDocument:
    """処理済みドキュメント"""
    source_path: str          # 元ファイルパス
    source_type: str          # ファイル形式
    markdown_content: str     # Markdown化された内容
    metadata: dict           # メタデータ
    needs_chunking: bool     # チャンキングが必要か
    chunk_strategy: Optional[str]  # チャンキング戦略


def determine_chunking_strategy(doc: ProcessedDocument) -> ProcessedDocument:
    """
    ファイルの性質に応じてチャンキング戦略を決定する
    
    チャンキングが必要なケース:
    - RAGで検索対象となるドキュメント
    - トークン上限を超える長文
    - 複数の独立したトピックを含む文書
    
    チャンキングが不要なケース:
    - 短いテキスト（500トークン未満）
    - 設定ファイルや定義書（分割すると意味が崩れる）
    - COBOLソースコード（関数/セクション単位で既に構造化済み）
    """
    
    tokens = count_tokens(doc.markdown_content)
    
    # 短いドキュメントはチャンキング不要
    if tokens < 500:
        doc.needs_chunking = False
        doc.chunk_strategy = None
        return doc
    
    # ファイル形式別の判定
    CHUNKING_RULES = {
        'excel': {
            'needs_chunking': True,
            'strategy': 'by_sheet_and_table',
            'reason': 'シート・テーブル単位が自然な分割境界'
        },
        'word': {
            'needs_chunking': True,
            'strategy': 'by_heading',
            'reason': '見出し階層で分割'
        },
        'pdf': {
            'needs_chunking': True,
            'strategy': 'by_page_or_section',
            'reason': 'ページまたはセクション単位'
        },
        'pptx': {
            'needs_chunking': True,
            'strategy': 'by_slide',
            'reason': 'スライド単位'
        },
        'bugless': {
            'needs_chunking': True,
            'strategy': 'by_program_section',
            'reason': 'プログラム・セクション単位'
        },
        'text': {
            'needs_chunking': tokens > 2000,
            'strategy': 'semantic' if tokens > 2000 else None,
            'reason': 'セマンティック分割'
        },
    }
    
    rule = CHUNKING_RULES.get(doc.source_type, {
        'needs_chunking': tokens > 2000,
        'strategy': 'fixed_size_with_overlap'
    })
    
    doc.needs_chunking = rule['needs_chunking']
    doc.chunk_strategy = rule.get('strategy')
    return doc


def chunk_document(doc: ProcessedDocument) -> list[dict]:
    """
    戦略に応じたチャンキングを実行
    """
    
    if not doc.needs_chunking:
        return [{
            'content': doc.markdown_content,
            'metadata': {
                **doc.metadata,
                'chunk_index': 0,
                'total_chunks': 1,
            }
        }]
    
    CHUNKERS = {
        'by_sheet_and_table': chunk_by_sheet_and_table,
        'by_heading': chunk_by_heading,
        'by_page_or_section': chunk_by_page_or_section,
        'by_slide': chunk_by_slide,
        'by_program_section': chunk_by_program_section,
        'semantic': semantic_chunk,
        'fixed_size_with_overlap': fixed_size_chunk,
    }
    
    chunker = CHUNKERS[doc.chunk_strategy]
    chunks = chunker(doc.markdown_content)
    
    # メタデータ付与
    result = []
    for i, chunk in enumerate(chunks):
        result.append({
            'content': chunk['text'],
            'metadata': {
                **doc.metadata,
                'chunk_index': i,
                'total_chunks': len(chunks),
                'chunk_title': chunk.get('title', ''),
                'chunk_strategy': doc.chunk_strategy,
            }
        })
    
    return result
```

### 出力形式

```python
import json

def write_output(chunks: list[dict], output_dir: str, doc_id: str):
    """
    チャンク結果を出力
    
    出力構造:
    output/
    ├── {doc_id}/
    │   ├── metadata.json      # ドキュメント全体のメタデータ
    │   ├── chunk_000.md       # チャンク本文
    │   ├── chunk_001.md
    │   └── chunks_index.json  # 全チャンクのインデックス
    """
    
    doc_dir = Path(output_dir) / doc_id
    doc_dir.mkdir(parents=True, exist_ok=True)
    
    index = []
    for chunk in chunks:
        chunk_file = f"chunk_{chunk['metadata']['chunk_index']:03d}.md"
        
        (doc_dir / chunk_file).write_text(
            chunk['content'], encoding='utf-8'
        )
        
        index.append({
            'file': chunk_file,
            'metadata': chunk['metadata'],
            'token_count': count_tokens(chunk['content']),
        })
    
    (doc_dir / 'chunks_index.json').write_text(
        json.dumps(index, ensure_ascii=False, indent=2),
        encoding='utf-8'
    )
```

---

## 8. 実行エンジン（全体のオーケストレーション）

```python
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

class DocumentPipeline:
    """ドキュメント処理パイプライン"""
    
    def __init__(self, config: dict):
        self.input_dir = config['input_dir']
        self.work_dir = config['work_dir']
        self.output_dir = config['output_dir']
        self.token_limit = config.get('token_limit', 80_000)
        self.openai_model = config.get('openai_model', 'gpt-4o-mini')
    
    def run(self):
        """パイプライン全体を実行"""
        
        logger.info("=== Phase 1: コピー & 準備 ===")
        classified = phase1_copy_and_classify(self.input_dir, self.work_dir)
        logger.info(f"分類結果: {
            {k: len(v) for k, v in classified.items()}
        }")
        
        logger.info("=== Phase 2: フォーマット正規化 ===")
        classified = phase2_normalize_formats(classified, self.work_dir)
        
        logger.info("=== Phase 3: ファイルサイズ検査 & 分割 ===")
        all_files = []
        for ftype, files in classified.items():
            for f in files:
                split_files = phase3_split_if_needed(f, ftype)
                for sf in split_files:
                    all_files.append({'path': sf, 'type': ftype})
        
        logger.info("=== Phase 4: Markdown変換 ===")
        processed_docs = []
        for file_info in all_files:
            md_content = self._convert_to_markdown(file_info)
            processed_docs.append(ProcessedDocument(
                source_path=file_info['path'],
                source_type=file_info['type'],
                markdown_content=md_content,
                metadata=self._build_metadata(file_info),
                needs_chunking=False,
                chunk_strategy=None,
            ))
        
        logger.info("=== Phase 5: 構造化 & ノイズ除去 ===")
        for doc in processed_docs:
            quality = assess_markdown_quality(
                doc.markdown_content, doc.source_type
            )
            logger.info(
                f"  {doc.source_path}: 品質={quality.value}"
            )
            doc.markdown_content = phase5_structure_and_denoise(
                doc.markdown_content, quality, doc.metadata
            )
        
        logger.info("=== Phase 6: チャンキング & 出力 ===")
        for doc in processed_docs:
            doc = determine_chunking_strategy(doc)
            chunks = chunk_document(doc)
            doc_id = self._generate_doc_id(doc)
            write_output(chunks, self.output_dir, doc_id)
            logger.info(
                f"  {doc.source_path}: {len(chunks)}チャンク出力"
            )
        
        logger.info("=== パイプライン完了 ===")
    
    def _convert_to_markdown(self, file_info: dict) -> str:
        """ファイル形式に応じた変換処理のディスパッチ"""
        
        converters = {
            'excel':   self._convert_excel,
            'word':    lambda f: convert_with_markitdown(f),
            'pptx':    lambda f: convert_with_markitdown(f),
            'pdf':     lambda f: convert_with_markitdown(f),
            'bugless': lambda f: parse_bugless_file(f),
            'trf':     lambda f: convert_with_pandoc(f, 'rtf'),
            'text':    lambda f: Path(f).read_text(encoding='utf-8',
                                                     errors='replace'),
        }
        
        converter = converters.get(
            file_info['type'],
            lambda f: convert_with_markitdown(f)  # デフォルト
        )
        return converter(file_info['path'])
    
    def _convert_excel(self, filepath: str) -> str:
        """Excel変換（環境に応じてxlwingsまたはopenpyxlを使用）"""
        try:
            import xlwings
            return excel_to_markdown_xlwings(filepath)
        except ImportError:
            logger.warning(
                "xlwings 未インストール。openpyxl で代替処理します。"
            )
            return excel_to_markdown_openpyxl(filepath)
```

---

## 9. コスト最適化のポイント

| 判断ポイント | 低コスト選択 | 高品質選択 |
|-------------|------------|-----------|
| LLM構造化 | 品質HIGH/MEDIUMはルールベースのみ | 全ファイルにLLM処理 |
| モデル選択 | gpt-4o-mini（入力$0.15/100万トークン） | gpt-4o（入力$2.50/100万トークン） |
| テーブル解析 | openpyxl + ルールベース | xlwings + LLMテーブル構造解析 |
| チャンキング | 固定サイズ+オーバーラップ | セマンティックチャンキング |

### コスト試算例（1000ファイル処理の場合）

```
想定: 平均ファイルサイズ 50KB → 平均 15,000トークン/ファイル

Phase 5（構造化）で LLM を使うファイル: 全体の30%と仮定
  = 300 ファイル × 15,000 トークン = 4,500,000 入力トークン

gpt-4o-mini の場合:
  入力: 4,500,000 × $0.15/1M = $0.675
  出力: 4,500,000 × $0.60/1M = $2.70
  合計: 約 $3.38

gpt-4o の場合:
  入力: 4,500,000 × $2.50/1M = $11.25
  出力: 4,500,000 × $10.00/1M = $45.00
  合計: 約 $56.25
```

---

## 10. 次のステップ

1. **Bugless ファイルのサンプル分析** — 実ファイルの構造を確認し、パーサのパターンを確定
2. **TRF ファイルの変換テスト** — LibreOffice での変換可否を検証
3. **Excel（BAGLES）サンプルでの PoC** — openpyxl/xlwings でテーブル検出精度を確認
4. **品質判定ロジックのチューニング** — 実データでのスコアリング閾値調整
5. **実行環境の決定** — Windows/Linux/Docker のどれで運用するかにより Phase 4 の実装が変わる
