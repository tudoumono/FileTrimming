"""Dify 自動チャンキング シミュレーター

Dify 1.13.0 の自動チャンキング挙動を再現し、
パイプライン出力の Markdown がどのように分割されるかを検証する。

Dify デフォルト設定 (AUTOMATIC_RULES):
  - delimiter (fixed_separator): "\\n"
  - max_tokens: 500
  - chunk_overlap: 50
  - 再帰分割 separators: ["\\n\\n", "\\n", " ", ""]
  - 前処理: 3+連続改行→2改行、2+連続空白→単一スペース

使い方:
  # 最新 run の全 Markdown を検証
  python tools/simulate_dify_chunking.py

  # 特定ファイルを検証
  python tools/simulate_dify_chunking.py output/20260315_190200/mixed_complex.md

  # カスタム設定で検証
  python tools/simulate_dify_chunking.py --max-tokens 1000 --overlap 100

  # チャンク内容を表示
  python tools/simulate_dify_chunking.py --show-chunks

  # 問題のあるチャンクのみ表示
  python tools/simulate_dify_chunking.py --show-problems
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path


# ---------------------------------------------------------------------------
# Dify 前処理の再現
# ---------------------------------------------------------------------------

def dify_clean(text: str) -> str:
    """Dify の CleanProcessor.clean() の remove_extra_spaces を再現する。

    - 3+ 連続改行 → 2 改行
    - 2+ 連続空白（タブ・スペース・Unicode空白）→ 単一スペース
    """
    # 3+ 連続改行を 2 改行に圧縮
    text = re.sub(r"\n{3,}", "\n\n", text)
    # 2+ 連続空白（改行以外）を単一スペースに圧縮
    text = re.sub(r"[^\S\n]{2,}", " ", text)
    return text


# ---------------------------------------------------------------------------
# チャンク分割の再現
# ---------------------------------------------------------------------------

def _merge_splits(splits: list[str], chunk_size: int, overlap: int) -> list[str]:
    """分割済みテキスト断片をチャンクサイズに収まるようにマージする。

    LangChain の RecursiveCharacterTextSplitter._merge_splits() を簡易再現。
    """
    chunks: list[str] = []
    current: list[str] = []
    current_len = 0

    for s in splits:
        s_len = len(s)
        if current and current_len + s_len + 1 > chunk_size:
            # 現在のバッファをチャンクとして確定
            chunk_text = "\n".join(current)
            if chunk_text.strip():
                chunks.append(chunk_text)
            # オーバーラップ: 末尾から overlap 文字分を残す
            if overlap > 0:
                overlap_parts: list[str] = []
                overlap_len = 0
                for part in reversed(current):
                    if overlap_len + len(part) + 1 > overlap:
                        break
                    overlap_parts.insert(0, part)
                    overlap_len += len(part) + 1
                current = overlap_parts
                current_len = overlap_len
            else:
                current = []
                current_len = 0
        current.append(s)
        current_len += s_len + (1 if current_len > 0 else 0)

    if current:
        chunk_text = "\n".join(current)
        if chunk_text.strip():
            chunks.append(chunk_text)

    return chunks


def _recursive_split(text: str, separators: list[str], chunk_size: int) -> list[str]:
    """再帰的にテキストを分割する。

    LangChain の RecursiveCharacterTextSplitter._split_text() を簡易再現。
    """
    if len(text) <= chunk_size:
        return [text] if text.strip() else []

    # 適用可能なセパレータを探す
    separator = ""
    remaining_seps = separators
    for i, sep in enumerate(separators):
        if sep == "":
            separator = sep
            remaining_seps = []
            break
        if sep in text:
            separator = sep
            remaining_seps = separators[i + 1:]
            break

    if separator:
        splits = text.split(separator)
    else:
        splits = list(text)

    # 各分割片が chunk_size 以内か確認、超過分は再帰分割
    good_splits: list[str] = []
    for s in splits:
        if len(s) <= chunk_size:
            good_splits.append(s)
        elif remaining_seps:
            # さらに細かいセパレータで再帰分割
            sub_splits = _recursive_split(s, remaining_seps, chunk_size)
            good_splits.extend(sub_splits)
        else:
            # 最終手段: そのまま追加（文字レベル分割は省略）
            good_splits.append(s)

    return good_splits


def simulate_chunking(
    text: str,
    max_tokens: int = 500,
    chunk_overlap: int = 50,
    fixed_separator: str = "\n",
    separators: list[str] | None = None,
) -> list[str]:
    """Dify の FixedRecursiveCharacterTextSplitter を再現する。

    Args:
        text: 入力 Markdown テキスト
        max_tokens: 最大チャンクサイズ（文字数で近似）
        chunk_overlap: チャンク間オーバーラップ（文字数で近似）
        fixed_separator: 初回分割のデリミタ
        separators: 再帰分割のセパレータリスト

    Returns:
        チャンクのリスト

    Note:
        Dify の max_tokens はトークン数だが、日本語では 1 トークン ≈ 1-2 文字。
        本スクリプトでは文字数ベースで近似する。
        --char-ratio オプションで調整可能（デフォルト: 1.5 文字/トークン）。
    """
    if separators is None:
        separators = ["\n\n", "\n", " ", ""]

    # Dify 前処理
    text = dify_clean(text)

    # 1. fixed_separator で初回分割
    initial_splits = text.split(fixed_separator) if fixed_separator else [text]

    # 2. 各分割片が max_tokens 以内か確認
    final_splits: list[str] = []
    for split in initial_splits:
        if len(split) <= max_tokens:
            final_splits.append(split)
        else:
            # 再帰的に分割
            sub = _recursive_split(split, separators, max_tokens)
            final_splits.extend(sub)

    # 3. マージ（小さい分割片を結合）
    chunks = _merge_splits(final_splits, max_tokens, chunk_overlap)

    return chunks


# ---------------------------------------------------------------------------
# チャンク品質分析
# ---------------------------------------------------------------------------

@dataclass
class ChunkAnalysis:
    """1チャンクの分析結果"""
    index: int
    text: str
    char_count: int
    line_count: int
    has_heading: bool
    headings: list[str]
    problems: list[str] = field(default_factory=list)


@dataclass
class FileAnalysis:
    """1ファイルの分析結果"""
    file_name: str
    total_chars: int
    chunk_count: int
    chunks: list[ChunkAnalysis] = field(default_factory=list)
    problems: list[str] = field(default_factory=list)


_TABLE_ROW_RE = re.compile(r"^\[行\d+\]$")
_LABEL_VALUE_RE = re.compile(r"^\s+\S+:\s")
_FLOW_RE = re.compile(r"^\[フロー図\]$")
_STEP_RE = re.compile(r"^\s+\d+\.\s")
_BOLD_LINE_RE = re.compile(r"^\*\*(.+)\*\*$")


def _first_non_empty_stripped(lines: list[str]) -> str:
    """最初の非空行を strip 済みで返す。"""
    for line in lines:
        stripped = line.strip()
        if stripped:
            return stripped
    return ""


def _last_non_empty_stripped(lines: list[str]) -> str:
    """最後の非空行を strip 済みで返す。"""
    for line in reversed(lines):
        stripped = line.strip()
        if stripped:
            return stripped
    return ""


def analyze_chunk(index: int, text: str) -> ChunkAnalysis:
    """チャンクの品質を分析する。"""
    lines = text.splitlines()
    headings = [l for l in lines if l.startswith("#")]

    analysis = ChunkAnalysis(
        index=index,
        text=text,
        char_count=len(text),
        line_count=len(lines),
        has_heading=len(headings) > 0,
        headings=headings,
    )

    # --- 問題検出 ---

    # 1. 見出しなしチャンク（コンテキスト不足の可能性）
    if not headings and len(text) > 100:
        analysis.problems.append("見出しなし: チャンクに見出しがなく、コンテキスト不足の可能性")

    # 2. 画像プレースホルダの分断
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("[画像"):
            # 画像の前後の説明文が別チャンクに飛んでいないか
            if i == 0 and index > 0:
                analysis.problems.append("画像分断: [画像] がチャンク先頭に出現（前の説明文と分離）")

    return analysis


def analyze_file(file_path: Path, chunks: list[str]) -> FileAnalysis:
    """ファイル全体のチャンク品質を分析する。"""
    text = file_path.read_text(encoding="utf-8")
    analysis = FileAnalysis(
        file_name=file_path.name,
        total_chars=len(text),
        chunk_count=len(chunks),
    )

    for i, chunk in enumerate(chunks):
        ca = analyze_chunk(i, chunk)
        analysis.chunks.append(ca)

    # チャンク境界を見て構造分断を判定する
    for prev_chunk, next_chunk in zip(analysis.chunks, analysis.chunks[1:]):
        prev_lines = prev_chunk.text.splitlines()
        next_lines = next_chunk.text.splitlines()
        prev_last = _last_non_empty_stripped(prev_lines)
        next_first = _first_non_empty_stripped(next_lines)

        if _TABLE_ROW_RE.match(prev_last) and _LABEL_VALUE_RE.match(next_first):
            prev_chunk.problems.append(
                "表の途中切断: [行N] の後のラベル:値が次のチャンクに分断"
            )

        if _FLOW_RE.match(prev_last) and _STEP_RE.match(next_first):
            prev_chunk.problems.append(
                "フロー図の途中切断: [フロー図] の後のステップが次のチャンクに分断"
            )

        if _BOLD_LINE_RE.match(prev_last) and (
            _TABLE_ROW_RE.match(next_first) or _LABEL_VALUE_RE.match(next_first)
        ):
            prev_chunk.problems.append(
                "キャプション分断: 太字キャプションが表と別チャンクに分離"
            )

    # ファイルレベルの問題集計
    problem_chunks = [c for c in analysis.chunks if c.problems]
    if problem_chunks:
        analysis.problems.append(
            f"{len(problem_chunks)}/{len(chunks)} チャンクに問題あり"
        )

    headingless = [c for c in analysis.chunks if not c.has_heading and c.char_count > 100]
    if headingless:
        analysis.problems.append(
            f"{len(headingless)}/{len(chunks)} チャンクに見出しなし"
        )

    return analysis


# ---------------------------------------------------------------------------
# レポート出力
# ---------------------------------------------------------------------------

def format_report(
    analyses: list[FileAnalysis],
    max_tokens: int,
    chunk_overlap: int,
    char_ratio: float,
    show_chunks: bool = False,
    show_problems: bool = False,
) -> str:
    """分析結果をテキストレポートに整形する。"""
    lines: list[str] = []
    lines.append("=" * 70)
    lines.append("Dify 自動チャンキング シミュレーション結果")
    lines.append("=" * 70)
    lines.append("")
    lines.append(f"設定: max_tokens={max_tokens}, chunk_overlap={chunk_overlap}, "
                 f"char_ratio={char_ratio}")
    lines.append(f"  -> 文字数換算: max_chars={int(max_tokens * char_ratio)}, "
                 f"overlap_chars={int(chunk_overlap * char_ratio)}")
    lines.append("")

    total_files = len(analyses)
    total_chunks = sum(a.chunk_count for a in analyses)
    total_problems = sum(len(c.problems) for a in analyses for c in a.chunks)
    problem_files = sum(1 for a in analyses if a.problems)

    lines.append(f"全体: {total_files} ファイル, {total_chunks} チャンク, "
                 f"{total_problems} 問題")
    if problem_files:
        lines.append(f"  問題ありファイル: {problem_files}/{total_files}")
    lines.append("")

    for analysis in analyses:
        # ファイルヘッダー
        char_sizes = [c.char_count for c in analysis.chunks]
        avg_size = sum(char_sizes) / len(char_sizes) if char_sizes else 0
        max_size = max(char_sizes) if char_sizes else 0
        min_size = min(char_sizes) if char_sizes else 0

        problem_count = sum(len(c.problems) for c in analysis.chunks)
        status = "OK" if problem_count == 0 else f"NG({problem_count}件)"

        lines.append(f"[{analysis.file_name}] {status}")
        lines.append(f"  元サイズ: {analysis.total_chars}文字 -> {analysis.chunk_count} チャンク")
        lines.append(f"  チャンクサイズ: 平均={avg_size:.0f}, 最小={min_size}, 最大={max_size}")

        # 問題サマリー
        if analysis.problems:
            for p in analysis.problems:
                lines.append(f"  * {p}")

        # チャンク詳細
        if show_chunks or show_problems:
            for ca in analysis.chunks:
                if show_problems and not ca.problems:
                    continue
                lines.append(f"  --- チャンク {ca.index + 1}/{analysis.chunk_count} "
                             f"({ca.char_count}文字, {ca.line_count}行) ---")
                if ca.headings:
                    for h in ca.headings:
                        lines.append(f"    見出し: {h}")
                if ca.problems:
                    for p in ca.problems:
                        lines.append(f"    !! {p}")
                if show_chunks:
                    # チャンク本文（先頭200文字）
                    preview = ca.text[:200]
                    if len(ca.text) > 200:
                        preview += "..."
                    for pl in preview.splitlines():
                        lines.append(f"    | {pl}")

        lines.append("")

    lines.append("=" * 70)
    if total_problems == 0:
        lines.append("結論: 全チャンクで問題なし")
    else:
        lines.append(f"結論: {total_problems} 件の潜在的問題を検出")
        # 問題種別の集計
        problem_types: dict[str, int] = {}
        for a in analyses:
            for c in a.chunks:
                for p in c.problems:
                    key = p.split(":")[0]
                    problem_types[key] = problem_types.get(key, 0) + 1
        for ptype, count in sorted(problem_types.items(), key=lambda x: -x[1]):
            lines.append(f"  {ptype}: {count} 件")
    lines.append("=" * 70)

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------

def _find_latest_output(output_base: Path) -> Path | None:
    """output/ 配下の最新タイムスタンプフォルダを検出する。"""
    if not output_base.exists():
        return None
    candidates = []
    for d in output_base.iterdir():
        if d.is_dir() and len(d.name) == 15 and d.name[8] == "_":
            candidates.append(d)
    return sorted(candidates, key=lambda d: d.name)[-1] if candidates else None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Dify 自動チャンキングのシミュレーション",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "files", nargs="*", type=Path,
        help="検証対象の Markdown ファイル。省略時は最新の output/ を使用",
    )
    parser.add_argument(
        "--max-tokens", type=int, default=500,
        help="Dify の max_tokens 設定 (default: 500)",
    )
    parser.add_argument(
        "--overlap", type=int, default=50,
        help="Dify の chunk_overlap 設定 (default: 50)",
    )
    parser.add_argument(
        "--char-ratio", type=float, default=1.5,
        help="1トークンあたりの文字数（日本語向け近似、default: 1.5）",
    )
    parser.add_argument(
        "--show-chunks", action="store_true",
        help="各チャンクの内容をプレビュー表示",
    )
    parser.add_argument(
        "--show-problems", action="store_true",
        help="問題のあるチャンクのみ表示",
    )
    parser.add_argument(
        "-o", "--report-file", type=Path, default=None,
        help="レポートをファイルに保存",
    )

    args = parser.parse_args()

    # ファイル一覧の解決
    md_files: list[Path] = []
    if args.files:
        md_files = [f for f in args.files if f.exists() and f.suffix == ".md"]
        if not md_files:
            print("エラー: 指定されたファイルが見つかりません。", file=sys.stderr)
            return 1
    else:
        output_dir = _find_latest_output(Path("output"))
        if not output_dir:
            print("エラー: output/ に結果が見つかりません。", file=sys.stderr)
            return 1
        md_files = sorted(output_dir.glob("**/*.md"))
        print(f"最新の出力を検出: {output_dir.name} ({len(md_files)} ファイル)")

    # トークン数 → 文字数の換算
    max_chars = int(args.max_tokens * args.char_ratio)
    overlap_chars = int(args.overlap * args.char_ratio)

    # 各ファイルを処理
    analyses: list[FileAnalysis] = []
    for md_file in md_files:
        text = md_file.read_text(encoding="utf-8")
        chunks = simulate_chunking(
            text,
            max_tokens=max_chars,
            chunk_overlap=overlap_chars,
            fixed_separator="\n",
        )
        analysis = analyze_file(md_file, chunks)
        analyses.append(analysis)

    # レポート出力
    report = format_report(
        analyses,
        max_tokens=args.max_tokens,
        chunk_overlap=args.overlap,
        char_ratio=args.char_ratio,
        show_chunks=args.show_chunks,
        show_problems=args.show_problems,
    )

    print(report)

    if args.report_file:
        args.report_file.write_text(report, encoding="utf-8")
        print(f"\nレポートを保存しました: {args.report_file}")

    # 問題があれば exit code 1
    total_problems = sum(len(c.problems) for a in analyses for c in a.chunks)
    return 1 if total_problems > 0 else 0


if __name__ == "__main__":
    sys.exit(main())
