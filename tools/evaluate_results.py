"""パイプライン実行結果の評価スクリプト

パイプライン実行後の中間 JSON と最終 Markdown を読み込み、
確認観点ごとに自動チェックしてレポートを出力する。

使い方:
  # Step2 → Step3 の結果を評価
  python tools/evaluate_results.py

  # パスを指定
  python tools/evaluate_results.py --extracted intermediate/02_extracted --transformed intermediate/03_transformed --output output

  # テキストレポートを保存
  python tools/evaluate_results.py -o evaluation_report.txt
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path


# ---------------------------------------------------------------------------
# 評価結果のデータ構造
# ---------------------------------------------------------------------------

@dataclass
class FileEvaluation:
    """1ファイルの評価結果"""
    file_name: str
    checks: list[tuple[str, bool, str]] = field(default_factory=list)
    # (観点名, pass/fail, 詳細メッセージ)

    def add(self, check_name: str, passed: bool, detail: str = "") -> None:
        self.checks.append((check_name, passed, detail))

    @property
    def pass_count(self) -> int:
        return sum(1 for _, p, _ in self.checks if p)

    @property
    def fail_count(self) -> int:
        return sum(1 for _, p, _ in self.checks if not p)


@dataclass
class EvaluationReport:
    """全体の評価レポート"""
    file_evals: list[FileEvaluation] = field(default_factory=list)
    summary_checks: list[tuple[str, bool, str]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Step2 (中間 JSON) の評価
# ---------------------------------------------------------------------------

def evaluate_json(json_path: Path) -> FileEvaluation:
    """1つの中間 JSON ファイルを評価する。"""
    ev = FileEvaluation(file_name=json_path.name)

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        ev.add("JSON読み込み", False, str(e))
        return ev

    ev.add("JSON読み込み", True)

    # --- メタデータ ---
    meta = data.get("metadata", {})
    ev.add("metadata存在", bool(meta), f"keys={list(meta.keys())}")
    ev.add("source_path", bool(meta.get("source_path")), meta.get("source_path", ""))
    ev.add("source_ext", bool(meta.get("source_ext")), meta.get("source_ext", ""))
    ev.add("doc_role_guess",
           meta.get("doc_role_guess") in ("spec_body", "change_history", "mixed", "unknown"),
           meta.get("doc_role_guess", "未設定"))

    # --- 要素 ---
    doc = data.get("document", {})
    elements = doc.get("elements", [])
    ev.add("elements存在", len(elements) > 0, f"{len(elements)}個")

    # 要素種別の内訳
    type_counts = Counter(e.get("type") for e in elements)
    ev.add("要素種別",
           True,
           " / ".join(f"{t}={c}" for t, c in sorted(type_counts.items())))

    # --- 疑似見出し検出 ---
    headings = [e for e in elements if e.get("type") == "heading"]
    ev.add("見出し検出", True, f"{len(headings)}個")
    if headings:
        methods = Counter(h["content"].get("detection_method", "?") for h in headings)
        ev.add("見出し検出方法", True, str(dict(methods)))

        levels = Counter(h["content"].get("level") for h in headings)
        ev.add("見出しレベル分布", True, str(dict(levels)))

    # --- 表の抽出 ---
    tables = [e for e in elements if e.get("type") == "table"]
    ev.add("表抽出", True, f"{len(tables)}個")

    if tables:
        # 結合セル
        merged = sum(1 for t in tables if t["content"].get("has_merged_cells"))
        ev.add("結合セルあり表", True, f"{merged}/{len(tables)}個")

        # 変更履歴テーブル
        ch_tables = [t for t in tables
                     if t["content"].get("fallback_reason") == "change_history_table"]
        ev.add("変更履歴テーブル", True, f"{len(ch_tables)}個")

        # 信頼度
        conf_counts = Counter(t["content"].get("confidence", "?") for t in tables)
        ev.add("表の信頼度", True, str(dict(conf_counts)))

        # セルの中身確認（空テーブルがないか）
        empty_tables = 0
        for t in tables:
            rows = t["content"].get("rows", [])
            if not rows or all(not any(c.get("text") for c in row) for row in rows):
                empty_tables += 1
        ev.add("空テーブルなし", empty_tables == 0, f"空テーブル={empty_tables}個")

    # --- 図形の抽出 ---
    shapes = [e for e in elements if e.get("type") == "shape"]
    ev.add("図形抽出", True, f"{len(shapes)}個")
    if shapes:
        with_text = sum(1 for s in shapes if s["content"].get("texts"))
        ev.add("テキスト付き図形", True, f"{with_text}/{len(shapes)}個")

    # --- 要素の出現順序 ---
    indices = [e.get("source_index", -1) for e in elements]
    is_ordered = indices == sorted(indices)
    ev.add("要素順序が単調増加", is_ordered,
           f"min={min(indices) if indices else '-'}, max={max(indices) if indices else '-'}")

    return ev


# ---------------------------------------------------------------------------
# Step3 (Markdown) の評価
# ---------------------------------------------------------------------------

def evaluate_markdown(md_path: Path) -> FileEvaluation:
    """1つの Markdown ファイルを評価する。"""
    ev = FileEvaluation(file_name=md_path.name)

    try:
        text = md_path.read_text(encoding="utf-8")
    except Exception as e:
        ev.add("MD読み込み", False, str(e))
        return ev

    ev.add("MD読み込み", True)
    size_kb = len(text.encode("utf-8")) / 1024
    ev.add("ファイルサイズ", True, f"{size_kb:.1f}KB")

    # 15MB チェック
    ev.add("15MB以下", size_kb < 15 * 1024,
           f"{size_kb:.1f}KB" if size_kb < 15 * 1024 else f"超過: {size_kb:.1f}KB")

    lines = text.splitlines()
    ev.add("行数", True, f"{len(lines)}行")

    # --- 見出し ---
    heading_lines = [l for l in lines if l.startswith("#")]
    ev.add("Markdown見出し", True, f"{len(heading_lines)}個")
    if heading_lines:
        level_counts = Counter(len(l.split()[0]) for l in heading_lines if l.strip())
        ev.add("見出しレベル分布", True,
               " / ".join(f"H{lv}={c}" for lv, c in sorted(level_counts.items())))

    # --- ラベル付きテキスト ---
    row_markers = [l for l in lines if l.strip().startswith("[行")]
    ev.add("ラベル付きテキスト行", True, f"{len(row_markers)}個")

    label_lines = [l for l in lines if ":" in l and l.strip().startswith(" ")]
    ev.add("ラベル:値 行", True, f"{len(label_lines)}個")

    # --- 図形プレースホルダ ---
    shape_lines = [l for l in lines if l.strip().startswith("[図形")]
    ev.add("図形プレースホルダ", True, f"{len(shape_lines)}個")

    # --- マーカーがないこと ---
    has_html_comment = any("<!--" in l for l in lines)
    ev.add("HTMLコメントなし", not has_html_comment,
           "HTMLコメントが検出された" if has_html_comment else "クリーン")

    # --- YAML front matter がないこと ---
    ev.add("YAML front matterなし", not text.startswith("---"),
           "front matter検出" if text.startswith("---") else "クリーン")

    # --- 空行の連続 ---
    max_consecutive_empty = 0
    current_empty = 0
    for l in lines:
        if l.strip() == "":
            current_empty += 1
            max_consecutive_empty = max(max_consecutive_empty, current_empty)
        else:
            current_empty = 0
    ev.add("連続空行", max_consecutive_empty <= 3,
           f"最大連続空行={max_consecutive_empty}")

    return ev


# ---------------------------------------------------------------------------
# ログファイルの評価
# ---------------------------------------------------------------------------

def evaluate_log(log_path: Path, step_name: str) -> list[tuple[str, bool, str]]:
    """ログファイル (JSONL) を評価する。"""
    checks: list[tuple[str, bool, str]] = []

    if not log_path.exists():
        checks.append((f"{step_name}ログ存在", False, f"{log_path} が見つかりません"))
        return checks

    checks.append((f"{step_name}ログ存在", True, str(log_path)))

    entries = []
    for line in log_path.read_text(encoding="utf-8").strip().splitlines():
        try:
            entries.append(json.loads(line))
        except json.JSONDecodeError:
            checks.append((f"{step_name}ログ形式", False, f"パース失敗: {line[:50]}"))
            return checks

    checks.append((f"{step_name}ログ件数", True, f"{len(entries)}件"))

    status_counts = Counter(e.get("status", "?") for e in entries)
    checks.append((f"{step_name}ステータス", True,
                    " / ".join(f"{s}={c}" for s, c in sorted(status_counts.items()))))

    errors = [e for e in entries if e.get("status") == "error"]
    checks.append((f"{step_name}エラーなし", len(errors) == 0,
                    f"{len(errors)}件のエラー" if errors else "クリーン"))

    if errors:
        for err in errors[:5]:  # 最大5件表示
            checks.append((f"  エラー詳細", False,
                            f"{err.get('file_path', '?')}: {err.get('message', '?')}"))

    # 処理時間
    durations = [e.get("duration_sec", 0) for e in entries]
    if durations:
        total = sum(durations)
        avg = total / len(durations)
        checks.append((f"{step_name}処理時間", True,
                        f"合計={total:.1f}s, 平均={avg:.2f}s, 最大={max(durations):.1f}s"))

    return checks


# ---------------------------------------------------------------------------
# レポート生成
# ---------------------------------------------------------------------------

def build_report(
    extracted_dir: Path,
    transformed_dir: Path,
    output_dir: Path,
) -> EvaluationReport:
    """全体の評価レポートを構築する。"""
    report = EvaluationReport()

    # --- Step2 JSON の評価 ---
    json_files = sorted(extracted_dir.glob("**/*.json"))
    json_files = [f for f in json_files if f.name != "extract_log.jsonl"]

    report.summary_checks.append(
        ("Step2 JSON ファイル数", len(json_files) > 0, f"{len(json_files)}個")
    )

    for jf in json_files:
        report.file_evals.append(evaluate_json(jf))

    # --- Step3 Markdown の評価 ---
    md_files = sorted(transformed_dir.glob("**/*.md"))

    report.summary_checks.append(
        ("Step3 Markdown ファイル数", len(md_files) > 0, f"{len(md_files)}個")
    )

    for mf in md_files:
        report.file_evals.append(evaluate_markdown(mf))

    # --- output/ の確認 ---
    output_mds = sorted(output_dir.glob("**/*.md")) if output_dir.exists() else []
    report.summary_checks.append(
        ("output/ ファイル数", len(output_mds) > 0, f"{len(output_mds)}個")
    )
    report.summary_checks.append(
        ("output/ と transformed/ の一致",
         len(output_mds) >= len(md_files),
         f"output={len(output_mds)}, transformed={len(md_files)}")
    )

    # --- ログファイルの評価 ---
    extract_log = extracted_dir / "extract_log.jsonl"
    report.summary_checks.extend(evaluate_log(extract_log, "Step2"))

    transform_log = transformed_dir / "transform_log.jsonl"
    report.summary_checks.extend(evaluate_log(transform_log, "Step3"))

    # --- doc_role 分布 ---
    roles: list[str] = []
    for jf in json_files:
        try:
            with open(jf, "r", encoding="utf-8") as f:
                data = json.load(f)
            roles.append(data.get("metadata", {}).get("doc_role_guess", "?"))
        except Exception:
            pass
    if roles:
        role_counts = Counter(roles)
        report.summary_checks.append(
            ("doc_role 分布", True,
             " / ".join(f"{r}={c}" for r, c in sorted(role_counts.items())))
        )

    return report


def _find_latest_run(intermediate_base: Path) -> str:
    """intermediate/ 配下の最新タイムスタンプフォルダを検出する。"""
    if not intermediate_base.exists():
        return ""
    # YYYYMMDD_HHMMSS 形式のフォルダを探す
    candidates = []
    for d in intermediate_base.iterdir():
        if d.is_dir() and len(d.name) == 15 and d.name[8] == "_":
            candidates.append(d.name)
    if not candidates:
        return ""
    return sorted(candidates)[-1]  # 辞書順で最新


def format_report(report: EvaluationReport, run_id: str = "") -> str:
    """レポートをテキストに整形する。"""
    lines: list[str] = []
    total_pass = 0
    total_fail = 0

    lines.append("=" * 70)
    title = "パイプライン実行結果 評価レポート"
    if run_id:
        title += f" [{run_id}]"
    lines.append(title)
    lines.append("=" * 70)
    lines.append("")

    # --- 全体サマリー ---
    lines.append("■ 全体サマリー")
    lines.append("-" * 50)
    for name, passed, detail in report.summary_checks:
        mark = "OK" if passed else "NG"
        lines.append(f"  {mark} {name}: {detail}")
        if passed:
            total_pass += 1
        else:
            total_fail += 1
    lines.append("")

    # --- ファイルごとの評価 ---
    lines.append("■ ファイルごとの評価")
    lines.append("-" * 50)

    for ev in report.file_evals:
        lines.append(f"\n  [{ev.file_name}] ({ev.pass_count}pass / {ev.fail_count}fail)")
        for name, passed, detail in ev.checks:
            mark = "OK" if passed else "NG"
            lines.append(f"    {mark} {name}: {detail}")
            if passed:
                total_pass += 1
            else:
                total_fail += 1

    lines.append("")
    lines.append("=" * 70)
    lines.append(f"合計: {total_pass} pass / {total_fail} fail")
    if total_fail == 0:
        lines.append("結果: 全チェック通過")
    else:
        lines.append(f"結果: {total_fail} 件の問題あり — 上記の ✗ を確認してください")
    lines.append("=" * 70)

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description="パイプライン実行結果の評価",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--run-id",
        default="",
        help="評価対象の実行ID。省略時は最新の実行を自動検出",
    )
    parser.add_argument(
        "--intermediate",
        type=Path, default=Path("intermediate"),
        help="中間成果物ベースフォルダ (default: intermediate)",
    )
    parser.add_argument(
        "--output",
        type=Path, default=Path("output"),
        help="最終出力ベースフォルダ (default: output)",
    )
    parser.add_argument(
        "-o", "--report-file",
        type=Path, default=None,
        help="レポートをファイルに保存 (指定しなければ標準出力のみ)",
    )

    args = parser.parse_args()

    # run_id の解決: 指定なしなら最新のタイムスタンプフォルダを検出
    run_id = args.run_id
    if not run_id:
        run_id = _find_latest_run(args.intermediate)
        if not run_id:
            print("エラー: intermediate/ に実行結果が見つかりません。", file=sys.stderr)
            print("パイプラインを先に実行してください:", file=sys.stderr)
            print("  python -m src.main --steps 2-3", file=sys.stderr)
            return 1
        print(f"最新の実行を検出: {run_id}")

    extracted_dir = args.intermediate / run_id / "02_extracted"
    transformed_dir = args.intermediate / run_id / "03_transformed"
    output_dir = args.output / run_id

    # 存在チェック
    missing = []
    if not extracted_dir.exists():
        missing.append(f"extracted: {extracted_dir}")
    if not transformed_dir.exists():
        missing.append(f"transformed: {transformed_dir}")
    if missing:
        print("エラー: 以下のディレクトリが見つかりません:", file=sys.stderr)
        for m in missing:
            print(f"  {m}", file=sys.stderr)
        print(f"\nrun_id={run_id} に対応する結果がありません。", file=sys.stderr)
        return 1

    report = build_report(extracted_dir, transformed_dir, output_dir)
    text = format_report(report, run_id=run_id)

    print(text)

    if args.report_file:
        args.report_file.write_text(text, encoding="utf-8")
        print(f"\nレポートを保存しました: {args.report_file}")

    # 失敗があれば exit code 1
    total_fail = sum(1 for _, p, _ in report.summary_checks if not p)
    total_fail += sum(ev.fail_count for ev in report.file_evals)
    return 1 if total_fail > 0 else 0


if __name__ == "__main__":
    sys.exit(main())
