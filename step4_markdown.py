"""
steps/step4_markdown.py - Phase 4: Markdown 変換

ファイル形式ごとに最適なツールでMarkdownに変換する。
  - Excel  -> xlwings (Windows COM)
  - Word   -> MarkItDown
  - PPT    -> MarkItDown
  - PDF    -> MarkItDown
  - BAGLES -> カスタムパーサ
  - TRF    -> pandoc -> MarkItDown フォールバック
  - テキスト -> そのまま
"""

import json
import logging
import re
from pathlib import Path

from .base import BaseStep
from utils.file_utils import detect_file_type, read_text_file

logger = logging.getLogger(__name__)


# --- 変換関数群 ---

def _convert_excel(filepath: Path) -> str:
    """Excel -> Markdown (xlwings)"""
    import xlwings as xw

    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(filepath.resolve()))
        md_parts = []

        for sheet in wb.sheets:
            md_parts.append(f"## シート: {sheet.name}\n")
            used = sheet.used_range
            if used is None or used.value is None:
                md_parts.append("（空のシート）\n")
                continue

            values = used.value
            if not isinstance(values, list):
                values = [[values]]
            elif values and not isinstance(values[0], list):
                values = [values]
            if not values:
                continue

            headers = [str(v).replace("\n", " ") if v is not None else "" for v in values[0]]
            md_parts.append("| " + " | ".join(headers) + " |")
            md_parts.append("| " + " | ".join(["---"] * len(headers)) + " |")

            for row in values[1:]:
                if row is None:
                    continue
                cells = [str(v).replace("\n", " / ") if v is not None else "" for v in row]
                cells = [c.replace("|", "\\|") for c in cells]
                md_parts.append("| " + " | ".join(cells) + " |")
            md_parts.append("")

        wb.close()
        return "\n".join(md_parts)
    finally:
        app.quit()


def _convert_with_markitdown(filepath: Path) -> str:
    """MarkItDown による汎用変換"""
    from markitdown import MarkItDown
    mid = MarkItDown()
    result = mid.convert(str(filepath))
    return result.text_content


def _convert_bagles(filepath: Path) -> str:
    """BAGLES定義書のMarkdown変換"""
    DEFINITION_KEYWORDS = [
        "チェック更新定義書", "業務用語定義書", "コード定義書",
        "フォーマット定義書", "ロジック定義書", "インタフェース定義書",
        "履歴定義書", "入出力定義書", "マクロ定義書",
    ]
    content = read_text_file(filepath)
    pattern = "|".join(re.escape(kw) for kw in DEFINITION_KEYWORDS)
    matches = list(re.finditer(pattern, content))

    if not matches:
        return f"# BAGLES定義書\n\n```\n{content}\n```"

    md_parts = ["# BAGLES定義書\n"]
    for i, match in enumerate(matches):
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(content)
        block = content[start:end].strip()
        defn_type = match.group()
        md_parts.append(f"## {defn_type}\n")
        lines = block.split("\n")
        if len(lines) > 1:
            md_parts.append("```")
            md_parts.extend(lines[1:])
            md_parts.append("```\n")

    return "\n".join(md_parts)


def _convert_text(filepath: Path) -> str:
    return read_text_file(filepath)


def _convert_trf(filepath: Path) -> str:
    """TRF: pandoc -> MarkItDown -> テキスト読み込み"""
    import subprocess
    try:
        result = subprocess.run(
            ["pandoc", "-f", "rtf", "-t", "markdown", str(filepath)],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout
    except Exception:
        pass
    try:
        return _convert_with_markitdown(filepath)
    except Exception:
        pass
    return read_text_file(filepath)


CONVERTER_MAP = {
    "excel":   _convert_excel,
    "word":    _convert_with_markitdown,
    "pptx":    _convert_with_markitdown,
    "pdf":     _convert_with_markitdown,
    "bagles":  _convert_bagles,
    "trf":     _convert_trf,
    "rtf":     _convert_trf,
    "text":    _convert_text,
}


class Step4Markdown(BaseStep):
    step_number = 4
    step_name = "Markdown 変換"

    def execute(self):
        prev_dir = self.config.paths.step_dir(3) / "files"
        if not prev_dir.exists():
            raise FileNotFoundError("Step 3 の出力が見つかりません。")

        out_dir = self.step_dir / "files"
        out_dir.mkdir(parents=True, exist_ok=True)

        conversion_log = []

        for src in sorted(prev_dir.rglob("*")):
            if not src.is_file() or src.name.startswith("."):
                continue

            rel = src.relative_to(prev_dir)
            ftype = detect_file_type(src)
            out_rel = rel.with_suffix(".md")
            dest = out_dir / out_rel

            if not self.should_process_file(str(rel), dest):
                conversion_log.append({"file": str(rel), "type": ftype, "status": "skipped"})
                continue

            converter = CONVERTER_MAP.get(ftype, _convert_text)

            try:
                md_content = converter(src)
                dest.parent.mkdir(parents=True, exist_ok=True)
                dest.write_text(md_content, encoding="utf-8")
                self.manifest.mark_file_done(str(rel))
                logger.info("  [%s] %s -> %s", ftype, rel, out_rel)
                conversion_log.append({
                    "file": str(rel), "type": ftype,
                    "output": str(out_rel), "status": "ok",
                })
            except Exception as e:
                logger.error("  変換失敗: %s (%s)", rel, e)
                self.manifest.mark_file_failed(str(rel), str(e))
                conversion_log.append({
                    "file": str(rel), "type": ftype,
                    "status": "failed", "error": str(e),
                })

        (self.step_dir / "conversion_log.json").write_text(
            json.dumps(conversion_log, ensure_ascii=False, indent=2), encoding="utf-8",
        )
        ok = sum(1 for e in conversion_log if e["status"] == "ok")
        fail = sum(1 for e in conversion_log if e["status"] == "failed")
        logger.info("  変換完了: %d 件成功, %d 件失敗", ok, fail)
