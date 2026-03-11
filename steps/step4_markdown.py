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

MARKDOWN_IMAGE_PATTERN = re.compile(r"!\[(?P<alt>[^\]]*)\]\((?P<src>[^)]+)\)", re.IGNORECASE)
HTML_IMAGE_PATTERN = re.compile(r"<img\b[^>]*>", re.IGNORECASE)


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


def _replace_visuals_with_placeholders(md_content: str) -> tuple[str, int]:
    """画像参照を [図N] プレースホルダへ置換する。"""
    counter = 0

    def build_placeholder(alt_text: str = "") -> str:
        nonlocal counter
        counter += 1
        alt_text = alt_text.strip()
        return f"[図{counter}: {alt_text}]" if alt_text else f"[図{counter}]"

    def replace_markdown(match: re.Match[str]) -> str:
        return build_placeholder(match.group("alt") or "")

    def replace_html(match: re.Match[str]) -> str:
        alt_match = re.search(r'alt=["\']([^"\']*)["\']', match.group(0), re.IGNORECASE)
        alt_text = alt_match.group(1) if alt_match else ""
        return build_placeholder(alt_text)

    md_content = MARKDOWN_IMAGE_PATTERN.sub(replace_markdown, md_content)
    md_content = HTML_IMAGE_PATTERN.sub(replace_html, md_content)
    return md_content, counter


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

    def _load_origin_map(self) -> dict[str, dict]:
        origin_map_path = self.config.paths.step_dir(2) / "normalization_log.json"
        if not origin_map_path.exists():
            return {}
        try:
            records = json.loads(origin_map_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {}
        return {
            entry["output"]: entry
            for entry in records
            if isinstance(entry, dict) and entry.get("output")
        }

    def execute(self):
        prev_dir = self.config.paths.step_dir(3) / "files"
        if not prev_dir.exists():
            raise FileNotFoundError("Step 3 の出力が見つかりません。")

        out_dir = self.step_dir / "files"
        out_dir.mkdir(parents=True, exist_ok=True)

        conversion_log = []
        origin_map = self._load_origin_map()
        sources = [
            src for src in sorted(prev_dir.rglob("*"))
            if src.is_file() and not src.name.startswith(".")
        ]
        total = len(sources)
        self.log_target_count(total, "Markdown 変換対象")

        for index, src in enumerate(sources, start=1):
            rel = src.relative_to(prev_dir)
            rel_str = str(rel)
            ftype = detect_file_type(src)
            source_meta = origin_map.get(rel_str, {})
            source_type = source_meta.get("source_type", ftype)
            out_rel = rel.with_suffix(".md")
            dest = out_dir / out_rel
            self.log_file_start(index, total, rel_str, "Markdown 変換")
            self.log_file_progress(index, total, rel_str, f"入力種別={ftype}, 元形式={source_type}")
            if source_type in {"rtf", "trf"}:
                self.log_file_progress(
                    index,
                    total,
                    rel_str,
                    "RTF/TRF 由来です。本文抽出を優先し、図やフロー図は [図N] プレースホルダへ寄せます",
                )

            if not self.should_process_file(rel_str, dest):
                self.log_file_skip(index, total, rel_str)
                conversion_log.append({
                    "file": rel_str,
                    "type": ftype,
                    "source_type": source_type,
                    "status": "skipped",
                })
                continue

            converter = CONVERTER_MAP.get(ftype, _convert_text)

            try:
                md_content = converter(src)
                md_content, visual_count = _replace_visuals_with_placeholders(md_content)
                if visual_count:
                    self.log_file_progress(
                        index,
                        total,
                        rel_str,
                        f"図・画像を {visual_count} 件検出し、プレースホルダに置換しました",
                    )
                if source_type in {"rtf", "trf"} and not md_content.strip():
                    self.log_file_progress(
                        index,
                        total,
                        rel_str,
                        "本文を十分に抽出できませんでした。画像/OCR 併用候補として確認してください",
                    )
                dest.parent.mkdir(parents=True, exist_ok=True)
                dest.write_text(md_content, encoding="utf-8")
                self.manifest.mark_file_done(rel_str)
                self.log_file_done(index, total, rel_str, f"{out_rel} を出力しました")
                conversion_log.append({
                    "file": rel_str,
                    "type": ftype,
                    "source_type": source_type,
                    "output": str(out_rel),
                    "status": "ok",
                    "visual_assets": visual_count,
                    "requires_visual_review": source_type in {"rtf", "trf"} or visual_count > 0,
                })
            except Exception as e:
                self.log_file_failed(index, total, rel_str, e)
                self.manifest.mark_file_failed(rel_str, str(e))
                conversion_log.append({
                    "file": rel_str,
                    "type": ftype,
                    "source_type": source_type,
                    "status": "failed",
                    "error": str(e),
                })

        (self.step_dir / "conversion_log.json").write_text(
            json.dumps(conversion_log, ensure_ascii=False, indent=2), encoding="utf-8",
        )
        ok = sum(1 for e in conversion_log if e["status"] == "ok")
        fail = sum(1 for e in conversion_log if e["status"] == "failed")
        logger.info("  変換完了: %d 件成功, %d 件失敗", ok, fail)
