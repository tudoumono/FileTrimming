"""
steps/step6_chunk.py - Phase 6: チャンキング & Dify向け出力

構造化済みMarkdownを適切なチャンクに分割し、
メタデータ付きでDify Knowledge Base向けに出力する。
"""

import json
import logging
import re
import shutil
from pathlib import Path

from .base import BaseStep
from utils.token_utils import estimate_tokens

logger = logging.getLogger(__name__)


def _chunk_by_heading(md_content: str, max_tokens: int) -> list[dict]:
    """見出し（##）単位でチャンキング"""
    sections = re.split(r"(?=^## )", md_content, flags=re.MULTILINE)
    chunks = []

    for section in sections:
        section = section.strip()
        if not section:
            continue

        tokens = estimate_tokens(section)
        if tokens <= max_tokens:
            title_match = re.match(r"^##\s+(.+)$", section, re.MULTILINE)
            title = title_match.group(1) if title_match else ""
            chunks.append({"text": section, "title": title})
        else:
            # セクション内をさらにパラグラフ単位で分割
            paragraphs = section.split("\n\n")
            current = ""
            for para in paragraphs:
                candidate = (current + "\n\n" + para).strip()
                if estimate_tokens(candidate) > max_tokens and current:
                    chunks.append({"text": current, "title": ""})
                    current = para
                else:
                    current = candidate
            if current.strip():
                chunks.append({"text": current, "title": ""})

    return chunks


def _chunk_fixed_size(md_content: str, max_tokens: int, overlap: int = 200) -> list[dict]:
    """固定サイズ＋オーバーラップでチャンキング"""
    lines = md_content.split("\n")
    chunks = []
    start = 0

    while start < len(lines):
        end = start + 1
        current = "\n".join(lines[start:end])

        while end < len(lines) and estimate_tokens(current) < max_tokens:
            end += 1
            current = "\n".join(lines[start:end])

        if estimate_tokens(current) > max_tokens and end > start + 1:
            end -= 1

        chunks.append({"text": "\n".join(lines[start:end]), "title": ""})

        # オーバーラップ分だけ戻す
        overlap_lines = max(1, overlap // 10)
        start = max(start + 1, end - overlap_lines)

    return chunks


class Step6Chunk(BaseStep):
    step_number = 6
    step_name = "チャンキング & 出力"

    # Dify向けのチャンクサイズ（トークン）
    CHUNK_MAX_TOKENS = 2000
    CHUNK_MIN_TOKENS = 100

    def execute(self):
        prev_dir = self.config.paths.step_dir(5) / "files"
        if not prev_dir.exists():
            raise FileNotFoundError("Step 5 の出力が見つかりません。")

        output_dir = self.config.paths.output_dir
        output_dir.mkdir(parents=True, exist_ok=True)

        chunk_log = []
        total_chunks = 0
        sources = sorted(prev_dir.rglob("*.md"))
        total = len(sources)
        self.log_target_count(total, "チャンキング対象")

        for index, src in enumerate(sources, start=1):
            rel = src.relative_to(prev_dir)
            rel_str = str(rel)
            md_content = src.read_text(encoding="utf-8")

            doc_id = rel.stem
            doc_out_dir = output_dir / doc_id
            completion_marker = doc_out_dir / ".done"
            self.log_file_start(index, total, rel_str, "チャンキング")

            # チャンキング戦略: 見出しがあれば見出し単位、なければ固定サイズ
            tokens = estimate_tokens(md_content)

            if tokens <= self.CHUNK_MIN_TOKENS:
                # 短すぎる場合はそのまま1チャンク
                chunks = [{"text": md_content, "title": ""}]
                strategy = "単一チャンク"
            elif re.search(r"^## ", md_content, re.MULTILINE):
                chunks = _chunk_by_heading(md_content, self.CHUNK_MAX_TOKENS)
                strategy = "見出しベース"
            else:
                chunks = _chunk_fixed_size(md_content, self.CHUNK_MAX_TOKENS)
                strategy = "固定サイズ"

            self.log_file_progress(index, total, rel_str, f"推定 {tokens} tokens / 戦略={strategy}")

            if not self.should_process_file(rel_str, completion_marker):
                self.log_file_skip(index, total, rel_str)
                chunk_log.append({
                    "file": rel_str,
                    "doc_id": doc_id,
                    "total_tokens": tokens,
                    "num_chunks": len(chunks),
                    "status": "skipped",
                })
                continue

            if doc_out_dir.exists():
                shutil.rmtree(doc_out_dir)
            doc_out_dir.mkdir(parents=True, exist_ok=True)

            # 各チャンクをファイル出力
            for i, chunk in enumerate(chunks):
                chunk_filename = f"{doc_id}_chunk{i:03d}.md"
                chunk_path = doc_out_dir / chunk_filename

                front_matter = (
                    f"---\n"
                    f"source_file: {rel}\n"
                    f"doc_id: {doc_id}\n"
                    f"chunk_index: {i}\n"
                    f"total_chunks: {len(chunks)}\n"
                    f"chunk_title: {chunk['title']}\n"
                    f"---\n\n"
                )

                chunk_path.write_text(front_matter + chunk["text"], encoding="utf-8")

            completion_marker.write_text("completed\n", encoding="utf-8")
            self.manifest.mark_file_done(rel_str)
            total_chunks += len(chunks)
            self.log_file_done(index, total, rel_str, f"{len(chunks)} チャンクを {doc_out_dir} に出力しました")

            chunk_log.append({
                "file": rel_str,
                "doc_id": doc_id,
                "total_tokens": tokens,
                "num_chunks": len(chunks),
                "strategy": strategy,
            })

        # チャンキングログ
        (self.step_dir / "chunk_log.json").write_text(
            json.dumps(chunk_log, ensure_ascii=False, indent=2), encoding="utf-8",
        )

        logger.info("  合計: %d ファイル -> %d チャンク", len(chunk_log), total_chunks)
        logger.info("  出力先: %s", output_dir)
