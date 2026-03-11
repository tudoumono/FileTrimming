"""
steps/step6_chunk.py - Phase 6: チャンキング & Dify 向け出力

処理内容:
    1. ファイル種別に応じたチャンキング戦略の決定
    2. セマンティックな境界（見出し、テーブル区切り等）でチャンク分割
    3. YAML Front Matter 付き .md ファイルとして出力
    4. 最終出力を output_dir にもコピー
"""

import re
import shutil
from pathlib import Path

from steps.base import BaseStep
from utils.token_counter import count_tokens


# チャンク1つあたりの目安トークン数（Dify のチャンクサイズに合わせて調整）
CHUNK_TARGET_TOKENS = 2000
CHUNK_MAX_TOKENS = 4000


class Step6Chunk(BaseStep):
    step_number = 6
    step_name = "chunked"

    def collect_inputs(self) -> list[Path]:
        step5_dir = self.config.step5_dir
        return sorted(
            f for f in step5_dir.rglob("*.md")
            if f.is_file() and f.name != "manifest.json"
        )

    def process_file(self, source: Path, dest_dir: Path) -> list[Path]:
        md_content = source.read_text(encoding="utf-8")
        tokens = count_tokens(md_content)

        # 短いファイルはチャンキング不要
        if tokens <= CHUNK_MAX_TOKENS:
            chunks = [{"title": source.stem, "content": md_content}]
        else:
            chunks = self._chunk_by_heading(md_content, source.stem)

        # 出力
        outputs: list[Path] = []
        doc_dir = dest_dir / source.stem
        doc_dir.mkdir(parents=True, exist_ok=True)

        for i, chunk in enumerate(chunks):
            front_matter = self._build_front_matter(
                source_file=source.name,
                chunk_index=i,
                total_chunks=len(chunks),
                chunk_title=chunk.get("title", ""),
            )
            content = f"---\n{front_matter}---\n\n{chunk['content']}"

            chunk_file = doc_dir / f"chunk_{i:03d}.md"
            chunk_file.write_text(content, encoding="utf-8")
            outputs.append(chunk_file)

        # 最終出力を output_dir にもコピー
        final_dir = self.config.output_dir / source.stem
        if final_dir.exists():
            shutil.rmtree(final_dir)
        shutil.copytree(doc_dir, final_dir)

        return outputs

    # ---- チャンキング戦略 ----

    def _chunk_by_heading(self, md_content: str, fallback_title: str) -> list[dict]:
        """
        Markdown の見出し（# / ## / ###）で分割。
        見出しがない場合は固定サイズ分割にフォールバック。
        """
        # 見出しで分割
        sections = re.split(r"(?=^#{1,3}\s)", md_content, flags=re.MULTILINE)
        sections = [s.strip() for s in sections if s.strip()]

        if len(sections) <= 1:
            # 見出しがない → 固定サイズ分割
            return self._chunk_fixed_size(md_content, fallback_title)

        # 見出しセクションをトークン上限に合わせて統合・分割
        chunks: list[dict] = []
        current_content: list[str] = []
        current_tokens = 0
        current_title = fallback_title

        for section in sections:
            section_tokens = count_tokens(section)

            # このセクション単体が max を超える場合はさらに分割
            if section_tokens > CHUNK_MAX_TOKENS:
                # 現在のバッファを flush
                if current_content:
                    chunks.append({
                        "title": current_title,
                        "content": "\n\n".join(current_content),
                    })
                    current_content = []
                    current_tokens = 0

                # 大きいセクションを固定サイズで分割
                title = self._extract_heading(section) or fallback_title
                sub_chunks = self._chunk_fixed_size(section, title)
                chunks.extend(sub_chunks)
                continue

            # バッファに追加して target を超えたら flush
            if current_tokens + section_tokens > CHUNK_TARGET_TOKENS and current_content:
                chunks.append({
                    "title": current_title,
                    "content": "\n\n".join(current_content),
                })
                current_content = []
                current_tokens = 0

            if not current_content:
                current_title = self._extract_heading(section) or fallback_title

            current_content.append(section)
            current_tokens += section_tokens

        if current_content:
            chunks.append({
                "title": current_title,
                "content": "\n\n".join(current_content),
            })

        return chunks

    def _chunk_fixed_size(self, text: str, title: str) -> list[dict]:
        """固定サイズ（行ベース）分割"""
        lines = text.split("\n")
        chunks: list[dict] = []
        current: list[str] = []
        current_tokens = 0

        for line in lines:
            lt = count_tokens(line)
            if current_tokens + lt > CHUNK_TARGET_TOKENS and current:
                chunks.append({
                    "title": f"{title} (part {len(chunks) + 1})",
                    "content": "\n".join(current),
                })
                current = []
                current_tokens = 0
            current.append(line)
            current_tokens += lt

        if current:
            chunks.append({
                "title": f"{title} (part {len(chunks) + 1})" if chunks else title,
                "content": "\n".join(current),
            })

        return chunks

    # ---- ヘルパー ----

    @staticmethod
    def _extract_heading(section: str) -> str:
        """セクション先頭の見出しテキストを抽出"""
        m = re.match(r"^#{1,3}\s+(.+)", section)
        return m.group(1).strip() if m else ""

    @staticmethod
    def _build_front_matter(
        source_file: str,
        chunk_index: int,
        total_chunks: int,
        chunk_title: str = "",
    ) -> str:
        lines = [
            f"source_file: {source_file}",
            f"chunk_index: {chunk_index}",
            f"total_chunks: {total_chunks}",
        ]
        if chunk_title:
            lines.append(f"chunk_title: \"{chunk_title}\"")
        return "\n".join(lines) + "\n"
