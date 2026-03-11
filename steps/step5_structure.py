"""
steps/step5_structure.py - Phase 5: 品質判定 & 構造化

Markdown変換後の品質をスコアリングし、
必要なファイルのみLLMで構造化・ノイズ除去を行う。
"""

import json
import logging
import re
from pathlib import Path

from .base import BaseStep

logger = logging.getLogger(__name__)


# --- 品質判定 ---

def _detect_garbage_ratio(text: str) -> float:
    """制御文字・文字化けの割合を検出"""
    if not text:
        return 0.0
    garbage = sum(1 for c in text if ord(c) < 32 and c not in "\n\r\t")
    return garbage / len(text)


def _has_valid_tables(text: str) -> bool:
    return bool(re.search(r"^\|.+\|$", text, re.MULTILINE))


def _has_heading_structure(text: str) -> bool:
    return bool(re.search(r"^#{1,4}\s+", text, re.MULTILINE))


def assess_quality(md_content: str, file_type: str, high_th: int, med_th: int) -> tuple[str, int]:
    """品質スコアを算出し、HIGH/MEDIUM/LOW を返す"""
    score = 100

    garbage = _detect_garbage_ratio(md_content)
    score -= int(garbage * 50)

    if not md_content.strip():
        return "LOW", 0

    if file_type == "excel":
        if not _has_valid_tables(md_content):
            score -= 30

    elif file_type == "word":
        img_count = md_content.count("[図")
        if img_count > 5:
            score -= 15
        if not _has_heading_structure(md_content):
            score -= 20

    elif file_type == "bagles":
        if "## チェック更新定義書" in md_content or "## ロジック定義書" in md_content:
            score += 10

    noise = len(re.findall(r"^[━═─┌┐└┘├┤┬┴┼│]+$", md_content, re.MULTILINE))
    score -= min(noise * 2, 20)

    if score >= high_th:
        return "HIGH", score
    elif score >= med_th:
        return "MEDIUM", score
    else:
        return "LOW", score


# --- ルールベース補正 ---

def rule_based_cleanup(text: str) -> str:
    """LLMを使わないルールベースのクリーンアップ"""
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"^[━═─┌┐└┘├┤┬┴┼│]+$", "", text, flags=re.MULTILINE)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    return text.strip()


# --- LLMプロンプト ---

PROMPTS = {
    "excel": (
        "以下はExcelのコード一覧から変換されたMarkdownです。"
        "テーブルが崩れている箇所を修正し、正しいMarkdownテーブルに再構成してください。"
        "元の情報は一切削除しないでください。出力はMarkdownのみ。"
    ),
    "word": (
        "以下はWord設計書から変換されたMarkdownです。"
        "見出し構造を適切に整理し、[図N]のプレースホルダに対して"
        "前後の文脈から図の内容を推測して簡潔な説明を追記してください。"
        "元の情報は一切削除しないでください。出力はMarkdownのみ。"
    ),
    "bagles": (
        "以下はBAGLES定義書（富士通の業務仕様書）から変換されたテキストです。"
        "BAGLES定義書は表形式の業務仕様書です。"
        "テーブル構造を正しいMarkdownテーブルに再構成してください。"
        "COBOL形式の型情報（９(３)等）はそのまま保持してください。出力はMarkdownのみ。"
    ),
    "default": (
        "以下はドキュメントから変換されたMarkdownです。"
        "レイアウト崩れを修正し、見出し・テーブル・リストを正しいMarkdown構造にしてください。"
        "元の情報は一切削除しないでください。出力はMarkdownのみ。"
    ),
}


class Step5Structure(BaseStep):
    step_number = 5
    step_name = "品質判定 & 構造化"

    def execute(self):
        prev_dir = self.config.paths.step_dir(4) / "files"
        if not prev_dir.exists():
            raise FileNotFoundError("Step 4 の出力が見つかりません。")

        out_dir = self.step_dir / "files"
        out_dir.mkdir(parents=True, exist_ok=True)

        high_th = self.config.processing.quality_threshold_high
        med_th = self.config.processing.quality_threshold_medium

        quality_log = []
        sources = sorted(prev_dir.rglob("*.md"))
        total = len(sources)
        self.log_target_count(total, "品質判定対象")

        for index, src in enumerate(sources, start=1):
            rel = src.relative_to(prev_dir)
            rel_str = str(rel)
            dest = out_dir / rel
            self.log_file_start(index, total, rel_str, "品質判定")

            if not self.should_process_file(rel_str, dest):
                self.log_file_skip(index, total, rel_str)
                quality_log.append({"file": rel_str, "quality": "skipped"})
                continue

            md_content = src.read_text(encoding="utf-8")

            # ファイル名からファイル種別を推定（Step4で変換元の拡張子情報が失われているため）
            ftype = self._guess_original_type(rel)
            quality, score = assess_quality(md_content, ftype, high_th, med_th)
            self.log_file_progress(index, total, rel_str, f"品質={quality}, score={score}, 推定元種別={ftype}")
            image_placeholder_count = md_content.count("[図")
            if image_placeholder_count:
                self.log_file_progress(
                    index,
                    total,
                    rel_str,
                    f"図プレースホルダを {image_placeholder_count} 件検出しました",
                )

            try:
                if quality == "HIGH":
                    result = rule_based_cleanup(md_content)
                elif quality == "MEDIUM":
                    result = rule_based_cleanup(md_content)
                    self.log_file_progress(index, total, rel_str, "ルールベース補正のみで構造を整えます")
                    # MEDIUM: 必要に応じてLLM補正（現時点ではルールベースのみ）
                    # 将来的には部分的なLLM補正をここに追加
                else:  # LOW
                    if self.llm:
                        self.log_file_progress(index, total, rel_str, "LLM で再構造化を実行します")
                        prompt = PROMPTS.get(ftype, PROMPTS["default"])
                        response = self.llm.chat(
                            system_prompt=prompt,
                            user_message=md_content[:60000],  # 安全なトランケーション
                        )
                        result = response.content
                        self.log_file_progress(
                            index,
                            total,
                            rel_str,
                            f"LLM 構造化を実行しました（使用トークン {response.total_tokens}）",
                        )
                    else:
                        self.log_file_progress(index, total, rel_str, "LLM 未設定のためルールベース補正のみ適用します")
                        result = rule_based_cleanup(md_content)

                dest.parent.mkdir(parents=True, exist_ok=True)
                dest.write_text(result, encoding="utf-8")
                self.manifest.mark_file_done(rel_str)
                self.log_file_done(index, total, rel_str, "構造化済み Markdown を出力しました")

                quality_log.append({
                    "file": rel_str, "quality": quality, "score": score,
                    "type": ftype, "llm_used": quality == "LOW" and self.llm is not None,
                })

            except Exception as e:
                self.log_file_failed(index, total, rel_str, e)
                self.manifest.mark_file_failed(rel_str, str(e))
                # 失敗時はクリーンアップのみ適用してコピー
                dest.parent.mkdir(parents=True, exist_ok=True)
                dest.write_text(rule_based_cleanup(md_content), encoding="utf-8")
                self.log_file_progress(index, total, rel_str, "ルールベース補正のみでフォールバック保存しました")
                quality_log.append({
                    "file": rel_str, "quality": quality, "score": score,
                    "status": "fallback", "error": str(e),
                })

        (self.step_dir / "quality_log.json").write_text(
            json.dumps(quality_log, ensure_ascii=False, indent=2), encoding="utf-8",
        )

        for q in ("HIGH", "MEDIUM", "LOW"):
            count = sum(1 for e in quality_log if e.get("quality") == q)
            if count:
                logger.info("  %s: %d 件", q, count)

    def _guess_original_type(self, rel: Path) -> str:
        """ファイル名から元のファイル種別を推定"""
        name = rel.stem.lower()
        if any(kw in name for kw in ("code", "一覧", "コード", "マスタ", "list")):
            return "excel"
        if any(kw in name for kw in ("設計", "仕様", "design", "spec")):
            return "word"
        if any(kw in name for kw in ("bagles", "定義書", "bgl")):
            return "bagles"
        return "unknown"
