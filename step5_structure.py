"""
steps/step5_structure.py - Phase 5: 品質判定 & 構造化

処理内容:
    1. Markdown 変換結果の品質を自動スコアリング
    2. HIGH  → ルールベースの軽微な補正のみ（LLM 不使用）
    3. MEDIUM → 問題箇所のみ LLM で補正
    4. LOW   → フル LLM 構造化
"""

import re
from pathlib import Path

from steps.base import BaseStep
from utils.token_counter import count_tokens


# ---- LLM 用プロンプト（ファイル種別ごと） ----

SYSTEM_PROMPTS = {
    "excel": (
        "以下はExcelのコード一覧から変換されたMarkdownです。\n"
        "テーブルが崩れている箇所を修正し、正しいMarkdownテーブルに再構成してください。\n"
        "元の情報は一切削除しないでください。出力はMarkdownのみ。"
    ),
    "word": (
        "以下はWord設計書から変換されたMarkdownです。\n"
        "見出し構造を適切に整理し、[図N]のプレースホルダに対して、\n"
        "前後の文脈から図の内容を推測して簡潔な説明を追記してください。\n"
        "元の情報は一切削除しないでください。出力はMarkdownのみ。"
    ),
    "bagles": (
        "以下はBAGLES定義書（富士通の業務仕様書）から変換されたテキストです。\n"
        "BAGLES定義書は表形式の業務仕様書で、以下の種類があります：\n"
        "チェック更新定義書、業務用語定義書、コード定義書、\n"
        "フォーマット定義書、ロジック定義書、インタフェース定義書。\n"
        "テーブル構造を正しいMarkdownテーブルに再構成してください。\n"
        "COBOL形式の型情報（９(３)等）はそのまま保持。出力はMarkdownのみ。"
    ),
    "default": (
        "以下のMarkdownテキストを整形してください。\n"
        "見出し・テーブル・リストの構造を正しいMarkdownに修正し、\n"
        "ノイズ（不要な装飾文字、文字化け等）を除去してください。\n"
        "元の情報は一切削除しないでください。出力はMarkdownのみ。"
    ),
}


class Step5Structure(BaseStep):
    step_number = 5
    step_name = "structured"

    def collect_inputs(self) -> list[Path]:
        step4_dir = self.config.step4_dir
        return sorted(
            f for f in step4_dir.rglob("*.md")
            if f.is_file() and f.name != "manifest.json"
        )

    def process_file(self, source: Path, dest_dir: Path) -> list[Path]:
        md_content = source.read_text(encoding="utf-8")
        ftype = self._infer_source_type(source)

        # 品質判定
        quality = self._assess_quality(md_content, ftype)

        # 品質に応じた処理
        if quality == "HIGH":
            result = self._rule_based_cleanup(md_content)
        elif quality == "MEDIUM":
            result = self._partial_llm_cleanup(md_content, ftype)
        else:  # LOW
            result = self._full_llm_restructure(md_content, ftype)

        dest = dest_dir / source.name
        dest.write_text(result, encoding="utf-8")
        return [dest]

    # ---- 品質判定 ----

    def _assess_quality(self, md_content: str, ftype: str) -> str:
        """品質を HIGH / MEDIUM / LOW で判定"""
        score = 100

        # 空コンテンツ
        if not md_content.strip():
            return "LOW"

        # 文字化け検出（制御文字の比率）
        garbage = sum(1 for c in md_content if ord(c) < 0x20 and c not in "\n\r\t")
        ratio = garbage / max(len(md_content), 1)
        score -= int(ratio * 80)

        # テーブル崩れ検出
        table_lines = [l for l in md_content.split("\n") if l.strip().startswith("|")]
        if table_lines:
            col_counts = [l.count("|") for l in table_lines]
            if len(set(col_counts)) > 2:  # 列数がバラバラ
                score -= 25

        # 見出し構造の有無
        if not re.search(r"^#{1,4}\s", md_content, re.MULTILINE):
            score -= 15

        # 図プレースホルダの多寡
        img_count = md_content.count("[図")
        if img_count > 5:
            score -= 10

        # BAGLES パーサ成功ボーナス
        if ftype == "bagles":
            for kw in ["チェック更新定義書", "ロジック定義書", "フォーマット定義書"]:
                if kw in md_content:
                    score += 5

        # 閾値判定
        if score >= self.config.quality_threshold_high:
            return "HIGH"
        elif score >= self.config.quality_threshold_medium:
            return "MEDIUM"
        else:
            return "LOW"

    # ---- ルールベース補正 ----

    def _rule_based_cleanup(self, md_content: str) -> str:
        """LLM を使わない軽微な補正"""
        # 連続空行の圧縮
        md_content = re.sub(r"\n{3,}", "\n\n", md_content)
        # 装飾罫線の除去
        md_content = re.sub(
            r"^[━═─┌┐└┘├┤┬┴┼│]+$", "", md_content, flags=re.MULTILINE
        )
        # 制御文字除去（改行・タブは保持）
        md_content = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", md_content)
        return md_content.strip()

    # ---- LLM 補正 ----

    def _partial_llm_cleanup(self, md_content: str, ftype: str) -> str:
        """品質 MEDIUM: ルールベース + 問題箇所のみ LLM"""
        md_content = self._rule_based_cleanup(md_content)

        if self.llm is None:
            return md_content

        # トークン超過時は LLM に投げない
        if count_tokens(md_content) > self.config.token_limit:
            return md_content

        prompt = SYSTEM_PROMPTS.get(ftype, SYSTEM_PROMPTS["default"])
        resp = self.llm.chat(
            system_prompt=prompt,
            user_message=md_content,
        )
        return resp.content

    def _full_llm_restructure(self, md_content: str, ftype: str) -> str:
        """品質 LOW: フル LLM 構造化"""
        md_content = self._rule_based_cleanup(md_content)

        if self.llm is None:
            return md_content

        if count_tokens(md_content) > self.config.token_limit:
            return md_content

        prompt = SYSTEM_PROMPTS.get(ftype, SYSTEM_PROMPTS["default"])
        resp = self.llm.chat(
            system_prompt=prompt,
            user_message=f"ファイル種別: {ftype}\n\n---\n\n{md_content}",
        )
        return resp.content

    # ---- ヘルパー ----

    def _infer_source_type(self, md_path: Path) -> str:
        """
        .md ファイル名から元のファイル種別を推定。
        将来的には前ステップの manifest から引くのが正確。
        """
        name = md_path.stem.lower()
        if any(kw in name for kw in ("定義書", "bagles", "bgl")):
            return "bagles"
        # Step 4 で元の拡張子をファイル名に残していれば判定可能
        # 暫定: manifest ベースの実装に拡張予定
        return "default"
