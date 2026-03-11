"""
steps/step1_copy.py - Phase 1: コピー & ファイル分類

入力フォルダの全ファイルを作業フォルダ（step1/）にコピーし、
ファイル種別を分類した classification.json を出力する。
"""

import json
import shutil
import logging
from pathlib import Path

from .base import BaseStep
from utils.file_utils import classify_files

logger = logging.getLogger(__name__)


class Step1Copy(BaseStep):
    step_number = 1
    step_name = "コピー & ファイル分類"

    def execute(self):
        input_dir = self.config.paths.input_dir
        if not input_dir.exists():
            raise FileNotFoundError(f"入力フォルダが見つかりません: {input_dir}")

        files_dir = self.step_dir / "files"
        files_dir.mkdir(parents=True, exist_ok=True)

        # ファイルのコピー
        copied_count = 0
        for src in sorted(input_dir.rglob("*")):
            if not src.is_file() or src.name.startswith("."):
                continue

            # 元のフォルダ構造を保持してコピー
            relative = src.relative_to(input_dir)
            dest = files_dir / relative
            dest.parent.mkdir(parents=True, exist_ok=True)

            if not self.should_process_file(str(relative), dest):
                continue

            try:
                shutil.copy2(src, dest)
                self.manifest.mark_file_done(str(relative))
                copied_count += 1
                logger.debug("  コピー: %s", relative)
            except Exception as e:
                logger.error("  コピー失敗: %s (%s)", relative, e)
                self.manifest.mark_file_failed(str(relative), str(e))

        logger.info("  %d ファイルをコピーしました", copied_count)

        # ファイル分類
        classified = classify_files(files_dir)
        summary = {ftype: [str(f.relative_to(files_dir)) for f in files] for ftype, files in classified.items()}

        classification_path = self.step_dir / "classification.json"
        classification_path.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        # サマリログ
        for ftype, files in summary.items():
            logger.info("  [%s] %d 件", ftype, len(files))
