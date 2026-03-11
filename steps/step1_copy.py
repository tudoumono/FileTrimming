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
        input_files = [
            src for src in sorted(input_dir.rglob("*"))
            if src.is_file() and not src.name.startswith(".")
        ]
        total = len(input_files)
        self.log_target_count(total, "コピー対象")

        # ファイルのコピー
        copied_count = 0
        for index, src in enumerate(input_files, start=1):
            # 元のフォルダ構造を保持してコピー
            relative = src.relative_to(input_dir)
            dest = files_dir / relative
            dest.parent.mkdir(parents=True, exist_ok=True)
            rel_str = str(relative)
            self.log_file_start(index, total, rel_str, "コピー")

            if not self.should_process_file(rel_str, dest):
                self.log_file_skip(index, total, rel_str)
                continue

            try:
                shutil.copy2(src, dest)
                self.manifest.mark_file_done(rel_str)
                copied_count += 1
                self.log_file_done(index, total, rel_str, "Step 1 作業領域へコピーしました")
            except Exception as e:
                self.log_file_failed(index, total, rel_str, e)
                self.manifest.mark_file_failed(rel_str, str(e))

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
