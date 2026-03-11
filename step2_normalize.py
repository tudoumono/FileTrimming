"""
steps/step2_normalize.py - Phase 2: フォーマット正規化 (Win32 COM)

レガシー形式 (.doc, .xls, .ppt, .rtf, .trf) を
Win32 COM Automation で現行形式に変換する。
"""

import json
import shutil
import logging
from pathlib import Path

from .base import BaseStep
from utils.file_utils import LEGACY_CONVERSIONS, detect_file_type

logger = logging.getLogger(__name__)

# Win32 COM の SaveAs FileFormat 定数
WORD_FORMAT_DOCX = 16         # wdFormatXMLDocument
EXCEL_FORMAT_XLSX = 51        # xlOpenXMLWorkbook
POWERPOINT_FORMAT_PPTX = 24   # ppSaveAsOpenXMLPresentation


def _convert_with_word(src: Path, dest: Path):
    """Word COM で .doc/.rtf/.trf → .docx"""
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(str(src.resolve()))
        doc.SaveAs2(str(dest.resolve()), FileFormat=WORD_FORMAT_DOCX)
        doc.Close(SaveChanges=False)
    finally:
        word.Quit()


def _convert_with_excel(src: Path, dest: Path):
    """Excel COM で .xls → .xlsx"""
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(src.resolve()))
        wb.SaveAs(str(dest.resolve()), FileFormat=EXCEL_FORMAT_XLSX)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


def _convert_with_powerpoint(src: Path, dest: Path):
    """PowerPoint COM で .ppt → .pptx"""
    import win32com.client
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    try:
        presentation = ppt.Presentations.Open(str(src.resolve()), WithWindow=False)
        presentation.SaveAs(str(dest.resolve()), FileFormat=POWERPOINT_FORMAT_PPTX)
        presentation.Close()
    finally:
        ppt.Quit()


# ファイル種別 → 変換関数 / 変換先拡張子
CONVERTERS = {
    "word_legacy": (_convert_with_word, ".docx"),
    "rtf":         (_convert_with_word, ".docx"),
    "trf":         (_convert_with_word, ".docx"),
    "excel_legacy": (_convert_with_excel, ".xlsx"),
    "pptx_legacy":  (_convert_with_powerpoint, ".pptx"),
}


class Step2Normalize(BaseStep):
    step_number = 2
    step_name = "フォーマット正規化"

    def execute(self):
        prev_files_dir = self.config.paths.step_dir(1) / "files"
        if not prev_files_dir.exists():
            raise FileNotFoundError("Step 1 の出力が見つかりません。Step 1 を先に実行してください。")

        out_dir = self.step_dir / "files"
        out_dir.mkdir(parents=True, exist_ok=True)

        # Step 1 の分類結果を読み込む
        classification_path = self.config.paths.step_dir(1) / "classification.json"
        classified = json.loads(classification_path.read_text(encoding="utf-8"))

        converted_count = 0
        passthrough_count = 0

        for ftype, rel_paths in classified.items():
            for rel_path_str in rel_paths:
                src = prev_files_dir / rel_path_str
                if not src.exists():
                    continue

                if ftype in CONVERTERS:
                    # レガシー形式 → 変換
                    converter_func, new_ext = CONVERTERS[ftype]
                    dest = out_dir / Path(rel_path_str).with_suffix(new_ext)
                    dest.parent.mkdir(parents=True, exist_ok=True)

                    if not self.should_process_file(rel_path_str, dest):
                        continue

                    try:
                        converter_func(src, dest)
                        self.manifest.mark_file_done(rel_path_str)
                        converted_count += 1
                        logger.info("  変換: %s → %s", src.name, dest.name)
                    except Exception as e:
                        logger.error("  変換失敗: %s (%s)", src.name, e)
                        self.manifest.mark_file_failed(rel_path_str, str(e))
                        # 変換失敗時は元ファイルをそのままコピー（後続ステップでフォールバック可能に）
                        fallback_dest = out_dir / rel_path_str
                        fallback_dest.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(src, fallback_dest)
                else:
                    # 変換不要 → そのままコピー
                    dest = out_dir / rel_path_str
                    dest.parent.mkdir(parents=True, exist_ok=True)

                    if not self.should_process_file(rel_path_str, dest):
                        continue

                    shutil.copy2(src, dest)
                    self.manifest.mark_file_done(rel_path_str)
                    passthrough_count += 1

        logger.info("  変換: %d 件、パススルー: %d 件", converted_count, passthrough_count)

        # 正規化後の分類結果を再生成
        from utils.file_utils import classify_files
        reclassified = classify_files(out_dir)
        summary = {ft: [str(f.relative_to(out_dir)) for f in fs] for ft, fs in reclassified.items()}

        (self.step_dir / "classification.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
