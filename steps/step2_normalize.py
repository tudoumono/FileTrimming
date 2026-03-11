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
from utils.file_utils import classify_files, detect_file_type

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

CONVERTER_LABELS = {
    "word_legacy": "Word COM",
    "rtf": "Word COM (RTF 正規化)",
    "trf": "Word COM (TRF 正規化)",
    "excel_legacy": "Excel COM",
    "pptx_legacy": "PowerPoint COM",
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
        targets = [
            (ftype, rel_path_str)
            for ftype, rel_paths in sorted(classified.items())
            for rel_path_str in rel_paths
        ]
        total = len(targets)
        self.log_target_count(total, "正規化対象")

        converted_count = 0
        passthrough_count = 0
        normalization_log = []

        for index, (ftype, rel_path_str) in enumerate(targets, start=1):
            src = prev_files_dir / rel_path_str
            self.log_file_start(index, total, rel_path_str, "正規化")

            if not src.exists():
                self.log_file_failed(index, total, rel_path_str, "Step 1 出力に対象ファイルがありません")
                normalization_log.append({
                    "source": rel_path_str,
                    "source_type": ftype,
                    "status": "missing",
                })
                continue

            if ftype in CONVERTERS:
                converter_func, new_ext = CONVERTERS[ftype]
                strategy = CONVERTER_LABELS[ftype]
                dest_rel = str(Path(rel_path_str).with_suffix(new_ext))
                dest = out_dir / dest_rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                self.log_file_progress(index, total, rel_path_str, f"{strategy} で {dest_rel} に変換します")

                if not self.should_process_file(rel_path_str, dest):
                    self.log_file_skip(index, total, rel_path_str)
                    normalization_log.append({
                        "source": rel_path_str,
                        "source_type": ftype,
                        "output": dest_rel,
                        "output_type": detect_file_type(dest),
                        "strategy": strategy,
                        "status": "skipped",
                    })
                    continue

                try:
                    converter_func(src, dest)
                    self.manifest.mark_file_done(rel_path_str)
                    converted_count += 1
                    if ftype in {"rtf", "trf"}:
                        self.log_file_progress(
                            index,
                            total,
                            rel_path_str,
                            "RTF/TRF 由来です。図・フロー図・埋め込みオブジェクトは後段で別途確認してください",
                        )
                    self.log_file_done(index, total, rel_path_str, f"{dest_rel} に変換しました")
                    normalization_log.append({
                        "source": rel_path_str,
                        "source_type": ftype,
                        "output": dest_rel,
                        "output_type": detect_file_type(dest),
                        "strategy": strategy,
                        "status": "converted",
                        "requires_visual_review": ftype in {"rtf", "trf"},
                    })
                except Exception as e:
                    self.log_file_failed(index, total, rel_path_str, e)
                    self.manifest.mark_file_failed(rel_path_str, str(e))

                    fallback_dest = out_dir / rel_path_str
                    fallback_dest.parent.mkdir(parents=True, exist_ok=True)
                    try:
                        shutil.copy2(src, fallback_dest)
                        self.log_file_done(
                            index,
                            total,
                            rel_path_str,
                            "変換失敗のため原本をコピーして後続処理へフォールバックしました",
                        )
                        normalization_log.append({
                            "source": rel_path_str,
                            "source_type": ftype,
                            "output": rel_path_str,
                            "output_type": detect_file_type(fallback_dest),
                            "strategy": strategy,
                            "status": "fallback_copy",
                            "error": str(e),
                            "requires_visual_review": ftype in {"rtf", "trf"},
                        })
                    except Exception as copy_error:
                        self.log_file_failed(index, total, rel_path_str, copy_error)
                        normalization_log.append({
                            "source": rel_path_str,
                            "source_type": ftype,
                            "strategy": strategy,
                            "status": "failed",
                            "error": str(copy_error),
                        })
            else:
                dest = out_dir / rel_path_str
                dest.parent.mkdir(parents=True, exist_ok=True)
                self.log_file_progress(index, total, rel_path_str, "変換不要のため原本をそのまま引き継ぎます")

                if not self.should_process_file(rel_path_str, dest):
                    self.log_file_skip(index, total, rel_path_str)
                    normalization_log.append({
                        "source": rel_path_str,
                        "source_type": ftype,
                        "output": rel_path_str,
                        "output_type": ftype,
                        "strategy": "copy",
                        "status": "skipped",
                    })
                    continue

                try:
                    shutil.copy2(src, dest)
                    self.manifest.mark_file_done(rel_path_str)
                    passthrough_count += 1
                    self.log_file_done(index, total, rel_path_str, "変換不要ファイルをコピーしました")
                    normalization_log.append({
                        "source": rel_path_str,
                        "source_type": ftype,
                        "output": rel_path_str,
                        "output_type": ftype,
                        "strategy": "copy",
                        "status": "passthrough",
                    })
                except Exception as e:
                    self.log_file_failed(index, total, rel_path_str, e)
                    self.manifest.mark_file_failed(rel_path_str, str(e))
                    normalization_log.append({
                        "source": rel_path_str,
                        "source_type": ftype,
                        "output": rel_path_str,
                        "output_type": ftype,
                        "strategy": "copy",
                        "status": "failed",
                        "error": str(e),
                    })

        logger.info("  変換: %d 件、パススルー: %d 件", converted_count, passthrough_count)
        (self.step_dir / "normalization_log.json").write_text(
            json.dumps(normalization_log, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        # 正規化後の分類結果を再生成
        reclassified = classify_files(out_dir)
        summary = {ft: [str(f.relative_to(out_dir)) for f in fs] for ft, fs in reclassified.items()}

        (self.step_dir / "classification.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
