"""Step1: ファイル正規化

.doc / .rtf → .docx への変換を Word COM 経由で行う。
.xls → .xlsx への変換を Excel COM 経由で行う。
.docx / .xlsx / .xlsm はそのままコピーする。

COM は STA (シングルスレッドアパートメント) 制約があるため逐次処理。
"""

from __future__ import annotations

import shutil
import time
from logging import getLogger
from pathlib import Path

from src.config import PipelineConfig
from src.models.metadata import ProcessStatus, StepResult

logger = getLogger(__name__)

# Word SaveAs2 の FileFormat 定数
WD_FORMAT_DOCX = 16  # wdFormatXMLDocument
# Excel SaveAs の FileFormat 定数
XL_FORMAT_XLSX = 51  # xlOpenXMLWorkbook


def _convert_word_via_com(src: Path, dst: Path) -> None:
    """Word COM で src を .docx として dst に保存する。

    Raises:
        RuntimeError: COM 初期化または変換に失敗した場合
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError(
            "pywin32 がインストールされていません。"
            "COM 変換には Windows + pywin32 が必要です。"
        ) from e

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(str(src.resolve()))
        dst.parent.mkdir(parents=True, exist_ok=True)
        doc.SaveAs2(str(dst.resolve()), FileFormat=WD_FORMAT_DOCX)
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _convert_excel_via_com(src: Path, dst: Path) -> None:
    """Excel COM で src を .xlsx として dst に保存する。

    Raises:
        RuntimeError: COM 初期化または変換に失敗した場合
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError(
            "pywin32 がインストールされていません。"
            "COM 変換には Windows + pywin32 が必要です。"
        ) from e

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(src.resolve()))
        dst.parent.mkdir(parents=True, exist_ok=True)
        wb.SaveAs(str(dst.resolve()), FileFormat=XL_FORMAT_XLSX)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


# Word COM で変換する拡張子
_WORD_COM_EXTS = {".doc", ".rtf"}
# Excel COM で変換する拡張子
_EXCEL_COM_EXTS = {".xls"}


def normalize_file(
    src: Path,
    dst_dir: Path,
    config: PipelineConfig,
) -> StepResult:
    """1ファイルを正規化する。

    Args:
        src: 元ファイルパス
        dst_dir: 出力先ディレクトリ (01_normalized/<機能フォルダ>/)
        config: パイプライン設定

    Returns:
        StepResult: 処理結果
    """
    ext = src.suffix.lower()
    t0 = time.perf_counter()

    if ext in config.com_normalize_exts:
        # COM 変換が必要な拡張子
        if ext in _EXCEL_COM_EXTS:
            # Excel COM: .xls → .xlsx
            dst = dst_dir / (src.stem + ".xlsx")
            try:
                _convert_excel_via_com(src, dst)
                elapsed = time.perf_counter() - t0
                logger.info("Excel COM 変換完了: %s → %s (%.1fs)", src.name, dst.name, elapsed)
                return StepResult(
                    file_path=str(src),
                    step="normalize",
                    status=ProcessStatus.SUCCESS,
                    message=f"Excel COM convert: {ext} → .xlsx",
                    duration_sec=round(elapsed, 2),
                )
            except Exception as e:
                elapsed = time.perf_counter() - t0
                logger.error("Excel COM 変換失敗: %s: %s", src.name, e)
                return StepResult(
                    file_path=str(src),
                    step="normalize",
                    status=ProcessStatus.ERROR,
                    message=str(e),
                    duration_sec=round(elapsed, 2),
                )
        else:
            # Word COM: .doc / .rtf → .docx
            dst = dst_dir / (src.stem + ".docx")
            try:
                _convert_word_via_com(src, dst)
                elapsed = time.perf_counter() - t0
                logger.info("Word COM 変換完了: %s → %s (%.1fs)", src.name, dst.name, elapsed)
                return StepResult(
                    file_path=str(src),
                    step="normalize",
                    status=ProcessStatus.SUCCESS,
                    message=f"Word COM convert: {ext} → .docx",
                    duration_sec=round(elapsed, 2),
                )
            except Exception as e:
                elapsed = time.perf_counter() - t0
                logger.error("Word COM 変換失敗: %s: %s", src.name, e)
                return StepResult(
                    file_path=str(src),
                    step="normalize",
                    status=ProcessStatus.ERROR,
                    message=str(e),
                    duration_sec=round(elapsed, 2),
                )

    elif ext in config.passthrough_exts:
        # パススルー: 元のファイル名をそのままコピー
        dst = dst_dir / src.name
        dst_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        elapsed = time.perf_counter() - t0
        logger.info("コピー: %s (%.1fs)", src.name, elapsed)
        return StepResult(
            file_path=str(src),
            step="normalize",
            status=ProcessStatus.SUCCESS,
            message=f"passthrough copy ({ext})",
            duration_sec=round(elapsed, 2),
        )

    else:
        elapsed = time.perf_counter() - t0
        logger.warning("未対応拡張子: %s", src.name)
        return StepResult(
            file_path=str(src),
            step="normalize",
            status=ProcessStatus.SKIPPED,
            message=f"unsupported extension: {ext}",
            duration_sec=round(elapsed, 2),
        )