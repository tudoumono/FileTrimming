"""Step1: ファイル正規化

.doc / .rtf → .docx への変換を Word COM 経由で行う。
.docx はそのままコピーする。

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


def _convert_via_com(src: Path, dst: Path) -> None:
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
    dst = dst_dir / (src.stem + ".docx")
    t0 = time.perf_counter()

    if ext in config.com_normalize_exts:
        try:
            _convert_via_com(src, dst)
            elapsed = time.perf_counter() - t0
            logger.info("COM 変換完了: %s → %s (%.1fs)", src.name, dst.name, elapsed)
            return StepResult(
                file_path=str(src),
                step="normalize",
                status=ProcessStatus.SUCCESS,
                message=f"COM convert: {ext} → .docx",
                duration_sec=round(elapsed, 2),
            )
        except Exception as e:
            elapsed = time.perf_counter() - t0
            logger.error("COM 変換失敗: %s: %s", src.name, e)
            return StepResult(
                file_path=str(src),
                step="normalize",
                status=ProcessStatus.ERROR,
                message=str(e),
                duration_sec=round(elapsed, 2),
            )

    elif ext in config.passthrough_exts:
        dst_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        elapsed = time.perf_counter() - t0
        logger.info("コピー: %s (%.1fs)", src.name, elapsed)
        return StepResult(
            file_path=str(src),
            step="normalize",
            status=ProcessStatus.SUCCESS,
            message="passthrough copy (.docx)",
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
