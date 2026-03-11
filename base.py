"""
steps/base.py - ステップの基底クラス

各ステップは以下のライフサイクルで動作する:
  1. run() が呼ばれる
  2. 中間出力フォルダ (work/stepN/) を確認
  3. manifest.json でどのファイルが処理済みかを管理
  4. skip/overwrite モードに応じて処理対象を決定
  5. execute() で実際の処理を実行
  6. manifest.json を更新

manifest.json の構造:
{
  "step": 1,
  "name": "copy_and_classify",
  "status": "completed",       # completed / partial / failed
  "processed_files": ["a.xlsx", "b.docx"],
  "failed_files": ["c.pdf"],
  "started_at": "2025-...",
  "completed_at": "2025-..."
}
"""

from abc import ABC, abstractmethod
from pathlib import Path
from datetime import datetime, timezone
import json
import logging
from typing import Optional

from config import AppConfig
from llm.base import LLMClient

logger = logging.getLogger(__name__)

MANIFEST_FILENAME = "manifest.json"


class StepManifest:
    """ステップの処理状態を管理するマニフェスト"""

    def __init__(self, step_dir: Path):
        self._path = step_dir / MANIFEST_FILENAME
        self._data: dict = self._load()

    def _load(self) -> dict:
        if self._path.exists():
            try:
                return json.loads(self._path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                return {}
        return {}

    def save(self):
        self._path.write_text(
            json.dumps(self._data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    @property
    def status(self) -> str:
        return self._data.get("status", "not_started")

    @property
    def processed_files(self) -> set[str]:
        return set(self._data.get("processed_files", []))

    @property
    def failed_files(self) -> set[str]:
        return set(self._data.get("failed_files", []))

    def is_file_processed(self, filename: str) -> bool:
        return filename in self.processed_files

    def mark_started(self, step_number: int, step_name: str):
        self._data.update({
            "step": step_number,
            "name": step_name,
            "status": "running",
            "started_at": datetime.now(timezone.utc).isoformat(),
        })
        self._data.setdefault("processed_files", [])
        self._data.setdefault("failed_files", [])
        self.save()

    def mark_file_done(self, filename: str):
        processed = list(self.processed_files)
        if filename not in processed:
            processed.append(filename)
        self._data["processed_files"] = processed
        # 失敗リストから除去
        failed = list(self.failed_files)
        if filename in failed:
            failed.remove(filename)
        self._data["failed_files"] = failed
        self.save()

    def mark_file_failed(self, filename: str, error: str):
        failed = list(self.failed_files)
        if filename not in failed:
            failed.append(filename)
        self._data["failed_files"] = failed
        self._data.setdefault("errors", {})[filename] = error
        self.save()

    def mark_completed(self):
        self._data["status"] = "completed"
        self._data["completed_at"] = datetime.now(timezone.utc).isoformat()
        self.save()

    def mark_partial(self):
        self._data["status"] = "partial"
        self._data["completed_at"] = datetime.now(timezone.utc).isoformat()
        self.save()

    def reset(self):
        """overwrite モードで最初からやり直す場合"""
        self._data = {}
        if self._path.exists():
            self._path.unlink()


class BaseStep(ABC):
    """パイプラインの各ステップの基底クラス"""

    # サブクラスで定義する
    step_number: int = 0
    step_name: str = ""

    def __init__(self, config: AppConfig, llm_client: Optional[LLMClient] = None):
        self.config = config
        self.llm = llm_client
        self.step_dir = config.paths.step_dir(self.step_number)
        self.manifest = StepManifest(self.step_dir)
        self.conflict_mode = config.execution.file_conflict_mode

    @property
    def prev_step_dir(self) -> Optional[Path]:
        """前ステップの出力フォルダ（Step1 は None）"""
        if self.step_number <= 1:
            return None
        return self.config.paths.step_dir(self.step_number - 1)

    def should_process_file(self, filename: str, output_path: Path) -> bool:
        """
        ファイルを処理すべきか判定する。

        - overwrite モード: 常に True
        - skip モード: マニフェスト上で処理済み かつ 出力ファイルが存在 → False
        """
        if self.conflict_mode == "overwrite":
            return True

        if self.manifest.is_file_processed(filename) and output_path.exists():
            logger.info("  [SKIP] %s (処理済み)", filename)
            return False

        return True

    def run(self):
        """ステップを実行する（テンプレートメソッド）"""
        logger.info("=" * 60)
        logger.info("Step %d: %s 開始", self.step_number, self.step_name)
        logger.info("=" * 60)

        if self.conflict_mode == "overwrite":
            self.manifest.reset()

        self.manifest.mark_started(self.step_number, self.step_name)

        try:
            self.execute()
        except Exception as e:
            logger.error("Step %d で致命的エラー: %s", self.step_number, e)
            self.manifest.mark_partial()
            raise

        if self.manifest.failed_files:
            logger.warning(
                "Step %d: %d 件のファイルで失敗。詳細は %s を参照",
                self.step_number,
                len(self.manifest.failed_files),
                self.step_dir / MANIFEST_FILENAME,
            )
            self.manifest.mark_partial()
        else:
            self.manifest.mark_completed()

        logger.info("Step %d: %s 完了", self.step_number, self.step_name)

    @abstractmethod
    def execute(self):
        """各ステップの実処理（サブクラスで実装）"""
        ...
