"""
pipeline.py - パイプライン・オーケストレーター

全ステップを順に実行する。
  - start_step 指定で途中から再開可能
  - 各ステップの中間出力は work/stepN/ に保存
  - manifest.json でファイル単位の処理状態を管理
  - file_conflict_mode (skip/overwrite) で再実行時の挙動を制御
"""

import logging
import time
from typing import Optional

from config import AppConfig
from llm import LLMClient, create_llm_client
from steps import ALL_STEPS, BaseStep

logger = logging.getLogger(__name__)


class Pipeline:
    """ドキュメント処理パイプライン"""

    def __init__(self, config: AppConfig):
        self.config = config
        self._llm_client: Optional[LLMClient] = None

    @property
    def llm_client(self) -> Optional[LLMClient]:
        """LLM クライアントの遅延初期化（Step 5 まで不要）"""
        if self._llm_client is None:
            try:
                self._llm_client = create_llm_client(self.config.llm)
                logger.info("LLM プロバイダ: %s", self._llm_client.provider_name())
            except Exception as e:
                logger.warning("LLM クライアント初期化失敗: %s（LLM を使わないステップは実行可能）", e)
        return self._llm_client

    def run(self, start_step: Optional[int] = None, end_step: Optional[int] = None):
        """
        パイプラインを実行する。

        Args:
            start_step: 開始ステップ番号（デフォルト: .env の START_STEP）
            end_step: 終了ステップ番号（デフォルト: 最終ステップ）
        """
        start = start_step or self.config.execution.start_step
        end = end_step or len(ALL_STEPS)
        selected_steps = [
            step_cls for step_cls in ALL_STEPS
            if start <= step_cls.step_number <= end
        ]

        self.config.paths.ensure_dirs()

        logger.info("=" * 60)
        logger.info("パイプライン開始 (Step %d → Step %d, 対象 %d ステップ)", start, end, len(selected_steps))
        logger.info("  入力: %s", self.config.paths.input_dir)
        logger.info("  作業: %s", self.config.paths.work_dir)
        logger.info("  出力: %s", self.config.paths.output_dir)
        logger.info("  モード: %s", self.config.execution.file_conflict_mode)
        logger.info("=" * 60)

        for step_index, step_cls in enumerate(selected_steps, start=1):
            step_num = step_cls.step_number
            t_step = time.perf_counter()
            logger.info("パイプライン進捗: %d/%d - Step %d %s", step_index, len(selected_steps), step_num, step_cls.step_name)

            # LLM が必要なステップ（Step 5）には LLM クライアントを渡す
            llm = self.llm_client if step_num >= 5 else None

            step: BaseStep = step_cls(config=self.config, llm_client=llm)

            try:
                step.run()
            except Exception as e:
                logger.error("=" * 60)
                logger.error("パイプライン中断: Step %d で失敗", step_num)
                logger.error("  エラー: %s", e)
                logger.error("  リトライ: python main.py --start-step %d", step_num)
                logger.error("=" * 60)
                raise
            else:
                elapsed = time.perf_counter() - t_step
                logger.info(
                    "パイプライン進捗: %d/%d - Step %d %s 完了 (%.1f秒)",
                    step_index,
                    len(selected_steps),
                    step_num,
                    step_cls.step_name,
                    elapsed,
                )

        logger.info("=" * 60)
        logger.info("パイプライン完了")
        logger.info("=" * 60)
