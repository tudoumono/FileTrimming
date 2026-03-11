"""
main.py - パイプライン実行エントリポイント
"""

from __future__ import annotations

import argparse
import logging
import sys
from dataclasses import replace
from pathlib import Path

# プロジェクトルートを sys.path に追加
sys.path.insert(0, str(Path(__file__).parent))

from config import load_config
from logging_utils import configure_logging


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="ドキュメント処理パイプライン",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
例:
  python main.py                                # 全ステップ実行
  python main.py --start-step 3                 # Step 3 から再開
  python main.py --start-step 2 --end-step 2    # Step 2 のみ実行
  python main.py --mode overwrite               # 全て上書き再処理
  python main.py --dry-run                      # 設定確認のみ
        """,
    )
    parser.add_argument("--start-step", type=int, default=None,
                        help="開始ステップ番号 (1-6)")
    parser.add_argument("--end-step", type=int, default=None,
                        help="終了ステップ番号 (1-6)")
    parser.add_argument("--mode", choices=["skip", "overwrite"], default=None,
                        help="中間ファイルの扱い")
    parser.add_argument("--dry-run", action="store_true",
                        help="設定を表示して終了")
    parser.add_argument("--log-level",
                        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
                        default=None, help="ログレベル上書き")
    parser.add_argument("--log-file", default=None,
                        help="ログファイルの出力先を上書き")
    parser.add_argument("--env", default=None,
                        help=".env ファイルのパスを指定")
    return parser.parse_args()


def main():
    args = parse_args()
    config = load_config(args.env)

    # CLI 引数で設定を上書き
    exec_updates = {}
    if args.mode:
        exec_updates["file_conflict_mode"] = args.mode
    if args.start_step:
        exec_updates["start_step"] = args.start_step
    if args.log_level:
        exec_updates["log_level"] = args.log_level
    if args.log_file:
        exec_updates["log_file"] = Path(args.log_file)
    if exec_updates:
        config = replace(config, execution=replace(config.execution, **exec_updates))

    log_path = configure_logging(
        level=config.execution.log_level,
        log_file=config.execution.log_file,
    )
    logger = logging.getLogger(__name__)
    if log_path is not None:
        logger.info("ログファイル: %s", log_path)

    if args.dry_run:
        logger.info("=== ドライラン（設定確認） ===")
        logger.info("入力: %s (存在: %s)", config.paths.input_dir, config.paths.input_dir.exists())
        logger.info("作業: %s", config.paths.work_dir)
        logger.info("出力: %s", config.paths.output_dir)
        llm_model = config.llm.openai_model if config.llm.provider == "openai" else config.llm.ollama_model
        logger.info("LLM: %s (%s)", config.llm.provider, llm_model)
        logger.info("トークン上限: %d", config.processing.token_limit)
        logger.info("競合モード: %s", config.execution.file_conflict_mode)
        logger.info("開始ステップ: %d", args.start_step or config.execution.start_step)
        logger.info("ログ出力先: %s", config.execution.log_file)
        return

    from pipeline import Pipeline

    pipeline = Pipeline(config)
    try:
        pipeline.run(start_step=args.start_step, end_step=args.end_step)
    except KeyboardInterrupt:
        logger.warning("ユーザーによる中断")
        sys.exit(130)
    except Exception as e:
        logger.exception("パイプライン失敗: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
