"""パイプライン CLI エントリーポイント

使い方:
  # 全ステップ実行
  python -m src.main --input input/ --output output/

  # Step2 から再実行（正規化済みデータを使用）
  python -m src.main --steps 2-3

  # Step3 のみ再実行（LLM バックエンドを変えて再変換）
  python -m src.main --steps 3
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from src.config import PipelineConfig
from src.models.metadata import ProcessStatus
from src.pipeline.folder_processor import run_pipeline


def _setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)-5s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="ドキュメント前処理パイプライン",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--input", "-i",
        type=Path, default=Path("input"),
        help="入力フォルダ (default: input/)",
    )
    parser.add_argument(
        "--output", "-o",
        type=Path, default=Path("output"),
        help="出力フォルダ (default: output/)",
    )
    parser.add_argument(
        "--intermediate",
        type=Path, default=Path("intermediate"),
        help="中間成果物フォルダ (default: intermediate/)",
    )
    parser.add_argument(
        "--steps", "-s",
        default="all",
        help="実行ステップ: all, 1, 2, 3, 1-2, 2-3, 1-3 (default: all)",
    )
    parser.add_argument(
        "--llm-backend",
        default="noop",
        choices=["noop", "openai", "local"],
        help="LLM バックエンド (default: noop)",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="詳細ログを出力",
    )

    args = parser.parse_args(argv)
    _setup_logging(args.verbose)

    config = PipelineConfig(
        input_dir=args.input,
        intermediate_dir=args.intermediate,
        output_dir=args.output,
        llm_backend=args.llm_backend,
    )
    # .env から API キー等を読み込み（CLI 引数の --llm-backend が優先）
    config.load_env()

    # 入力フォルダの存在チェック
    if "1" in args.steps or args.steps == "all":
        if not config.input_dir.exists():
            print(f"エラー: 入力フォルダが見つかりません: {config.input_dir}", file=sys.stderr)
            return 1

    logger = logging.getLogger(__name__)
    logger.info("パイプライン開始: steps=%s, input=%s, output=%s", args.steps, config.input_dir, config.output_dir)

    results = run_pipeline(config, steps=args.steps)

    # サマリー出力
    print("\n" + "=" * 60)
    print("パイプライン実行結果")
    print("=" * 60)

    total_ok = 0
    total_err = 0
    total_skip = 0
    total_warn = 0

    for step_name, step_results in results.items():
        ok = sum(1 for r in step_results if r.status == ProcessStatus.SUCCESS)
        err = sum(1 for r in step_results if r.status == ProcessStatus.ERROR)
        skip = sum(1 for r in step_results if r.status == ProcessStatus.SKIPPED)
        warn = sum(1 for r in step_results if r.status == ProcessStatus.WARNING)
        print(f"  {step_name}: {ok} 成功, {warn} 警告, {err} 失敗, {skip} スキップ")
        total_ok += ok
        total_err += err
        total_skip += skip
        total_warn += warn

    print(f"  合計: {total_ok} 成功, {total_warn} 警告, {total_err} 失敗, {total_skip} スキップ")
    print("=" * 60)

    if total_err > 0:
        print(f"\n{total_err} 件のエラーがあります。ログを確認してください。")

    return 1 if total_err > 0 else 0


if __name__ == "__main__":
    sys.exit(main())
