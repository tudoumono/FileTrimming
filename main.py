"""
main.py - ドキュメント処理パイプライン メインオーケストレーター

Usage:
    # 全ステップ実行
    python main.py

    # ステップ3から再開（失敗ファイルのみ再処理）
    python main.py --from 3

    # ステップ4だけ、全ファイル上書きモードで再実行
    python main.py --from 4 --to 4 --mode overwrite

    # 設定確認（ドライラン）
    python main.py --dry-run

環境変数:
    .env ファイルで設定（.env.example を参照）
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

from config import PipelineConfig
from llm_client import BaseLLMClient, create_llm_client
from steps import ALL_STEPS
from steps.base import BaseStep


def setup_logging(config: PipelineConfig) -> None:
    """ログ設定"""
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]

    if config.log_file:
        log_path = Path(config.log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_path, encoding="utf-8"))

    logging.basicConfig(
        level=getattr(logging, config.log_level.upper(), logging.INFO),
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
    )


def build_step(
    step_cls: type[BaseStep],
    config: PipelineConfig,
    llm: BaseLLMClient | None,
) -> BaseStep:
    """ステップインスタンスを生成"""
    return step_cls(config=config, llm=llm)


def run_pipeline(config: PipelineConfig) -> None:
    """パイプラインを実行"""
    logger = logging.getLogger("pipeline")

    # ディレクトリ作成
    config.ensure_dirs()

    # LLM クライアント初期化（Step 5 で使用）
    llm: BaseLLMClient | None = None
    if config.step_from <= 5 <= config.step_to:
        try:
            llm = create_llm_client(config)
        except Exception as e:
            logger.warning("LLM クライアント初期化失敗: %s（LLM なしで続行）", e)

    # パイプライン実行記録
    pipeline_record = {
        "started_at": datetime.now(timezone.utc).isoformat(),
        "config": {
            "resume_mode": config.resume_mode,
            "step_from": config.step_from,
            "step_to": config.step_to,
            "llm_provider": config.llm_provider,
            "token_limit": config.token_limit,
        },
        "steps": {},
    }

    logger.info("=" * 60)
    logger.info("パイプライン開始")
    logger.info("  入力: %s", config.input_dir)
    logger.info("  作業: %s", config.work_dir)
    logger.info("  出力: %s", config.output_dir)
    logger.info("  ステップ: %d → %d  モード: %s",
                config.step_from, config.step_to, config.resume_mode)
    logger.info("=" * 60)

    t_total = time.time()

    for step_cls in ALL_STEPS:
        step_num = step_cls.step_number

        # 範囲外のステップはスキップ
        if step_num < config.step_from or step_num > config.step_to:
            continue

        t_step = time.time()
        step = build_step(step_cls, config, llm)

        try:
            manifest = step.run()
            elapsed = time.time() - t_step
            pipeline_record["steps"][step_num] = {
                "name": step.step_name,
                "status": manifest.status,
                "elapsed_sec": round(elapsed, 1),
                "summary": manifest.to_dict()["summary"],
            }

            # partial（一部失敗）の場合、後続ステップも実行する
            # （成功したファイルだけで続行）
            if manifest.status == "partial":
                logger.warning(
                    "Step %d: 一部ファイルが失敗しました。"
                    "成功分で後続ステップを続行します。",
                    step_num,
                )

        except Exception as exc:
            elapsed = time.time() - t_step
            logger.error("Step %d: 致命的エラー — %s", step_num, exc, exc_info=True)
            pipeline_record["steps"][step_num] = {
                "name": step_cls.step_name,
                "status": "fatal_error",
                "error": str(exc),
                "elapsed_sec": round(elapsed, 1),
            }
            logger.error(
                "パイプラインを中断します。修正後に --from %d で再開してください。",
                step_num,
            )
            break

    # 完了記録
    total_elapsed = time.time() - t_total
    pipeline_record["finished_at"] = datetime.now(timezone.utc).isoformat()
    pipeline_record["total_elapsed_sec"] = round(total_elapsed, 1)

    # パイプライン全体の manifest を保存
    manifest_path = config.manifest_path
    manifest_path.write_text(
        json.dumps(pipeline_record, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    logger.info("=" * 60)
    logger.info("パイプライン完了 (%.1f秒)", total_elapsed)
    logger.info("実行記録: %s", manifest_path)
    logger.info("=" * 60)


def dry_run(config: PipelineConfig) -> None:
    """設定内容を表示して終了"""
    print("=== パイプライン設定（ドライラン） ===\n")
    print(f"入力ディレクトリ : {config.input_dir}")
    print(f"作業ディレクトリ : {config.work_dir}")
    print(f"出力ディレクトリ : {config.output_dir}")
    print(f"LLM プロバイダー : {config.llm_provider}")
    print(f"LLM モデル       : {config.openai_model}")
    print(f"トークン上限     : {config.token_limit:,}")
    print(f"品質閾値 (H/M)   : {config.quality_threshold_high}/{config.quality_threshold_medium}")
    print(f"実行範囲         : Step {config.step_from} → {config.step_to}")
    print(f"再開モード       : {config.resume_mode}")
    print()

    # 入力ファイル一覧
    if config.input_dir.exists():
        files = list(config.input_dir.rglob("*"))
        files = [f for f in files if f.is_file()]
        print(f"入力ファイル数   : {len(files)}")
        for f in files[:20]:
            print(f"  - {f.relative_to(config.input_dir)}")
        if len(files) > 20:
            print(f"  ... 他 {len(files) - 20} 件")
    else:
        print(f"⚠ 入力ディレクトリが存在しません: {config.input_dir}")

    # 中間ファイルの状態
    print()
    for step_cls in ALL_STEPS:
        step_dir = config.work_dir / f"step{step_cls.step_number}_{step_cls.step_name}"
        manifest_file = step_dir / "manifest.json"
        if manifest_file.exists():
            data = json.loads(manifest_file.read_text(encoding="utf-8"))
            summary = data.get("summary", {})
            print(
                f"Step {step_cls.step_number} ({step_cls.step_name}): "
                f"{data.get('status', '?')} "
                f"— 成功={summary.get('success', 0)}, "
                f"失敗={summary.get('failed', 0)}, "
                f"スキップ={summary.get('skipped', 0)}"
            )
        else:
            print(f"Step {step_cls.step_number} ({step_cls.step_name}): 未実行")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="ドキュメント処理パイプライン",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--from", dest="step_from", type=int, default=None,
        help="開始ステップ番号 (1-6)",
    )
    parser.add_argument(
        "--to", dest="step_to", type=int, default=None,
        help="終了ステップ番号 (1-6)",
    )
    parser.add_argument(
        "--mode", choices=["skip", "overwrite"], default=None,
        help="中間ファイルの扱い (skip: スキップ, overwrite: 上書き)",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="設定内容を表示して終了",
    )
    parser.add_argument(
        "--env", default=".env",
        help=".env ファイルのパス (デフォルト: .env)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    # 設定読み込み（.env → コマンドライン引数で上書き）
    config = PipelineConfig(_env_file=args.env)

    if args.step_from is not None:
        config.step_from = args.step_from
    if args.step_to is not None:
        config.step_to = args.step_to
    if args.mode is not None:
        config.resume_mode = args.mode

    setup_logging(config)

    if args.dry_run:
        dry_run(config)
        return

    run_pipeline(config)


if __name__ == "__main__":
    main()
