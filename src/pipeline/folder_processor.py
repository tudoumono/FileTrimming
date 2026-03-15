"""フォルダスキャン → 各ファイルを個別処理

入力フォルダを再帰的にスキャンし、拡張子に応じた処理を行う。
各ステップは前のステップの出力だけを入力とする。
"""

from __future__ import annotations

import json
import shutil
import time
from logging import getLogger
from pathlib import Path

from src.config import PipelineConfig
from src.extractors.registry import get_extractor
from src.models.metadata import ProcessStatus, StepResult
from src.pipeline.normalizer import normalize_file
from src.pipeline.splitter import split_if_needed
from src.transform.to_markdown import transform_file

logger = getLogger(__name__)


def _write_log(log_path: Path, results: list[StepResult]) -> None:
    """ステップ結果を JSONL ログに書き出す。"""
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with open(log_path, "w", encoding="utf-8") as f:
        for r in results:
            f.write(json.dumps(r.to_dict(), ensure_ascii=False) + "\n")


def _collect_files(input_dir: Path, exts: set[str]) -> list[Path]:
    """対象拡張子のファイルを再帰的に収集する。"""
    files: list[Path] = []
    for p in sorted(input_dir.rglob("*")):
        if p.is_file() and p.suffix.lower() in exts:
            files.append(p)
    return files


def run_step1_normalize(config: PipelineConfig) -> list[StepResult]:
    """Step1: 正規化 — input/ → intermediate/01_normalized/"""
    logger.info("=== Step1: 正規化 開始 ===")
    t0 = time.perf_counter()

    target_exts = config.com_normalize_exts | config.passthrough_exts
    files = _collect_files(config.input_dir, target_exts)
    logger.info("対象ファイル: %d 件", len(files))

    results: list[StepResult] = []
    for src in files:
        rel = src.relative_to(config.input_dir)
        dst_dir = config.normalized_dir / rel.parent
        result = normalize_file(src, dst_dir, config)
        results.append(result)

    _write_log(config.normalized_dir / "normalize_log.jsonl", results)

    elapsed = time.perf_counter() - t0
    ok = sum(1 for r in results if r.status == ProcessStatus.SUCCESS)
    err = sum(1 for r in results if r.status == ProcessStatus.ERROR)
    logger.info("=== Step1 完了: %d成功, %d失敗, %.1fs ===", ok, err, elapsed)
    return results


def run_step2_extract(config: PipelineConfig) -> list[StepResult]:
    """Step2: 構造抽出 — 01_normalized/ → 02_extracted/"""
    logger.info("=== Step2: 構造抽出 開始 ===")
    t0 = time.perf_counter()

    # レジストリに登録済みの拡張子を全て対象にする (.docx, .xlsx 等)
    from src.extractors.registry import supported_extensions
    target_exts = supported_extensions()
    files = _collect_files(config.normalized_dir, target_exts)
    logger.info("対象ファイル: %d 件 (拡張子: %s)", len(files), ", ".join(sorted(target_exts)))

    results: list[StepResult] = []
    for file_path in files:
        rel = file_path.relative_to(config.normalized_dir)

        source_path = str(rel)
        source_ext = file_path.suffix.lower()

        extractor = get_extractor(source_ext)
        if extractor is None:
            results.append(StepResult(
                file_path=source_path, step="extract",
                status=ProcessStatus.SKIPPED,
                message=f"no extractor for {source_ext}",
            ))
            continue

        record, result = extractor(file_path, source_path, source_ext, config)
        results.append(result)

        # JSON 出力
        if result.status != ProcessStatus.ERROR:
            json_path = config.extracted_dir / rel.with_suffix(".json")
            json_path.parent.mkdir(parents=True, exist_ok=True)
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(record.to_dict(), f, ensure_ascii=False, indent=2)

    _write_log(config.extracted_dir / "extract_log.jsonl", results)

    elapsed = time.perf_counter() - t0
    ok = sum(1 for r in results if r.status in (ProcessStatus.SUCCESS, ProcessStatus.WARNING))
    err = sum(1 for r in results if r.status == ProcessStatus.ERROR)
    logger.info("=== Step2 完了: %d成功, %d失敗, %.1fs ===", ok, err, elapsed)
    return results


def run_step3_transform(config: PipelineConfig) -> list[StepResult]:
    """Step3: 変換 — 02_extracted/ → 03_transformed/ → output/"""
    logger.info("=== Step3: Markdown 変換 開始 ===")
    t0 = time.perf_counter()

    json_files = _collect_files(config.extracted_dir, {".json"})
    # ログファイルを除外
    json_files = [f for f in json_files if f.name != "extract_log.jsonl"]
    logger.info("対象ファイル: %d 件", len(json_files))

    results: list[StepResult] = []
    for json_path in json_files:
        rel = json_path.relative_to(config.extracted_dir)
        md_rel = rel.with_suffix(".md")

        # 03_transformed/ への出力
        transformed_path = config.transformed_dir / md_rel
        result = transform_file(json_path, transformed_path)
        results.append(result)

        if result.status != ProcessStatus.ERROR:
            # 15MB 超の分割チェック
            split_results = split_if_needed(transformed_path, config)
            results.extend(split_results)

            # output/ へコピー
            output_path = config.output_dir / md_rel
            output_path.parent.mkdir(parents=True, exist_ok=True)
            if transformed_path.exists():
                shutil.copy2(transformed_path, output_path)
                # 分割ファイルもコピー
                for sr in split_results:
                    if sr.status == ProcessStatus.SUCCESS:
                        split_file = Path(sr.file_path)
                        if split_file.exists():
                            out_split = config.output_dir / split_file.relative_to(config.transformed_dir)
                            out_split.parent.mkdir(parents=True, exist_ok=True)
                            shutil.copy2(split_file, out_split)

    _write_log(config.transformed_dir / "transform_log.jsonl", results)

    elapsed = time.perf_counter() - t0
    ok = sum(1 for r in results if r.status == ProcessStatus.SUCCESS)
    err = sum(1 for r in results if r.status == ProcessStatus.ERROR)
    logger.info("=== Step3 完了: %d成功, %d失敗, %.1fs ===", ok, err, elapsed)
    return results


def run_pipeline(config: PipelineConfig, steps: str = "all") -> dict[str, list[StepResult]]:
    """パイプライン全体またはステップ指定で実行する。

    Args:
        config: パイプライン設定
        steps: "all", "1", "2", "3", "1-2", "2-3" 等

    Returns:
        ステップ名 → 結果リストの辞書
    """
    all_results: dict[str, list[StepResult]] = {}

    run_1 = steps in ("all", "1", "1-2", "1-3")
    run_2 = steps in ("all", "2", "1-2", "2-3", "1-3")
    run_3 = steps in ("all", "3", "2-3", "1-3")

    if run_1:
        all_results["step1"] = run_step1_normalize(config)
    if run_2:
        all_results["step2"] = run_step2_extract(config)
    if run_3:
        all_results["step3"] = run_step3_transform(config)

    return all_results
