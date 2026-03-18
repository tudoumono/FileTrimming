"""フォルダスキャン → 各ファイルを個別処理

入力フォルダを再帰的にスキャンし、拡張子に応じた処理を行う。
各ステップは前のステップの出力だけを入力とする。
文書単位の並列化は設定で有効化できる。
"""

from __future__ import annotations

import json
import shutil
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from logging import getLogger
from pathlib import Path

from src.config import PipelineConfig
from src.extractors.registry import get_extractor
from src.llm import create_backend
from src.llm.base import LLMBackend
from src.models.metadata import ProcessStatus, StepResult
from src.pipeline.normalizer import normalize_file
from src.pipeline.splitter import split_if_needed
from src.transform.to_markdown import transform_file

logger = getLogger(__name__)

_transform_backend_local = threading.local()


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


def _collect_json_files(extracted_dir: Path) -> list[Path]:
    """Step3 対象の JSON ファイルを収集する。"""
    return [f for f in _collect_files(extracted_dir, {".json"}) if f.name != "extract_log.jsonl"]


def _normalized_output_path(
    src: Path,
    dst_dir: Path,
    config: PipelineConfig,
) -> Path | None:
    """Step1 後に生成される正規化済みファイルパスを返す。"""
    ext = src.suffix.lower()
    if ext in config.com_normalize_exts:
        dst_ext = ".xlsx" if ext == ".xls" else ".docx"
        return dst_dir / f"{src.stem}{dst_ext}"
    if ext in config.passthrough_exts:
        return dst_dir / src.name
    return None


def _step_result_from_worker_error(
    file_path: str,
    step: str,
    exc: BaseException,
) -> StepResult:
    """ワーカークラッシュを StepResult に丸める。"""
    return StepResult(
        file_path=file_path,
        step=step,
        status=ProcessStatus.ERROR,
        message=f"worker error: {exc}",
    )


def _normalize_worker(
    src: Path,
    config: PipelineConfig,
) -> tuple[Path | None, StepResult]:
    """Step1 を 1 ファイルぶん実行する。"""
    rel = src.relative_to(config.input_dir)
    dst_dir = config.normalized_dir / rel.parent
    result = normalize_file(src, dst_dir, config)

    normalized_path = None
    if result.status in (ProcessStatus.SUCCESS, ProcessStatus.WARNING):
        candidate = _normalized_output_path(src, dst_dir, config)
        if candidate is not None and candidate.exists():
            normalized_path = candidate

    return normalized_path, result


def _extract_worker(
    file_path: Path,
    config: PipelineConfig,
) -> tuple[Path | None, StepResult]:
    """Step2 を 1 ファイルぶん実行する。"""
    t0 = time.perf_counter()

    try:
        rel = file_path.relative_to(config.normalized_dir)
        source_path = str(rel)
        source_ext = file_path.suffix.lower()

        extractor = get_extractor(source_ext)
        if extractor is None:
            return None, StepResult(
                file_path=source_path,
                step="extract",
                status=ProcessStatus.SKIPPED,
                message=f"no extractor for {source_ext}",
            )

        record, result = extractor(file_path, source_path, source_ext, config)
        json_path = None
        if result.status != ProcessStatus.ERROR:
            json_path = config.extracted_dir / rel.with_suffix(".json")
            json_path.parent.mkdir(parents=True, exist_ok=True)
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(record.to_dict(), f, ensure_ascii=False, indent=2)

        return json_path, result
    except Exception as exc:
        elapsed = time.perf_counter() - t0
        logger.exception("Step2 worker unexpected error: %s", file_path)
        return None, StepResult(
            file_path=str(file_path),
            step="extract",
            status=ProcessStatus.ERROR,
            message=f"worker error: {exc}",
            duration_sec=round(elapsed, 2),
        )


def _transform_backend_key(config: PipelineConfig) -> tuple[object, ...]:
    """Step3 のスレッドローカル LLM バックエンド識別子を返す。"""
    return (
        config.llm_backend,
        config.openai_api_key,
        config.openai_base_url,
        config.openai_model,
        config.ollama_base_url,
        config.ollama_model,
        config.llm_proxy_url,
        config.llm_skip_ssl_verify,
        config.llm_observation_only,
    )


def _resolve_transform_backend(config: PipelineConfig) -> LLMBackend:
    """現在のスレッド用に LLM バックエンドを初期化または再利用する。"""
    cache_key = _transform_backend_key(config)
    cached_key = getattr(_transform_backend_local, "cache_key", None)
    cached_backend = getattr(_transform_backend_local, "backend", None)

    if cached_backend is None or cached_key != cache_key:
        cached_backend = create_backend(config)
        _transform_backend_local.cache_key = cache_key
        _transform_backend_local.backend = cached_backend

    return cached_backend


def _transform_document(
    json_path: Path,
    config: PipelineConfig,
    backend: LLMBackend | None,
) -> list[StepResult]:
    """Step3 を 1 ファイルぶん実行する。"""
    rel = json_path.relative_to(config.extracted_dir)
    md_rel = rel.with_suffix(".md")

    observation_path = None
    if backend is not None and backend.supports_table_interpretation():
        review_suffix = (
            ".llm_observation.json"
            if config.llm_observation_only
            else ".llm_review.json"
        )
        observation_path = config.review_dir / rel.with_suffix(review_suffix)

    transformed_path = config.transformed_dir / md_rel
    result = transform_file(
        json_path,
        transformed_path,
        backend=backend,
        observation_only=config.llm_observation_only,
        observation_path=observation_path,
    )
    results = [result]

    if result.status == ProcessStatus.ERROR:
        return results

    split_results = split_if_needed(transformed_path, config)
    results.extend(split_results)

    output_path = config.output_dir / md_rel
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if transformed_path.exists():
        shutil.copy2(transformed_path, output_path)
        for sr in split_results:
            if sr.status != ProcessStatus.SUCCESS:
                continue
            split_file = Path(sr.file_path)
            if not split_file.exists():
                continue
            out_split = config.output_dir / split_file.relative_to(config.transformed_dir)
            out_split.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(split_file, out_split)

    return results


def _transform_worker(
    json_path: Path,
    config: PipelineConfig,
) -> list[StepResult]:
    """Step3 を 1 ファイルぶん並列ワーカーで実行する。"""
    t0 = time.perf_counter()

    try:
        backend = _resolve_transform_backend(config)
        return _transform_document(json_path, config, backend)
    except Exception as exc:
        elapsed = time.perf_counter() - t0
        logger.exception("Step3 worker unexpected error: %s", json_path)
        return [StepResult(
            file_path=str(json_path),
            step="transform",
            status=ProcessStatus.ERROR,
            message=f"worker error: {exc}",
            duration_sec=round(elapsed, 2),
        )]


def run_step1_normalize(config: PipelineConfig) -> list[StepResult]:
    """Step1: 正規化 — input/ → intermediate/01_normalized/"""
    logger.info("=== Step1: 正規化 開始 ===")
    t0 = time.perf_counter()

    target_exts = config.com_normalize_exts | config.passthrough_exts
    files = _collect_files(config.input_dir, target_exts)
    logger.info(
        "対象ファイル: %d 件, workers=%d",
        len(files),
        config.normalize_workers,
    )

    results: list[StepResult] = []
    if config.normalize_workers <= 1 or len(files) <= 1:
        for src in files:
            _, result = _normalize_worker(src, config)
            results.append(result)
    else:
        with ThreadPoolExecutor(max_workers=config.normalize_workers) as executor:
            future_to_src = {
                executor.submit(_normalize_worker, src, config): src
                for src in files
            }
            for future in as_completed(future_to_src):
                src = future_to_src[future]
                try:
                    _, result = future.result()
                except Exception as exc:
                    logger.exception("Step1 worker crash: %s", src)
                    result = _step_result_from_worker_error(str(src), "normalize", exc)
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

    from src.extractors.registry import supported_extensions

    target_exts = supported_extensions()
    files = _collect_files(config.normalized_dir, target_exts)
    logger.info(
        "対象ファイル: %d 件 (拡張子: %s), workers=%d",
        len(files),
        ", ".join(sorted(target_exts)),
        config.extract_workers,
    )

    results: list[StepResult] = []
    if config.extract_workers <= 1 or len(files) <= 1:
        for file_path in files:
            _, result = _extract_worker(file_path, config)
            results.append(result)
    else:
        with ThreadPoolExecutor(max_workers=config.extract_workers) as executor:
            future_to_file = {
                executor.submit(_extract_worker, file_path, config): file_path
                for file_path in files
            }
            for future in as_completed(future_to_file):
                file_path = future_to_file[future]
                try:
                    _, result = future.result()
                except Exception as exc:
                    logger.exception("Step2 worker crash: %s", file_path)
                    result = _step_result_from_worker_error(
                        str(file_path.relative_to(config.normalized_dir)),
                        "extract",
                        exc,
                    )
                results.append(result)

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

    try:
        backend = create_backend(config)
    except Exception as exc:
        elapsed = time.perf_counter() - t0
        logger.exception("LLM バックエンド初期化失敗: backend=%s", config.llm_backend)
        return [StepResult(
            file_path="step3",
            step="transform",
            status=ProcessStatus.ERROR,
            message=f"LLM backend init error: {exc}",
            duration_sec=round(elapsed, 2),
        )]

    logger.info(
        "Step3 LLM バックエンド: %s, workers=%d",
        config.llm_backend,
        config.transform_workers,
    )

    json_files = _collect_json_files(config.extracted_dir)
    logger.info("対象ファイル: %d 件", len(json_files))

    results: list[StepResult] = []
    if config.transform_workers <= 1 or len(json_files) <= 1:
        for json_path in json_files:
            results.extend(_transform_document(json_path, config, backend))
    else:
        with ThreadPoolExecutor(max_workers=config.transform_workers) as executor:
            future_to_json = {
                executor.submit(_transform_worker, json_path, config): json_path
                for json_path in json_files
            }
            for future in as_completed(future_to_json):
                json_path = future_to_json[future]
                try:
                    results.extend(future.result())
                except Exception as exc:
                    logger.exception("Step3 worker crash: %s", json_path)
                    results.append(_step_result_from_worker_error(
                        str(json_path),
                        "transform",
                        exc,
                    ))

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
