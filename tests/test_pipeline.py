"""Step2 → Step3 統合テスト

確認観点:
  - .docx → 中間 JSON → Markdown の一連の流れが正常に動作する
  - JSON ファイルが生成される
  - Markdown ファイルが生成される
  - output/ に最終成果物がコピーされる
  - ログファイルが生成される
  - 15MB 分割 (閾値を下げてテスト)
"""

import json
from pathlib import Path

import pytest

from src.config import PipelineConfig
from src.extractors.word import extract_docx
from src.llm.base import LLMBackend, ReconstructionUnit, TableInterpretationResult
from src.models.metadata import ProcessStatus
from src.pipeline.folder_processor import run_step2_extract, run_step3_transform
from src.pipeline.splitter import split_if_needed
from src.transform.to_markdown import transform_file, transform_to_markdown


class _FakeTableBackend(LLMBackend):
    def generate(self, prompt: str, system: str = "") -> str:
        return ""

    def supports_table_interpretation(self) -> bool:
        return True

    def interpret_table(
        self, unit: ReconstructionUnit, system: str = "",
    ) -> TableInterpretationResult:
        return TableInterpretationResult(
            schema_version="1.0",
            unit_id=unit.unit_id,
            table_type="data_table",
            render_strategy="data_table",
            self_assessment={"confidence": "medium"},
        )


class TestStep2ToStep3:
    """Step2 → Step3 の一連の流れ"""

    def test_docx_to_json_to_markdown(
        self, simple_docx: Path, config: PipelineConfig, tmp_path: Path,
    ):
        """1ファイルの .docx → JSON → Markdown 変換"""
        # Step2: 抽出
        record, result = extract_docx(
            simple_docx, "simple.docx", ".docx", config,
        )
        assert result.status in (ProcessStatus.SUCCESS, ProcessStatus.WARNING)

        # JSON 書き出し
        json_path = tmp_path / "extracted" / "simple.json"
        json_path.parent.mkdir(parents=True, exist_ok=True)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(record.to_dict(), f, ensure_ascii=False, indent=2)
        assert json_path.exists()

        # Step3: Markdown 変換
        md_path = tmp_path / "output" / "simple.md"
        transform_result = transform_file(json_path, md_path)
        assert transform_result.status == ProcessStatus.SUCCESS
        assert md_path.exists()

        # Markdown の内容確認
        md_text = md_path.read_text(encoding="utf-8")
        assert len(md_text) > 0
        assert "概要" in md_text  # 見出しが含まれる
        assert "テスト用の仕様書" in md_text  # 段落が含まれる

    def test_change_history_docx_pipeline(
        self, change_history_docx: Path, config: PipelineConfig, tmp_path: Path,
    ):
        """変更履歴 .docx の Step2 → Step3"""
        record, result = extract_docx(
            change_history_docx, "ch.docx", ".docx", config,
        )

        # doc_role が change_history
        assert record.metadata.doc_role_guess == "change_history"

        # Markdown 変換（マーカーなし、データは出力される）
        md_text = transform_to_markdown(record.to_dict())
        assert "<!--" not in md_text  # マーカーなし
        assert "種別: 追加" in md_text
        assert "2025/01" in md_text  # データは出力される

    def test_mixed_docx_pipeline(
        self, mixed_docx: Path, config: PipelineConfig,
    ):
        """混在 .docx の Step2 → Step3"""
        record, result = extract_docx(
            mixed_docx, "mixed.docx", ".docx", config,
        )
        assert record.metadata.doc_role_guess == "mixed"

        md_text = transform_to_markdown(record.to_dict())
        # 仕様書部分
        assert "機能仕様書" in md_text
        # マーカーなし（品質情報は中間 JSON に記録済み）
        assert "<!--" not in md_text


class TestFolderProcessor:
    """フォルダプロセッサの統合テスト"""

    def test_step2_creates_json(
        self, simple_docx: Path, config: PipelineConfig,
    ):
        """run_step2_extract が JSON を生成すること"""
        # 01_normalized/ にファイルを配置
        norm_dir = config.normalized_dir
        norm_dir.mkdir(parents=True, exist_ok=True)
        import shutil
        shutil.copy2(simple_docx, norm_dir / "simple.docx")

        results = run_step2_extract(config)
        assert len(results) == 1
        assert results[0].status in (ProcessStatus.SUCCESS, ProcessStatus.WARNING)

        # JSON が生成されていること
        json_path = config.extracted_dir / "simple.json"
        assert json_path.exists()

        # JSON の中身を確認
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        assert "metadata" in data
        assert "document" in data
        assert len(data["document"]["elements"]) > 0

    def test_step3_creates_markdown(
        self, simple_docx: Path, config: PipelineConfig,
    ):
        """run_step3_transform が Markdown を生成すること"""
        # Step2 を先に実行
        norm_dir = config.normalized_dir
        norm_dir.mkdir(parents=True, exist_ok=True)
        import shutil
        shutil.copy2(simple_docx, norm_dir / "simple.docx")
        run_step2_extract(config)

        # Step3 実行
        results = run_step3_transform(config)
        ok_results = [r for r in results if r.status == ProcessStatus.SUCCESS]
        assert len(ok_results) >= 1

        # Markdown が生成されていること
        md_path = config.transformed_dir / "simple.md"
        assert md_path.exists()

        # output/ にもコピーされていること
        output_md = config.output_dir / "simple.md"
        assert output_md.exists()

    def test_log_files_created(
        self, simple_docx: Path, config: PipelineConfig,
    ):
        """ログファイルが生成されること"""
        norm_dir = config.normalized_dir
        norm_dir.mkdir(parents=True, exist_ok=True)
        import shutil
        shutil.copy2(simple_docx, norm_dir / "simple.docx")

        run_step2_extract(config)
        run_step3_transform(config)

        extract_log = config.extracted_dir / "extract_log.jsonl"
        transform_log = config.transformed_dir / "transform_log.jsonl"
        assert extract_log.exists()
        assert transform_log.exists()

        # JSONL の各行がパース可能であること
        for line in extract_log.read_text(encoding="utf-8").strip().splitlines():
            entry = json.loads(line)
            assert "status" in entry
            assert "step" in entry

    def test_step3_writes_llm_review_file_in_apply_mode(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path, config: PipelineConfig,
    ):
        """LLM適用モードでもレビューJSONが生成されること"""
        config.llm_backend = "openai"

        sample_record = {
            "metadata": {
                "source_path": "sample.xlsx",
                "source_ext": ".xlsx",
                "doc_role_guess": "unknown",
            },
            "document": {
                "elements": [
                    {"type": "heading", "content": {
                        "text": "申請書", "level": 2, "detection_method": "sheet_name",
                    }},
                    {"type": "table", "content": {
                        "rows": [
                            [
                                {"text": "項目", "row": 0, "col": 0, "is_header": True},
                                {"text": "値", "row": 0, "col": 1, "is_header": True},
                            ],
                            [
                                {"text": "件名", "row": 1, "col": 0},
                                {"text": "サンプル", "row": 1, "col": 1},
                            ],
                        ],
                        "caption": "",
                        "has_merged_cells": False,
                        "confidence": "medium",
                        "fallback_reason": "",
                    }},
                ],
            },
        }

        json_path = config.extracted_dir / "sample.json"
        json_path.parent.mkdir(parents=True, exist_ok=True)
        json_path.write_text(
            json.dumps(sample_record, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        monkeypatch.setattr(
            "src.pipeline.folder_processor.create_backend",
            lambda _config: _FakeTableBackend(),
        )

        results = run_step3_transform(config)
        assert any(r.status == ProcessStatus.SUCCESS for r in results)

        review_path = config.review_dir / "sample.llm_review.json"
        assert review_path.exists()
        review = json.loads(review_path.read_text(encoding="utf-8"))
        assert review["observation_only"] is False
        assert len(review["tables"]) == 1

    def test_step3_parallel_workers_process_multiple_files(
        self, monkeypatch: pytest.MonkeyPatch, config: PipelineConfig,
    ):
        """Step3 が複数ファイルを並列 worker で処理できること"""
        config.llm_backend = "openai"
        config.transform_workers = 2

        sample_record_a = {
            "metadata": {
                "source_path": "sample_a.xlsx",
                "source_ext": ".xlsx",
                "doc_role_guess": "unknown",
            },
            "document": {
                "elements": [
                    {"type": "heading", "content": {
                        "text": "申請書A", "level": 2, "detection_method": "sheet_name",
                    }},
                    {"type": "table", "content": {
                        "rows": [
                            [
                                {"text": "項目", "row": 0, "col": 0, "is_header": True},
                                {"text": "値", "row": 0, "col": 1, "is_header": True},
                            ],
                            [
                                {"text": "件名", "row": 1, "col": 0},
                                {"text": "サンプルA", "row": 1, "col": 1},
                            ],
                        ],
                        "caption": "",
                        "has_merged_cells": False,
                        "confidence": "medium",
                        "fallback_reason": "",
                    }},
                ],
            },
        }
        sample_record_b = {
            "metadata": {
                "source_path": "sample_b.xlsx",
                "source_ext": ".xlsx",
                "doc_role_guess": "unknown",
            },
            "document": {
                "elements": [
                    {"type": "heading", "content": {
                        "text": "申請書B", "level": 2, "detection_method": "sheet_name",
                    }},
                    {"type": "table", "content": {
                        "rows": [
                            [
                                {"text": "項目", "row": 0, "col": 0, "is_header": True},
                                {"text": "値", "row": 0, "col": 1, "is_header": True},
                            ],
                            [
                                {"text": "件名", "row": 1, "col": 0},
                                {"text": "サンプルB", "row": 1, "col": 1},
                            ],
                        ],
                        "caption": "",
                        "has_merged_cells": False,
                        "confidence": "medium",
                        "fallback_reason": "",
                    }},
                ],
            },
        }

        for name, record in {
            "sample_a.json": sample_record_a,
            "sample_b.json": sample_record_b,
        }.items():
            json_path = config.extracted_dir / name
            json_path.parent.mkdir(parents=True, exist_ok=True)
            json_path.write_text(
                json.dumps(record, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

        monkeypatch.setattr(
            "src.pipeline.folder_processor.create_backend",
            lambda _config: _FakeTableBackend(),
        )

        results = run_step3_transform(config)
        ok_results = [
            r for r in results
            if r.step == "transform" and r.status == ProcessStatus.SUCCESS
        ]
        assert len(ok_results) == 2
        assert (config.transformed_dir / "sample_a.md").exists()
        assert (config.transformed_dir / "sample_b.md").exists()


class TestConfigLoading:
    def test_load_env_reads_worker_settings(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
    ):
        """並列 worker 設定が .env から読み込まれること"""
        for key in (
            "NORMALIZE_WORKERS",
            "EXTRACT_WORKERS",
            "TRANSFORM_WORKERS",
            "LLM_OBSERVATION_ONLY",
        ):
            monkeypatch.delenv(key, raising=False)

        env_path = tmp_path / ".env"
        env_path.write_text(
            "\n".join([
                "NORMALIZE_WORKERS=2",
                "EXTRACT_WORKERS=3",
                "TRANSFORM_WORKERS=4",
                "LLM_OBSERVATION_ONLY=true",
            ]),
            encoding="utf-8",
        )

        config = PipelineConfig()
        config.load_env(env_path)
        config.validate()

        assert config.normalize_workers == 2
        assert config.extract_workers == 3
        assert config.transform_workers == 4
        assert config.llm_observation_only is True


class TestSplitter:
    """15MB 分割テスト"""

    def test_no_split_needed(self, tmp_path: Path, config: PipelineConfig):
        """15MB 以下は分割されないこと"""
        md_path = tmp_path / "small.md"
        md_path.write_text("# テスト\n\n本文です。\n", encoding="utf-8")

        results = split_if_needed(md_path, config)
        assert len(results) == 0  # 分割不要

    def test_split_by_heading(self, tmp_path: Path):
        """閾値を下げて分割が動作することを確認"""
        # 小さい閾値で強制的に分割をテスト
        config = PipelineConfig(max_file_size_bytes=100)

        # 100バイト超のファイルを作成
        content = "# セクション1\n\n" + "あ" * 30 + "\n\n"
        content += "# セクション2\n\n" + "い" * 30 + "\n\n"
        content += "# セクション3\n\n" + "う" * 30 + "\n\n"

        md_path = tmp_path / "large.md"
        md_path.write_text(content, encoding="utf-8")

        results = split_if_needed(md_path, config)
        assert len(results) >= 2, "分割ファイルが2つ以上生成されること"

        # 分割ファイルが存在すること
        for r in results:
            assert Path(r.file_path).exists()
            assert r.status == ProcessStatus.SUCCESS

        # 分割ファイル名のパターン
        parts = list(tmp_path.glob("large_part*.md"))
        assert len(parts) >= 2

    def test_no_heading_warning(self, tmp_path: Path):
        """見出しなしの大きいファイルで WARNING が返ること"""
        config = PipelineConfig(max_file_size_bytes=50)

        content = "あ" * 100  # 見出しなしで閾値超え
        md_path = tmp_path / "noheading.md"
        md_path.write_text(content, encoding="utf-8")

        results = split_if_needed(md_path, config)
        assert len(results) == 1
        assert results[0].status == ProcessStatus.WARNING
        assert "cannot split" in results[0].message
