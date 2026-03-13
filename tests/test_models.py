"""中間表現モデルのテスト

確認観点:
  - IntermediateDocument の要素追加と JSON シリアライズ
  - 空段落のスキップ
  - FileMetadata / StepResult のシリアライズ
"""

from src.models.intermediate import (
    CellData,
    Confidence,
    ElementType,
    IntermediateDocument,
)
from src.models.metadata import (
    ExtractedFileRecord,
    FileMetadata,
    ProcessStatus,
    StepResult,
)


class TestIntermediateDocument:
    def test_add_heading(self):
        doc = IntermediateDocument()
        doc.add_heading(level=1, text="テスト見出し", detection_method="style")

        assert len(doc.elements) == 1
        assert doc.elements[0].type == ElementType.HEADING
        assert doc.elements[0].content.text == "テスト見出し"
        assert doc.elements[0].content.level == 1

    def test_add_paragraph(self):
        doc = IntermediateDocument()
        doc.add_paragraph("本文テキスト")

        assert len(doc.elements) == 1
        assert doc.elements[0].type == ElementType.PARAGRAPH
        assert doc.elements[0].content.text == "本文テキスト"

    def test_skip_empty_paragraph(self):
        """空段落はスキップされること"""
        doc = IntermediateDocument()
        doc.add_paragraph("")
        doc.add_paragraph("  ")
        doc.add_paragraph("\t")

        assert len(doc.elements) == 0

    def test_add_table(self):
        rows = [
            [CellData(text="ヘッダ", row=0, col=0)],
            [CellData(text="データ", row=1, col=0)],
        ]
        doc = IntermediateDocument()
        doc.add_table(rows=rows, caption="テスト表")

        assert len(doc.elements) == 1
        assert doc.elements[0].type == ElementType.TABLE
        assert doc.elements[0].content.caption == "テスト表"

    def test_add_shape(self):
        doc = IntermediateDocument()
        doc.add_shape(
            shape_type="text_box",
            texts=["開始", "処理A", "終了"],
            confidence=Confidence.MEDIUM,
        )

        assert len(doc.elements) == 1
        assert doc.elements[0].type == ElementType.SHAPE
        assert len(doc.elements[0].content.texts) == 3

    def test_to_dict_roundtrip(self):
        """to_dict() で JSON シリアライズ可能な dict が返ること"""
        doc = IntermediateDocument()
        doc.add_heading(1, "見出し", "style")
        doc.add_paragraph("段落")
        doc.add_table(
            rows=[[CellData(text="A", row=0, col=0)]],
            has_merged_cells=True,
            confidence=Confidence.MEDIUM,
        )

        d = doc.to_dict()
        assert "elements" in d
        assert len(d["elements"]) == 3
        assert d["elements"][0]["type"] == "heading"
        assert d["elements"][1]["type"] == "paragraph"
        assert d["elements"][2]["type"] == "table"

    def test_element_order_preserved(self):
        """要素の追加順序が保持されること"""
        doc = IntermediateDocument()
        doc.add_heading(1, "H1", "style", source_index=0)
        doc.add_paragraph("P1", source_index=1)
        doc.add_heading(2, "H2", "style", source_index=2)
        doc.add_paragraph("P2", source_index=3)

        indices = [e.source_index for e in doc.elements]
        assert indices == [0, 1, 2, 3]


class TestMetadata:
    def test_file_metadata_to_dict(self):
        meta = FileMetadata(
            source_path="機能A/仕様書.docx",
            source_ext=".docx",
            source_size_bytes=12345,
            doc_role_guess="spec_body",
        )
        d = meta.to_dict()
        assert d["source_path"] == "機能A/仕様書.docx"
        assert d["source_ext"] == ".docx"
        assert d["doc_role_guess"] == "spec_body"

    def test_step_result_to_dict(self):
        result = StepResult(
            file_path="test.docx",
            step="extract",
            status=ProcessStatus.SUCCESS,
            message="ok",
            duration_sec=1.5,
        )
        d = result.to_dict()
        assert d["status"] == "success"
        assert d["step"] == "extract"
        assert d["duration_sec"] == 1.5

    def test_extracted_file_record(self):
        meta = FileMetadata(source_path="test.docx", source_ext=".docx")
        doc = IntermediateDocument()
        doc.add_paragraph("テスト")
        record = ExtractedFileRecord(metadata=meta, document=doc.to_dict())

        d = record.to_dict()
        assert "metadata" in d
        assert "document" in d
        assert len(d["document"]["elements"]) == 1
