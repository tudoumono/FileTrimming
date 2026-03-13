"""Step3 Markdown 変換テスト

確認観点:
  - 見出しの ## レベルが正しい
  - 表がラベル付きテキスト形式 ([行N] ラベル: 値) に変換される
  - 品質マーカーが Markdown に含まれないこと（Dify ノイズ防止）
  - テキストなし図形がスキップされること
  - 段落がそのまま残る
  - 空行の整理
"""

from src.models.intermediate import CellData, Confidence, IntermediateDocument
from src.models.metadata import ExtractedFileRecord, FileMetadata
from src.transform.to_markdown import transform_to_markdown


def _make_record(doc: IntermediateDocument) -> dict:
    meta = FileMetadata(source_path="test.docx", source_ext=".docx")
    return ExtractedFileRecord(metadata=meta, document=doc.to_dict()).to_dict()


class TestHeadingRendering:
    def test_heading_levels(self):
        """見出しの ## レベルが正しく出力されること"""
        doc = IntermediateDocument()
        doc.add_heading(1, "レベル1", "style")
        doc.add_heading(2, "レベル2", "style")
        doc.add_heading(3, "レベル3", "font_size")

        md = transform_to_markdown(_make_record(doc))
        assert "# レベル1" in md
        assert "## レベル2" in md
        assert "### レベル3" in md

    def test_heading_max_level(self):
        """レベル6を超えないこと"""
        doc = IntermediateDocument()
        doc.add_heading(6, "レベル6", "style")

        md = transform_to_markdown(_make_record(doc))
        assert "###### レベル6" in md


class TestParagraphRendering:
    def test_plain_paragraph(self):
        """段落がそのまま出力されること"""
        doc = IntermediateDocument()
        doc.add_paragraph("テスト段落です。")

        md = transform_to_markdown(_make_record(doc))
        assert "テスト段落です。" in md

    def test_list_item(self):
        """リスト項目が - 形式で出力されること"""
        doc = IntermediateDocument()
        doc.add_paragraph("項目A", is_list_item=True, list_level=0)
        doc.add_paragraph("項目B", is_list_item=True, list_level=1)

        md = transform_to_markdown(_make_record(doc))
        assert "- 項目A" in md
        assert "  - 項目B" in md


class TestTableRendering:
    def test_labeled_text_format(self):
        """表がラベル付きテキスト形式に変換されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [CellData(text="項目", row=0, col=0, is_header=True),
                 CellData(text="内容", row=0, col=1, is_header=True)],
                [CellData(text="機能A", row=1, col=0),
                 CellData(text="入力チェック", row=1, col=1)],
                [CellData(text="機能B", row=2, col=0),
                 CellData(text="データ更新", row=2, col=1)],
            ],
        )

        md = transform_to_markdown(_make_record(doc))

        # ラベル: 値 形式
        assert "項目: 機能A" in md
        assert "内容: 入力チェック" in md
        assert "項目: 機能B" in md
        assert "内容: データ更新" in md

        # 行番号
        assert "[行2]" in md
        assert "[行3]" in md

    def test_table_caption(self):
        """表のキャプションが出力されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[[CellData(text="A", row=0, col=0)]],
            caption="エラーコード一覧",
        )

        md = transform_to_markdown(_make_record(doc))
        assert "**エラーコード一覧**" in md

    def test_no_change_history_marker_in_md(self):
        """変更履歴テーブルのマーカーが Markdown に含まれないこと（品質情報は中間 JSON に記録済み）"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [CellData(text="ページ", row=0, col=0),
                 CellData(text="種別", row=0, col=1),
                 CellData(text="年月", row=0, col=2),
                 CellData(text="記事", row=0, col=3)],
                [CellData(text="1", row=1, col=0),
                 CellData(text="新規", row=1, col=1),
                 CellData(text="2025/01", row=1, col=2),
                 CellData(text="初版", row=1, col=3)],
            ],
            confidence=Confidence.HIGH,
            fallback_reason="change_history_table",
        )

        md = transform_to_markdown(_make_record(doc))
        assert "<!--" not in md  # HTML コメント形式のマーカーなし
        assert "種別: 新規" in md  # データ自体は出力される

    def test_no_low_confidence_marker_in_md(self):
        """LOW_CONFIDENCE マーカーが Markdown に含まれないこと"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[[CellData(text="A", row=0, col=0)]],
            confidence=Confidence.LOW,
            fallback_reason="complex_merged_cells",
        )

        md = transform_to_markdown(_make_record(doc))
        assert "LOW_CONFIDENCE" not in md
        assert "<!--" not in md

    def test_header_only_table(self):
        """ヘッダーのみの表が出力されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[[
                CellData(text="A", row=0, col=0),
                CellData(text="B", row=0, col=1),
            ]],
        )

        md = transform_to_markdown(_make_record(doc))
        assert "A" in md
        assert "B" in md


class TestShapeRendering:
    def test_shape_with_texts(self):
        """テキスト付き図形がリスト形式で出力されること"""
        doc = IntermediateDocument()
        doc.add_shape(
            shape_type="text_box",
            texts=["開始", "処理A", "終了"],
        )

        md = transform_to_markdown(_make_record(doc))
        assert "[図形]" in md
        assert "- 開始" in md
        assert "- 処理A" in md
        assert "- 終了" in md

    def test_shape_no_text_keeps_placeholder(self):
        """テキストなし図形でも存在を示すプレースホルダが出力されること"""
        doc = IntermediateDocument()
        doc.add_shape(
            shape_type="vml",
            texts=[],
            confidence=Confidence.LOW,
            fallback_reason="no_text_content",
        )

        md = transform_to_markdown(_make_record(doc))
        assert "[図形: vml]" in md
        assert "LOW_CONFIDENCE" not in md
        assert "<!--" not in md

    def test_shape_with_description(self):
        """LLM 生成の説明文がある場合はそれが使われること"""
        doc = IntermediateDocument()
        doc.add_shape(
            shape_type="flowchart",
            texts=["開始", "終了"],
            description="開始から終了までの一本道のフロー図です。",
        )

        md = transform_to_markdown(_make_record(doc))
        assert "開始から終了までの一本道のフロー図です。" in md
        # texts のリストは出力されない（description が優先）
        assert "- 開始" not in md


class TestOutputFormat:
    def test_no_yaml_frontmatter(self):
        """YAML front matter が含まれないこと (Dify が認識しないため)"""
        doc = IntermediateDocument()
        doc.add_heading(1, "テスト", "style")
        doc.add_paragraph("本文")

        md = transform_to_markdown(_make_record(doc))
        assert not md.startswith("---")

    def test_ends_with_newline(self):
        """末尾が改行で終わること"""
        doc = IntermediateDocument()
        doc.add_paragraph("テスト")

        md = transform_to_markdown(_make_record(doc))
        assert md.endswith("\n")
