"""Step3 Markdown 変換テスト

確認観点:
  - 見出しの ## レベルが正しい
  - 表がラベル付きテキスト形式 ([行N] ラベル: 値) に変換される
  - 品質マーカーが Markdown に含まれないこと（Dify ノイズ防止）
  - テキストなし図形がスキップされること
  - 段落がそのまま残る
  - 空行の整理
"""

from src.llm.base import LLMBackend, ReconstructionUnit, TableInterpretationResult
from src.models.intermediate import CellData, Confidence, IntermediateDocument
from src.models.metadata import ExtractedFileRecord, FileMetadata
from src.transform.to_markdown import transform_to_markdown


def _make_record(doc: IntermediateDocument) -> dict:
    meta = FileMetadata(source_path="test.docx", source_ext=".docx")
    return ExtractedFileRecord(metadata=meta, document=doc.to_dict()).to_dict()


class _RenderPlanBackend(LLMBackend):
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
            table_type="form",
            render_strategy="form_grid",
            render_plan={"row_roles": ["field_pairs"]},
            self_assessment={"confidence": "medium"},
        )


class _SummaryLabelsBackend(LLMBackend):
    def generate(self, prompt: str, system: str = "") -> str:
        return ""

    def supports_table_interpretation(self) -> bool:
        return True

    def interpret_table(
        self, unit: ReconstructionUnit, system: str = "",
    ) -> TableInterpretationResult:
        previous = unit.context.get("previous_table", {})
        previous_labels = previous.get("column_labels_by_col", [])
        label_by_col = {
            item.get("col"): item.get("label", "")
            for item in previous_labels
            if isinstance(item, dict)
        }

        bounds = unit.context.get("source_bounds", {})
        base_col = bounds.get("col_start", 1)
        first_row = unit.rows[0] if unit.rows else []
        summary_labels: list[str] = []
        for cell in first_row[1:]:
            absolute_col = base_col + int(cell.get("col", 0))
            label = label_by_col.get(absolute_col, "")
            if label:
                summary_labels.append(label)

        return TableInterpretationResult(
            schema_version="1.0",
            unit_id=unit.unit_id,
            table_type="data_table",
            render_strategy="data_table",
            render_plan={"summary_labels": summary_labels},
            self_assessment={"confidence": "medium"},
        )


class _SummaryContextOnlyBackend(LLMBackend):
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


class _DataTablePlanBackend(LLMBackend):
    def generate(self, prompt: str, system: str = "") -> str:
        return ""

    def supports_table_interpretation(self) -> bool:
        return True

    def interpret_table(
        self, unit: ReconstructionUnit, system: str = "",
    ) -> TableInterpretationResult:
        row_roles = ["field_pairs", "", "", "field_pairs", "skip"]
        return TableInterpretationResult(
            schema_version="1.0",
            unit_id=unit.unit_id,
            table_type="data_table",
            render_strategy="data_table",
            header_rows=[1],
            data_start_row=2,
            render_plan={"row_roles": row_roles},
            self_assessment={"confidence": "medium"},
        )


class _MarkdownLinesBackend(LLMBackend):
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
            table_type="form",
            render_strategy="form_grid",
            render_plan={
                "markdown_lines": [
                    "チェック項目",
                    "",
                    "- □ 予算取得済み",
                ],
            },
            self_assessment={"confidence": "medium"},
        )


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
        assert "[行1]" in md
        assert "[行2]" in md

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


class TestMergedCellRendering:
    """結合セル表の描画テスト（P1 改善）"""

    def test_banner_row_renders_as_single_line(self):
        """横結合で全列スパンの行がバナー（太字1行）として出力されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [CellData(text="項目", row=0, col=0, is_header=True),
                 CellData(text="設定値", row=0, col=1, is_header=True),
                 CellData(text="デフォルト", row=0, col=2, is_header=True),
                 CellData(text="備考", row=0, col=3, is_header=True)],
                # バナー行: 1セルが全4列スパン
                [CellData(text="■ 接続設定", row=1, col=0, colspan=4)],
                [CellData(text="ホスト名", row=2, col=0),
                 CellData(text="db-server01", row=2, col=1),
                 CellData(text="localhost", row=2, col=2),
                 CellData(text="", row=2, col=3)],
            ],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))

        # バナー行は太字1行として出力（ラベル:値 形式ではない）
        assert "**■ 接続設定**" in md
        # バナーのテキストがラベル付きで重複しないこと
        assert md.count("■ 接続設定") == 1
        # 通常行はラベル形式
        assert "項目: ホスト名" in md
        assert "設定値: db-server01" in md

    def test_banner_row_all_same_text(self):
        """全セル同一テキストの行（抽出時の重複残骸）もバナーとして扱うこと"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [CellData(text="品目", row=0, col=0, is_header=True),
                 CellData(text="数量", row=0, col=1, is_header=True),
                 CellData(text="単価", row=0, col=2, is_header=True),
                 CellData(text="金額", row=0, col=3, is_header=True)],
                [CellData(text="サーバー", row=1, col=0),
                 CellData(text="2", row=1, col=1),
                 CellData(text="500,000", row=1, col=2),
                 CellData(text="1,000,000", row=1, col=3)],
                # 横結合の合計行（全セル同一テキスト）
                [CellData(text="小計", row=2, col=0),
                 CellData(text="小計", row=2, col=1),
                 CellData(text="小計", row=2, col=2),
                 CellData(text="2,100,000", row=2, col=3)],
            ],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))

        # 通常行は正常出力
        assert "品目: サーバー" in md
        # 合計行: 全セル同一ではない（最後が異なる）のでバナーにはならない
        # → ラベル形式で出力される
        assert "金額: 2,100,000" in md

    def test_full_banner_row(self):
        """全セルが完全同一テキストの行はバナーになること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [CellData(text="A", row=0, col=0, is_header=True),
                 CellData(text="B", row=0, col=1, is_header=True)],
                [CellData(text="区切り", row=1, col=0),
                 CellData(text="区切り", row=1, col=1)],
                [CellData(text="x", row=2, col=0),
                 CellData(text="y", row=2, col=1)],
            ],
        )

        md = transform_to_markdown(_make_record(doc))
        assert "**区切り**" in md
        assert md.count("区切り") == 1  # 重複なし

    def test_multilevel_header(self):
        """多段ヘッダー（colspan + サブヘッダー）が親/子ラベルに結合されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                # 親ヘッダー: テスト項目 | テスト環境(colspan=2) | 本番環境(colspan=2)
                [CellData(text="テスト項目", row=0, col=0, is_header=True),
                 CellData(text="テスト環境", row=0, col=1, is_header=True, colspan=2),
                 CellData(text="本番環境", row=0, col=3, is_header=True, colspan=2)],
                # サブヘッダー
                [CellData(text="テスト項目", row=1, col=0),
                 CellData(text="Windows", row=1, col=1),
                 CellData(text="Linux", row=1, col=2),
                 CellData(text="Windows", row=1, col=3),
                 CellData(text="Linux", row=1, col=4)],
                # データ行
                [CellData(text="機能テスト", row=2, col=0),
                 CellData(text="OK", row=2, col=1),
                 CellData(text="OK", row=2, col=2),
                 CellData(text="NG", row=2, col=3),
                 CellData(text="OK", row=2, col=4)],
            ],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))

        # 多段ヘッダーが結合ラベルになること
        assert "テスト環境/Windows: OK" in md
        assert "テスト環境/Linux: OK" in md
        assert "本番環境/Windows: NG" in md
        assert "本番環境/Linux: OK" in md
        # テスト項目は親子同一なので単一ラベル
        assert "テスト項目: 機能テスト" in md

    def test_colspan_in_data_row(self):
        """データ行の colspan セルが重複出力されないこと"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [CellData(text="品目", row=0, col=0, is_header=True),
                 CellData(text="数量", row=0, col=1, is_header=True),
                 CellData(text="金額", row=0, col=2, is_header=True)],
                # 合計行: 品目+数量のスパン + 金額
                [CellData(text="合計", row=1, col=0, colspan=2),
                 CellData(text="2,310,000", row=1, col=2)],
            ],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))
        # 合計は1回だけ出力（colspan 展開位置はスキップ）
        assert md.count("合計") == 1
        assert "品目: 合計" in md
        assert "金額: 2,310,000" in md

    def test_single_row_header_strip(self):
        """承認欄のような 1 行の短いヘッダー列を key-value にしないこと"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[[
                CellData(text="設備購入稟議書", row=0, col=0, colspan=8, is_header=True),
                CellData(text="担当課長", row=0, col=8, colspan=2, is_header=True),
                CellData(text="部長", row=0, col=10, colspan=2, is_header=True),
                CellData(text="役員", row=0, col=12, colspan=2, is_header=True),
            ]],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))
        assert "設備購入稟議書 | 担当課長 | 部長 | 役員" in md
        assert "設備購入稟議書: 担当課長" not in md

    def test_approval_request_form_grid_regression(self):
        """承認欄・フォーム行・チェック行が混在しても意味関係を崩さないこと"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [
                    CellData(text="設備購入稟議書", row=0, col=0, colspan=8, is_header=True),
                    CellData(text="担当課長", row=0, col=8, colspan=2, is_header=True),
                    CellData(text="部長", row=0, col=10, colspan=2, is_header=True),
                    CellData(text="役員", row=0, col=12, colspan=2, is_header=True),
                ],
                [
                    CellData(text="起案部署", row=1, col=0, colspan=3),
                    CellData(text="情報システム部", row=1, col=3, colspan=3),
                    CellData(text="起案者", row=1, col=6, colspan=3),
                    CellData(text="鈴木 花子", row=1, col=9, colspan=4),
                ],
                [
                    CellData(text="チェック項目", row=2, col=0, colspan=3),
                    CellData(text="□ 予算取得済み", row=2, col=3, colspan=10),
                ],
            ],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))
        assert "設備購入稟議書 | 担当課長 | 部長 | 役員" in md
        assert "起案部署: 情報システム部" in md
        assert "起案者: 鈴木 花子" in md
        assert "チェック項目: □ 予算取得済み" in md
        assert "設備購入稟議書: 担当課長" not in md

    def test_rowspan_header_calendar(self):
        """rowspan を含む 2 段ヘッダーが親子ラベルに結合されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [CellData(text="社員番号", row=0, col=0, rowspan=2),
                 CellData(text="氏名", row=0, col=1, rowspan=2),
                 CellData(text="1", row=0, col=2),
                 CellData(text="2", row=0, col=3)],
                [CellData(text="水", row=1, col=2),
                 CellData(text="木", row=1, col=3)],
                [CellData(text="E001", row=2, col=0),
                 CellData(text="田中", row=2, col=1),
                 CellData(text="出", row=2, col=2),
                 CellData(text="休", row=2, col=3)],
            ],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))
        assert "社員番号: E001" in md
        assert "氏名: 田中" in md
        assert "1/水: 出" in md
        assert "2/木: 休" in md

    def test_summary_header_only_table(self):
        """1行だけの集計表が重複ラベルなしで出力されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[[
                CellData(text="日別集計", row=0, col=0, colspan=3, is_header=True),
                CellData(text="3", row=0, col=3, is_header=True),
                CellData(text="4", row=0, col=4, is_header=True),
                CellData(text="5", row=0, col=5, is_header=True),
            ]],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))
        assert "日別集計: 3 | 4 | 5" in md
        assert "日別集計 | 日別集計 | 日別集計" not in md

    def test_sectioned_two_column_memo_table(self):
        """部門別メモのような2列補足表は deterministic に KV 出力すること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [
                    CellData(text="営業", row=0, col=0),
                    CellData(text="売上データは案件単位で管理", row=0, col=1, colspan=5),
                ],
                [
                    CellData(text="経理", row=1, col=0),
                    CellData(text="入金区分ごとに消込運用が異なる", row=1, col=1, colspan=5),
                ],
                [
                    CellData(text="運用", row=2, col=0),
                    CellData(text="問い合わせ履歴は別システムとも二重管理", row=2, col=1, colspan=5),
                ],
            ],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))
        assert "営業: 売上データは案件単位で管理" in md
        assert "経理: 入金区分ごとに消込運用が異なる" in md
        assert "運用: 問い合わせ履歴は別システムとも二重管理" in md
        assert "**営業 / 売上データは案件単位で管理**" not in md

    def test_single_row_merged_field_pair_table(self):
        """分離抽出された1行の merged field row が form_grid として描画されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[[
                CellData(text="件名", row=0, col=0, rowspan=2, colspan=3, is_header=True),
                CellData(text="受注 CSV 取込レイアウト変更", row=0, col=3, rowspan=2, colspan=11, is_header=True),
            ]],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))
        assert "件名: 受注 CSV 取込レイアウト変更" in md
        assert "**件名 / 受注 CSV 取込レイアウト変更**" not in md

    def test_two_row_merged_field_pairs_table(self):
        """記入例のような2行の merged field row 群が form_grid で出ること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [
                    CellData(text="項目", row=0, col=0, rowspan=2, colspan=2, is_header=True),
                    CellData(text="値", row=0, col=2, rowspan=2, colspan=8, is_header=True),
                ],
                [
                    CellData(text="画面名", row=2, col=0, rowspan=2, colspan=2),
                    CellData(text="受注一括登録", row=2, col=2, rowspan=2, colspan=8),
                ],
            ],
            has_merged_cells=True,
        )

        md = transform_to_markdown(_make_record(doc))
        assert "項目: 値" in md
        assert "画面名: 受注一括登録" in md
        assert "**項目 / 値**" not in md

    def test_llm_render_plan_can_override_form_grid_row_role(self):
        """LLM の render_plan が form_grid の行描画方針へ反映されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[[
                CellData(text="件名", row=0, col=0, colspan=7, is_header=True),
                CellData(text="受注 CSV 取込レイアウト変更", row=0, col=7, colspan=1, is_header=True),
            ]],
            has_merged_cells=True,
        )

        observations: list[dict] = []
        md = transform_to_markdown(
            _make_record(doc),
            backend=_RenderPlanBackend(),
            observation_records=observations,
        )
        assert "件名: 受注 CSV 取込レイアウト変更" in md
        assert observations
        assert observations[0]["decision"]["used_for_rendering"] is True

    def test_llm_summary_labels_can_use_previous_table_context(self):
        """LLM が直前テーブルの列ラベル文脈を使って集計行の出力方針を返せること"""
        doc = IntermediateDocument()
        doc.add_heading(2, "2026年04月勤怠", detection_method="sheet_name")
        doc.add_table(
            rows=[
                [
                    CellData(text="社員番号", row=0, col=0, rowspan=2, is_header=True),
                    CellData(text="氏名", row=0, col=1, rowspan=2, is_header=True),
                    CellData(text="所属", row=0, col=2, rowspan=2, is_header=True),
                    CellData(text="1", row=0, col=3, is_header=True),
                    CellData(text="2", row=0, col=4, is_header=True),
                    CellData(text="3", row=0, col=5, is_header=True),
                ],
                [
                    CellData(text="水", row=1, col=3),
                    CellData(text="木", row=1, col=4),
                    CellData(text="金", row=1, col=5),
                ],
                [
                    CellData(text="E001", row=2, col=0),
                    CellData(text="田中", row=2, col=1),
                    CellData(text="営業1課", row=2, col=2),
                    CellData(text="出", row=2, col=3),
                    CellData(text="休", row=2, col=4),
                    CellData(text="出", row=2, col=5),
                ],
            ],
            has_merged_cells=True,
            source_col_start=1,
            source_col_end=6,
        )
        doc.add_table(
            rows=[[
                CellData(text="日別集計", row=0, col=0, colspan=3, is_header=True),
                CellData(text="=COUNTIF(D5:D5,\"出\")", row=0, col=3, is_header=True),
                CellData(text="=COUNTIF(E5:E5,\"出\")", row=0, col=4, is_header=True),
                CellData(text="=COUNTIF(F5:F5,\"出\")", row=0, col=5, is_header=True),
            ]],
            has_merged_cells=True,
            source_col_start=1,
            source_col_end=6,
        )

        observations: list[dict] = []
        md = transform_to_markdown(
            _make_record(doc),
            backend=_SummaryLabelsBackend(),
            observation_records=observations,
        )

        assert "日別集計\n  1/水: =COUNTIF(D5:D5,\"出\")" in md
        assert "  2/木: =COUNTIF(E5:E5,\"出\")" in md
        assert "  3/金: =COUNTIF(F5:F5,\"出\")" in md
        assert observations[-1]["decision"]["used_for_rendering"] is True

    def test_llm_summary_can_fall_back_to_previous_table_context(self):
        """LLM が summary_labels を省略しても直前テーブル文脈で補完できること"""
        doc = IntermediateDocument()
        doc.add_heading(2, "2026年04月勤怠", detection_method="sheet_name")
        doc.add_table(
            rows=[
                [
                    CellData(text="社員番号", row=0, col=0, rowspan=2, is_header=True),
                    CellData(text="氏名", row=0, col=1, rowspan=2, is_header=True),
                    CellData(text="所属", row=0, col=2, rowspan=2, is_header=True),
                    CellData(text="1", row=0, col=3, is_header=True),
                    CellData(text="2", row=0, col=4, is_header=True),
                    CellData(text="3", row=0, col=5, is_header=True),
                ],
                [
                    CellData(text="水", row=1, col=3),
                    CellData(text="木", row=1, col=4),
                    CellData(text="金", row=1, col=5),
                ],
                [
                    CellData(text="E001", row=2, col=0),
                    CellData(text="田中", row=2, col=1),
                    CellData(text="営業1課", row=2, col=2),
                    CellData(text="出", row=2, col=3),
                    CellData(text="休", row=2, col=4),
                    CellData(text="出", row=2, col=5),
                ],
            ],
            has_merged_cells=True,
            source_col_start=1,
            source_col_end=6,
        )
        doc.add_table(
            rows=[[
                CellData(text="日別集計", row=0, col=0, colspan=3, is_header=True),
                CellData(text="=COUNTIF(D5:D5,\"出\")", row=0, col=3, is_header=True),
                CellData(text="=COUNTIF(E5:E5,\"出\")", row=0, col=4, is_header=True),
                CellData(text="=COUNTIF(F5:F5,\"出\")", row=0, col=5, is_header=True),
            ]],
            has_merged_cells=True,
            source_col_start=1,
            source_col_end=6,
        )

        observations: list[dict] = []
        md = transform_to_markdown(
            _make_record(doc),
            backend=_SummaryContextOnlyBackend(),
            observation_records=observations,
        )

        assert "日別集計\n  1/水: =COUNTIF(D5:D5,\"出\")" in md
        assert "  2/木: =COUNTIF(E5:E5,\"出\")" in md
        assert "  3/金: =COUNTIF(F5:F5,\"出\")" in md
        assert observations[-1]["decision"]["used_for_rendering"] is True

    def test_llm_data_table_plan_can_change_preheader_and_mixed_rows(self):
        """LLM の data_table render_plan が pre-header / mixed-row / skip へ反映されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [
                    CellData(text="請求先", row=0, col=0, colspan=2),
                    CellData(text="株式会社サンプル", row=0, col=2, colspan=2),
                ],
                [
                    CellData(text="品目", row=1, col=0, is_header=True),
                    CellData(text="数量", row=1, col=1, is_header=True),
                    CellData(text="単価", row=1, col=2, is_header=True),
                    CellData(text="金額", row=1, col=3, is_header=True),
                ],
                [
                    CellData(text="サーバ", row=2, col=0),
                    CellData(text="2", row=2, col=1),
                    CellData(text="500,000", row=2, col=2),
                    CellData(text="1,000,000", row=2, col=3),
                ],
                [
                    CellData(text="備考", row=3, col=0),
                    CellData(text="保守込み", row=3, col=1, colspan=3),
                ],
                [
                    CellData(text="社内メモ", row=4, col=0, colspan=4),
                ],
            ],
            has_merged_cells=True,
        )

        observations: list[dict] = []
        md = transform_to_markdown(
            _make_record(doc),
            backend=_DataTablePlanBackend(),
            observation_records=observations,
        )

        assert "請求先: 株式会社サンプル" in md
        assert "備考: 保守込み" in md
        assert "社内メモ" not in md
        assert "[行1]" in md
        assert observations[-1]["decision"]["used_for_rendering"] is True

    def test_llm_markdown_lines_can_replace_table_body(self):
        """LLM の markdown_lines がテーブル本文として採用されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [
                    CellData(text="チェック項目", row=0, col=0, colspan=3, is_header=True),
                    CellData(text="□ 予算取得済み", row=0, col=3, colspan=5, is_header=True),
                ],
            ],
            has_merged_cells=True,
        )

        observations: list[dict] = []
        md = transform_to_markdown(
            _make_record(doc),
            backend=_MarkdownLinesBackend(),
            observation_records=observations,
        )

        assert "チェック項目" in md
        assert "- □ 予算取得済み" in md
        assert observations[-1]["decision"]["used_for_rendering"] is True


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


class TestWorkflowRendering:
    """フロー図描画テスト（P3）"""

    def test_workflow_numbered_list(self):
        """フロー図が番号付きリストで出力されること"""
        doc = IntermediateDocument()
        doc.add_shape(
            shape_type="workflow",
            texts=["申請者 / 申請書作成", "上長 / 内容確認", "部長 / 承認判断"],
        )

        md = transform_to_markdown(_make_record(doc))
        assert "[フロー図]" in md
        assert "1. 申請者 / 申請書作成" in md
        assert "2. 上長 / 内容確認" in md
        assert "3. 部長 / 承認判断" in md

    def test_workflow_with_description(self):
        """description がある場合はそちらが優先されること"""
        doc = IntermediateDocument()
        doc.add_shape(
            shape_type="workflow",
            texts=["ステップA", "ステップB"],
            description="申請→承認→処理のフロー",
        )

        md = transform_to_markdown(_make_record(doc))
        assert "申請→承認→処理のフロー" in md
        assert "1. ステップA" not in md


class TestCaptionNoDuplication:
    """キャプション重複なしの E2E テスト（P5）"""

    def test_caption_appears_once_in_markdown(self):
        """キャプションが Markdown に 1 回だけ出力されること"""
        doc = IntermediateDocument()
        doc.add_table(
            rows=[
                [CellData(text="項目", row=0, col=0, is_header=True),
                 CellData(text="値", row=0, col=1, is_header=True)],
                [CellData(text="A", row=1, col=0),
                 CellData(text="1", row=1, col=1)],
            ],
            caption="機能一覧表",
        )

        md = transform_to_markdown(_make_record(doc))
        assert md.count("機能一覧表") == 1
        assert "**機能一覧表**" in md


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
