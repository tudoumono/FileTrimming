"""テストデータ .xlsx 生成スクリプト

openpyxl で作成可能なパターンのテストデータを生成する。

使い方:
  python tools/generate_excel_test_data.py [--output-dir input/]

生成ファイル:
  1. many_tables.xlsx            - 複数シートに業務表を配置したブック
  2. multiple_tables_sheet.xlsx  - 1 シート内に複数表が混在するブック
  3. merged_cells.xlsx           - 縦結合・横結合・複合ヘッダを含むブック
  4. formulas_and_formats.xlsx   - 数式、表示形式、条件付き書式を含むブック
  5. comments_and_annotations.xlsx - コメント、色、リンク、入力規則を含むブック
  6. many_images.xlsx            - 画像を多数配置したブック
  7. mixed_complex.xlsx          - 複数パターンを混在させた総合ブック
  8. large_workbook.xlsx         - 大きめのシートを含むブック
  9. change_history.xlsx         - 変更履歴主体のブック
"""

from __future__ import annotations

import argparse
import io
from datetime import date, timedelta
from functools import lru_cache
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError as exc:
    raise SystemExit("openpyxl が必要です。`pip install openpyxl` を実行してください。") from exc

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError as exc:
    raise SystemExit("Pillow が必要です。`pip install Pillow` を実行してください。") from exc


HEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
SUBHEADER_FILL = PatternFill("solid", fgColor="E2F0D9")
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")
OK_FILL = PatternFill("solid", fgColor="E2F0D9")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
NG_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN_SIDE = Side(style="thin", color="B7C9E2")
ALL_BORDER = Border(
    left=THIN_SIDE,
    right=THIN_SIDE,
    top=THIN_SIDE,
    bottom=THIN_SIDE,
)
JP_FONT_CANDIDATES = [
    Path("C:/Windows/Fonts/msgothic.ttc"),
    Path("C:/Windows/Fonts/YuGothM.ttc"),
    Path("C:/Windows/Fonts/meiryo.ttc"),
    Path("C:/Windows/Fonts/MSMINCHO.TTC"),
]


def _style_cell(cell, *, bold: bool = False, fill=None, number_format: str | None = None) -> None:
    cell.font = Font(bold=bold)
    cell.fill = fill or PatternFill(fill_type=None)
    cell.border = ALL_BORDER
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    if number_format:
        cell.number_format = number_format


def _autosize_columns(ws, *, max_width: int = 36, sample_rows: int = 200) -> None:
    row_limit = min(ws.max_row, sample_rows)
    for col_idx in range(1, ws.max_column + 1):
        width = 10
        for row_idx in range(1, row_limit + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            width = max(width, min(max_width, len(str(value)) + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


@lru_cache(maxsize=16)
def _load_japanese_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for font_path in JP_FONT_CANDIDATES:
        if not font_path.exists():
            continue
        try:
            return ImageFont.truetype(str(font_path), size)
        except OSError:
            continue
    return ImageFont.load_default()


def _draw_centered_label(
    draw: ImageDraw.ImageDraw,
    *,
    box: tuple[int, int, int, int],
    text: str,
    fill: str | tuple[int, int, int],
    font_size: int,
) -> None:
    font = _load_japanese_font(font_size)
    spacing = 4
    left, top, right, bottom = box
    bbox = draw.multiline_textbbox((0, 0), text, font=font, spacing=spacing, align="center")
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = left + max(0, (right - left - text_width) // 2)
    y = top + max(0, (bottom - top - text_height) // 2)
    draw.multiline_text((x, y), text, fill=fill, font=font, spacing=spacing, align="center")


def _write_table(
    ws,
    *,
    start_row: int,
    start_col: int,
    headers: list[str],
    rows: list[list[object]],
    table_name: str,
    title: str | None = None,
) -> int:
    current_row = start_row
    if title:
        ws.cell(row=current_row, column=start_col, value=title)
        _style_cell(ws.cell(row=current_row, column=start_col), bold=True, fill=SUBHEADER_FILL)
        ws.merge_cells(
            start_row=current_row,
            start_column=start_col,
            end_row=current_row,
            end_column=start_col + len(headers) - 1,
        )
        current_row += 1

    for offset, header in enumerate(headers):
        cell = ws.cell(row=current_row, column=start_col + offset, value=header)
        _style_cell(cell, bold=True, fill=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_offset, row_values in enumerate(rows, start=1):
        for col_offset, value in enumerate(row_values):
            cell = ws.cell(row=current_row + row_offset, column=start_col + col_offset, value=value)
            _style_cell(cell)

    end_row = current_row + len(rows)
    end_col = start_col + len(headers) - 1
    table = Table(
        displayName=table_name,
        ref=f"{get_column_letter(start_col)}{current_row}:{get_column_letter(end_col)}{end_row}",
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)
    return end_row


def _make_labeled_image(label: str, size: tuple[int, int], color: tuple[int, int, int]) -> XLImage:
    img = Image.new("RGB", size, color)
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((4, 4, size[0] - 4, size[1] - 4), radius=12, outline="white", width=3)
    _draw_centered_label(
        draw,
        box=(12, 12, size[0] - 12, size[1] - 12),
        text=label,
        fill="white",
        font_size=max(14, min(size) // 7),
    )
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)
    xl_image = XLImage(buffer)
    xl_image._buffer = buffer
    xl_image.width, xl_image.height = size
    return xl_image


def generate_many_tables(output_dir: Path) -> Path:
    wb = Workbook()
    sheet_specs = [
        (
            "ユーザー管理",
            ["ID", "氏名", "メール", "権限"],
            [
                ["001", "田中太郎", "tanaka@example.com", "管理者"],
                ["002", "鈴木花子", "suzuki@example.com", "一般"],
                ["003", "佐藤次郎", "sato@example.com", "閲覧者"],
                ["004", "高橋美咲", "takahashi@example.com", "一般"],
            ],
            "UserTable",
        ),
        (
            "エラーコード",
            ["コード", "種別", "メッセージ", "対処法"],
            [
                ["E001", "入力", "必須項目が未入力です", "項目を入力してください"],
                ["E002", "入力", "形式が不正です", "正しい形式で入力してください"],
                ["E003", "DB", "接続に失敗しました", "DB サーバーの状態を確認"],
                ["E004", "DB", "タイムアウト", "ネットワーク状態を確認"],
                ["E005", "認証", "トークンが無効です", "再ログインしてください"],
            ],
            "ErrorTable",
        ),
        (
            "画面一覧",
            ["画面ID", "画面名", "概要"],
            [
                ["SC001", "ログイン画面", "認証を行う"],
                ["SC002", "ダッシュボード", "集計情報を表示"],
                ["SC003", "ユーザー一覧", "ユーザーの検索・表示"],
                ["SC004", "ユーザー詳細", "ユーザー情報の編集"],
                ["SC005", "帳票出力", "PDF/Excel 出力"],
            ],
            "ScreenTable",
        ),
        (
            "API一覧",
            ["Method", "Path", "説明", "認証"],
            [
                ["GET", "/api/users", "ユーザー一覧取得", "要"],
                ["POST", "/api/users", "ユーザー作成", "要"],
                ["GET", "/api/users/{id}", "ユーザー詳細取得", "要"],
                ["PUT", "/api/users/{id}", "ユーザー更新", "要"],
                ["DELETE", "/api/users/{id}", "ユーザー削除", "要（管理者）"],
            ],
            "ApiTable",
        ),
        (
            "環境情報",
            ["環境", "サーバー", "DB", "URL"],
            [
                ["開発", "dev-app01", "dev-db01", "https://dev.example.com"],
                ["検証", "stg-app01", "stg-db01", "https://stg.example.com"],
                ["本番", "prd-app01/02", "prd-db01(M)/02(S)", "https://www.example.com"],
            ],
            "EnvironmentTable",
        ),
        (
            "設定値",
            ["パラメータ", "デフォルト値", "説明"],
            [
                ["MAX_RETRY", "3", "リトライ回数上限"],
                ["TIMEOUT_SEC", "30", "タイムアウト秒数"],
                ["BATCH_SIZE", "1000", "バッチ処理単位"],
                ["LOG_LEVEL", "INFO", "ログレベル"],
                ["CACHE_TTL", "3600", "キャッシュ有効秒数"],
            ],
            "ConfigTable",
        ),
    ]

    for idx, (title, headers, rows, table_name) in enumerate(sheet_specs):
        ws = wb.active if idx == 0 else wb.create_sheet()
        ws.title = title
        ws.freeze_panes = "A2"
        _write_table(ws, start_row=1, start_col=1, headers=headers, rows=rows, table_name=table_name)
        note = ws.cell(row=len(rows) + 4, column=1, value=f"補足: {title} のテスト用データ。")
        _style_cell(note, bold=True, fill=NOTE_FILL)
        ws.merge_cells(start_row=len(rows) + 4, start_column=1, end_row=len(rows) + 4, end_column=len(headers))
        _autosize_columns(ws)

    path = output_dir / "many_tables.xlsx"
    wb.save(path)
    print(f"  生成: {path}")
    return path


def generate_multiple_tables_sheet(output_dir: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "複数表混在"
    ws["A1"] = "1 シート内に表が点在するパターン"
    _style_cell(ws["A1"], bold=True, fill=NOTE_FILL)
    ws.merge_cells("A1:H1")

    _write_table(
        ws,
        start_row=3,
        start_col=1,
        headers=["帳票ID", "帳票名", "出力形式"],
        rows=[
            ["RP001", "売上日報", "PDF"],
            ["RP002", "月次実績", "Excel"],
            ["RP003", "在庫一覧", "Excel"],
        ],
        table_name="ReportTable",
        title="帳票一覧",
    )

    _write_table(
        ws,
        start_row=3,
        start_col=6,
        headers=["バッチID", "処理名", "実行時刻", "担当"],
        rows=[
            ["B001", "日次集計", "02:00", "運用A"],
            ["B002", "月次締め", "01:00", "運用B"],
            ["B003", "データ連携", "06:00/18:00", "運用A"],
        ],
        table_name="BatchTable",
        title="バッチ一覧",
    )

    _write_table(
        ws,
        start_row=13,
        start_col=2,
        headers=["項目", "内容"],
        rows=[
            ["備考", "2 つの表の間に空白列がある。"],
            ["確認観点", "used_range が広がっても個別の表を認識できるか。"],
            ["注意点", "途中のメモ行や離れたセルも混ざる。"],
        ],
        table_name="MemoTable",
        title="補足メモ",
    )

    ws["J12"] = "孤立セル"
    _style_cell(ws["J12"], bold=True, fill=WARN_FILL)
    ws["J13"] = "右下に単独のメモを配置"
    _style_cell(ws["J13"])
    ws.freeze_panes = "A3"
    _autosize_columns(ws)

    ws2 = wb.create_sheet("右寄せ配置")
    ws2["C2"] = "右寄せ・下寄せの表配置"
    _style_cell(ws2["C2"], bold=True, fill=NOTE_FILL)
    ws2.merge_cells("C2:I2")
    _write_table(
        ws2,
        start_row=5,
        start_col=4,
        headers=["機能ID", "機能名", "状態"],
        rows=[
            ["F001", "受注登録", "稼働中"],
            ["F002", "出荷指示", "テスト中"],
            ["F003", "在庫照会", "企画中"],
        ],
        table_name="FunctionTable",
    )
    _write_table(
        ws2,
        start_row=18,
        start_col=10,
        headers=["キー", "値"],
        rows=[
            ["TOKEN_LIMIT", "8000"],
            ["MAX_RETRY", "3"],
            ["TIMEOUT", "30"],
        ],
        table_name="ParameterTable",
    )
    _autosize_columns(ws2)

    path = output_dir / "multiple_tables_sheet.xlsx"
    wb.save(path)
    print(f"  生成: {path}")
    return path


def generate_merged_cells(output_dir: Path) -> Path:
    wb = Workbook()

    ws = wb.active
    ws.title = "階層構造"
    headers = ["大分類", "中分類", "項目", "値"]
    for idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=idx, value=header)
        _style_cell(cell, bold=True, fill=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    data_rows = [
        ["入力系", "ファイル入力", "形式", "CSV"],
        ["", "", "文字コード", "UTF-8"],
        ["", "DB入力", "接続先", "PostgreSQL"],
        ["", "", "タイムアウト", "30秒"],
        ["出力系", "帳票", "形式", "PDF"],
        ["", "", "用紙サイズ", "A4"],
        ["", "ログ", "出力先", "/var/log/app"],
    ]
    for row_idx, row_values in enumerate(data_rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            _style_cell(cell)

    ws.merge_cells("A2:A5")
    ws.merge_cells("B2:B3")
    ws.merge_cells("B4:B5")
    ws.merge_cells("A6:A8")
    ws.merge_cells("B6:B7")
    for ref in ("A2", "B2", "B4", "A6", "B6"):
        ws[ref].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws2 = wb.create_sheet("複合ヘッダ")
    for idx, header in enumerate(["コード", "名称", "入力", "", "出力", ""], start=1):
        cell = ws2.cell(row=1, column=idx, value=header)
        _style_cell(cell, bold=True, fill=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws2.merge_cells("A1:A2")
    ws2.merge_cells("B1:B2")
    ws2.merge_cells("C1:D1")
    ws2.merge_cells("E1:F1")
    subheaders = ["ファイル", "DB", "画面", "帳票"]
    for idx, header in enumerate(subheaders, start=3):
        cell = ws2.cell(row=2, column=idx, value=header)
        _style_cell(cell, bold=True, fill=SUBHEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    rows2 = [
        ["F001", "受注処理", "CSV", "SELECT/INSERT", "一覧画面", "受注伝票"],
        ["F002", "出荷処理", "-", "SELECT/UPDATE", "詳細画面", "出荷指示書"],
        ["F003", "在庫照会", "-", "SELECT", "照会画面", "-"],
        ["F004", "月次集計", "CSV", "SELECT", "-", "月次レポート"],
    ]
    for row_idx, row_values in enumerate(rows2, start=3):
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws2.cell(row=row_idx, column=col_idx, value=value)
            _style_cell(cell)

    ws3 = wb.create_sheet("不規則結合")
    headers3 = ["品目", "数量", "単価", "金額"]
    for idx, header in enumerate(headers3, start=1):
        cell = ws3.cell(row=1, column=idx, value=header)
        _style_cell(cell, bold=True, fill=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    rows3 = [
        ["サーバー", 2, 500000, 1000000],
        ["ストレージ", 4, 200000, 800000],
        ["ネットワーク機器", 1, 300000, 300000],
        ["小計", "", "", 2100000],
        ["消費税（10%）", "", "", 210000],
        ["合計", "", "", 2310000],
    ]
    for row_idx, row_values in enumerate(rows3, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws3.cell(row=row_idx, column=col_idx, value=value)
            _style_cell(cell, number_format="#,##0")
    for ref in ("A5:C5", "A6:C6", "A7:C7"):
        ws3.merge_cells(ref)
    for ref in ("A5", "A6", "A7"):
        ws3[ref].alignment = Alignment(horizontal="right", vertical="center")

    for sheet in wb.worksheets:
        _autosize_columns(sheet)

    path = output_dir / "merged_cells.xlsx"
    wb.save(path)
    print(f"  生成: {path}")
    return path


def generate_formulas_and_formats(output_dir: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "売上集計"
    headers = ["日付", "部門", "担当", "数量", "単価", "売上", "消費税", "税込売上", "達成率", "判定"]
    for idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=idx, value=header)
        _style_cell(cell, bold=True, fill=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    base_date = date(2026, 4, 1)
    departments = ["営業1課", "営業2課", "営業3課", "営業4課"]
    reps = ["田中", "鈴木", "佐藤", "高橋", "伊藤", "渡辺"]

    for row_idx in range(2, 14):
        ws.cell(row=row_idx, column=1, value=base_date + timedelta(days=row_idx - 2))
        ws.cell(row=row_idx, column=2, value=departments[(row_idx - 2) % len(departments)])
        ws.cell(row=row_idx, column=3, value=reps[(row_idx - 2) % len(reps)])
        ws.cell(row=row_idx, column=4, value=10 + (row_idx * 2))
        ws.cell(row=row_idx, column=5, value=12000 + (row_idx * 350))
        ws.cell(row=row_idx, column=6, value=f"=D{row_idx}*E{row_idx}")
        ws.cell(row=row_idx, column=7, value=f"=F{row_idx}*0.1")
        ws.cell(row=row_idx, column=8, value=f"=F{row_idx}+G{row_idx}")
        ws.cell(row=row_idx, column=9, value=f"=F{row_idx}/250000")
        ws.cell(row=row_idx, column=10, value=f'=IF(I{row_idx}>=1,"達成","未達")')
        for col_idx in range(1, 11):
            _style_cell(ws.cell(row=row_idx, column=col_idx))

    for row_idx in range(2, 14):
        ws.cell(row=row_idx, column=1).number_format = "yyyy/mm/dd"
        ws.cell(row=row_idx, column=4).number_format = "#,##0"
        for col_idx in (5, 6, 7, 8):
            ws.cell(row=row_idx, column=col_idx).number_format = '"¥"#,##0'
        ws.cell(row=row_idx, column=9).number_format = "0.0%"

    total_row = 15
    ws.cell(row=total_row, column=1, value="合計")
    _style_cell(ws.cell(row=total_row, column=1), bold=True, fill=SUBHEADER_FILL)
    for col_idx in range(4, 9):
        col_letter = get_column_letter(col_idx)
        ws.cell(row=total_row, column=col_idx, value=f"=SUM({col_letter}2:{col_letter}13)")
        _style_cell(ws.cell(row=total_row, column=col_idx), bold=True, fill=SUBHEADER_FILL)
    ws.cell(row=total_row, column=9, value="=AVERAGE(I2:I13)")
    _style_cell(ws.cell(row=total_row, column=9), bold=True, fill=SUBHEADER_FILL)
    ws.cell(row=total_row, column=9).number_format = "0.0%"

    ws.conditional_formatting.add(
        "I2:I13",
        CellIsRule(operator="lessThan", formula=["1"], stopIfTrue=True, fill=WARN_FILL),
    )
    ws.conditional_formatting.add(
        "I2:I13",
        CellIsRule(operator="greaterThanOrEqual", formula=["1"], stopIfTrue=True, fill=OK_FILL),
    )
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:J13"
    _autosize_columns(ws)

    ws2 = wb.create_sheet("KPI")
    headers2 = ["KPI", "目標", "実績", "差分", "達成率"]
    rows2 = [
        ["処理件数", 1500, 1620, "=C2-B2", "=C2/B2"],
        ["障害件数", 5, 3, "=C3-B3", "=C3/B3"],
        ["平均応答(ms)", 700, 640, "=C4-B4", "=C4/B4"],
        ["帳票出力件数", 420, 390, "=C5-B5", "=C5/B5"],
    ]
    _write_table(ws2, start_row=1, start_col=1, headers=headers2, rows=rows2, table_name="KpiTable")
    for row_idx in range(2, 6):
        for col_idx in (2, 3, 4):
            ws2.cell(row=row_idx, column=col_idx).number_format = "#,##0"
        ws2.cell(row=row_idx, column=5).number_format = "0.0%"
    ws2["G2"] = "備考"
    _style_cell(ws2["G2"], bold=True, fill=NOTE_FILL)
    ws2["G3"] = "数式セルは Excel で開いたときに再計算される。"
    _style_cell(ws2["G3"])
    _autosize_columns(ws2)

    path = output_dir / "formulas_and_formats.xlsx"
    wb.save(path)
    print(f"  生成: {path}")
    return path


def generate_comments_and_annotations(output_dir: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "レビュー管理"
    headers = ["項目ID", "項目名", "状態", "ルール", "参照URL", "担当"]
    rows = [
        ["I001", "顧客ID", "要確認", "半角英数字10桁", "https://example.com/spec/customer", "田中"],
        ["I002", "契約日", "OK", "未来日不可", "https://example.com/spec/contract", "鈴木"],
        ["I003", "利用料金", "NG", "0以上、上限なし", "https://example.com/spec/fee", "佐藤"],
        ["I004", "ステータス", "要確認", "コード表と一致", "https://example.com/spec/status", "高橋"],
    ]
    _write_table(ws, start_row=1, start_col=1, headers=headers, rows=rows, table_name="ReviewTable")

    status_fills = {"OK": OK_FILL, "要確認": WARN_FILL, "NG": NG_FILL}
    for row_idx in range(2, 6):
        status = ws.cell(row=row_idx, column=3).value
        ws.cell(row=row_idx, column=3).fill = status_fills[status]
        ws.cell(row=row_idx, column=4).comment = Comment(
            f"{ws.cell(row=row_idx, column=2).value} の補足仕様。外部IFとの整合確認が必要。",
            "review-bot",
        )
        ws.cell(row=row_idx, column=5).hyperlink = ws.cell(row=row_idx, column=5).value
        ws.cell(row=row_idx, column=5).style = "Hyperlink"

    dv = DataValidation(type="list", formula1='"OK,要確認,NG"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C2:C20")
    ws["H2"] = "内部メモ"
    _style_cell(ws["H2"], bold=True, fill=NOTE_FILL)
    ws["H3"] = "H 列は非表示。抽出時に hidden 列をどう扱うか確認用。"
    _style_cell(ws["H3"])
    ws.column_dimensions["H"].hidden = True
    ws.row_dimensions[9].hidden = True
    _autosize_columns(ws)

    ws2 = wb.create_sheet("色分け一覧")
    ws2["A1"] = "状態別の色分け"
    _style_cell(ws2["A1"], bold=True, fill=NOTE_FILL)
    ws2.merge_cells("A1:D1")
    palette_rows = [
        ["正常", "緑", "E2F0D9", "処理可能"],
        ["注意", "黄", "FFF2CC", "確認が必要"],
        ["異常", "赤", "FCE4D6", "修正が必要"],
    ]
    _write_table(
        ws2,
        start_row=3,
        start_col=1,
        headers=["状態", "表示名", "色コード", "説明"],
        rows=palette_rows,
        table_name="PaletteTable",
    )
    ws2["B6"].comment = Comment("色だけでなくラベルも保持したい。", "review-bot")
    _autosize_columns(ws2)

    path = output_dir / "comments_and_annotations.xlsx"
    wb.save(path)
    print(f"  生成: {path}")
    return path


def generate_many_images(output_dir: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "画面一覧"
    ws["A1"] = "画像が複数配置されたシート"
    _style_cell(ws["A1"], bold=True, fill=NOTE_FILL)
    ws.merge_cells("A1:H1")

    image_specs = [
        ("ログイン画面", "B3", (280, 150), (70, 130, 180)),
        ("ダッシュボード", "J3", (280, 150), (46, 139, 87)),
        ("ユーザー一覧", "B18", (280, 150), (220, 53, 69)),
        ("帳票出力", "J18", (280, 150), (147, 112, 219)),
    ]
    for title, anchor, size, color in image_specs:
        ws.add_image(_make_labeled_image(title, size, color), anchor)
        cell = ws[anchor]
        cell.value = title
        _style_cell(cell, bold=True, fill=SUBHEADER_FILL)

    for col in ("B", "J"):
        ws.column_dimensions[col].width = 22
    for row in (3, 18):
        ws.row_dimensions[row].height = 26
    ws["A32"] = "備考"
    _style_cell(ws["A32"], bold=True, fill=NOTE_FILL)
    ws["B32"] = "画像アンカーと近くのセルテキストを一緒に扱えるか確認用。"
    _style_cell(ws["B32"])
    _autosize_columns(ws)

    ws2 = wb.create_sheet("アイコン一覧")
    headers = ["種別", "説明", "アイコン"]
    rows = [
        ["ユーザー", "ユーザー関連操作", ""],
        ["設定", "システム設定", ""],
        ["警告", "エラー・警告表示", ""],
    ]
    _write_table(ws2, start_row=1, start_col=1, headers=headers, rows=rows, table_name="IconTable")
    icon_specs = [
        ("U", "C2", (48, 48), (70, 130, 180)),
        ("S", "C3", (48, 48), (46, 139, 87)),
        ("!", "C4", (48, 48), (220, 53, 69)),
    ]
    for label, anchor, size, color in icon_specs:
        ws2.add_image(_make_labeled_image(label, size, color), anchor)
    _autosize_columns(ws2)

    path = output_dir / "many_images.xlsx"
    wb.save(path)
    print(f"  生成: {path}")
    return path


def generate_mixed_complex(output_dir: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "概要"
    ws["A1"] = "機能仕様書（Excel 総合テスト）"
    _style_cell(ws["A1"], bold=True, fill=NOTE_FILL)
    ws.merge_cells("A1:F1")
    meta_rows = [
        ["システム名", "販売管理システム"],
        ["版数", "2.4"],
        ["作成日", date(2026, 3, 15)],
        ["担当", "業務設計チーム"],
    ]
    _write_table(
        ws,
        start_row=3,
        start_col=1,
        headers=["項目", "内容"],
        rows=meta_rows,
        table_name="MetaTable",
    )
    ws["B6"].number_format = "yyyy/mm/dd"
    ws["D3"] = "レビューコメント"
    _style_cell(ws["D3"], bold=True, fill=SUBHEADER_FILL)
    ws["D4"] = "シートごとに別の観点を持たせる。"
    _style_cell(ws["D4"])
    ws["D4"].comment = Comment("構造保持と Markdown 化の両方を見る。", "architect")
    _autosize_columns(ws)

    ws2 = wb.create_sheet("機能一覧")
    _write_table(
        ws2,
        start_row=1,
        start_col=1,
        headers=["機能ID", "機能名", "種別", "状態", "備考"],
        rows=[
            ["F001", "ログイン", "画面", "稼働中", "認証基盤連携あり"],
            ["F002", "受注一覧", "画面", "稼働中", "検索条件が多い"],
            ["F003", "CSV 出力", "帳票", "テスト中", "列順固定"],
            ["F004", "日次集計", "バッチ", "開発中", "深夜実行"],
        ],
        table_name="FeatureTable",
    )
    dv = DataValidation(type="list", formula1='"画面,帳票,バッチ,API"', allow_blank=False)
    ws2.add_data_validation(dv)
    dv.add("C2:C20")
    ws2.auto_filter.ref = "A1:E5"
    ws2.freeze_panes = "A2"
    _autosize_columns(ws2)

    ws3 = wb.create_sheet("フロー")
    ws3["A1"] = "ワークフロー図イメージ"
    _style_cell(ws3["A1"], bold=True, fill=NOTE_FILL)
    ws3.merge_cells("A1:G1")
    ws3.add_image(_make_labeled_image("申請 -> 承認 -> 完了", (420, 120), (91, 155, 213)), "B3")
    ws3["B11"] = "図の横に補足メモ"
    _style_cell(ws3["B11"], bold=True, fill=SUBHEADER_FILL)
    ws3["C11"] = "画像だけでなく近接テキストも保持したい。"
    _style_cell(ws3["C11"])
    _autosize_columns(ws3)

    ws4 = wb.create_sheet("設定")
    headers4 = ["パラメータ", "現在値", "初期値", "判定"]
    rows4 = [
        ["TOKEN_LIMIT", 8000, 6000, '=IF(B2>=C2,"OK","要確認")'],
        ["MAX_RETRY", 3, 3, '=IF(B3=C3,"OK","要確認")'],
        ["TIMEOUT_SEC", 45, 30, '=IF(B4<=60,"OK","要確認")'],
        ["BATCH_SIZE", 1500, 1000, '=IF(B5>=C5,"OK","要確認")'],
    ]
    _write_table(ws4, start_row=1, start_col=1, headers=headers4, rows=rows4, table_name="SettingsTable")
    for row_idx in range(2, 6):
        status_cell = ws4.cell(row=row_idx, column=4)
        _style_cell(status_cell, fill=OK_FILL)
    ws4["A8"] = "複合ヘッダ"
    _style_cell(ws4["A8"], bold=True, fill=SUBHEADER_FILL)
    ws4.merge_cells("A8:D8")
    ws4["A9"] = "出力"
    ws4["C9"] = "入力"
    _style_cell(ws4["A9"], bold=True, fill=HEADER_FILL)
    _style_cell(ws4["C9"], bold=True, fill=HEADER_FILL)
    ws4.merge_cells("A9:B9")
    ws4.merge_cells("C9:D9")
    _autosize_columns(ws4)

    hidden = wb.create_sheet("マスタ")
    hidden.sheet_state = "hidden"
    hidden["A1"] = "状態"
    hidden["A2"] = "稼働中"
    hidden["A3"] = "テスト中"
    hidden["A4"] = "開発中"

    path = output_dir / "mixed_complex.xlsx"
    wb.save(path)
    print(f"  生成: {path}")
    return path


def generate_large_workbook(output_dir: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "取引明細"
    headers = [
        "取引ID",
        "取引日",
        "顧客コード",
        "顧客名",
        "商品コード",
        "数量",
        "単価",
        "金額",
        "ステータス",
        "担当",
        "拠点",
        "備考",
    ]
    for idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=idx, value=header)
        _style_cell(cell, bold=True, fill=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    statuses = ["受付", "処理中", "完了", "差戻し"]
    offices = ["東京", "大阪", "名古屋", "福岡"]
    reps = ["田中", "鈴木", "佐藤", "高橋", "伊藤", "渡辺"]

    for row_idx in range(2, 4002):
        qty = (row_idx % 9) + 1
        price = 1200 + ((row_idx * 17) % 850)
        ws.cell(row=row_idx, column=1, value=f"TX{row_idx:06d}")
        ws.cell(row=row_idx, column=2, value=date(2025, 1, 1) + timedelta(days=row_idx % 365))
        ws.cell(row=row_idx, column=3, value=f"C{(row_idx % 500):04d}")
        ws.cell(row=row_idx, column=4, value=f"顧客{row_idx % 500:04d}")
        ws.cell(row=row_idx, column=5, value=f"P{(row_idx % 900):05d}")
        ws.cell(row=row_idx, column=6, value=qty)
        ws.cell(row=row_idx, column=7, value=price)
        ws.cell(row=row_idx, column=8, value=qty * price)
        ws.cell(row=row_idx, column=9, value=statuses[row_idx % len(statuses)])
        ws.cell(row=row_idx, column=10, value=reps[row_idx % len(reps)])
        ws.cell(row=row_idx, column=11, value=offices[row_idx % len(offices)])
        ws.cell(
            row=row_idx,
            column=12,
            value=f"案件{row_idx:06d} の処理内容。関連資料ID=DOC-{row_idx % 120:03d}。",
        )

    for row_idx in range(2, 4002):
        ws.cell(row=row_idx, column=2).number_format = "yyyy/mm/dd"
        for col_idx in (6, 7, 8):
            ws.cell(row=row_idx, column=col_idx).number_format = "#,##0"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L{ws.max_row}"
    width_overrides = {
        "A": 14,
        "B": 12,
        "C": 12,
        "D": 14,
        "E": 12,
        "F": 8,
        "G": 10,
        "H": 12,
        "I": 10,
        "J": 10,
        "K": 10,
        "L": 42,
    }
    for col, width in width_overrides.items():
        ws.column_dimensions[col].width = width

    ws2 = wb.create_sheet("月次集計")
    _write_table(
        ws2,
        start_row=1,
        start_col=1,
        headers=["月", "取引件数", "売上金額", "備考"],
        rows=[
            ["2025-01", 330, 5120000, "四半期末調整あり"],
            ["2025-02", 305, 4890000, "通常運用"],
            ["2025-03", 352, 5450000, "キャンペーン実施"],
            ["2025-04", 311, 4980000, "通常運用"],
            ["2025-05", 327, 5210000, "大型案件計上"],
            ["2025-06", 339, 5330000, "通常運用"],
        ],
        table_name="MonthlySummaryTable",
    )
    for row_idx in range(2, 8):
        ws2.cell(row=row_idx, column=2).number_format = "#,##0"
        ws2.cell(row=row_idx, column=3).number_format = "#,##0"
    _autosize_columns(ws2)

    path = output_dir / "large_workbook.xlsx"
    wb.save(path)
    print(f"  生成: {path}")
    return path


def generate_change_history(output_dir: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "変更履歴"
    headers = ["ページ", "種別", "年　月", "記　　事"]
    rows = [
        ["全", "新規", "2024/10", "初版作成"],
        ["3", "追加", "2025/01", "コンポーネント設計追加"],
        ["5-8", "修正", "2025/03", "API 設計見直し"],
        ["10", "削除", "2025/06", "旧バッチ処理削除"],
        ["12-14", "修正", "2025/09", "Excel 出力仕様を更新"],
    ]
    _write_table(ws, start_row=1, start_col=1, headers=headers, rows=rows, table_name="ChangeHistoryTable")
    ws["A9"] = "変更履歴は全角スペースを含む見出しにしてある。"
    _style_cell(ws["A9"], bold=True, fill=NOTE_FILL)
    ws.merge_cells("A9:D9")
    _autosize_columns(ws)

    ws2 = wb.create_sheet("設計概要")
    _write_table(
        ws2,
        start_row=1,
        start_col=1,
        headers=["項目", "内容"],
        rows=[
            ["文書名", "ドキュメント処理パイプライン詳細設計書"],
            ["アーキテクチャ", "マイクロサービス"],
            ["DB", "PostgreSQL"],
            ["帳票", "PDF / Excel"],
        ],
        table_name="SummaryTable",
    )
    ws2["A8"] = "本文と変更履歴が別シートに分かれるケース。"
    _style_cell(ws2["A8"], bold=True, fill=NOTE_FILL)
    ws2.merge_cells("A8:B8")
    _autosize_columns(ws2)

    path = output_dir / "change_history.xlsx"
    wb.save(path)
    print(f"  生成: {path}")
    return path


def main() -> None:
    parser = argparse.ArgumentParser(description="テストデータ .xlsx 生成")
    parser.add_argument(
        "--output-dir",
        "-o",
        type=Path,
        default=Path("input"),
        help="出力先ディレクトリ (default: input/)",
    )
    args = parser.parse_args()

    output_dir: Path = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"テストデータ生成先: {output_dir}/")
    print()

    generate_many_tables(output_dir)
    generate_multiple_tables_sheet(output_dir)
    generate_merged_cells(output_dir)
    generate_formulas_and_formats(output_dir)
    generate_comments_and_annotations(output_dir)
    generate_many_images(output_dir)
    generate_mixed_complex(output_dir)
    generate_large_workbook(output_dir)
    generate_change_history(output_dir)

    print()
    print("完了: 9 ファイル生成しました。")


if __name__ == "__main__":
    main()
