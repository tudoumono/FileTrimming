"""固定回帰セットの入力ディレクトリを作るユーティリティ。

使い方:
  python tools/prepare_regression_subset.py
  python tools/prepare_regression_subset.py --set excel_llm_core --output-dir tmp/my_subset
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path


REGRESSION_SETS: dict[str, list[str]] = {
    "excel_llm_core": [
        "approval_request.xlsx",
        "excel_form_grid.xlsx",
        "comments_and_annotations.xlsx",
        "timesheet_calendar.xlsx",
    ],
}


def prepare_subset(input_dir: Path, output_dir: Path, file_names: list[str]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    copied: list[str] = []
    missing: list[str] = []

    for file_name in file_names:
        src = input_dir / file_name
        dst = output_dir / file_name
        if not src.exists():
            missing.append(file_name)
            continue
        shutil.copy2(src, dst)
        copied.append(file_name)

    print(f"input_dir:  {input_dir}")
    print(f"output_dir: {output_dir}")
    print(f"copied:     {len(copied)}")
    for name in copied:
        print(f"  - {name}")

    if missing:
        print(f"missing:    {len(missing)}")
        for name in missing:
            print(f"  - {name}")
        raise SystemExit(1)


def main() -> None:
    parser = argparse.ArgumentParser(description="固定回帰セットの入力ディレクトリを作成する")
    parser.add_argument(
        "--set",
        default="excel_llm_core",
        choices=sorted(REGRESSION_SETS.keys()),
        help="準備する回帰セット名",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path("input/excel"),
        help="元ファイルのある入力ディレクトリ",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("tmp/excel_llm_core_subset"),
        help="コピー先ディレクトリ",
    )
    args = parser.parse_args()

    prepare_subset(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        file_names=REGRESSION_SETS[args.set],
    )


if __name__ == "__main__":
    main()
