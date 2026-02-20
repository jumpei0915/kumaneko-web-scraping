#!/usr/bin/env python3
"""Read reservation Excel files and print a quick preview."""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


DEFAULT_PATH = Path("/home/jumpei_private/workspace/熊猫カンパニー/reservation")
EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}


def resolve_excel_file(path: Path) -> Path:
    if path.is_file():
        if path.suffix.lower() not in EXCEL_SUFFIXES:
            raise ValueError(f"Not an Excel file: {path}")
        return path

    if not path.is_dir():
        raise FileNotFoundError(f"Path does not exist: {path}")

    candidates = [p for p in path.iterdir() if p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES]
    if not candidates:
        raise FileNotFoundError(f"No Excel files found in directory: {path}")

    return max(candidates, key=lambda p: p.stat().st_mtime)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Read a reservation Excel file and print a preview.",
    )
    parser.add_argument(
        "--path",
        type=Path,
        default=DEFAULT_PATH,
        help="Excel file path or directory containing Excel files.",
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="Sheet name or index (default: 0 for first sheet).",
    )
    parser.add_argument(
        "--rows",
        type=int,
        default=20,
        help="Number of rows to print from the top (default: 20).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Optional output file (.csv or .json).",
    )
    args = parser.parse_args()

    excel_file = resolve_excel_file(args.path.expanduser())

    sheet: str | int
    if isinstance(args.sheet, str) and args.sheet.isdigit():
        sheet = int(args.sheet)
    else:
        sheet = args.sheet

    df = pd.read_excel(excel_file, sheet_name=sheet)

    print(f"File: {excel_file}")
    print(f"Sheet: {args.sheet}")
    print(f"Rows: {len(df)}, Columns: {len(df.columns)}")
    print("Columns:", ", ".join(map(str, df.columns)))
    print()
    print(df.head(args.rows).to_string(index=False))

    if args.output:
        output_path = args.output.expanduser()
        if output_path.suffix.lower() == ".csv":
            df.to_csv(output_path, index=False)
        elif output_path.suffix.lower() == ".json":
            df.to_json(output_path, orient="records", force_ascii=False, indent=2)
        else:
            raise ValueError("Output file must end with .csv or .json")
        print()
        print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
