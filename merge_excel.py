#!/usr/bin/env python3
"""Merge Excel workbooks, clean text fields, and handle duplicate rows.

Examples:
    python merge_excel.py ./input --output merged.xlsx
    python merge_excel.py ./input --sheet Sales --keys Email --keep last
    python merge_excel.py ./input --sheet 1 --keys "Customer ID" Date --keep none
"""

from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import pandas as pd

EXCEL_PATTERNS = ("*.xlsx", "*.xls", "*.xlsm")
SOURCE_COLUMN = "_source_file"
SHEET_NAME_MERGED = "merged"
SHEET_NAME_DUPLICATES = "duplicates_removed"
KeepPolicy = Literal["first", "last", "none"]


@dataclass(frozen=True)
class MergeSummary:
    files_found: int
    files_read: int
    files_skipped: int
    rows_read: int
    blank_rows_removed: int
    duplicate_rows_removed: int
    final_rows: int


def configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )


def parse_sheet(value: str) -> str | int:
    try:
        return int(value)
    except ValueError:
        return value


def normalize_column_name(column: object) -> str:
    return str(column).strip()


def find_excel_files(folder: Path, excluded_file: Path | None = None) -> list[Path]:
    excluded = excluded_file.resolve() if excluded_file and excluded_file.exists() else None
    files = {
        path.resolve()
        for pattern in EXCEL_PATTERNS
        for path in folder.glob(pattern)
        if path.is_file()
        and not path.name.startswith("~$")
        and path.resolve() != excluded
    }
    return sorted(files, key=lambda path: path.name.lower())


def ensure_unique_columns(columns: Sequence[object], filename: str) -> list[str]:
    normalized = [normalize_column_name(column) for column in columns]
    duplicates = sorted({name for name in normalized if normalized.count(name) > 1})
    if duplicates:
        raise ValueError(
            f"'{filename}' contains duplicate column names after trimming: {duplicates}"
        )
    if SOURCE_COLUMN in normalized:
        raise ValueError(
            f"'{filename}' already contains the reserved column '{SOURCE_COLUMN}'."
        )
    return normalized


def read_excel_files(
    files: Sequence[Path], sheet: str | int
) -> tuple[pd.DataFrame, int]:
    frames: list[pd.DataFrame] = []
    skipped = 0

    for file_path in files:
        try:
            frame = pd.read_excel(file_path, sheet_name=sheet)
            if not isinstance(frame, pd.DataFrame):
                raise ValueError("The selected sheet did not produce a table.")
            frame.columns = ensure_unique_columns(frame.columns, file_path.name)
            frame[SOURCE_COLUMN] = file_path.name
        except (OSError, ValueError, ImportError, KeyError) as error:
            skipped += 1
            logging.warning("Skipped %s: %s", file_path.name, error)
            continue
        except Exception as error: 
            skipped += 1
            logging.warning("Skipped %s: %s", file_path.name, error)
            continue

        logging.info(
            "Read %s: %s rows, %s data columns",
            file_path.name,
            len(frame),
            len(frame.columns) - 1,
        )
        frames.append(frame)

    if not frames:
        raise RuntimeError("No Excel files could be read successfully.")

    return pd.concat(frames, ignore_index=True, sort=False), skipped


def clean_dataframe(dataframe: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    cleaned = dataframe.copy()
    data_columns = [column for column in cleaned.columns if column != SOURCE_COLUMN]

    for column in data_columns:
        if pd.api.types.is_object_dtype(cleaned[column]) or pd.api.types.is_string_dtype(
            cleaned[column]
        ):
            cleaned[column] = cleaned[column].map(
                lambda value: value.strip() if isinstance(value, str) else value
            )
            cleaned[column] = cleaned[column].replace("", pd.NA)

    before = len(cleaned)
    cleaned = cleaned.dropna(how="all", subset=data_columns).reset_index(drop=True)
    return cleaned, before - len(cleaned)


def validate_keys(dataframe: pd.DataFrame, keys: Sequence[str] | None) -> list[str]:
    available = [column for column in dataframe.columns if column != SOURCE_COLUMN]
    if not available:
        raise ValueError("No data columns are available for duplicate detection.")

    if not keys:
        return available

    normalized_keys = [key.strip() for key in keys]
    empty_keys = [key for key in normalized_keys if not key]
    missing = [key for key in normalized_keys if key not in available]
    repeated = sorted({key for key in normalized_keys if normalized_keys.count(key) > 1})

    if empty_keys or missing or repeated:
        details: list[str] = []
        if empty_keys:
            details.append("empty key names are not allowed")
        if missing:
            details.append(f"missing columns: {missing}")
        if repeated:
            details.append(f"repeated keys: {repeated}")
        raise ValueError(
            f"Invalid --keys argument ({'; '.join(details)}). "
            f"Available columns: {available}"
        )

    return normalized_keys


def remove_duplicates(
    dataframe: pd.DataFrame, keys: Sequence[str] | None, keep: KeepPolicy
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    subset = validate_keys(dataframe, keys)
    pandas_keep: Literal["first", "last", False] = False if keep == "none" else keep
    duplicate_mask = dataframe.duplicated(subset=subset, keep=pandas_keep)

    removed = dataframe.loc[duplicate_mask].copy().reset_index(drop=True)
    retained = dataframe.loc[~duplicate_mask].copy().reset_index(drop=True)
    return retained, removed, subset


def validate_output_path(output: Path) -> Path:
    output = output.expanduser().resolve()
    if output.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Output must use the .xlsx or .xlsm extension.")
    output.parent.mkdir(parents=True, exist_ok=True)
    return output


def autosize_worksheet(worksheet: object) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_cells in worksheet.columns:
        column_letter = column_cells[0].column_letter
        max_length = max(
            (len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells),
            default=0,
        )
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 45)


def write_output(
    output: Path,
    merged: pd.DataFrame,
    duplicates: pd.DataFrame,
    summary: MergeSummary,
) -> None:
    summary_frame = pd.DataFrame(
        {
            "Metric": [
                "Files found",
                "Files read",
                "Files skipped",
                "Rows read",
                "Blank rows removed",
                "Duplicate rows removed",
                "Final rows",
            ],
            "Value": [
                summary.files_found,
                summary.files_read,
                summary.files_skipped,
                summary.rows_read,
                summary.blank_rows_removed,
                summary.duplicate_rows_removed,
                summary.final_rows,
            ],
        }
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name=SHEET_NAME_MERGED, index=False)
        duplicates.to_excel(writer, sheet_name=SHEET_NAME_DUPLICATES, index=False)
        summary_frame.to_excel(writer, sheet_name="summary", index=False)

        for worksheet in writer.book.worksheets:
            autosize_worksheet(worksheet)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Merge Excel files, clean text values, and remove duplicate rows.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "folder",
        type=Path,
        help="Folder containing the Excel files to merge (non-recursive).",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("merged.xlsx"),
        help="Output workbook path.",
    )
    parser.add_argument(
        "--sheet",
        default="0",
        help="Worksheet name or zero-based worksheet index to read.",
    )
    parser.add_argument(
        "--keys",
        nargs="+",
        metavar="COLUMN",
        help="Columns used for duplicate detection. Defaults to all data columns.",
    )
    parser.add_argument(
        "--keep",
        choices=("first", "last", "none"),
        default="first",
        help="Duplicate policy: retain first, retain last, or retain no duplicate rows.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow replacing an existing output workbook.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Show debug-level diagnostic messages.",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    configure_logging(args.verbose)

    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        logging.error("Input folder does not exist or is not a directory: %s", folder)
        return 2

    try:
        output = validate_output_path(args.output)
    except ValueError as error:
        logging.error("%s", error)
        return 2

    if output.exists() and not args.overwrite:
        logging.error(
            "Output file already exists: %s. Use --overwrite to replace it.", output
        )
        return 2

    files = find_excel_files(folder, excluded_file=output)
    if not files:
        logging.error("No supported Excel files were found in: %s", folder)
        return 1

    logging.info("Found %s Excel file(s) in %s", len(files), folder)

    try:
        raw, skipped_files = read_excel_files(files, parse_sheet(args.sheet))
        cleaned, blank_rows_removed = clean_dataframe(raw)
        merged, duplicates, keys = remove_duplicates(cleaned, args.keys, args.keep)

        summary = MergeSummary(
            files_found=len(files),
            files_read=len(files) - skipped_files,
            files_skipped=skipped_files,
            rows_read=len(raw),
            blank_rows_removed=blank_rows_removed,
            duplicate_rows_removed=len(duplicates),
            final_rows=len(merged),
        )
        write_output(output, merged, duplicates, summary)
    except (RuntimeError, ValueError, OSError, ImportError) as error:
        logging.error("Merge failed: %s", error)
        return 1
    except Exception as error:
        logging.exception("Unexpected error while merging Excel files: %s", error)
        return 1

    key_label = ", ".join(keys)
    logging.info(
        "Completed successfully. Final rows: %s | Removed duplicates: %s | Keys: %s",
        summary.final_rows,
        summary.duplicate_rows_removed,
        key_label,
    )
    logging.info("Output workbook created: %s", output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
