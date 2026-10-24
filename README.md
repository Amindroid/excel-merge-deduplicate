# Excel Merge & Deduplication Tool

A Python command-line tool for merging Excel files, cleaning data,
and removing duplicate records with source-file tracking.

## Overview

This tool combines spreadsheet data from multiple Excel workbooks
into a single output workbook.

It also saves removed duplicate rows separately and generates
a processing summary for reviewing the results.

## Features

- Read `.xlsx`, `.xls`, and `.xlsm` input files.
- Select a worksheet by name or zero-based index.
- Trim leading and trailing whitespace from text values.
- Remove rows that are empty across all data columns.
- Detect duplicates using all data columns or selected key columns.
- Keep the first occurrence, keep the last occurrence,
  or remove every occurrence of duplicated records.
- Preserve the source workbook name in a `_source_file` column.
- Export removed duplicates to a separate worksheet.
- Generate processing statistics.
- Protect existing output files unless `--overwrite` is provided.

## Technologies

- Python
- pandas
- openpyxl
- xlrd
- argparse
- pathlib
- logging

## Installation

Python 3.10 or newer is required.

Download or clone this repository, then open a terminal
inside the project folder.

Create a virtual environment on Windows:

```powershell
py -m venv .venv
```

Install the dependencies:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Usage

Create a folder named `input_files` and place your Excel files inside it.

Merge the first worksheet from each workbook:

```powershell
.\.venv\Scripts\python.exe merge_excel.py ./input_files --output ./output/merged.xlsx
```

Read a worksheet by name:

```powershell
.\.venv\Scripts\python.exe merge_excel.py ./input_files --sheet Sales --output ./output/merged.xlsx
```

Remove duplicates using a column named `Email`:

```powershell
.\.venv\Scripts\python.exe merge_excel.py ./input_files --keys Email --output ./output/merged.xlsx
```

Keep the last occurrence of each email address:

```powershell
.\.venv\Scripts\python.exe merge_excel.py ./input_files --keys Email --keep last --output ./output/merged.xlsx
```

Remove all occurrences of duplicated email addresses:

```powershell
.\.venv\Scripts\python.exe merge_excel.py ./input_files --keys Email --keep none --output ./output/merged.xlsx
```

The column names and worksheet names in these examples must exist
in your input workbooks.

To replace an existing output workbook, add `--overwrite`.

View all command-line options:

```powershell
.\.venv\Scripts\python.exe merge_excel.py --help
```

## Output Workbook

| Worksheet | Contents |
| --- | --- |
| `merged` | Retained records and their source workbook names |
| `duplicates_removed` | Records removed by the selected duplicate policy |
| `summary` | File counts, row counts, and removal statistics |

The `_source_file` column is excluded from duplicate detection.

## Project Structure

```text
excel-merge-deduplicate/
├── merge_excel.py
├── requirements.txt
├── README.md
└── .gitignore
```

Input and output folders are created locally and are not included
in this structure.

## Current Limitations

- Only files directly inside the input folder are scanned.
- One selected worksheet is read from each workbook.
- The tool does not preserve original workbook formatting or macros.
- Identifiers may be inferred as numbers, removing leading zeros.
- Missing values in duplicate-key columns require careful review.
- `--keep last` follows processing order, not automatic date sorting.
- Text beginning with `=` may be exported as an Excel formula.
- Large files and all supported Excel formats have not been fully tested.

Review generated output before using it for important workflows.

## Verification

The current implementation was manually exercised with three sample
`.xlsx` workbooks using the `first`, `last`, and `none` duplicate policies.

These checks were performed with Python 3.12.8, pandas 2.2.3,
and openpyxl 3.1.5. Compatibility with other versions and file formats
has not been fully verified.

An automated test suite has not yet been added.

## Planned Improvements

- Preserve leading zeros in identifier columns.
- Define explicit handling for missing duplicate keys.
- Prevent unintended formula interpretation in exported text.
- Improve output-path error handling.
- Add safer output-file replacement.
- Add automated tests.

## Data Privacy

Use fictional data for public examples.

Do not commit real employee, customer, payroll, or other sensitive
spreadsheet data to a public repository.
