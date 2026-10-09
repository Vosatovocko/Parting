# PARTING

PARTING is a small Tkinter-based data-exploration tool for searching and filtering a linguistic dataset stored in an Excel workbook. The project was created for the Linguistics Department of the University of Stuttgart and provides a GUI for selecting linguistic features and viewing matching records from the database.

## Overview

This repository contains:

- `UI_PARTING_Windows.py` — Windows version of the GUI
- `UI_PARTING_Mac.py` — macOS version of the GUI
- `db_PARTING.xlsx` — the source dataset used by both scripts
- `README.md` — project documentation

Both scripts read the same Excel workbook and build a filter interface based on grouped linguistic fields such as:

- `SPEC_FEAT`
- `PRO_Nominal`
- `REL_Type`
- `OUT_SYN`
- `OUT_SEM`
- `INN_Syn`
- `ADJ`
- `INN_Lex`
- `INN_ctxt`

The interface lets the user choose values from dropdown menus, apply filtering logic, and display matching rows. The scripts also include a set of business rules that disable some selectors when earlier choices make them irrelevant (for example, related fields are disabled if a given condition is met).

## Data model

The project expects the Excel workbook to contain the columns used by the GUI, including:

- `INDEX`
- `Meaning`
- `Context`
- `NP`
- plus the feature columns referenced in the `label_groups` dictionaries in the scripts.

The scripts load the workbook with `pandas.read_excel()`, enumerate unique values from each column, and present those values as dropdown options. Rows with missing data are additionally exposed as the special option `None` in the UI.

## Prerequisites

Before running the project, install the required Python dependencies:

- Python 3.8+
- `pandas`
- `openpyxl` (used to read `.xlsx` files)
- `tkinter` (included with standard Python installations on Windows and macOS)

Install them with:

```bash
pip install pandas openpyxl
```

## Setup

1. Clone this repository.
2. Make sure `db_PARTING.xlsx` is present in the repository root.
3. Confirm the Excel file name matches the script's file path expectation.

Important: both scripts currently contain hard-coded absolute paths to the Excel file, for example:

- Windows/macOS examples in the code point to a local developer path such as `/Users/erickvillanueva/Documents/GitHub/Parting/db_PARTING.xlsx`

If you are running the project from another machine, update the `file_path` variable in the script before executing it.

Example:

```python
file_path = 'db_PARTING.xlsx'
```

or an absolute path appropriate to your local environment:

```python
file_path = '/path/to/Parting/db_PARTING.xlsx'
```

## Usage

### Windows

Run:

```bash
python UI_PARTING_Windows.py
```

### macOS

Run:

```bash
python3 UI_PARTING_Mac.py
```

## How the GUI works

After launching the app:

1. Select one or more values from the feature dropdowns.
2. Use `ALL` to ignore a filter.
3. Use `None` to search for missing values in a given field.
4. Press `Submit` to filter the dataset.
5. Press `Reset` to clear all selections and return to the default state.

The app applies a set of dependency rules between fields. For example:

- if a value is selected in `PRO_Nominal`, related groups to the right are disabled
- if `DET` is set to `None`, related fields are disabled
- if `Case` is not `PP`, certain preposition-related selections are disabled
- if `Adj` is `None`, adjective-related filters are disabled

These rules help prevent logically inconsistent combinations.

## Output

When the user submits a query:

- matching rows are displayed in the GUI
- the total number of matches is shown
- a text file is written with the selected filters and matching records

The Windows version writes a file named:

- `parting_search.txt`

The macOS version writes:

- `search.txt`

The result output is formatted around these columns:

- `INDEX`
- `NP`
- `Context`
- `Meaning`

## Notes

- The project is a local GUI utility rather than a web application.
- The scripts rely on the workbook structure and column names defined in the project.
- If the dataset schema changes, both the GUI and the filtering logic may need to be updated accordingly.

## Typical workflow

A common workflow is:

1. Open the script matching your OS.
2. Verify/update the database path.
3. Launch the GUI.
4. Select linguistic categories.
5. Filter the dataset and inspect matches.
6. Save or reuse the exported output text file.

This repository provides a lightweight search tool for analyzing linguistic pattern combinations stored in the Excel workbook.
