# LX_O_REP — LX Report Generator

## Overview

`LX_O_REP` is the report generation module for the LX Programs workspace. It reads `Record.csv` as the source-of-truth matrix, validates the matrix, resolves dynamic path placeholders, and generates Excel reports using frame modules from `LX_O_REP/frames/`.

## Purpose

- Provide a Tkinter-based GUI for managing report generation.
- Load and edit the `Record.csv` matrix.
- Validate matrix integrity against `Record_Contract.csv`.
- Run `Orchestrator` to generate output workbooks via `xlsxwriter`.
- Resolve path placeholders through `LX_CommCentre`.

## Important Files

- `LX_O_REP_GUI.py` — main GUI entrypoint.
- `LX_Orchestrator.py` — core report processing and frame dispatch.
- `LX_CommCentre.py` — placeholder resolver and shared runtime state.
- `LX_BaseFrame.py` — base frame contract and lock-state interface.
- `frames/` — frame modules used by the orchestrator.
- `Record.csv` — source-of-truth row definitions for report generation.
- `Record_Contract.csv` — type alias contract used by validation.

## Requirements

- Python 3.11
- `tkinter` (built-in on Windows)
- `xlsxwriter`
- `openpyxl`

Install required packages:

```powershell
pip install xlsxwriter openpyxl
```

## Running the GUI

From the workspace root:

```powershell
python LX_O_REP/LX_O_REP_GUI.py
```

The GUI automatically searches for `Record.csv` in:

1. Current working directory
2. `LX_O_REP/` script directory
3. The workspace root parent folder

If the file cannot be found, a warning banner appears with a browse option.

## How It Works

1. `LX_O_REP_GUI.py` loads `Record.csv` into a tree view.
2. The user can validate the matrix or start report generation.
3. `validate_matrix()` checks required fields, alias contract validity, mirror pointers, and cell coordinates.
4. `run_generation()` writes any CSV edits, updates `LX_CommCentre` state, and launches `Orchestrator.run_sequence()` in a background thread.

## CSV Expectations

### `Record.csv`

The report matrix expects columns including, but not limited to:

- `Index`
- `LX`
- `Type`
- `M`
- `D_File_Name`
- `D_Sheet`
- `D_Path`
- `S_Start_Cell_Range`
- `S_End_Cell_Range`
- `D_Start_Cell_Range`
- `D_End_Cell_Range`

### `Record_Contract.csv`

This contract file should define valid `Type` aliases using an `Alias` column.
When provided, `validate_matrix()` verifies that each `Type` value exists in the contract.

## Path Resolution

`LX_CommCentre.resolve_path()` supports dynamic placeholders:

- `PCB_ID` — replaced with the current PCB identifier.
- `P0` through `P9` — replaced with configured path values from the UI.

Examples:

- `P0/Results/PCB_ID` → `C:\some\path\Results\PROJECT_VER_BRDNAM`
- `PCB_ID/Output.xlsx`

## Frame Execution

`LX_Orchestrator.py`:

- Loads `Record.csv` rows.
- Groups rows by destination workbook path.
- Opens an `xlsxwriter.Workbook` for each destination file.
- Executes frames in the order: `IMG`, `MIR`, `HDC`, `VAL`, `TXT`, `STR`.

Frame modules are registered by `Type`:

- `IMG` → `frames/F_IMGR.py`
- `MIR` → `frames/F_MIR.py`
- `HDC` → `frames/F_HDCREP.py`
- `VAL` → `frames/F_REPRF.py`
- `TXT` → `frames/F_REPRF.py`
- `STR` → `frames/F_REPT.py`

Each frame receives a row dictionary and should perform workbook/sheet updates using the provided `_wb` and `_ws_dict` keys.

## Extending Frames

To add a new frame type:

1. Create a new module in `LX_O_REP/frames/`.
2. Add a class that implements `execute(row: dict) -> bool` and uses `get_lock_status()`.
3. Register the new `Type` alias in `Orchestrator.registry`.

Frame modules should inherit from `LX_BaseFrame` and respect `LXLockState` status semantics.

## Notes

- The GUI disables the `START REPORT GENERATION` and `VALIDATE MATRIX` buttons while generation is active.
- The application uses thread-safe logging into the UI log area.
- `openpyxl` is imported only for future mirroring and workbook compatibility, while report output is produced with `xlsxwriter`.

## Validation Status

✔ `python -m py_compile LX_O_REP/LX_O_REP_GUI.py` passes with no syntax errors.
✔ `python -m py_compile LXDS/LXDS.py` also passes.
