# LX_O_REP AI Context

## Purpose

This file documents the current AI-facing architecture and integration points for the `LX_O_REP` module of the LX Programs workspace.
It is intended to provide developers and AI copilots with a compact reference for how the report generator is wired together.

## Module Responsibilities

- `LX_O_REP_GUI.py` — Primary user interface for report generation.
- `LX_Orchestrator.py` — Command router that reads `Record.csv`, groups output rows, and executes frame modules.
- `LX_CommCentre.py` — Shared runtime state manager for PCB ID, tester/reviewer credentials, and placeholder path resolution.
- `LX_BaseFrame.py` — Base frame abstraction with lock state semantics used by report modules.
- `frames/` — Concrete frame implementations for report content generation.

## Integration Points

### `LX_O_REP_GUI.py`

- Auto-detects `Record.csv` and `Record_Contract.csv`.
- Provides actions:
  - `load_csv()` — Load and display the source matrix.
  - `save_csv()` — Persist matrix edits back to `Record.csv`.
  - `validate_matrix()` — Check matrix integrity and alias contract validity.
  - `run_generation()` — Update `LX_CommCentre` state and start `Orchestrator.run_sequence()`.

- `run_generation()` passes the current `record_path` into the orchestrator.
- It uses `threading.Thread(..., daemon=True)` so the GUI remains responsive during generation.

### `LX_Orchestrator.py`

- Reads all rows from `Record.csv`.
- Skips rows with `LX == X`.
- Resolves each row's `D_Path` via `LX_CommCentre.resolve_path()`.
- Groups rows by final destination workbook path.
- Creates one `xlsxwriter.Workbook` per destination file.
- Dispatches row execution to frame handlers according to `Type`.

### `LX_CommCentre.py`

- Tracks `PCB_ID`, `tester_name`, `reviewed_by`, and P0–P9 path slots.
- `resolve_path(path_str)` replaces placeholders with configured values.
- This is a singleton instance available as `LX_CommCentre`.

## Frame Execution Flow

- Frame registry map in `Orchestrator.__init__()` defines type-to-frame class associations.
- Execution order is hard-coded as: `IMG`, `MIR`, `HDC`, `VAL`, `TXT`, `STR`.
- Each row is passed to a frame via `_execute_row()`.
- Frames should update workbook/sheet objects by using `_wb` and `_ws_dict` set on the row.

## Expected Data Contract

### `Record.csv`

Rows should include at minimum:

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

Must define a valid `Alias` column.
The GUI validation uses this contract to verify `Type` values.

## Current Implementation Notes

- Output generation is currently handled with `xlsxwriter`, not openpyxl.
- `openpyxl` is present for compatibility or future XLSX read/write support.
- The current frame set includes:
  - `F_IMGR` — image-related output
  - `F_MIR` — mirror / reference data
  - `F_HDCREP` — header/repository block generation
  - `F_REPRF` — validation frame and text frame support
  - `F_REPT` — string frame output

## Known Good Commands

```powershell
python LX_O_REP/LX_O_REP_GUI.py
python LXDS/LXDS.py
```

## Diagnostic Checks

- `python -m py_compile LX_O_REP/LX_O_REP_GUI.py`
- `python -m py_compile LXDS/LXDS.py`

Both checks currently pass successfully.
