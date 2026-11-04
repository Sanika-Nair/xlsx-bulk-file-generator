# Bulk File Generator

Create many Excel (.xlsx) files from one template, with unique TxnIDs and fixed dates.
Every file is an exact copy of the template (all sheets, formatting, widths, number formats,
formulas, merged cells, frozen panes). Only the columns and rows you choose are changed.
Runs as a single Windows `.exe`: double-click it and the app opens in its own window
(no address bar, its own taskbar button and icon). Closing the window stops the app, after any
batch that is still running has finished.
Nothing leaves your computer.

## Get the .exe (no Python needed)

GitHub builds the app automatically on every change.

- **Latest build:** open the **Actions** tab → click the newest green **Build Windows .exe** run →
  under **Artifacts**, download **BulkFileGenerator-exe** (a zip containing `BulkFileGenerator.exe`).
- **Releases:** if a version has been published, download `BulkFileGenerator.exe` from the
  **Releases** section on the right of the repo page. (To publish one, create a tag like `v1.0.0`.)

## File size

Generated files are exactly the same size as the template, byte for byte, and updated files
stay exactly the size they were. The app packs the changed sheet so it fits the original size
and fills any small gap with standard zip padding that Excel ignores. The only exception is a
template whose TxnID or date cells are empty: filling them adds real data, so those files are
a little larger.

## Two tasks

- **Create new bulk files**: upload one template and create many files from it.
- **Update file names / column data**: point the app at a folder of existing files and change
  their file names, the data in chosen columns, or both. Choose to replace the files in that
  folder or to save updated copies elsewhere. Renaming leaves file contents byte-for-byte
  unchanged; data updates keep every file exactly its original size.

## Use it

1. Put `BulkFileGenerator.exe` in a folder you can write to (e.g. Desktop or Documents).
2. Double-click it. A small console window opens and the app opens in your browser.
3. Follow the steps: **Project → Task → Extension → Upload template → Batch settings →
   Columns to update → Generate**.
4. Download the ZIP or open the output folder. Click **Quit app** (or close the console) when done.

If Windows shows "Windows protected your PC", click **More info → Run anyway** (the app isn't code-signed).

## What people using the app see

Only `BulkFileGenerator.exe`, plus the `Tracker` and `Output` folders it creates next to itself.
The code, the page and the project list are built into the .exe, so no `.py`, `.json` or
`.html` files are visible, not even in Windows' temporary folder while the app runs.

**Changing the projects:** edit `config.json` in the build folder (the one with
`build_exe.bat`), run `build_exe.bat` again, and share the new .exe.

## Files created next to the .exe

| Path | What it is |
|---|---|
| `Tracker\<Project>\filename_Numbercounter.txt` | Next file number (counts down from 9999999). |
| `Tracker\<Project>\TxnID_columnNumberCounter.txt` | Next TxnID number (counts down from 9999999). |
| `Output\<Project>\` | Default folder for generated files. |

To continue from existing counters, copy your two tracker `.txt` files into `Tracker\<Project>\`.

Number rules: every file number and TxnID is used once only. The counters are moved past the
whole batch before any file is written, so even a failed batch never reuses numbers. A batch
that would need a number below 1 is refused, so nothing negative is ever created.

## For developers

- Run from source: `pip install -r requirements.txt` then `python BulkFileGenerator.py`
- Build the .exe locally on Windows: double-click `build_exe.bat` (output in `dist\`). It runs
  `make_bundle.py` first, which packs `config.json` and `ui.html` into the app.

| File | Role |
|---|---|
| `BulkFileGenerator.py` | Starts the local web server and opens the browser |
| `bulk_engine.py` | Template scan, tracker, validation and parallel file generation |
| `ui.html` | The app's page |
| `.github/workflows/build-exe.yml` | Automatic Windows .exe build |
