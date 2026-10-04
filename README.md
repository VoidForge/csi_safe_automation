# CSI SAFE automation tools

A set of VBA modules and Python scripts for automating CSI SAFE analysis from Excel.
They read and write SAFE database tables, run the analysis, push a pile layout into the
model, collect reactions and pile utilization, and interpolate scattered points onto a
surface.

Every module attaches to a SAFE instance that already has the model loaded. None of
them start SAFE, open a model, or close your session. The workbooks, models and
reference material they run against stay out of this repository.

## What's here

| File | What it is |
| ---- | ---------- |
| `SAFE_Library.bas` | The library module (import this one): read one SAFE table at a time into a 2-D array, print arrays to worksheets, write edits back. |
| `SAFE_Use.bas` | Companion, site-specific subs that call into the library. |
| `Interpolation2D.bas` | Standalone module: linear interpolation on scattered 2-D points. Calls nothing outside itself. |
| `auto_analysis.py` | Python (xlwings) driver that runs the workbook's own macros in sequence and records the results. |
| `lloyd_cvt.py` | Python: centroidal Voronoi tessellation of a convex polygon by Lloyd's algorithm. |
| `rename_repo.bat` | Repairs the `origin` remote URL after the GitHub repository was renamed. |

## How to use

1. Start SAFE and open the model you want to work on. Leave SAFE running.
2. In the Excel VBA IDE (`Alt+F11`), open `Tools → References…` and tick `SAFEv1`,
   or browse to `SAFEv1.tlb` in the SAFE installation folder. The reference comes
   from this IDE setting, and no file path is hardcoded in the modules.
3. Import the modules into the workbook with `File → Import File…`:
   `SAFE_Library.bas`, plus `SAFE_Use.bas` for the site subs and
   `Interpolation2D.bas` if you need interpolation.
4. Run a demo (`DemoExport`, `InterpolationDemo`), one of the site subs in
   `SAFE_Use.bas`, or the Python driver below.

`SAFE_Use.bas` and `Interpolation2D.bas` keep message boxes off by default and
write every message to their own log instead (`ShowLog`, `InterpShowLog`, or the
Immediate window with `Ctrl+G`). Set the module flag to `True` to see the boxes
again.

## `SAFE_Library.bas`

The read and the write are separate halves, so a caller can process the data in VBA in
between: `ExportSAFETables` returns a table as a 2-D array, and `PrintTable` writes a
2-D array to a sheet.

```vba
Dim hdrs() As String, data As Variant, warn As String, failed As Boolean

data = ExportSAFETables("Element Forces - Area Shells", hdrs, , , warn, failed)
If failed Then
    MarkTableFailed "Forces", "B2", "Element Forces - Area Shells", warn
Else
    PrintTable data, "Forces", "B2", "Element Forces - Area Shells", hdrs
End If
```

- `ExportSAFETables(TableKey, Headers, [LoadCases], [LoadCombos], [Warning],
  [Failed])` reads one table into a 1-based 2-D `Variant` array; `Headers()` comes
  back with SAFE's column keys in SAFE's order, and the data array never carries a
  header row. Nothing is written to a sheet here. `LoadCases` and `LoadCombos` are
  SAFE's display filters, and they only affect result tables: load cases and load
  combinations are two separate SAFE lists, so a case name belongs in `LoadCases` and
  a combination name in `LoadCombos`; empty means all of them. The selection in force
  is saved first and put back on the normal exit and on the error path. A name SAFE
  does not accept fails that read, and the log names the rejected entry. Branch on
  `Failed`, never on `Empty`: an empty table is a normal answer, a failed read is not.
- `PrintTable(Data, SheetName, StartCell, [Title], [Headers], [WriteTitle],
  [WriteHeader], [StackHorizontally], [NextRow], [NextCol])` writes one array, normally
  the one `ExportSAFETables` returned, though any 1-based 2-D array works, a
  `Range.Value` array included. The destination is the worksheet `SheetName`, created
  if it does not exist, at `StartCell`. `Title` and `Headers` are the optional title
  line and column-header row; `WriteTitle:=False` with `WriteHeader:=False` writes
  neither, so the data lands flush on `StartCell`. `NextRow` and `NextCol` come back as
  the cursor for the next block, which is how a loop stacks several tables or places
  them side by side. It returns the number of data rows written, `0` for an empty block
  (title and/or header rows only, or a `no data returned` marker) and `-1` when the
  block could not be written. A block too big for one worksheet continues on
  `<SheetName>_2`, `_3`, … with the title and header row repeated; a block wider than
  the sheet, or a `StartCell` with no room left, is refused with a bold red marker and
  a failure count, so nothing is truncated silently.
- `ListSAFETables(SheetName, StartCell)` dumps every table SAFE reports
  (`GetAllTables`: key, name, import type), a superset of the tables merely available
  for display, and the log prints both counts so the difference is visible. These are
  the exact strings to pass to `ExportSAFETables`.
- `SAFEReadEditingTable(TableKey, Headers, Data)` is the public read bridge for
  companion modules: an editing-table read into `Headers()` plus a 1-based 2-D array
  without the header row, which is the form `WriteSAFETable` takes back. It only reads,
  and never touches the model's lock.
- `WriteSAFETable(TableKey, Data, [UnlockModel])` writes a 2-D array back into SAFE and
  applies it. `GroupName` is inactive in this SAFE release, so pass `""`. The model is
  unlocked only when SAFE reports the table as interactively importable (`ImportType`
  2), and the lock state found on entry is restored on every exit path. The column
  count must match. `ApplyEditedTables` is called with `FillImportLog:=True`, and its
  error and fatal counts are checked, with `CancelTableEditing` on a failure. Save the
  model before writing back.
- `MarkTableFailed(SheetName, StartCell, TableKey, Reason, …)` writes the bold red
  `Table '<key>' : FAILED - <reason>` marker for a table whose read failed (`Failed` =
  `True`; pass the `Warning` text as the reason), so the sheet shows that the table was
  asked for and did not come back instead of an empty gap. It returns the same cursor
  as `PrintTable`, so the two can be swapped inside one loop. `PrintTable` writes its
  own marker when the data was read but the worksheet could not hold the block.
- `TableColumnIndex(Headers, ColumnKey)` gives the 1-based index of a column key in a
  `Headers()` array (`0` = the table does not report it), case- and space-insensitive.
  The same number indexes the data array, so `col = TableColumnIndex(hdrs, "M11")`
  addresses the column by name and keeps working when SAFE reorders its columns.
- `ResetExportFailures()` and `GetLastExportFailures()` start and read the failure
  count for a batch (0 = none): each failed read and each block the worksheet could not
  hold counts once.
- `SAFEConnect()` and `SAFEDisconnect()` attach and release the running SAFE; `gSAFE`,
  `gSapModel`, `gDB` and `gConnected` expose the API objects for calls this library does
  not wrap, so call `SAFEConnect()` first.
- `ShowLog()`, `ClearLog()`, `GetLog()` and `LogMsg(msg)` are the diagnostics, and the
  error state is readable through `LastErrorNumber`, `LastErrorDescription`,
  `LastErrorSource`, `LastErrorContext` and `LastErrorText`. `LogMsg` is public, so a
  companion module writes into the same log.
- Demos: `DemoExport` (the chain in a batch loop), `DemoExportSingle`,
  `DemoExportFiltered`, `DemoExportFilteredCombo`, `DemoProcessTable` (find the largest
  value of one column, then print the same array) and `DemoListTables`.

The module targets the `SAFEv1` COM API shipped with SAFE 20. Its header comment covers
the API quirks it works around: a nonzero read code told apart from an empty table by
whether column headers came back; the data array arriving flattened row by row and
being rebuilt from the column count; `GroupName` inactive for editing tables;
`TableVersion` being an output item, so pass 0; `ApplyEditedTables` needing
`FillImportLog:=True`; and Excel's sheet limits read at runtime so the spill sheets work
for `.xls` as well as `.xlsm`.

## `SAFE_Use.bas`

Each entry point is a small wrapper that holds its own settings as local constants
(sheet, range, table key, load cases) and hands them to a worker, so adding another
area is another two-line wrapper with one place to edit per area.

- `ApplyPileCoordinates1` calls `SetPointCoordinates(SheetName, RangeAddress)`. It
  reads a three-column sheet table (`Prefix | X | Y`) from the given sheet and range,
  then sets the X and Y of every point in SAFE's `Point Object Connectivity` whose name
  begins with one of those prefixes. Matching is case-insensitive, and the first
  matching prefix in the sheet wins. It reads the editing table, makes one pass over
  its rows, hands the rebuilt table to `WriteSAFETable`, then reads back and reports the
  points written per prefix, the prefixes that matched nothing, and any values that do
  not match the requested coordinates. Coordinates are written as given, with no unit
  conversion, and the units SAFE reports for X are printed in the report. A blank row is
  ignored; a row with a non-numeric X or Y is skipped with a log line, so a header row
  inside the range is harmless. Applying the edit makes existing analysis results stale,
  and this sub neither runs the analysis nor saves the model.
- `RunAnalysis1`: attaches to the running SAFE and calls `Analyze.RunAnalysis()`. It
  does not save the model, touch the lock, or change the display filters.
- `ReadResultTableForCases(TableKey, LoadCases, SheetName, TopLeftCell,
  [IncludeHeader])`: a generic worker that runs the two library halves in chain, but
  first clears SAFE's display load combinations and puts them back afterwards, on the
  normal path and on the error path, so a result table cannot pick up combination rows
  by accident. An empty `LoadCases` (`""`, or omitted) reads every load case the model
  reports: SAFE's filter has no "all cases" value, so the names are enumerated with
  `Analyze.GetRunCaseFlag` and passed on as an ordinary list. A non-empty argument with
  no usable entry is refused as a mistake. `IncludeHeader` defaults to `False` and
  suppresses the whole block header, so the data lands flush on `TopLeftCell`; pass
  `True` for the labelled block, which moves the first data row two rows down.
- `WriteNodalReactions1`: clears the display combinations, reads `Joint Reactions` (a
  result table, so it can be read here but never written back), projects it onto a fixed
  column list (`Node | OutputCase | Fx | Fy | Fz | Mx | My | Mz`), drops every row whose
  `Fz` is zero, divides the force and moment columns by `FORCE_DIV` (1000 in the shipped
  settings) and prints the block flush at `REACT_TOPLEFT`. Its settings are local
  constants: `REACT_SHEET` and `REACT_TOPLEFT` (the top-left corner only; a block too
  big for one sheet continues on `<sheet>_2`, `_3`, …), `REACT_LOADCASES` (cases, never
  combinations; empty = every load case the model reports) and `REACT_TABLE`. A key that
  is wrong for the model writes nothing and marks the block bold red instead of
  exporting plausible-looking data, and a required output column the model does not
  report does the same while naming the key it looked for. Columns are resolved by name
  against the keys SAFE returned, so SAFE may reorder them freely; each output column
  has a `"a|b"` candidate list, so a renamed key still resolves (`Node` falls back from
  `UniqueName` to `Label`, and the substitution is logged). A cell that does not read as
  a number is written exactly as SAFE returned it.

## `Interpolation2D.bas`

Linear interpolation on an unstructured 2-D mesh. It builds a Delaunay triangulation of
the mesh points once, then interpolates a query point by finding the triangle that
contains it and blending that triangle's three corner values with barycentric weights,
so the result is exactly linear in X and Y inside a triangle and continuous across the
edges between them. The mesh does not have to be a grid, and it does not have to be
ordered: any set of points that is not (nearly) collinear will do.

The module calls nothing outside itself and needs no reference beyond Excel. It carries
its own log and its own sheet writer, so it can be dropped into any workbook.

- Sheet formats, all read from their first row downwards with no header row and a blank
  row ending the table: mesh points `X | Y | V1 | V2 | …` (three or more columns, with
  several value columns interpolated in one pass), samples `Name | X | Y` (the name
  column is optional, so exactly two columns means `X | Y`), and output
  `Name? | X | Y | V1 | V2 | …` written with no title and no header row, flush on the
  top-left corner given.
- The calls go in one order, because the module holds the mesh, the triangulation and
  the sample points between calls: `ReadMeshPoints` (the mesh first, since interpolation
  needs it), then `ReadSamplePoints` (samples second, which builds their values), then
  `WriteSampleTable` (prints them). `RunInterpolation(MeshSheet, MeshRange, SampleSheet,
  SampleRange, OutSheet, OutCell, [WriteName])` runs those three in that order, and
  `InterpolationDemo` is the ready-made entry point whose local constants are the only
  place to edit. Re-running any step is harmless: `ReadSamplePoints` rebuilds the values
  against whatever mesh is loaded, and reading a new mesh re-interpolates the sample
  points already there.
- The triangulation is Bowyer-Watson incremental Delaunay, seeded with a super triangle
  that is deleted again at the end. Duplicated or near-coincident mesh points are
  dropped (the first one wins) and reported. Every geometric test carries a tolerance
  relative to the size of the coordinates, so a degenerate or collinear input degrades
  into "nearest point" instead of producing garbage triangles. Point location is a plain
  loop over the triangles, which is deliberate because the data here is small.
- Outside the mesh the answer is the nearest mesh point's value, so those rows are
  counted, listed in the log (first 10) and flagged per row by `SampleUsedNearest`.
- State and accessors for a caller in another module: `MeshIsLoaded`, `MeshPointCount`,
  `MeshTriangleCount`, `MeshValueCount`, `SamplePointCount`, `SampleValueAt`,
  `SampleUsedNearest`, `OutsideCount`, `ClearInterpState`, the public `ReadMeshPoints`,
  `ReadSamplePoints`, `BuildInterpolatedTable`, `WriteSampleTable`, `Interpolate` and
  `InterpolateEx`, and its own log (`InterpLogMsg`, `InterpGetLog`, `InterpShowLog`,
  `InterpClearLog`).

## `auto_analysis.py`

An xlwings wrapper that runs the workbook's own macros (module `SAFE_Use`, which needs
`SAFE_Library`). The workbook must already be open in Excel, and SAFE must be running
with the model loaded. The workbook name is a module constant at the top of the file,
and every worksheet, cell and macro name is a constant inside the function that uses
it.

```python
write_coords(coords)  ->  run_analysis()  ->  write_reactions()  ->  get_utilization()
```

- `write_coords(coords: list[list[float]])` writes the (x, y) pairs into the pile
  coordinate sheet, in columns B and C downwards, and leaves the name column in A
  alone. It then runs `SAFE_Use.ApplyPileCoordinates1`.
- `run_analysis()` runs `SAFE_Use.RunAnalysis1`.
- `write_reactions()` runs `SAFE_Use.WriteNodalReactions1`.
- `get_utilization() -> list[list[float]]` reads one triple per pile (utilization
  without wind, utilization with wind, tension surplus) for the rows written by the last
  `write_coords` call. A blank or non-numeric cell raises.
- `write_test_record() -> list[float]` appends the workbook's current pile values as a
  new row at the bottom of the run-records CSV, one row per record, after checking them
  against the CSV's own header: the number of values has to match the column count, and
  each pile prefix has to line up with its columns, so a misaligned row is refused. It
  returns the row it appended.
- `print_vba_log() -> str` prints the library's shared log (`SAFE_Library.GetLog`) to
  stdout and returns it. It only reads, and never clears the log.

A workbook that is not open, a worksheet or cell that cannot be read, or a macro that
cannot be run raises a Python exception. The `SAFE_Use` subs log their own failures
instead of raising, so `print_vba_log()` is where a run that did nothing shows up.

## Table keys

Use the exact strings shown in SAFE *Display → Show Tables*, for example:

- `Point Object Connectivity` (import type 2, editable)
- `Area Load Assignments - Uniform` (2, editable)
- `Load Combination Definitions` (2, editable)
- `Element Forces - Area Shells` (0, result table)
- `Joint Displacements` (0, result table)
- `Joint Reactions` (0, result table); also `Base Reactions` and `Integrated Wall Reactions`

The key list shipped with SAFE 20 tags each table with an Import Type: 0 = not
importable (a result table), 1 = importable but not interactively, 2 = interactively
importable when the model is unlocked, 3 = importable when locked or unlocked.
Reading works for all of them; only types 2 and 3 can go back through `WriteSAFETable`.

Run `DemoListTables` or `ListSAFETables` to see the keys the open model actually
reports, since a model can omit tables and the shipped list is a superset.

## `lloyd_cvt.py`

Centroidal Voronoi tessellation of a convex polygon by Lloyd's algorithm. The outline is
partitioned into `n` subregions, and each site is moved to its own cell's area centroid
until the layout stops moving. It needs only numpy and scipy (Qhull through
`scipy.spatial.Voronoi`).

- `lloyd_cvt(polygon, n, *, seed=None, tol=1e-9, max_iter=1000)` returns the centroid
  of each subregion as an `(x, y)` tuple, sorted by descending y and then ascending x.
  `polygon` is an `(M, 2)` boundary with `M >= 3`; a clockwise ring is re-oriented and a
  repeated closing vertex dropped. `seed=None` starts from a deterministic hexagonal
  lattice, an int starts from a random layout with that seed, and either way one input
  gives one output. `tol` is relative to the polygon's bounding-box diagonal, and
  reaching `max_iter` returns the current sites as they are.
- `best_cvt(polygon, n, k=8, *, criterion="spread", max_iter=1000)` runs the lattice
  start plus `k` seeded random starts and keeps the best one that passes `validate`.
  `criterion` is `"spread"` (smallest area spread) or `"energy"` (lowest CVT energy). It
  raises when none of the `k + 1` runs validates.
- `validate(polygon, sites, tol=1e-3)` lists what is wrong with a tessellation, and an
  empty list means it is fine: site count, sites distinct and inside the polygon, no
  collapsed cell, cells that tile the polygon, and each site within `tol` of its own
  cell's centroid.
- `area_spread(polygon, sites)` returns the cell areas and their spread,
  `(max - min) / (polygon area / number of cells)`. `energy(polygon, sites)` returns
  the CVT energy, the sum over the cells of the integral of `|x - site|² dA`.

## `rename_repo.bat`

Repairs the `origin` remote URL of the repo after the GitHub repository was renamed.
Put the `.bat` in the parent folder that contains the repo subfolder and run it, either
by double-clicking or by passing the folder name; `/y` skips the confirmation prompt.
`OWNER`, `NEWNAME` and `OLDNAME` are the settings at the top of the file. If the folder
name given is not a repository it looks for the one `git` repository beside it, and it
says so when there is more than one. An `origin` that already points at the new URL is
left alone. Exit codes: 0 = ok or already correct, 1 = error, 2 = aborted by the user.

## Notes and limitations

- The modules attach to a running SAFE and never call `ApplicationExit`, which would
  close your session. None of them save the model. Connecting tries
  `GetObject(, "CSI.SAFE.API.ETABSObject")` first and falls back to the
  `CSI.SAFE.API.Helper`, then proves the link with a cheap read before reporting
  success, so a stale entry left by a closed or crashed SAFE is caught.
- Reading works whether the model is locked or unlocked. Only the write-back may need an
  unlock, and only for a table SAFE reports as `ImportType` 2; the lock state found on
  entry is restored on every exit path.
- Applying an edit makes existing analysis results stale, so save the model in SAFE
  before writing back and re-run the analysis afterwards.
- Table keys are per model and SAFE version. A key that is wrong for the model writes
  nothing and leaves a bold red `FAILED` marker rather than plausible-looking data.
- Behaviour comes from the SAFE 20 `SAFEv1` API. A later SAFE version may rename tables
  or change the API.
- A missing `SAFEv1` reference, or SAFE not open or busy analysing, is the usual cause of
  a "438 or 5" style failure. The last error, its 8-digit HRESULT and a hint are
  readable through `LastErrorText` and `ShowLog`.
- Every message goes to the shared log whether or not message boxes are on, so
  `ShowLog()` in the Immediate window (`Ctrl+G`) is the first place to look when a run
  did nothing.
