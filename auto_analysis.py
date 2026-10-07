"""xlwings driver for the Auto Piling workbook.

Runs the workbook's own SAFE macros (module SAFE_Use, which needs SAFE_Library)
through xlwings, so the workbook must ALREADY be open in Excel and SAFE must be
running with the model loaded. Intended order of calls:

    write_coords(coords)  ->  run_analysis()  ->  write_reactions()  ->  get_utilization()

init_pile_layout(polygon, weights=None) generates `coords` with weighted_cvt and
can stand in for the write_coords() call.

get_utilization() hands the results back; write_test_record() appends them to the
records CSV; print_vba_log() prints the workbook's own log.

The workbook name is a module constant; the worksheet, cells and macro names are
hardcoded at the top of each function.
"""

from __future__ import annotations

import csv
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator, Sequence

import xlwings as xw

WORKBOOK = "Auto Piling v20260916-0421.xlsm"    # must already be open in Excel

RECORDS_CSV = Path(__file__).resolve().parent / "data" / "SWTKT_test_records.csv"
                                # appended to by write_test_record(), one row per record

_wb: xw.Book | None = None      # workbook found on the first call, kept for the session
_n_coords = 0                   # rows written by the last write_coords() call


# ---------------------------------------------------------------------------
# Excel OLE call log
# ---------------------------------------------------------------------------
def _stamp() -> str:
    """Local wall-clock time of an OLE call, to the millisecond."""
    return datetime.now().strftime("%H:%M:%S.%f")[:-3]


@contextmanager
def _ole(label: str) -> Iterator[None]:
    """Log one Excel OLE call to stdout with its time and duration.

    One line is printed when the call goes out and one when it comes back, so a
    call that is still running (a long SAFE analysis, say) is visible rather
    than silent. A call that raises is marked '!!' with the exception type.
    """
    started = time.perf_counter()
    print(f"{_stamp()}  OLE -> {label}", flush=True)
    try:
        yield
    except BaseException as exc:
        print(f"{_stamp()}  OLE !! {label}  "
              f"({time.perf_counter() - started:.3f}s, {type(exc).__name__})",
              flush=True)
        raise
    print(f"{_stamp()}  OLE <- {label}  ({time.perf_counter() - started:.3f}s)",
          flush=True)


# ---------------------------------------------------------------------------
# workbook / macro plumbing
# ---------------------------------------------------------------------------
def _book(workbook: str) -> xw.Book:
    """Return the ALREADY OPEN workbook called `workbook` (cached)."""
    global _wb
    if _wb is not None:
        try:
            _wb.name                        # cheap liveness probe: a closed book raises
            return _wb
        except Exception:
            _wb = None

    with _ole(f"attach to Excel workbook '{workbook}'"):
        try:
            apps = list(xw.apps)
        except Exception as exc:
            raise RuntimeError(
                "no running Excel instance found - open the workbook first"
            ) from exc

        for app in apps:
            for book in app.books:
                if book.name.lower() == workbook.lower():
                    _wb = book
                    return _wb
        raise RuntimeError(
            f"workbook '{workbook}' is not open in Excel - open it and call again"
        )


def _run_macro(book: xw.Book, macro_name: str) -> None:
    """Run one VBA macro; a macro that cannot be run raises RuntimeError."""
    try:
        with _ole(f"run macro '{macro_name}'"):
            book.macro(macro_name)()
    except Exception as exc:
        raise RuntimeError(
            f"macro '{macro_name}' failed: {exc}. The SAFE_Use subs log their own "
            "errors instead of raising, so check the log in the workbook (ShowLog) "
            "for the reason."
        ) from exc


# ---------------------------------------------------------------------------
# value helpers
# ---------------------------------------------------------------------------
def _number(value: object, where: str) -> float:
    """float(value); a blank or non-numeric value raises ValueError."""
    if value is None or isinstance(value, bool):
        raise ValueError(f"{where} is not a number: {value!r}")
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        raise ValueError(f"{where} is blank, expected a number")
    try:
        return float(text)
    except ValueError as exc:
        raise ValueError(f"{where} is not a number: {value!r}") from exc


def _grid(values: Any, n_rows: int, n_cols: int) -> list[list[Any]]:
    """Normalize a rectangular xlwings .value read into a list of row lists."""
    if values is None:                      # whole range empty
        return [[None] * n_cols for _ in range(n_rows)]
    if n_rows == 1:
        return [list(values)] if n_cols > 1 else [[values]]
    return [list(row) for row in values]


def _cell(rng: xw.Range, r: int, c: int) -> str:
    """Address of the cell r rows down and c columns right of rng's top-left."""
    return xw.utils.col_name(rng.column + c) + str(rng.row + r)


def _pile_prefixes(sheet: xw.Sheet, first_row: int) -> list[str]:
    """Pile prefixes in column A from `first_row` down the contiguous block.

    The block ends at the first blank cell, which is also the row count a
    coordinate write covers, so a blank row can never leave a prefix behind.
    """
    raw = sheet.range(f"A{first_row}").expand("down").value
    raw = raw if isinstance(raw, list) else [raw]  # one pile comes back as a scalar
    prefixes = ["" if p is None else str(p).strip() for p in raw]
    if not prefixes or not all(prefixes):
        raise RuntimeError(
            f"no pile prefixes found in '{sheet.name}'!A{first_row} downwards - the "
            "coordinate table is expected in A (prefix), B (x), C (y)"
        )
    return prefixes


def _write_coords_range(book: xw.Book, sheet_name: str, first_cell: str,
                        coords: list[list[float]]) -> None:
    """Write (x, y) rows into sheet_name at first_cell, under one OLE log line."""
    with _ole(f"write '{sheet_name}'!{first_cell} ({len(coords)} pile(s), x|y)"):
        book.sheets[sheet_name].range(first_cell).resize(len(coords), 2).value = coords


def _utilization(book: xw.Book, sheet_name: str, n_rows: int) -> list[list[float]]:
    """Read n_rows of sheet_name's H:J as (util_nw, util_w, tension) triples."""
    with _ole(f"read '{sheet_name}'!H3:J{n_rows + 2} utilization"):
        rng = book.sheets[sheet_name].range("H3").resize(n_rows, 3)   # H util w/o wind,
        grid = _grid(rng.value, n_rows, 3)              # I util w/ wind, J tension (kN)
        return [[_number(grid[r][c], _cell(rng, r, c)) for c in range(3)]
                for r in range(n_rows)]


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------
def write_coords(coords: list[list[float]]) -> None:
    """Write (x, y) pairs to 'Pile Coords'!B3:C, then apply them in SAFE."""
    SHEET = "Pile Coords"
    FIRST_CELL = "B3"                                # x in B, y in C, downwards
    MACRO = "SAFE_Use.ApplyPileCoordinates1"         # reads A2:C50 (Prefix | X | Y)

    global _n_coords
    if not coords:
        raise ValueError("coords is empty - nothing to write")

    rows = []
    for i, pair in enumerate(coords, start=1):
        if not isinstance(pair, (list, tuple)) or len(pair) != 2:
            raise ValueError(f"coords[{i}] is not an (x, y) pair: {pair!r}")
        rows.append([_number(pair[0], f"coords[{i}][0]"),
                     _number(pair[1], f"coords[{i}][1]")])

    book = _book(WORKBOOK)
    _write_coords_range(book, SHEET, FIRST_CELL, rows)
    _run_macro(book, MACRO)
    _n_coords = len(rows)


def init_pile_layout(polygon: Sequence[Sequence[float]],
                     weights: Sequence[float] | None = None, *,
                     n_piles: int | None = None,
                     seed: int | None = None,
                     best_of: int = 0,
                     area_tol: float = 1e-3,
                     allow_nonconverged: bool = False,
                     apply: bool = True) -> list[list[float]]:
    """Generate a pile layout on `polygon` and write it to the workbook.

    `weighted_cvt` builds the layout: `polygon` is the convex cap outline, and
    `weights` holds one area weight per pile, so a larger weight gives a larger
    cell and a lighter pile. `weights=None` gives every pile the same area; the
    pile count is then `n_piles`, or the prefixes already in 'Pile Coords' when
    `n_piles` is None. Sites arrive ordered by descending y and then ascending x,
    and the function writes them top to bottom onto the existing rows, so the
    count has to match the prefix column. The prefixes themselves are not touched.

    `best_of` and `area_tol` go through to `weighted_cvt` / `best_weighted_cvt`
    (`best_of=0` is one deterministic start). A layout whose worst cell misses its
    target area by more than `area_tol` of the cap area raises unless
    `allow_nonconverged` is True. With `apply=True` (the default) the function
    writes the coordinates and pushes them into SAFE through
    `SAFE_Use.ApplyPileCoordinates1`. With `apply=False` it writes the sheet only
    and leaves `_n_coords` alone, so `get_utilization()` stays tied to the last
    `write_coords()` call.

    Returns the coordinates written, in `write_coords()`'s own shape.
    """
    import weighted_cvt as wcvt              # numpy/scipy only when this runs

    SHEET = "Pile Coords"
    FIRST_CELL = "B3"                        # x in B, y in C, downwards
    FIRST_ROW = 3                            # sheet row of the first pile

    book = _book(WORKBOOK)
    with _ole(f"read '{SHEET}' pile prefixes"):
        prefixes = _pile_prefixes(book.sheets[SHEET], FIRST_ROW)

    area_weights = None if weights is None else [float(w) for w in weights]
    if area_weights is not None and n_piles is not None and n_piles != len(area_weights):
        raise ValueError(
            f"n_piles={n_piles} disagrees with the {len(area_weights)} weight(s)"
        )
    n = n_piles if area_weights is None else len(area_weights)
    if n is None:
        n = len(prefixes)                    # one pile per existing prefix
    if n <= 0:
        raise ValueError(f"a layout needs at least one pile, got {n}")
    if n != len(prefixes):
        raise ValueError(
            f"the layout holds {n} pile(s) but '{SHEET}'!A{FIRST_ROW} holds "
            f"{len(prefixes)} prefix(es) - refusing to write a layout that does not "
            "line up with the sheet"
        )
    if area_weights is None:
        area_weights = [1.0] * n             # equal areas

    if best_of > 0:
        result = wcvt.best_weighted_cvt(polygon, area_weights, k=best_of,
                                        seed=seed, area_tol=area_tol)
    else:
        result = wcvt.weighted_cvt(polygon, area_weights,
                                   seed=seed, area_tol=area_tol)

    if not result.converged and not allow_nonconverged:
        raise RuntimeError(
            f"the layout did not reach the requested areas: the worst cell misses its "
            f"target by {result.max_area_error:.3g} of the cap area after "
            f"{result.iterations} iteration(s) (area_tol={area_tol:g}) - nothing was "
            "written. Loosen area_tol, raise best_of, or pass allow_nonconverged=True."
        )

    print(f"{_stamp()}  init_pile_layout: {len(result.sites)} pile(s), "
          f"{result.iterations} iteration(s), converged={result.converged}, worst "
          f"area error {result.max_area_error:.3g} of the cap area; cells "
          f"{min(result.areas):.6g}..{max(result.areas):.6g} vs targets "
          f"{min(result.targets):.6g}..{max(result.targets):.6g}", flush=True)

    coords = [[float(x), float(y)] for x, y in result.sites]
    if apply:
        write_coords(coords)                 # sheet write + ApplyPileCoordinates1
    else:
        _write_coords_range(book, SHEET, FIRST_CELL, coords)
    return coords


def run_analysis() -> None:
    """Run SAFE's analysis for the open model."""
    MACRO = "SAFE_Use.RunAnalysis1"

    _run_macro(_book(WORKBOOK), MACRO)


def write_reactions() -> None:
    """Write SAFE's nodal reactions into the workbook."""
    MACRO = "SAFE_Use.WriteNodalReactions1"

    _run_macro(_book(WORKBOOK), MACRO)


def get_utilization() -> list[list[float]]:
    """Read every pile's (util_nw, util_w, tension) triple, top to bottom."""
    SHEET = "Pile Coords"

    if _n_coords <= 0:
        raise RuntimeError(
            "write_coords() has not run yet, so the number of utilization rows is "
            "unknown - call write_coords() first"
        )

    return _utilization(_book(WORKBOOK), SHEET, _n_coords)


def read_pile_values() -> tuple[list[str], list[float]]:
    """Read 'Pile Coords' into (pile prefixes, flat CSV-ordered values).

    One round trip for the whole table: the prefix column A downwards, the x and
    y in B:C, and the utilization w/o wind, utilization w/ wind and tension
    surplus in H:J - five numbers per pile, in the records CSV's own column
    order. Nothing is written anywhere and no CSV is touched.
    """
    SHEET = "Pile Coords"
    FIRST_ROW = 3                                    # sheet row of the first pile
    XY = "B"                                         # x in B, y in C (A = prefix)

    book = _book(WORKBOOK)

    with _ole(f"read '{SHEET}' pile table"):
        sheet = book.sheets[SHEET]
        prefixes = _pile_prefixes(sheet, FIRST_ROW)

        last = FIRST_ROW + len(prefixes) - 1
        xy_rng = sheet.range(f"{XY}{FIRST_ROW}:C{last}")
        xy = _grid(xy_rng.value, len(prefixes), 2)

    ut = _utilization(book, SHEET, len(prefixes))     # the same read get_utilization() uses

    values: list[float] = []
    for i in range(len(prefixes)):
        values += [_number(xy[i][0], _cell(xy_rng, i, 0)),
                   _number(xy[i][1], _cell(xy_rng, i, 1)),
                   *ut[i]]
    return prefixes, values


def append_test_record(values: list[float], prefixes: list[str] | None = None,
                       path: Path | str | None = None) -> list[float]:
    """Append one flat row of `values` to the BOTTOM of a records CSV.

    `values` is the CSV-ordered row that read_pile_values() returns (five numbers
    per pile); `path` is the CSV to append to and defaults to the workbook's own
    data/SWTKT_test_records.csv. The row is checked against that file's own header
    first: the number of values has to match the column count and, when
    `prefixes` is given, each pile prefix has to line up with its columns, so a
    misaligned row is refused rather than appended. Returns the row it appended.
    """
    PER_PILE = 5                                     # x, y, util_nw, util_w, tension

    csv_path = RECORDS_CSV if path is None else Path(path)
    text = csv_path.read_text(encoding="utf-8")
    header = next(csv.reader(text.splitlines()), None)
    if header is None:
        raise RuntimeError(f"'{csv_path}' has no header row to line up with")
    if len(values) != len(header):
        raise ValueError(
            f"{len(values)} value(s) to append but '{csv_path.name}' has "
            f"{len(header)} column(s) - refusing to append a misaligned row"
        )
    if prefixes is not None:
        for i, prefix in enumerate(prefixes):
            if not header[PER_PILE * i].startswith(prefix):
                raise ValueError(
                    f"pile {i + 1} is '{prefix}' but column "
                    f"{PER_PILE * i + 1} of '{csv_path.name}' is "
                    f"'{header[PER_PILE * i]}' - refusing to append a misaligned row"
                )

    with open(csv_path, "a", newline="", encoding="utf-8") as csv_file:
        if text and not text.endswith("\n"):          # never glue onto the last row
            csv_file.write("\n")
        csv.writer(csv_file).writerow(values)
    return values


def write_test_record() -> list[float]:
    """Append the workbook's current pile values as a new row in the records CSV.

    Reads 'Pile Coords' (A = pile prefix, B:C = x and y, H:J = utilization w/o
    wind, utilization w/ wind and tension surplus) and appends the values to the
    BOTTOM of data/SWTKT_test_records.csv - five per pile, in the CSV's own column
    order, which the header is checked against first. Returns the row appended.
    """
    prefixes, values = read_pile_values()
    return append_test_record(values, prefixes)


PileRecord = tuple[str, float, float, float, float, float]
                                # (pile name, x, y, util_nw, util_w, tension)


def read_test_records(path: Path | str | None = None) -> list[list[PileRecord]]:
    """Read the records CSV back as one tuple per pile per row.

    Returns one list per DATA row (header excluded, blank lines skipped), each
    holding one tuple per pile in the CSV's column order:
    (pile name, x, y, util_nw, util_w, tension) - the name being the column key
    before '~' ('BP01~x' -> 'BP01'). Defaults to data/SWTKT_test_records.csv;
    a header-only file gives []; values arrive as floats, never as strings.
    """
    PER_PILE = 5                                     # x, y, util_nw, util_w, tension

    csv_path = RECORDS_CSV if path is None else Path(path)
    with open(csv_path, newline="", encoding="utf-8") as csv_file:
        rows = list(csv.reader(csv_file))
    if not rows:
        raise RuntimeError(f"'{csv_path}' is empty - it has no header row")

    header = rows[0]
    if len(header) % PER_PILE:
        raise ValueError(
            f"'{csv_path.name}' has {len(header)} column(s), which is not a "
            f"multiple of {PER_PILE}"
        )
    names = [header[g * PER_PILE].split("~")[0] for g in range(len(header) // PER_PILE)]

    out: list[list[PileRecord]] = []
    for line, cells in enumerate(rows[1:], start=2):  # the header is line 1
        if not cells:
            continue                                 # blank line - nothing to read
        if len(cells) != len(header):
            raise ValueError(
                f"line {line} of '{csv_path.name}' has {len(cells)} column(s) but "
                f"the header has {len(header)}"
            )
        piles: list[PileRecord] = []
        for g, name in enumerate(names):
            first = g * PER_PILE
            values = [_number(cells[first + k],
                              f"{csv_path.name} line {line} column {first + k + 1}")
                      for k in range(PER_PILE)]
            x, y, util_nw, util_w, tension = values
            piles.append((name, x, y, util_nw, util_w, tension))
        out.append(piles)
    return out


def print_vba_log() -> str:
    """Print the workbook's shared VBA log (SAFE_Library) to stdout; return it too.

    The SAFE_Use subs report failures by logging instead of raising, so a macro
    that ran but did nothing is only visible here. Read-only - it never clears.
    """
    MACRO = "SAFE_Library.GetLog"                    # returns the whole log as one string

    try:
        with _ole(f"run macro '{MACRO}'"):
            text = _book(WORKBOOK).macro(MACRO)() or ""
    except Exception as exc:
        raise RuntimeError(
            f"macro '{MACRO}' failed: {exc}. The log lives in SAFE_Library, so "
            "that module has to be imported into the workbook."
        ) from exc

    print(text)
    return text


# ---------------------------------------------------------------------------
# Excel quiet mode
# ---------------------------------------------------------------------------
@contextmanager
def quiet_excel() -> Iterator[None]:
    """Silence Excel's screen updates and alerts for the duration of a block.

    Long stretches of SAFE work keep Excel parked inside an outbound OLE call,
    and while it is parked any message its window receives makes it put up
    "Microsoft Excel is waiting for another application to complete an OLE
    action". Dropping repaint and alert traffic removes most of the messages
    that trigger it. Only these two settings are touched - events, calculation
    mode and the model are left alone - and the values found on entry are put
    back on every exit path, so a failure inside the block cannot leave Excel
    stuck looking frozen.
    """
    app = _book(WORKBOOK).app
    with _ole("silence Excel screen updates and alerts"):
        saved = (app.api.ScreenUpdating, app.api.DisplayAlerts)
        app.api.ScreenUpdating = False
        app.api.DisplayAlerts = False
    try:
        yield
    finally:
        with _ole("restore Excel screen updates and alerts"):
            app.api.ScreenUpdating = saved[0]
            app.api.DisplayAlerts = saved[1]