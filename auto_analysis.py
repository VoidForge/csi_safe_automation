"""xlwings driver for the Auto Piling workbook.

Runs the workbook's own SAFE macros (module SAFE_Use, which needs SAFE_Library)
through xlwings, so the workbook must ALREADY be open in Excel and SAFE must be
running with the model loaded. Intended order of calls:

    write_coords(coords)  ->  run_analysis()  ->  write_reactions()  ->  get_utilization()

get_utilization() hands the results back; write_test_record() appends them to the
records CSV; print_vba_log() prints the workbook's own log.

The workbook name is a module constant; the worksheet, cells and macro names are
hardcoded at the top of each function.
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

import xlwings as xw

WORKBOOK = "Auto Piling v20260916-0421.xlsm"    # must already be open in Excel

RECORDS_CSV = Path(__file__).resolve().parent / "data" / "SWTKT_test_records.csv"
                                # appended to by write_test_record(), one row per record

_wb: xw.Book | None = None      # workbook found on the first call, kept for the session
_n_coords = 0                   # rows written by the last write_coords() call


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


def _utilization(sheet: xw.Sheet, n_rows: int) -> list[list[float]]:
    """Read n_rows of 'Pile Coords' H:J as (util_nw, util_w, tension) triples."""
    rng = sheet.range("H3").resize(n_rows, 3)         # H util w/o wind, I util w/ wind,
    grid = _grid(rng.value, n_rows, 3)                # J tension surplus (kN)
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
    book.sheets[SHEET].range(FIRST_CELL).resize(len(rows), 2).value = rows
    _run_macro(book, MACRO)
    _n_coords = len(rows)


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

    return _utilization(_book(WORKBOOK).sheets[SHEET], _n_coords)


def write_test_record() -> list[float]:
    """Append the workbook's current pile values as a new row in the records CSV.

    Reads 'Pile Coords' (A = pile prefix, B:C = x and y, H:J = utilization w/o
    wind, utilization w/ wind and tension surplus) and appends the values to the
    BOTTOM of data/SWTKT_test_records.csv - five per pile, in the CSV's own column
    order, which the header is checked against first. Returns the row appended.
    """
    SHEET = "Pile Coords"
    FIRST_ROW = 3                                    # sheet row of the first pile
    XY = "B"                                         # x in B, y in C (A = prefix)
    PER_PILE = 5                                     # x, y, util_nw, util_w, tension

    sheet = _book(WORKBOOK).sheets[SHEET]

    raw = sheet.range(f"A{FIRST_ROW}").expand("down").value
    raw = raw if isinstance(raw, list) else [raw]     # one pile comes back as a scalar
    prefixes = ["" if p is None else str(p).strip() for p in raw]
    if not prefixes or not all(prefixes):
        raise RuntimeError(
            f"no pile prefixes found in '{SHEET}'!A{FIRST_ROW} downwards - the "
            "coordinate table is expected in A (prefix), B (x), C (y)"
        )

    last = FIRST_ROW + len(prefixes) - 1
    xy_rng = sheet.range(f"{XY}{FIRST_ROW}:C{last}")
    xy = _grid(xy_rng.value, len(prefixes), 2)
    ut = _utilization(sheet, len(prefixes))           # the same read get_utilization() uses

    values: list[float] = []
    for i in range(len(prefixes)):
        values += [_number(xy[i][0], _cell(xy_rng, i, 0)),
                   _number(xy[i][1], _cell(xy_rng, i, 1)),
                   *ut[i]]

    text = RECORDS_CSV.read_text(encoding="utf-8")
    header = next(csv.reader(text.splitlines()), None)
    if header is None:
        raise RuntimeError(f"'{RECORDS_CSV}' has no header row to line up with")
    if len(values) != len(header):
        raise ValueError(
            f"{len(values)} value(s) read from '{SHEET}' but '{RECORDS_CSV.name}' "
            f"has {len(header)} column(s) - refusing to append a misaligned row"
        )
    for i, prefix in enumerate(prefixes):
        if not header[PER_PILE * i].startswith(prefix):
            raise ValueError(
                f"pile {i + 1} in '{SHEET}' is '{prefix}' but column "
                f"{PER_PILE * i + 1} of '{RECORDS_CSV.name}' is "
                f"'{header[PER_PILE * i]}' - refusing to append a misaligned row"
            )

    with open(RECORDS_CSV, "a", newline="", encoding="utf-8") as csv_file:
        if text and not text.endswith("\n"):          # never glue onto the last row
            csv_file.write("\n")
        csv.writer(csv_file).writerow(values)
    return values


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
        text = _book(WORKBOOK).macro(MACRO)() or ""
    except Exception as exc:
        raise RuntimeError(
            f"macro '{MACRO}' failed: {exc}. The log lives in SAFE_Library, so "
            "that module has to be imported into the workbook."
        ) from exc

    print(text)
    return text