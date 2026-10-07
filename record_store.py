"""The records CSV and the store of its rows.

Every read and write of the run-records CSV happens in this module. The file holds
five numbers per pile (x, y, util_nw, util_w, tension) under a header of `~`-joined
keys, and `RecordStore` is how callers reach it: `load(path)` reads the file and
`append()` writes a trial to it and to memory in one call. `auto_analysis.py`
reads the sheet and hands the flat row here.

The readings the optimizer takes from a row are here too: `record_coords`,
`record_utilization`, `record_peak` and `best_record`.

A store is always given the CSV path, so the file rows were read from and the file
new rows are appended to cannot drift apart. Nothing here has a default path: the
caller's model glue (for `trivial_optimizer.py`, its `RECORDS_CSV`) decides.
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Sequence

DUP_DECIMALS = 6                                # rounding used for the "already tried" test
CSV_FIELDS_PER_PILE = 5                         # x, y, util_nw, util_w, tension per CSV row
HEADER_SEP = "~"                                # 'BP01~x' -> pile prefix 'BP01'

PILE_NAME, PILE_X, PILE_Y, PILE_UTIL_NW, PILE_UTIL_W, PILE_TENSION = range(6)
                                # the field order of PileRecord

PileRecord = tuple[str, float, float, float, float, float]
                                # (pile name, x, y, util_nw, util_w, tension)


def record_coords(record: list[PileRecord]) -> list[list[float]]:
    """The (x, y) coordinates of one records-CSV row."""
    return [[float(pile[PILE_X]), float(pile[PILE_Y])] for pile in record]


def record_utilization(record: list[PileRecord]) -> list[list[float]]:
    """The (util_nw, util_w, tension) triples of one records-CSV row."""
    return [[float(pile[PILE_UTIL_NW]), float(pile[PILE_UTIL_W]),
             float(pile[PILE_TENSION])] for pile in record]


def record_max_util(record: list[PileRecord]) -> float:
    """Highest governing utilization of one records-CSV row."""
    return max(max(float(pile[PILE_UTIL_NW]), float(pile[PILE_UTIL_W]))
               for pile in record)


def best_record(rows: list[list[PileRecord]]) -> list[PileRecord]:
    """The row with the lowest highest-utilization (ties keep the earliest)."""
    return min(rows, key=record_max_util)


def key_of(coords: list[list[float]]) -> tuple[float, ...]:
    """Hashable identity of a layout: every coordinate, rounded to DUP_DECIMALS.

    Coordinates make a round trip through Excel and the CSV text, so an exact
    float comparison would miss a repeat that differs in the last bit; rounding
    well below the model's precision makes the test reliable.
    """
    return tuple(round(value, DUP_DECIMALS) for pair in coords for value in pair)


# ---------------------------------------------------------------------------
# reading and writing the records CSV
# ---------------------------------------------------------------------------
def _names(header: Sequence[str]) -> list[str]:
    """The pile prefixes a header describes ('BP01~x' -> 'BP01')."""
    return [header[g * CSV_FIELDS_PER_PILE].split(HEADER_SEP)[0]
            for g in range(len(header) // CSV_FIELDS_PER_PILE)]


def _csv_float(text: str, where: str) -> float:
    """float(text) for one CSV cell; a blank or non-numeric cell raises."""
    stripped = text.strip()
    if not stripped:
        raise ValueError(f"{where} is blank, expected a number")
    try:
        return float(stripped)
    except ValueError as exc:
        raise ValueError(f"{where} is not a number: {text!r}") from exc


def _row(values: Sequence[float], names: Sequence[str]) -> list[PileRecord]:
    """Pair a flat CSV-order row with `names`; a wrong number of values raises."""
    if len(values) != CSV_FIELDS_PER_PILE * len(names):
        raise ValueError(
            f"{len(values)} recorded value(s) but {len(names)} pile(s) of "
            f"{CSV_FIELDS_PER_PILE} values were expected - refusing a misaligned "
            "row"
        )
    row: list[PileRecord] = []
    for i, name in enumerate(names):
        base = CSV_FIELDS_PER_PILE * i
        x, y, util_nw, util_w, tension = values[base:base + CSV_FIELDS_PER_PILE]
        row.append((str(name), float(x), float(y), float(util_nw), float(util_w),
                    float(tension)))
    return row


def _read(path: Path) -> tuple[list[str], list[list[PileRecord]]]:
    """Read a records CSV into (its header, one parsed row per data row).

    A header-only file gives an empty row list and a blank line is skipped; a
    non-numeric cell, or a row that does not match the header's width, raises.
    """
    with open(path, newline="", encoding="utf-8") as csv_file:
        rows = list(csv.reader(csv_file))
    if not rows:
        raise RuntimeError(f"'{path}' is empty - it has no header row")

    header = rows[0]
    if len(header) % CSV_FIELDS_PER_PILE:
        raise ValueError(
            f"'{path.name}' has {len(header)} column(s), which is not a multiple "
            f"of {CSV_FIELDS_PER_PILE}"
        )
    names = _names(header)

    out: list[list[PileRecord]] = []
    for line, cells in enumerate(rows[1:], start=2):  # the header is line 1
        if not cells:
            continue                                 # blank line: nothing to read
        if len(cells) != len(header):
            raise ValueError(
                f"line {line} of '{path.name}' has {len(cells)} column(s) but the "
                f"header has {len(header)}"
            )
        piles: list[PileRecord] = []
        for g, name in enumerate(names):
            first = g * CSV_FIELDS_PER_PILE
            values = [_csv_float(cells[first + k],
                                 f"{path.name} line {line} column {first + k + 1}")
                      for k in range(CSV_FIELDS_PER_PILE)]
            x, y, util_nw, util_w, tension = values
            piles.append((name, x, y, util_nw, util_w, tension))
        out.append(piles)
    return header, out


def _write_row(values: Sequence[float], path: Path, *,
               prefixes: Sequence[str] | None = None,
               expect_header: Sequence[str] | None = None) -> list[PileRecord]:
    """Check a flat row against `path`'s own header, then append it at the bottom.

    The row has to match the file's column count and, when `prefixes` is given,
    every pile prefix has to line up with its columns. Anything else raises
    ValueError without touching the file. `expect_header` is the header the caller
    believes the file carries, and a file that no longer matches it raises too. The
    appended row comes back parsed.
    """
    text = path.read_text(encoding="utf-8")
    header = next(csv.reader(text.splitlines()), None)
    if header is None:
        raise RuntimeError(f"'{path}' has no header row to line up with")
    if expect_header is not None and header != list(expect_header):
        raise ValueError(
            f"'{path.name}' no longer has the header this store was loaded with - "
            "refusing to append a row that would not line up"
        )
    if len(values) != len(header):
        raise ValueError(
            f"{len(values)} value(s) to append but '{path.name}' has "
            f"{len(header)} column(s) - refusing to append a misaligned row"
        )
    if prefixes is not None:
        for i, prefix in enumerate(prefixes):
            if not header[CSV_FIELDS_PER_PILE * i].startswith(prefix):
                raise ValueError(
                    f"pile {i + 1} is '{prefix}' but column "
                    f"{CSV_FIELDS_PER_PILE * i + 1} of '{path.name}' is "
                    f"'{header[CSV_FIELDS_PER_PILE * i]}' - refusing to append a "
                    "misaligned row"
                )

    with open(path, "a", newline="", encoding="utf-8") as csv_file:
        if text and not text.endswith("\n"):         # never glue onto the last row
            csv_file.write("\n")
        csv.writer(csv_file).writerow(values)
    return _row(values, _names(header))


class RecordStore:
    """Every recorded trial in memory, indexed by the layouts already tried.

    The records CSV is the durable log; `append()` writes a trial to it and to this
    copy in one call, so an iteration never re-reads the file. Rows are keyed by
    their whole pile-coordinate set (`key_of`), so a candidate whose layout has
    already been analysed is recognised without touching Excel again.
    """

    def __init__(self, rows: list[list[PileRecord]], *,
                 path: Path | str | None = None,
                 header: Sequence[str] | None = None) -> None:
        """`path` and `header` describe the CSV the rows came from, if any.

        Without a `header` the pile names come from `rows[0]`, so a store built
        from rows alone needs at least one row; one loaded from a CSV can hold
        none, because its header already says which piles it holds.
        """
        self.path: Path | None = None if path is None else Path(path)
        self._header: list[str] | None = None if header is None else list(header)

        if self._header is not None:
            names = _names(self._header)
        elif rows:
            names = [str(pile[PILE_NAME]) for pile in rows[0]]
        else:
            raise ValueError(
                "a store built from rows needs at least one row to take the pile "
                "names from - pass header= for a file that is only a header"
            )
        self.names: list[str] = names

        self.rows: list[list[PileRecord]] = [list(row) for row in rows]
        self._index: dict[tuple[float, ...], list[PileRecord]] = {}
        repeats: set[tuple[float, ...]] = set()
        for row in self.rows:
            if len(row) != len(self.names):
                raise ValueError(
                    f"a recorded row holds {len(row)} pile(s) but the first one "
                    f"holds {len(self.names)} - refusing to index them together"
                )
            if self._index_row(row):
                repeats.add(key_of(record_coords(row)))
        self.duplicate_keys = len(repeats)

    @classmethod
    def load(cls, path: Path | str) -> "RecordStore":
        """Read the records CSV at `path` into a store.

        `path` is required, so a run cannot read one CSV and append to another.
        """
        csv_path = Path(path)
        header, rows = _read(csv_path)
        return cls(rows, path=csv_path, header=header)

    def __len__(self) -> int:
        return len(self.rows)

    def best(self) -> list[PileRecord]:
        """The recorded row with the lowest highest-utilization (ties: earliest)."""
        if not self.rows:
            raise RuntimeError(
                "no recorded trial to work from - the records CSV needs at least "
                "one row before the optimizer can start"
            )
        return best_record(self.rows)

    def find(self, coords: list[list[float]]) -> list[PileRecord] | None:
        """The recorded row for these exact coordinates, or None if untried."""
        return self._index.get(key_of(coords))

    def add(self, values: Sequence[float]) -> list[PileRecord]:
        """Add a flat CSV-order row to memory; nothing is written to the CSV."""
        row = _row(values, self.names)
        self.rows.append(row)
        if self._index_row(row):
            self.duplicate_keys += 1
        return row

    def append(self, values: Sequence[float], *,
               prefixes: Sequence[str] | None = None) -> list[PileRecord]:
        """Append a trial to this store's CSV and to memory, in one call.

        The row has to match the file's own header (its column count and, when
        `prefixes` is given, every pile prefix) and the header this store was loaded
        with, so a CSV edited since then raises ValueError before anything is
        written. Returns the row as it is held in memory now.
        """
        if self.path is None:
            raise RuntimeError(
                "this store was built from rows, not from a CSV, so it has no "
                "file to append to - build it with RecordStore.load(path)"
            )
        row = _write_row(values, self.path, prefixes=prefixes,
                         expect_header=self._header)
        self.rows.append(row)
        if self._index_row(row):
            self.duplicate_keys += 1
        return row

    def _index_row(self, row: list[PileRecord]) -> bool:
        """Index one row; True when that layout was already in the index."""
        key = key_of(record_coords(row))
        held = self._index.get(key)
        if held is None:
            self._index[key] = row
            return False
        if record_max_util(row) < record_max_util(held):
            self._index[key] = row         # of repeats, keep the best result
        return True
