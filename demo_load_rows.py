"""Demo: load the first recorded layout, pause, then load the best one.

Loads row 1 of a records CSV into 'Pile Coords' and pushes it into SAFE, waits
for you to type anything (look at Excel and SAFE in the meantime), then loads the
best recorded row the same way. Both layouts come straight from the CSV, so the
peaks printed here are the recorded ones and no new trial is appended. It ends
with the net x and y shift of every pile that moved between the two rows.

`RUN_ANALYSIS` decides whether each load also re-runs SAFE's analysis, which is
what refreshes the utilization column. It is off by default: the load is then
geometry only, and the sheet's utilization column still holds the numbers of the
layout loaded before, not of the one now in the model.

Run it with:  python demo_load_rows.py [records.csv]
"""

from __future__ import annotations

import sys
from pathlib import Path

import auto_analysis as aa
from record_store import (PILE_NAME, PileRecord, RecordStore, record_coords,
                          record_max_util)

RECORDS_CSV: Path = Path(__file__).resolve().parent / "data" / "SWTKT_test_records.csv"

RUN_ANALYSIS = False                            # True: re-run SAFE per load (slow)


def load(row: list[PileRecord]) -> None:
    """Write one recorded row's coordinates into the sheet and into SAFE."""
    aa.write_coords(record_coords(row))         # sheet write + ApplyPileCoordinates1
    if RUN_ANALYSIS:
        aa.run_analysis()
        aa.write_reactions()


def report_shifts(first: list[PileRecord], best: list[PileRecord]) -> None:
    """Print every pile's net (x, y) shift from `first` to `best`.

    Only piles that actually moved are listed, so the output is the layout change
    itself rather than one line per pile. Both rows come from the same CSV, so
    they hold the same piles in the same order.
    """
    shifts = [(str(name), best_xy[0] - first_xy[0], best_xy[1] - first_xy[1])
              for name, first_xy, best_xy in zip(
                  (pile[PILE_NAME] for pile in first),
                  record_coords(first), record_coords(best))]
    moved = [shift for shift in shifts if shift[1] or shift[2]]
    if not moved:
        print("no pile moved: row 1 and the best row hold the same coordinates")
        return
    print(f"{len(moved)} pile(s) moved from row 1 to the best row:")
    for name, dx, dy in moved:
        print(f"  {name}: dx {dx:.6g}, dy {dy:.6g}")


def main() -> None:
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else RECORDS_CSV
    store = RecordStore.load(path)
    if not len(store):
        raise SystemExit(f"'{path.name}' is header-only: there is no row to load")

    first = store.rows[0]
    best = store.best()
    print(f"{path.name}: {len(store)} recorded layout(s), {len(store.names)} pile(s)")

    print(f"loading recorded row 1 (highest recorded utilization "
          f"{record_max_util(first):.6g})")
    load(first)

    input("row 1 is loaded; press Enter (any text) to load the best row ...")

    if best is first:
        print("the first recorded row is already the best one; loading it again")
    print(f"loading the best recorded row (highest recorded utilization "
          f"{record_max_util(best):.6g})")
    load(best)

    report_shifts(first, best)


if __name__ == "__main__":
    main()
