"""Greedy pile-layout optimizer for the Auto Piling workbook.

Every iteration moves ONE pile and keeps the best of eight trial moves:

  1. Pick the pile with the highest governing utilization, where a pile's
     governing utilization is the worse of its two cases:
     ``max(utilization without wind, utilization with wind)``.
  2. Move that pile by ``STEP_SIZE`` in each of the eight compass directions. A
     candidate whose exact coordinates are already in the record store is
     skipped - its result is known, so no analysis runs and no row is appended.
     Every other candidate is analysed and its row is appended to `RECORDS_CSV`
     immediately (`auto_analysis.append_test_record`), so an interrupted run
     keeps every trial it had reached.
  3. The `RecordStore` mirrors that CSV in memory: each trial's row is added to it
     as it is recorded, and the next baseline is its best row (lowest highest
     utilization). Nothing is re-read from disk during a run.
  4. If that best row is no better than the baseline the iteration started from,
     raise `NoImprovementError` - the layout is a local optimum at this step
     size. That is also what an iteration whose candidates were all tried before
     reports.

The pile count and pile order come from the records CSV header, so nothing here
assumes 22 piles. All model-specific knowledge is confined to the "model glue"
block below - `RECORDS_CSV`, `STEP_SIZE`, `DIRECTIONS` and `evaluate` - so
swapping that block is all it takes to point the optimizer at another model.

A recorded row is only valid for the model it was measured against, so a model
edit or a changed STEP_SIZE makes older rows stale - they can be both the
baseline and the thing that marks a candidate "already tried".

Requirements: the workbook must already be open in Excel and SAFE must be
running with the model loaded.

Run it with:  python trivial_optimizer.py
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Sequence

import auto_analysis as aa

# ---------------------------------------------------------------------------
# model glue - the only part that knows about the Auto Piling model.
# ---------------------------------------------------------------------------
RECORDS_CSV: Path = aa.RECORDS_CSV              # run-records log: read back, appended to

STEP_SIZE = 500.0                               # trial move distance, model units

DIRECTIONS: tuple[tuple[int, int], ...] = (     # eight compass directions
    (1, 0),                                     # E
    (1, 1),                                     # NE
    (0, 1),                                     # N
    (-1, 1),                                    # NW
    (-1, 0),                                    # W
    (-1, -1),                                   # SW
    (0, -1),                                    # S
    (1, -1),                                    # SE
)

DUP_DECIMALS = 6                                # rounding used for the "already tried" test
CSV_FIELDS_PER_PILE = 5                         # x, y, util_nw, util_w, tension per CSV row

PILE_NAME, PILE_X, PILE_Y, PILE_UTIL_NW, PILE_UTIL_W, PILE_TENSION = range(6)
                                # the field order of auto_analysis.PileRecord


def evaluate(coords: list[list[float]]) -> list[float]:
    """Push `coords` into the model, analyse it, and record the trial.

    Returns the flat row just appended to `RECORDS_CSV` - five numbers per pile,
    in the CSV's own order. The utilization comes from that row instead of a
    second sheet read, so a trial costs one Excel read. The caller puts the row
    into the `RecordStore`.
    """
    aa.write_coords(coords)
    aa.run_analysis()
    aa.write_reactions()
    prefixes, values = aa.read_pile_values()    # one Excel round trip for the table
    return aa.append_test_record(values, prefixes, path=RECORDS_CSV)
# ---------------------------------------------------------------------------
# end model glue
# ---------------------------------------------------------------------------


def key_of(coords: list[list[float]]) -> tuple[float, ...]:
    """Hashable identity of a layout: every coordinate, rounded to DUP_DECIMALS.

    Coordinates make a round trip through Excel and the CSV text, so an exact
    float comparison would miss a repeat that differs in the last bit; rounding
    well below the model's precision makes the test reliable.
    """
    return tuple(round(value, DUP_DECIMALS) for pair in coords for value in pair)


class RecordStore:
    """Every recorded trial in memory, indexed by the layouts already tried.

    The records CSV stays the durable log - every trial is still appended to it -
    and this is the copy the optimizer works from, so an iteration never re-reads
    the file. Rows are keyed by their whole pile-coordinate set (`key_of`), so a
    candidate whose layout has already been analysed is recognised without
    touching Excel again.
    """

    def __init__(self, rows: list[list[aa.PileRecord]]) -> None:
        if not rows:
            raise RuntimeError(
                "no recorded trial to work from - the records CSV needs at least "
                "one row before the optimizer can start"
            )
        self.rows: list[list[aa.PileRecord]] = [list(row) for row in rows]
        self.names: list[str] = [str(pile[PILE_NAME]) for pile in self.rows[0]]

        self._index: dict[tuple[float, ...], list[aa.PileRecord]] = {}
        counts: dict[tuple[float, ...], int] = {}
        for row in self.rows:
            if len(row) != len(self.names):
                raise ValueError(
                    f"a recorded row holds {len(row)} pile(s) but the first one "
                    f"holds {len(self.names)} - refusing to index them together"
                )
            key = key_of(record_coords(row))
            counts[key] = counts.get(key, 0) + 1
            held = self._index.get(key)
            if held is None or record_peak(row) < record_peak(held):
                self._index[key] = row             # of repeats, keep the best result
        self.duplicate_keys = sum(1 for n in counts.values() if n > 1)

    @classmethod
    def load(cls, path: Path | str = RECORDS_CSV) -> "RecordStore":
        """Read a records CSV (default `RECORDS_CSV`) into a store."""
        return cls(aa.read_test_records(path))

    def __len__(self) -> int:
        return len(self.rows)

    def best(self) -> list[aa.PileRecord]:
        """The recorded row with the lowest highest-utilization (ties: earliest)."""
        return best_record(self.rows)

    def find(self, coords: list[list[float]]) -> list[aa.PileRecord] | None:
        """The recorded row for these exact coordinates, or None if untried."""
        return self._index.get(key_of(coords))

    def add(self, values: list[float]) -> list[aa.PileRecord]:
        """Add a flat CSV-order row - the one just appended to the records CSV."""
        if len(values) != CSV_FIELDS_PER_PILE * len(self.names):
            raise ValueError(
                f"{len(values)} recorded value(s) but the store holds "
                f"{len(self.names)} pile(s) of {CSV_FIELDS_PER_PILE} values - "
                "refusing to add a misaligned row"
            )

        row: list[aa.PileRecord] = []
        for i, name in enumerate(self.names):
            base = CSV_FIELDS_PER_PILE * i
            x, y, util_nw, util_w, tension = values[base:base + CSV_FIELDS_PER_PILE]
            row.append((name, x, y, util_nw, util_w, tension))

        self.rows.append(row)
        key = key_of(record_coords(row))
        held = self._index.get(key)
        if held is None:
            self._index[key] = row
        else:
            self.duplicate_keys += 1
            if record_peak(row) < record_peak(held):
                self._index[key] = row
        return row


def record_coords(record: list[aa.PileRecord]) -> list[list[float]]:
    """The (x, y) coordinates of one records-CSV row."""
    return [[float(pile[PILE_X]), float(pile[PILE_Y])] for pile in record]


def record_utilization(record: list[aa.PileRecord]) -> list[list[float]]:
    """The (util_nw, util_w, tension) triples of one records-CSV row."""
    return [[float(pile[PILE_UTIL_NW]), float(pile[PILE_UTIL_W]),
             float(pile[PILE_TENSION])] for pile in record]


def record_peak(record: list[aa.PileRecord]) -> float:
    """Highest governing utilization of one records-CSV row."""
    return max(max(float(pile[PILE_UTIL_NW]), float(pile[PILE_UTIL_W]))
               for pile in record)


def best_record(rows: list[list[aa.PileRecord]]) -> list[aa.PileRecord]:
    """The row with the lowest highest-utilization (ties keep the earliest)."""
    return min(rows, key=record_peak)


def governing(utilization: list[list[float]]) -> list[float]:
    """Per pile, the worse of the no-wind and with-wind utilization."""
    return [max(triple[0], triple[1]) for triple in utilization]


def peak(utilization: list[list[float]]) -> float:
    """Highest governing utilization across every pile - the objective."""
    return max(governing(utilization))


def move_pile(coords: list[list[float]], index: int,
              direction: tuple[int, int], step_size: float) -> list[list[float]]:
    """Copy `coords` with pile `index` stepped `step_size` along `direction`."""
    moved = [list(pair) for pair in coords]
    moved[index][0] += direction[0] * step_size
    moved[index][1] += direction[1] * step_size
    return moved


class NoImprovementError(RuntimeError):
    """None of the eight trials beat the best layout the records CSV held.

    Carries the best state found (``coords``, ``utilization``, ``pile_index``,
    ``iterations``) so the caller can carry on from it.
    """

    def __init__(self, pile_index: int, coords: list[list[float]],
                 utilization: list[list[float]], iterations: int,
                 best_trial_peak: float) -> None:
        self.pile_index = pile_index
        self.coords = coords
        self.utilization = utilization
        self.iterations = iterations
        self.best_trial_peak = best_trial_peak
        super().__init__(
            f"no trial move improved the layout after {iterations} accepted "
            f"move(s): moving pile {pile_index + 1} in any of the eight "
            f"directions left the best recorded utilization at "
            f"{peak(utilization):.6g} (best trial this iteration "
            f"{best_trial_peak:.6g}). The layout is a local optimum for this "
            "step size."
        )


def optimize(*,
             store: RecordStore | None = None,
             step_size: float = STEP_SIZE,
             directions: Sequence[tuple[int, int]] = DIRECTIONS,
             evaluate_fn: Callable[[list[list[float]]], list[float]] = evaluate,
             max_iterations: int | None = None,
             log: Callable[[str], None] = print,
             ) -> tuple[list[list[float]], list[list[float]], int]:
    """Hill-climb the pile layout, one pile and one step per iteration.

    Starts from the store's best row and works from the store throughout, so a
    layout that has already been analysed is never analysed again. Returns the
    best ``(coords, utilization, iterations)``; raises `NoImprovementError` when
    the iteration's candidates all fail to beat the baseline.
    """
    if store is None:
        store = RecordStore.load()
    iterations = 0

    while max_iterations is None or iterations < max_iterations:
        baseline = store.best()
        coords = record_coords(baseline)
        utilization = record_utilization(baseline)
        current_peak = record_peak(baseline)
        pile_index = max(range(len(coords)),
                         key=governing(utilization).__getitem__)
        log(f"iteration {iterations + 1}: baseline highest utilization "
            f"{current_peak:.6g}, pile {pile_index + 1} is the worst")

        for direction in directions:
            trial_coords = move_pile(coords, pile_index, direction, step_size)
            known = store.find(trial_coords)
            if known is not None:
                log(f"  {direction}: already tried (highest utilization "
                    f"{record_peak(known):.6g}) - skipped")
                continue
            record = store.add(evaluate_fn(trial_coords))
            log(f"  {direction}: highest utilization {record_peak(record):.6g}")

        accepted = store.best()                          # the store is the memory
        if record_peak(accepted) >= current_peak:
            raise NoImprovementError(pile_index, record_coords(accepted),
                                     record_utilization(accepted), iterations,
                                     record_peak(accepted))
        iterations += 1
        log(f"  accepted the best of the eight -> highest utilization "
            f"{record_peak(accepted):.6g}")

    best = store.best()
    return record_coords(best), record_utilization(best), iterations


def main() -> None:
    """Optimize from the best layout the records CSV already holds."""
    store = RecordStore.load()
    print(f"{RECORDS_CSV.name}: {len(store.names)} piles, {len(store)} recorded "
          f"layout(s), {store.duplicate_keys} repeated, best highest utilization "
          f"{record_peak(store.best()):.6g}, step size {STEP_SIZE:g}")

    try:
        with aa.quiet_excel():                       # fewer Excel messages: fewer OLE prompts
            coords, utilization, iterations = optimize(store=store)
    except NoImprovementError as exc:
        print(exc)
        print(f"stopped after {exc.iterations} accepted move(s) at highest "
              f"utilization {peak(exc.utilization):.6g}")
        raise

    print(f"finished after {iterations} accepted move(s): "
          f"highest utilization {peak(utilization):.6g}")


if __name__ == "__main__":
    main()
