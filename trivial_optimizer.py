"""Greedy pile-layout optimizer for the Auto Piling workbook.

Every iteration moves ONE pile and keeps the best of eight trial moves:

  1. Pick the pile with the highest governing utilization, where a pile's
     governing utilization is the worse of its two cases:
     ``max(utilization without wind, utilization with wind)``.
  2. Move that pile by ``STEP_SIZE`` in each of the eight compass directions
     and analyse every candidate. A trial is appended to the run-records CSV
     the moment it has been analysed (`auto_analysis.write_test_record`), so an
     interrupted run keeps every trial it had reached.
  3. Finish the iteration by reading the records CSV back and taking its best
     row (lowest highest-utilization) as the next baseline. Nothing has to be
     re-analysed to register a baseline: the accepted layout is already a row.
  4. If that best row is no better than the baseline the iteration started
     from, raise `NoImprovementError` - the layout is a local optimum at this
     step size.

The pile count and pile order come from the records CSV header, so nothing here
assumes 22 piles. All model-specific knowledge is confined to the "model glue"
block below - swapping that block (STEP_SIZE, DIRECTIONS, evaluate) is all it
takes to point the optimizer at another model.

The CSV is both the log and the working memory: every iteration starts from the
best row it already holds, and normally that is the layout the workbook shows.

Requirements: the workbook must already be open in Excel and SAFE must be
running with the model loaded.

Run it with:  python trivial_optimizer.py
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Callable, Sequence

import auto_analysis as aa

# ---------------------------------------------------------------------------
# model glue - the only part that knows about the Auto Piling model.
# ---------------------------------------------------------------------------
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

RECORDS_CSV = aa.RECORDS_CSV                    # the run-records log (one row per trial)
PER_PILE = 5                                    # x, y, util_nw, util_w, tension
FIELD_NAMES = ("x", "y", "util_nw", "util_w", "tension")   # per pile, in that order


def evaluate(coords: list[list[float]]) -> list[list[float]]:
    """Push `coords` into the model, analyse it, record it, and return results.

    Returns one ``(util_nw, util_w, tension)`` triple per pile, in pile order.
    The trial row is appended to the records CSV right after the analysis, so
    every analysed candidate is on disk before the next one starts.
    """
    aa.write_coords(coords)
    aa.run_analysis()
    aa.write_reactions()
    utilization = aa.get_utilization()
    aa.write_test_record()                      # each trial is recorded immediately
    return utilization
# ---------------------------------------------------------------------------
# end model glue
# ---------------------------------------------------------------------------


def read_records(path: Path = RECORDS_CSV) -> list[list[list[float]]]:
    """Read every row of the records CSV as per-pile value groups.

    The header fixes the pile count and the field order, so both come from the
    data rather than from an assumption. Each row comes back as one
    ``[x, y, util_nw, util_w, tension]`` list per pile. A row of the wrong width
    or a cell that is not a number is refused rather than guessed at.
    """
    with open(path, newline="", encoding="utf-8") as csv_file:
        rows = list(csv.reader(csv_file))
    if not rows:
        raise RuntimeError(f"'{path}' has no header row")

    header = rows[0]
    if not header or len(header) % PER_PILE:
        raise ValueError(
            f"'{path.name}' header has {len(header)} column(s), not a positive "
            f"multiple of {PER_PILE} - cannot work out the piles"
        )
    for i, name in enumerate(header):
        if name.rsplit("~", 1)[-1] != FIELD_NAMES[i % PER_PILE]:
            raise ValueError(
                f"'{path.name}' column {i + 1} is '{name}', expected a "
                f"'~{FIELD_NAMES[i % PER_PILE]}' column - refusing to read it"
            )
    n_piles = len(header) // PER_PILE

    records: list[list[list[float]]] = []
    for row_number, row in enumerate(rows[1:], start=2):
        if not row or all(cell.strip() == "" for cell in row):
            continue
        if len(row) != len(header):
            raise ValueError(
                f"'{path.name}' row {row_number} has {len(row)} value(s) but the "
                f"header has {len(header)} - refusing to read a misaligned row"
            )
        values: list[float] = []
        for cell in row:
            try:
                values.append(float(cell))
            except ValueError as exc:
                raise ValueError(
                    f"'{path.name}' row {row_number} has a non-numeric value "
                    f"{cell!r}"
                ) from exc
        records.append([values[PER_PILE * p:PER_PILE * (p + 1)]
                        for p in range(n_piles)])
    if not records:
        raise RuntimeError(f"'{path}' holds no records to optimize from")
    return records


def record_coords(record: list[list[float]]) -> list[list[float]]:
    """The (x, y) coordinates of one records-CSV row."""
    return [[pile[0], pile[1]] for pile in record]


def record_utilization(record: list[list[float]]) -> list[list[float]]:
    """The (util_nw, util_w, tension) triples of one records-CSV row."""
    return [[pile[2], pile[3], pile[4]] for pile in record]


def record_peak(record: list[list[float]]) -> float:
    """Highest governing utilization of one records-CSV row."""
    return max(max(pile[2], pile[3]) for pile in record)


def best_record(records: list[list[list[float]]]) -> list[list[float]]:
    """The record with the lowest highest-utilization (ties keep the earliest)."""
    return min(records, key=record_peak)


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
             step_size: float = STEP_SIZE,
             directions: Sequence[tuple[int, int]] = DIRECTIONS,
             evaluate_fn: Callable[[list[list[float]]], list[list[float]]] = evaluate,
             read_fn: Callable[[], list[list[list[float]]]] = read_records,
             max_iterations: int | None = None,
             log: Callable[[str], None] = print,
             ) -> tuple[list[list[float]], list[list[float]], int]:
    """Hill-climb the pile layout, one pile and one step per iteration.

    Every iteration starts from the best row the records CSV holds and finishes
    by reading that CSV back, so the accepted layout never has to be re-analysed
    to register a baseline. Returns the best ``(coords, utilization,
    iterations)``; raises `NoImprovementError` when an iteration's eight trials
    fail to beat the baseline.
    """
    iterations = 0

    while max_iterations is None or iterations < max_iterations:
        baseline = best_record(read_fn())
        coords = record_coords(baseline)
        utilization = record_utilization(baseline)
        current_peak = record_peak(baseline)
        pile_index = max(range(len(coords)),
                         key=governing(utilization).__getitem__)
        log(f"iteration {iterations + 1}: baseline highest utilization "
            f"{current_peak:.6g}, pile {pile_index + 1} is the worst")

        for direction in directions:
            trial_coords = move_pile(coords, pile_index, direction, step_size)
            trial_peak = peak(evaluate_fn(trial_coords))
            log(f"  {direction}: highest utilization {trial_peak:.6g}")

        accepted = best_record(read_fn())                # the CSV is the memory
        if record_peak(accepted) >= current_peak:
            raise NoImprovementError(pile_index, record_coords(accepted),
                                     record_utilization(accepted), iterations,
                                     record_peak(accepted))
        iterations += 1
        log(f"  accepted the best of the eight -> highest utilization "
            f"{record_peak(accepted):.6g}")

    best = best_record(read_fn())
    return record_coords(best), record_utilization(best), iterations


def main() -> None:
    """Optimize from the best layout the records CSV already holds."""
    baseline = best_record(read_records())
    print(f"{len(baseline)} piles, best recorded highest utilization "
          f"{record_peak(baseline):.6g}, step size {STEP_SIZE:g}")

    try:
        with aa.quiet_excel():                       # fewer Excel messages: fewer OLE prompts
            coords, utilization, iterations = optimize()
    except NoImprovementError as exc:
        print(exc)
        print(f"stopped after {exc.iterations} accepted move(s) at highest "
              f"utilization {peak(exc.utilization):.6g}")
        raise

    print(f"finished after {iterations} accepted move(s): "
          f"highest utilization {peak(utilization):.6g}")


if __name__ == "__main__":
    main()
