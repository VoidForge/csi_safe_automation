"""Greedy pile-layout optimizer for the Auto Piling workbook.

Every iteration moves one pile and keeps the best of eight trial moves:

  1. Pick the pile with the highest governing utilization, where a pile's
     governing utilization is the worse of its two cases:
     ``max(utilization without wind, utilization with wind)``.
  2. Move that pile by ``STEP_SIZE`` in each of the eight compass directions. A
     candidate whose exact coordinates are already in the record store is
     skipped: its result is known, so no analysis runs and no row is appended.
     Every other candidate is analysed and its row is appended to `RECORDS_CSV`
     immediately (`record_store.RecordStore`), so an interrupted run keeps every
     trial it had reached.
  3. The `RecordStore` mirrors that CSV in memory: each trial's row is added to it
     as it is recorded, and the next baseline is its best row (lowest highest
     utilization). Nothing is re-read from disk during a run.
  4. If that best row is no better than the baseline the iteration started from,
     raise `NoImprovementError`: the layout is a local optimum at this step
     size. That is also what an iteration whose candidates were all tried before
     reports.

The pile count and pile order come from the records CSV header, so nothing here
assumes 22 piles. All model-specific knowledge is confined to the "model glue"
block below (`RECORDS_CSV`, `STEP_SIZE`, `DIRECTIONS` and `evaluate`), so
swapping that block is all it takes to point the optimizer at another model.

`record_store.py` holds the recorded rows, the readings taken from them and the
`RecordStore` itself. The `RECORDS_CSV` below is the path this module gives it.

A recorded row is only valid for the model it was measured against, so a model
edit or a changed STEP_SIZE makes older rows stale. They can be both the
baseline and the thing that marks a candidate "already tried".

Requirements: the workbook must already be open in Excel and SAFE must be
running with the model loaded.

Run it with:  python trivial_optimizer.py
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Sequence

import auto_analysis as aa
from record_store import (RecordStore, record_coords, record_max_util,
                          record_utilization)

# ---------------------------------------------------------------------------
# model glue: the only part that knows about the Auto Piling model.
# ---------------------------------------------------------------------------
# RECORDS_CSV: Path = Path(__file__).resolve().parent / "data" / "SWTKT_test_records.csv"
RECORDS_CSV: Path = Path(__file__).resolve().parent / "data" / "SWTKT_test_records_2026-10-11_02.csv"
                                # the run-records log: read back and appended to

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


def evaluate(coords: list[list[float]]) -> list[float]:
    """Push `coords` into the model, analyse it, and record the trial.

    Returns the flat row just appended to `RECORDS_CSV`: five numbers per pile,
    in the CSV's own order. The utilization comes from that row instead of a
    second sheet read, so a trial costs one Excel read. The caller puts the row
    into the `RecordStore`.

    `record_store` does the append and checks the row against the file's own
    header. This glue has no store, so it loads the file for that.
    """
    aa.write_coords(coords)
    aa.run_analysis()
    aa.write_reactions()
    prefixes, values = aa.read_pile_values()    # one Excel round trip for the table
    RecordStore.load(RECORDS_CSV).append(values, prefixes=prefixes)   # durable row
    return values
# ---------------------------------------------------------------------------
# end model glue
# ---------------------------------------------------------------------------


def governing(utilization: list[list[float]]) -> list[float]:
    """Per pile, the worse of the no-wind and with-wind utilization."""
    return [max(triple[0], triple[1]) for triple in utilization]


def peak(utilization: list[list[float]]) -> float:
    """Highest governing utilization across every pile: the objective."""
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
    ``iterations``) so the caller can carry on from it, and ``best_trial_peak``:
    the lowest highest-utilization that iteration's trials reached, which is the
    baseline's own value when the iteration ran no trial at all.
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
    the iteration's candidates all fail to beat the baseline, carrying the best
    result those candidates reached.
    """
    if store is None:
        store = RecordStore.load(RECORDS_CSV)
    iterations = 0

    while max_iterations is None or iterations < max_iterations:
        baseline = store.best()
        coords = record_coords(baseline)
        utilization = record_utilization(baseline)
        current_peak = record_max_util(baseline)
        pile_index = max(range(len(coords)),
                         key=governing(utilization).__getitem__)
        log(f"iteration {iterations + 1}: baseline highest utilization "
            f"{current_peak:.6g}, pile {pile_index + 1} is the worst")

        best_trial_peak: float | None = None     # best result any trial of this iteration reached
        for direction in directions:
            trial_coords = move_pile(coords, pile_index, direction, step_size)
            trial = store.find(trial_coords)
            if trial is not None:
                log(f"  {direction}: already tried (highest utilization "
                    f"{record_max_util(trial):.6g}), skipped")
            else:
                trial = store.add(evaluate_fn(trial_coords))
                log(f"  {direction}: highest utilization {record_max_util(trial):.6g}")
            trial_peak = record_max_util(trial)
            if best_trial_peak is None or trial_peak < best_trial_peak:
                best_trial_peak = trial_peak

        accepted = store.best()                          # the store is the memory
        if record_max_util(accepted) >= current_peak:
            raise NoImprovementError(
                pile_index, record_coords(accepted), record_utilization(accepted),
                iterations,
                current_peak if best_trial_peak is None else best_trial_peak,
            )
        iterations += 1
        log(f"  accepted the best of the eight -> highest utilization "
            f"{record_max_util(accepted):.6g}")

    best = store.best()
    return record_coords(best), record_utilization(best), iterations


def main() -> None:
    """Optimize from the best layout the records CSV already holds."""
    # ---- initial pile layout: the config init_pile_layout() reads ----------
    CAP_POLYGON: list[list[float]] = [[0.0,0.0],[56600,0],[56600,101000],[0,101000]]         
    # convex cap outline in (x, y) model
    # units, counter-clockwise, M >= 3.
    # Empty: no layout is generated
    PILE_WEIGHTS = [6.25,	9,	10.89,	10.89,	9,	10.89,	9,	9,	9,	10.89,	10.89,	9,	9,	9,	9,	9,	6.25,	9,	6.25,	6.25,	6.25,	6.25]

    # one target area weight per pile;
    # None gives every pile the same area
    N_PILES: int | None = None                  
    # pile count; None takes the count from
    # 'Pile Coords'!A, and either way it has
    # to match that prefix column
    MARGIN = 3500

    # ---- write that layout into the sheet and into SAFE -------------------
    init_coords: list[list[float]]  = aa.init_pile_layout(CAP_POLYGON, PILE_WEIGHTS, n_piles=N_PILES, seed=42, best_of=7, margin=MARGIN)
    print(init_coords)
    store = RecordStore.load(RECORDS_CSV)

    try:
        with aa.quiet_excel():
            if not len(store):
                print("records CSV is header-only: seeding the starting layout")
                store.add(evaluate(init_coords))

            print(f"{RECORDS_CSV.name}: {len(store.names)} piles, {len(store)} recorded "
                f"layout(s), {store.duplicate_keys} repeated, best highest utilization "
                f"{record_max_util(store.best()):.6g}, step size {STEP_SIZE:g}")

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
