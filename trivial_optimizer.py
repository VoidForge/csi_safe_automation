"""Greedy pile-layout optimizer for the Auto Piling workbook.

Every iteration moves ONE pile and keeps the best of eight trial moves:

  1. Pick the pile with the highest governing utilization, where a pile's
     governing utilization is the worse of its two cases:
     ``max(utilization without wind, utilization with wind)``.
  2. Move that pile by ``STEP_SIZE`` in each of the eight compass directions
     and analyse every candidate. A trial is appended to the run-records CSV
     the moment it has been analysed (`auto_analysis.write_test_record`), so an
     interrupted run keeps every trial it had reached.
  3. Keep the candidate whose highest utilization across ALL piles is lowest.
     If none of the eight beats the configuration it started from, raise
     `NoImprovementError` and stop - the layout is a local optimum at this step.

The number of piles is inferred from the pile-prefix column of 'Pile Coords',
so nothing here assumes 22 piles. All model-specific knowledge is confined to
the "model glue" block below - swapping that block (STEP_SIZE, DIRECTIONS,
read_baseline, evaluate) is all it takes to point the optimizer at another
model.

Requirements: the workbook must already be open in Excel, SAFE must be running
with the model loaded, and the 'Pile Coords' sheet must hold the results of the
layout it currently shows (that is what `read_baseline` reads).

Run it with:  python trivial_optimizer.py
"""

from __future__ import annotations

from typing import Callable, Sequence

import xlwings as xw

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

PILE_SHEET = "Pile Coords"                      # A = pile prefix, B:C = x, y
FIRST_ROW = 3                                   # first pile row in that sheet


def _sheet() -> xw.Sheet:
    """Return the already-open workbook's 'Pile Coords' sheet."""
    for app in xw.apps:
        for book in app.books:
            if book.name.lower() == aa.WORKBOOK.lower():
                return book.sheets[PILE_SHEET]
    raise RuntimeError(
        f"workbook '{aa.WORKBOOK}' is not open in Excel - open it and try again"
    )


def _rows(values: object, n: int) -> list[list[float]]:
    """Normalize a multi-row xlwings read into a list of row lists of floats.

    xlwings returns a flat list when the range is a single row, so that case is
    wrapped here; every cell must read as a number.
    """
    if n == 1:
        values = [values]
    return [[float(cell) for cell in row] for row in values]      # type: ignore[union-attr]


def read_baseline() -> tuple[list[list[float]], list[list[float]]]:
    """Read the sheet's current (x, y) coordinates and utilization triples.

    The number of piles comes from the pile-prefix column, so a model with more
    or fewer piles needs no change here. The utilization columns (H:J) are read
    as they stand, i.e. the sheet must hold the results of the layout shown in
    B:C.
    """
    sheet = _sheet()
    prefixes = sheet.range(f"A{FIRST_ROW}").expand("down").value
    prefixes = prefixes if isinstance(prefixes, list) else [prefixes]
    if not prefixes or any(p is None or not str(p).strip() for p in prefixes):
        raise RuntimeError(
            f"no pile prefixes found in '{PILE_SHEET}'!A{FIRST_ROW} downwards - "
            "the coordinate table is expected in A (prefix), B (x), C (y)"
        )
    n = len(prefixes)

    last = FIRST_ROW + n - 1
    coords = _rows(sheet.range(f"B{FIRST_ROW}:C{last}").value, n)
    utilization = _rows(sheet.range(f"H{FIRST_ROW}:J{last}").value, n)   # nw, w, tension
    return coords, utilization


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
    """None of the eight trial moves beat the layout the iteration started from.

    Carries the state the optimizer had reached (``coords``, ``utilization``,
    ``pile_index``, ``direction``, ``iterations``) so the caller can carry on
    from the best layout found.
    """

    def __init__(self, pile_index: int, direction: tuple[int, int],
                 coords: list[list[float]], utilization: list[list[float]],
                 iterations: int, best_peak: float) -> None:
        self.pile_index = pile_index
        self.direction = direction
        self.coords = coords
        self.utilization = utilization
        self.iterations = iterations
        self.best_peak = best_peak
        super().__init__(
            f"no trial move improved the layout after {iterations} accepted "
            f"move(s): moving pile {pile_index + 1} in any of the eight "
            f"directions left the highest utilization at or above "
            f"{peak(utilization):.6g} (best trial {best_peak:.6g}). The layout "
            "is a local optimum for this step size."
        )


def optimize(start_coords: list[list[float]] | None = None,
             start_utilization: list[list[float]] | None = None, *,
             step_size: float = STEP_SIZE,
             directions: Sequence[tuple[int, int]] = DIRECTIONS,
             evaluate_fn: Callable[[list[list[float]]], list[list[float]]] = evaluate,
             max_iterations: int | None = None,
             log: Callable[[str], None] = print,
             ) -> tuple[list[list[float]], list[list[float]], int]:
    """Hill-climb the pile layout, one pile and one step per iteration.

    Returns the accepted ``(coords, utilization, iterations)``. Raises
    `NoImprovementError` when no trial move improves the layout. The starting
    layout is read from the workbook when it is not supplied.
    """
    if start_coords is None or start_utilization is None:
        start_coords, start_utilization = read_baseline()
    coords = [list(pair) for pair in start_coords]
    utilization = start_utilization
    iterations = 0

    while max_iterations is None or iterations < max_iterations:
        worst = governing(utilization)
        current_peak = max(worst)
        pile_index = max(range(len(coords)), key=worst.__getitem__)
        log(f"iteration {iterations + 1}: pile {pile_index + 1} is the worst "
            f"(governing utilization {current_peak:.6g})")

        best_peak = float("inf")
        best_direction = directions[0]
        best_coords = coords
        for direction in directions:
            trial_coords = move_pile(coords, pile_index, direction, step_size)
            trial_utilization = evaluate_fn(trial_coords)
            trial_peak = peak(trial_utilization)
            log(f"  {direction}: highest utilization {trial_peak:.6g}")
            if trial_peak < best_peak:                   # ties keep the first direction
                best_peak = trial_peak
                best_direction = direction
                best_coords = trial_coords

        if best_peak >= current_peak:
            raise NoImprovementError(pile_index, best_direction, coords,
                                     utilization, iterations, best_peak)

        # The last trial analysed is not necessarily the best one, so leave the
        # model - and the CSV's last row - on the accepted layout.
        coords = best_coords
        utilization = evaluate_fn(coords)
        if peak(utilization) >= current_peak:            # analysis disagreed with the trial
            raise NoImprovementError(pile_index, best_direction, coords,
                                     utilization, iterations, peak(utilization))
        iterations += 1
        log(f"  accepted {best_direction} -> highest utilization "
            f"{peak(utilization):.6g}")

    return coords, utilization, iterations


def main() -> None:
    """Optimize the workbook's current layout, recording every trial."""
    coords, utilization = read_baseline()
    print(f"{len(coords)} piles, baseline highest utilization {peak(utilization):.6g}, "
          f"step size {STEP_SIZE:g}")

    try:
        coords, utilization, iterations = optimize(coords, utilization)
    except NoImprovementError as exc:
        print(exc)
        print(f"best layout after {exc.iterations} accepted move(s): "
              f"highest utilization {peak(exc.utilization):.6g}")
        raise

    print(f"finished after {iterations} accepted move(s): "
          f"highest utilization {peak(utilization):.6g}")


if __name__ == "__main__":
    main()
