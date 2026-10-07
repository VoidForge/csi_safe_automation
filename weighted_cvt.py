"""Prescribed-area (capacity-constrained) CVT through power / Laguerre diagrams.

Each site carries an additive weight and owns the set of points closer to it in the
power metric ``|x - p_i|^2 - w_i``. The weights are solved so that the cells come out in
the requested area ratio, and every site is relaxed onto its own cell's centroid.

The caller's area weight is a target, and ``A_i / A = area_weights[i] /
sum(area_weights)``. The Laguerre weights that realise it are solved internally and come
back in ``WeightedCVT.weights``; they follow the requested ratios monotonically, though
they are not proportional to them.

Public interface: `weighted_cvt` to generate a layout, `best_weighted_cvt` to pick the
best of several starts, and `power_cells` / `cell_areas` / `validate_weighted` to
inspect one. numpy and scipy only; the geometry primitives come from `lloyd_cvt`.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import NamedTuple, TypeAlias

import numpy as np
from scipy.spatial import ConvexHull, QhullError

from lloyd_cvt import Points, _area, _centroid, _clip, _inside, _ring, lloyd_cvt

# One number per site: the area targets a caller passes in, or the Laguerre weights
# that come back. Like Points, this accepts a plain sequence or a numpy array.
Values: TypeAlias = Sequence[float] | np.ndarray

__all__ = [
    "WeightedCVT",
    "weighted_cvt",
    "best_weighted_cvt",
    "power_cells",
    "cell_areas",
    "validate_weighted",
]


class WeightedCVT(NamedTuple):
    """Result of :func:`weighted_cvt`.

    ``sites``, ``weights``, ``areas`` and ``targets`` are parallel lists, all sorted by
    descending y and then ascending x. ``weights`` are the solved Laguerre potentials
    (gauge-fixed to sum to zero), ``areas`` the realised cell areas, and ``targets`` the
    areas those ratios ask for. ``max_area_error`` is
    ``max |areas - targets| / total area`` and ``converged`` records whether the run
    reached the requested ``area_tol``.
    """

    sites: list[tuple[float, float]]
    weights: list[float]
    areas: list[float]
    targets: list[float]
    iterations: int
    max_area_error: float
    converged: bool


def _is_convex(poly: np.ndarray, diag: float) -> bool:
    """Whether the ring `poly` is convex (any winding)."""
    edge = np.roll(poly, -1, axis=0) - poly
    cross = edge[:, 0] * np.roll(edge[:, 1], -1) - edge[:, 1] * np.roll(edge[:, 0], -1)
    tol = 1e-9 * diag * diag
    return bool((cross >= -tol).all() or (cross <= tol).all())


def _lifted_neighbours(sites: np.ndarray, lam: np.ndarray) -> list[tuple[int, int]] | None:
    """Edges of the regular (weighted Delaunay) triangulation, or None if degenerate.

    Two power cells share a face exactly when their sites are joined in the regular
    triangulation, which is the projection of the lower hull of the sites lifted onto
    ``z = |p|^2 - w``. Those are the only bisectors that need clipping.
    """
    if len(sites) < 4:
        return None
    lifted = np.column_stack((sites, (sites**2).sum(axis=1) - lam))
    try:
        hull = ConvexHull(lifted, qhull_options="QJ")  # QJ breaks cocircular ties
    except (QhullError, ValueError):
        return None
    pairs: set[tuple[int, int]] = set()
    for simplex, equation in zip(hull.simplices, hull.equations):
        if equation[2] >= 0.0:  # keep the lower hull (outward normal pointing down)
            continue
        a, b, c = (int(simplex[0]), int(simplex[1]), int(simplex[2]))
        for i, j in ((a, b), (b, c), (c, a)):
            pairs.add((i, j) if i < j else (j, i))
    return sorted(pairs) or None


def _neighbours(sites: np.ndarray, lam: np.ndarray, mode: str) -> tuple[list[set[int]], bool]:
    """For every site, the set of sites it must be clipped against, and whether that set
    came from the regular triangulation. A site with no triangulation edge is dominated,
    so its power cell is empty rather than the whole polygon."""
    n = len(sites)
    if mode == "all" or n < 4:
        return [set(range(n)) - {i} for i in range(n)], False
    pairs = _lifted_neighbours(sites, lam)
    if pairs is None:
        if mode == "hull":
            raise RuntimeError("the regular triangulation could not be built")
        return [set(range(n)) - {i} for i in range(n)], False
    out = [set() for _ in range(n)]
    for i, j in pairs:
        out[i].add(j)
        out[j].add(i)
    return out, True


def _power_cells(
    poly: np.ndarray, sites: np.ndarray, lam: np.ndarray, mode: str = "auto"
) -> list[np.ndarray]:
    """Power cell of every site, clipped to `poly` (empty array if it collapsed)."""
    diag = float(np.hypot(*np.ptp(poly, axis=0)))
    neighbours, from_hull = _neighbours(sites, lam, mode)
    cells = []
    for i, others in enumerate(neighbours):
        if from_hull and not others:  # dominated site: the cell is empty
            cells.append(np.empty((0, 2)))
            continue
        cell = poly
        for j in sorted(others):
            d = sites[j] - sites[i]
            if d @ d <= (1e-12 * diag) ** 2:  # coincident sites have no bisector
                continue
            # cell i keeps (x - mid).d <= (lam[i] - lam[j]) / 2; a larger lam grows it
            s = (cell - 0.5 * (sites[i] + sites[j])) @ d - 0.5 * (lam[i] - lam[j])
            if s.max() <= 0.0:  # bisector misses the cell
                continue
            cell = _clip(cell, s)
            if len(cell) < 3:
                break
        cells.append(cell if len(cell) >= 3 else np.empty((0, 2)))
    return cells


def _jacobian(cells: list[np.ndarray], sites: np.ndarray, lam: np.ndarray, diag: float) -> np.ndarray:
    """d A_i / d lam_j for the current cells.

    Off the diagonal this is ``-|e_ij| / (2 |p_i - p_j|)``, the area a shared face gains
    or loses per unit of weight, and the diagonal is minus its row sum because the areas
    always add up to the polygon. The matrix is symmetric and singular along the
    constants: shifting every weight leaves the diagram unchanged.
    """
    n = len(sites)
    jac = np.zeros((n, n))
    p_sq = (sites**2).sum(axis=1)
    fac = 1e-9 * diag * diag
    seen: set[tuple[int, int]] = set()
    for i, cell in enumerate(cells):
        if len(cell) < 3:
            continue
        nxt = np.roll(cell, -1, axis=0)
        mid = 0.5 * (cell + nxt)
        length = np.hypot(*(cell - nxt).T)
        # power distance from every site to every edge midpoint; a face's pair ties
        power = ((mid[:, None, :] - sites[None, :, :]) ** 2).sum(axis=2) - lam[None, :]
        for e in range(len(mid)):
            row = power[e]
            partner: int | None = None
            for idx in np.argsort(row):
                idx = int(idx)
                if idx == i:
                    continue
                if abs(row[idx] - row[i]) <= fac:
                    partner = idx
                break
            if partner is None:  # an edge of the original polygon, no counterpart
                continue
            key = (i, partner) if i < partner else (partner, i)
            if key in seen:
                continue
            seen.add(key)
            dist = float(np.linalg.norm(sites[i] - sites[partner]))
            if dist <= 1e-12 * diag:
                continue
            wgt = float(length[e]) / (2.0 * dist)
            jac[i, partner] -= wgt
            jac[partner, i] -= wgt
            jac[i, i] += wgt
            jac[partner, partner] += wgt
    return jac


def _solve_weights(
    poly: np.ndarray,
    sites: np.ndarray,
    lam: np.ndarray,
    targets: np.ndarray,
    area_tol: float,
    max_iter: int,
    mode: str,
) -> tuple[np.ndarray, list[np.ndarray], np.ndarray]:
    """Damped Newton on the Laguerre weights until the cell areas hit `targets`.

    A site whose cell has been squeezed to nothing has no faces, so its Jacobian row is
    empty and a plain Newton step leaves it there. When no undamped step helps, the solve
    escalates a Levenberg damping term, which gives every empty row a step that grows the
    cell back. The damping changes only the step size, so the areas still converge to
    `targets`.
    """
    diag = float(np.hypot(*np.ptp(poly, axis=0)))
    total = float(targets.sum())
    lam = np.asarray(lam, dtype=float)
    lam = lam - float(lam.mean())
    cells = _power_cells(poly, sites, lam, mode)
    areas = np.array([_area(c) for c in cells])
    for _ in range(max(1, max_iter)):
        if float(np.abs(areas - targets).max()) <= area_tol * total:
            break
        residual = areas - targets
        base = float(np.linalg.norm(residual))
        jac = _jacobian(cells, sites, lam, diag)
        jscale = max(float(np.abs(jac).max()), 1e-12)
        improved = False
        for mult in (0.0, 1e-2, 1e-1):  # plain Newton first, then Levenberg damping
            matrix = jac if mult == 0.0 else jac + mult * jscale * np.eye(len(lam))
            try:
                step = np.linalg.lstsq(matrix, -residual, rcond=None)[0]
            except np.linalg.LinAlgError:
                continue
            if not np.isfinite(step).all() or not np.any(step != 0.0):
                continue
            shrink = 1.0
            for _ in range(20):  # backtrack until the step helps and nothing collapses
                trial = lam + shrink * step
                trial = trial - float(trial.mean())
                trial_cells = _power_cells(poly, sites, trial, mode)
                trial_areas = np.array([_area(c) for c in trial_cells])
                if trial_areas.min() > 0.0 and float(np.linalg.norm(trial_areas - targets)) < base:
                    lam, cells, areas = trial, trial_cells, trial_areas
                    improved = True
                    break
                shrink *= 0.5
            if improved:
                break
        if not improved:
            break  # no damping level helped, so let the outer loop decide
    return lam, cells, areas


def _centroid_error(poly: np.ndarray, sites: np.ndarray, lam: np.ndarray) -> float:
    """Worst |site - own cell centroid| as a fraction of the polygon's diagonal."""
    diag = float(np.hypot(*np.ptp(poly, axis=0)))
    cells = _power_cells(poly, sites, lam, "auto")
    worst = max(
        (
            float(np.linalg.norm(_centroid(cell, sites[i]) - sites[i]))
            for i, cell in enumerate(cells)
            if _area(cell) > 0.0
        ),
        default=0.0,
    )
    return worst / diag


def weighted_cvt(
    polygon: Points,
    area_weights: Values,
    *,
    seed: int | None = None,
    tol: float = 1e-9,
    area_tol: float = 1e-3,
    max_iter: int = 200,
    weight_iter: int = 50,
    damping: float = 0.5,
    neighbour_mode: str = "auto",
) -> WeightedCVT:
    """Tessellate the convex `polygon` so cell areas follow `area_weights`.

    polygon:        (M, 2) boundary coordinates, M >= 3, counter-clockwise. A clockwise
                    ring is re-oriented and a repeated closing vertex is dropped. The
                    polygon must be convex.
    area_weights:   one positive weight per site; cell `i` is driven towards
                    ``polygon area * area_weights[i] / sum(area_weights)``. The number of
                    sites is the length of this sequence.
    seed:           None -> the deterministic equal-area CVT is the start; an int ->
                    a random start from that seed. Either way the same input gives the
                    same output.
    tol:            stop once the largest site move in an outer step falls below `tol`
                    times the polygon's bounding-box diagonal.
    area_tol:       stop once the largest relative area error falls below this.
    max_iter:       cap on the outer (weights then Lloyd) iterations.
    weight_iter:    cap on the inner damped-Newton weight solves.
    damping:        fraction of the way to its cell centroid a site moves each outer
                    step. Full steps make the weight solve and the site move fight and
                    collapse cells, so the default stops short of 1.
    neighbour_mode: "auto" (regular triangulation, falling back to all pairs), "hull"
                    (regular triangulation only) or "all" (clip against every site).

    Returns a :class:`WeightedCVT`. Every site sits on its own cell's centroid and the
    areas match `area_weights` to `area_tol` unless `converged` says otherwise.
    """
    poly, poly_area2 = _ring(polygon)
    total = poly_area2 / 2.0
    diag = float(np.hypot(*np.ptp(poly, axis=0)))
    if not _is_convex(poly, diag):
        raise ValueError("polygon must be convex")

    weights = np.asarray(area_weights, dtype=float)
    if weights.ndim != 1 or weights.size == 0:
        raise ValueError("area_weights must be a 1-D sequence of positive numbers")
    if not np.isfinite(weights).all() or (weights <= 0.0).any():
        raise ValueError("area_weights must all be positive and finite")
    n = int(weights.size)
    targets = total * weights / weights.sum()

    sites = np.asarray(lloyd_cvt(poly, n, seed=seed), dtype=float).reshape(n, 2)
    eps = tol * diag
    lam = np.zeros(n)
    cells: list[np.ndarray] = []
    areas = np.zeros(n)
    iterations = 0
    for iterations in range(1, max_iter + 1):
        lam, cells, areas = _solve_weights(poly, sites, lam, targets, area_tol, weight_iter, neighbour_mode)
        shifted = sites.copy()
        for i, cell in enumerate(cells):
            if _area(cell) <= 1e-12 * total:  # collapsed cell has no centroid to chase
                continue
            shifted[i] = _centroid(cell, sites[i])
        delta = shifted - sites
        moved = float(np.hypot(*delta.T).max())  # full distance to the centroid, before damping
        sites = sites + damping * delta
        if moved <= eps and float(np.abs(areas - targets).max()) <= area_tol * total:
            break

    lam, cells, areas = _solve_weights(poly, sites, lam, targets, area_tol, weight_iter, neighbour_mode)
    error = float(np.abs(areas - targets).max() / total)
    order = sorted(range(n), key=lambda i: (-sites[i, 1], sites[i, 0]))
    return WeightedCVT(
        sites=[(float(sites[i, 0]), float(sites[i, 1])) for i in order],
        weights=[float(lam[i]) for i in order],
        areas=[float(areas[i]) for i in order],
        targets=[float(targets[i]) for i in order],
        iterations=int(iterations),
        max_area_error=error,
        converged=bool(error <= area_tol),
    )


def best_weighted_cvt(
    polygon: Points,
    area_weights: Values,
    k: int = 8,
    *,
    criterion: str = "area_error",
    seed: int | None = None,
    max_iter: int = 200,
    weight_iter: int = 50,
    damping: float = 0.5,
    neighbour_mode: str = "auto",
    area_tol: float = 1e-3,
) -> WeightedCVT:
    """Run the deterministic start plus `k` seeded starts and keep the best result.

    `criterion` is "area_error" (smallest `max_area_error`, default) or "centroid"
    (smallest worst |site - cell centroid| relative to the diagonal).
    """
    if k < 1:
        raise ValueError("k must be >= 1")
    if criterion not in ("area_error", "centroid"):
        raise ValueError("criterion must be 'area_error' or 'centroid'")
    poly, _ = _ring(polygon)
    seeds: list[int | None] = [seed]
    if seed is None:
        seeds.extend(range(k))
    else:
        seeds.extend(s for s in range(k) if s != seed)
    best: WeightedCVT | None = None
    best_score = np.inf
    for start in seeds:
        result = weighted_cvt(
            polygon,
            area_weights,
            seed=start,
            area_tol=area_tol,
            max_iter=max_iter,
            weight_iter=weight_iter,
            damping=damping,
            neighbour_mode=neighbour_mode,
        )
        score = (
            result.max_area_error
            if criterion == "area_error"
            else _centroid_error(poly, np.asarray(result.sites), np.asarray(result.weights))
        )
        if score < best_score:
            best, best_score = result, score
    assert best is not None  # seeds is never empty, so the loop always assigns
    return best


def power_cells(
    polygon: Points,
    sites: Points,
    weights: Values | None = None,
) -> list[np.ndarray]:
    """Power cell of every site, clipped to the polygon (empty array if it collapsed)."""
    poly, _ = _ring(polygon)
    pts = np.asarray(sites, dtype=float)
    if pts.ndim != 2 or pts.shape[1] != 2 or len(pts) == 0:
        raise ValueError("sites must be an (n, 2) array of coords")
    lam = np.zeros(len(pts)) if weights is None else np.asarray(weights, dtype=float)
    if lam.shape != (len(pts),):
        raise ValueError("weights must have one entry per site")
    return _power_cells(poly, pts, lam, "auto")


def cell_areas(
    polygon: Points,
    sites: Points,
    weights: Values | None = None,
) -> list[float]:
    """Area of each power cell, in the order the sites were given."""
    return [_area(cell) for cell in power_cells(polygon, sites, weights)]


def validate_weighted(
    polygon: Points,
    sites: Points,
    weights: Values,
    area_weights: Values,
    tol: float = 1e-3,
    area_tol: float = 1e-2,
) -> list[str]:
    """List what is wrong with a weighted tessellation; an empty list means it is fine.

    Checks the site count, that the sites are distinct and inside the polygon, that no
    cell collapsed, that the cells tile the polygon, that each site sits within `tol`
    (times the diagonal) of its own cell's centroid, and that the areas are within
    `area_tol` of the ratio `area_weights` asks for. `weights` are the Laguerre
    potentials, as returned in :attr:`WeightedCVT.weights`.
    """
    poly, poly_area2 = _ring(polygon)
    total = poly_area2 / 2.0
    diag = float(np.hypot(*np.ptp(poly, axis=0)))
    pts = np.asarray(sites, dtype=float)
    if pts.ndim != 2 or pts.shape[1] != 2 or len(pts) == 0:
        return [f"sites must be an (n, 2) array of coords, got shape {pts.shape}"]
    n = len(pts)
    lam = np.asarray(weights, dtype=float)
    ratios = np.asarray(area_weights, dtype=float)
    if lam.shape != (n,) or ratios.shape != (n,):
        return [f"weights and area_weights must each have {n} entries"]

    bad: list[str] = []
    if not np.isfinite(ratios).all() or (ratios <= 0.0).any():
        bad.append("area_weights must all be positive and finite")
    if not _inside(poly, pts).all():
        bad.append("site(s) outside the polygon")
    if n > 1:
        gap = np.hypot(pts[:, 0, None] - pts[None, :, 0], pts[:, 1, None] - pts[None, :, 1])
        np.fill_diagonal(gap, np.inf)
        if gap.min() <= 1e-9 * diag:
            bad.append(f"coincident sites, closest pair {gap.min():.6g} apart")

    cells = _power_cells(poly, pts, lam, "auto")
    areas = np.array([_area(c) for c in cells])
    alive = np.flatnonzero(areas > 1e-12 * total)
    if len(alive) < len(areas):
        dead = np.setdiff1d(np.arange(len(areas)), alive).tolist()
        bad.append(f"collapsed cell(s) at site(s) {dead}")
    if abs(areas.sum() - total) > 1e-9 * total:
        bad.append(f"cells do not tile the polygon: sum/area = {areas.sum() / total:.12f}")
    residual = max(
        (float(np.linalg.norm(_centroid(cells[i], pts[i]) - pts[i])) for i in alive),
        default=0.0,
    )
    if residual > tol * diag:
        bad.append(f"not centroidal: worst |site - cell centroid| = {residual:.6g}, limit {tol * diag:.6g}")
    targets = total * ratios / ratios.sum()
    error = float(np.abs(areas - targets).max() / total)
    if error > area_tol:
        bad.append(f"areas off target: worst |A - target|/total = {error:.6g}, limit {area_tol:.6g}")
    return bad
