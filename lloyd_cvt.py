"""Centroidal Voronoi tessellation of a convex polygon by Lloyd's algorithm.

Public interface: `lloyd_cvt` to generate a tessellation, `best_cvt` to pick the best of
several starts, and `validate` / `area_spread` / `energy` to inspect one. Uses numpy and
scipy (Qhull via `scipy.spatial.Voronoi`) only.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import TypeAlias

import numpy as np
from scipy.spatial import QhullError, Voronoi

# Points reach these functions either as the plain list of pairs a caller builds or as
# the numpy array an earlier step handed back. A numpy array is not a Sequence, so the
# public annotations have to name both.
Points: TypeAlias = Sequence[Sequence[float]] | np.ndarray


def _ring(polygon: Points) -> tuple[np.ndarray, float]:
    """Counter-clockwise boundary ring of `polygon` and its doubled area."""
    poly = np.asarray(polygon, dtype=float)
    if poly.ndim != 2 or poly.shape[1] != 2 or len(poly) < 3:
        raise ValueError("polygon must be an (M, 2) array of coords with M >= 3")
    if np.all(poly[0] == poly[-1]):
        poly = poly[:-1]
    x, y = poly[:, 0], poly[:, 1]
    area2 = float((x * np.roll(y, -1) - np.roll(x, -1) * y).sum())
    if area2 == 0.0:
        raise ValueError("polygon has zero area")
    if area2 < 0.0:
        poly, area2 = poly[::-1], -area2
    return poly, area2


def _area(cell: np.ndarray) -> float:
    """Unsigned area of a polygon ring, 0.0 if it has collapsed."""
    if len(cell) < 3:
        return 0.0
    local = cell - cell[0]  # cell-sized coords, or the shoelace sums cancel
    x, y = local[:, 0], local[:, 1]
    return abs(float((x * np.roll(y, -1) - np.roll(x, -1) * y).sum())) / 2.0


def _centroid(cell: np.ndarray, origin: np.ndarray) -> np.ndarray:
    """Area centroid of the convex `cell`, summed relative to `origin`."""
    local = cell - origin  # cell-sized coords, or the shoelace sums cancel
    nxt = np.roll(local, -1, axis=0)
    cross = local[:, 0] * nxt[:, 1] - nxt[:, 0] * local[:, 1]
    return origin + np.array(
        [
            ((local[:, 0] + nxt[:, 0]) * cross).sum(),
            ((local[:, 1] + nxt[:, 1]) * cross).sum(),
        ]
    ) / (3.0 * float(cross.sum()))


def _cells(poly: np.ndarray, sites: np.ndarray) -> list[np.ndarray]:
    """Voronoi cell of every site, clipped to `poly` (empty array if it collapsed)."""
    diag = float(np.hypot(*np.ptp(poly, axis=0)))
    try:
        ridge = Voronoi(sites, qhull_options="QJ").ridge_points
    except QhullError:  # degenerate layout (collinear / cocircular sites)
        ridge = np.array(np.triu_indices(len(sites), 1)).T
    neighbour: list[list[int]] = [[] for _ in sites]
    for i, j in ridge:
        neighbour[i].append(j)
        neighbour[j].append(i)

    cells = []
    for i, others in enumerate(neighbour):
        cell = poly
        for j in others:
            d = sites[j] - sites[i]
            if d @ d <= (1e-12 * diag) ** 2:  # coincident sites have no bisector
                continue
            s = (cell - 0.5 * (sites[i] + sites[j])) @ d
            if s.max() <= 0.0:  # bisector misses the cell
                continue
            cell = _clip(cell, s)
            if len(cell) < 3:
                break
        cells.append(cell if len(cell) >= 3 else np.empty((0, 2)))
    return cells


def _inside(poly: np.ndarray, pts: np.ndarray) -> np.ndarray:
    """Mask of the points in `pts` that are inside the CCW convex `poly`."""
    edge = np.roll(poly, -1, axis=0) - poly
    rel = pts[:, None, :] - poly[None, :, :]
    return (
        (edge[None, :, 0] * rel[:, :, 1] - edge[None, :, 1] * rel[:, :, 0]) >= 0.0
    ).all(axis=1)


def _clip(pts: np.ndarray, s: np.ndarray) -> np.ndarray:
    """Sutherland-Hodgman clip of the convex polygon `pts` to the half-plane `s <= 0`.

    `s[k]` is the half-plane value at `pts[k]`; winding is preserved, and a cell
    clipped away to nothing comes back with fewer than 3 points.
    """
    inside = s <= 0.0
    cut = inside != np.roll(inside, -1)
    den = s - np.roll(s, -1)
    t = np.divide(s, den, out=np.zeros_like(s), where=cut & (den != 0.0))
    exits = pts + t[:, None] * (np.roll(pts, -1, axis=0) - pts)
    slots = np.empty((len(pts), 2, 2))
    slots[:, 0] = pts
    slots[:, 1] = exits
    return slots.reshape(-1, 2)[np.column_stack((inside, cut)).ravel()]


def lloyd_cvt(
    polygon: Points,
    n: int,
    *,
    seed: int | None = None,
    tol: float = 1e-9,
    max_iter: int = 1000,
) -> list[tuple[float, float]]:
    """Partition the convex `polygon` into `n` centroidal Voronoi cells.

    polygon:  (M, 2) boundary coordinates, M >= 3, counter-clockwise. A clockwise
              ring is re-oriented and a repeated closing vertex is dropped.
    n:        number of subregions, >= 1.
    seed:     None -> deterministic hexagonal-lattice start; an int -> random start
              from that seed. Either way one input always gives one output.
    tol:      stop when the largest site move in an iteration falls below
              `tol` times the polygon's bounding-box diagonal.
    max_iter: iteration cap; if it is reached the current sites are returned as is.

    Returns the centroid of each subregion as (x, y) tuples, sorted by descending
    y and then ascending x.
    """
    poly, poly_area2 = _ring(polygon)
    if n < 1:
        raise ValueError("n must be >= 1")

    lo, hi = poly.min(axis=0), poly.max(axis=0)
    eps = tol * float(np.hypot(*(hi - lo)))

    if seed is None:
        # Hexagonal lattice spaced for ~1 point per cell area, thinned to exactly n.
        s = float(np.sqrt(poly_area2 / (n * np.sqrt(3.0))))
        sites = np.empty((0, 2))
        for _ in range(200):
            xs = np.arange(lo[0], hi[0] + s, s)
            rows = int((hi[1] - lo[1]) / (s * np.sqrt(3.0) / 2.0)) + 2
            grid = np.vstack(
                [
                    np.column_stack(
                        (
                            xs + (k % 2) * 0.5 * s,
                            np.full(xs.size, lo[1] + k * s * np.sqrt(3.0) / 2.0),
                        )
                    )
                    for k in range(rows)
                ]
            )
            sites = grid[_inside(poly, grid)]
            if len(sites) >= n:
                break
            s *= 0.95
        if len(sites) > n:  # drop the most crowded points first
            gap = np.hypot(
                sites[:, 0, None] - sites[None, :, 0],
                sites[:, 1, None] - sites[None, :, 1],
            )
            np.fill_diagonal(gap, np.inf)
            crowding = gap.min(axis=1)
            keep = np.lexsort((sites[:, 0], -sites[:, 1], crowding))[len(sites) - n :]
            sites = sites[keep]
    else:
        # Uniform over the polygon via its fan triangulation from vertex 0.
        tri = np.stack(
            [np.tile(poly[0], (len(poly) - 2, 1)), poly[1:-1], poly[2:]], axis=1
        )
        ab, ac = tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0]
        area = np.abs(ab[:, 0] * ac[:, 1] - ab[:, 1] * ac[:, 0])
        rng = np.random.default_rng(seed)
        pick = rng.choice(len(tri), size=n, p=area / area.sum())
        r, v = np.sqrt(rng.random(n)), rng.random(n)
        sites = (
            (1.0 - r)[:, None] * tri[pick, 0]
            + (r * (1.0 - v))[:, None] * tri[pick, 1]
            + (r * v)[:, None] * tri[pick, 2]
        )

    for _ in range(max_iter):
        shifted = sites.copy()
        for i, cell in enumerate(_cells(poly, sites)):
            if _area(cell) <= 1e-12 * (poly_area2 / 2.0):  # collapsed cell
                continue
            shifted[i] = _centroid(cell, sites[i])

        moved = float(np.hypot(*(shifted - sites).T).max())
        sites = shifted
        if moved < eps:
            break

    return sorted(
        ((float(p[0]), float(p[1])) for p in sites), key=lambda q: (-q[1], q[0])
    )


def validate(
    polygon: Points,
    sites: Points,
    tol: float = 1e-3,
) -> list[str]:
    """List what is wrong with `sites` as a tessellation of `polygon`; empty means fine.

    Checks the site count, that the sites are distinct and inside the polygon, that no
    cell collapsed, that the cells tile the polygon, and that no site is further than
    `tol` times the bounding-box diagonal from its own cell's centroid.
    """
    poly, poly_area2 = _ring(polygon)
    pts = np.asarray(sites, dtype=float)
    if pts.ndim != 2 or pts.shape[1] != 2 or len(pts) == 0:
        return [f"sites must be an (n, 2) array of coords, got shape {pts.shape}"]

    area = poly_area2 / 2.0
    diag = float(np.hypot(*np.ptp(poly, axis=0)))
    bad: list[str] = []
    if not _inside(poly, pts).all():
        bad.append("site(s) outside the polygon")
    if len(pts) > 1:
        gap = np.hypot(
            pts[:, 0, None] - pts[None, :, 0],
            pts[:, 1, None] - pts[None, :, 1],
        )
        np.fill_diagonal(gap, np.inf)
        if gap.min() <= 1e-9 * diag:
            bad.append(f"coincident sites, closest pair {gap.min():.6g} apart")

    cells = _cells(poly, pts)
    areas = np.array([_area(c) for c in cells])
    alive = np.flatnonzero(areas > 1e-12 * area)
    if len(alive) < len(areas):
        bad.append(f"collapsed cell(s) at site(s) {np.setdiff1d(np.arange(len(areas)), alive).tolist()}")
    if abs(areas.sum() - area) > 1e-9 * area:
        bad.append(f"cells do not tile the polygon: sum/area = {areas.sum() / area:.12f}")
    resid = max(
        (float(np.linalg.norm(_centroid(cells[i], pts[i]) - pts[i])) for i in alive),
        default=0.0,
    )
    if resid > tol * diag:
        bad.append(f"not centroidal: worst |site - cell centroid| = {resid:.6g}, limit {tol * diag:.6g}")
    return bad


def area_spread(polygon: Points, sites: Points) -> tuple[list[float], float]:
    """Cell areas and their spread (max - min) / (polygon area / number of cells)."""
    poly, poly_area2 = _ring(polygon)
    areas = np.array([_area(c) for c in _cells(poly, np.asarray(sites, dtype=float))])
    return areas.tolist(), float((areas.max() - areas.min()) / (poly_area2 / 2.0 / len(areas)))


def best_cvt(
    polygon: Points,
    n: int,
    k: int = 8,
    *,
    criterion: str = "spread",
    max_iter: int = 1000,
) -> list[tuple[float, float]]:
    """Run the lattice start plus `k` seeded random starts and keep the best valid one.

    `criterion` is "spread" (smallest area spread, default) or "energy" (lowest CVT
    energy). Raises if none of the k+1 runs passes `validate`.
    """
    if k < 1:
        raise ValueError("k must be >= 1")
    if criterion not in ("spread", "energy"):
        raise ValueError("criterion must be 'spread' or 'energy'")
    best: list[tuple[float, float]] = []
    best_score = np.inf
    failure = "none"
    for seed in (None, *range(k)):
        sites = lloyd_cvt(polygon, n, seed=seed, max_iter=max_iter)
        bad = validate(polygon, sites)
        if bad:
            failure = "; ".join(bad)
            continue
        score = area_spread(polygon, sites)[1] if criterion == "spread" else energy(polygon, sites)
        if score < best_score:
            best, best_score = sites, score
    if not best:
        raise RuntimeError(f"none of the {k + 1} runs validated: {failure}")
    return best


def energy(polygon: Points, sites: Points) -> float:
    """CVT energy of the tessellation, sum_i integral over cell i of |x - site_i|^2 dA."""
    poly, _ = _ring(polygon)
    pts = np.asarray(sites, dtype=float)
    tot = 0.0
    for i, cell in enumerate(_cells(poly, pts)):
        local = cell - pts[i]  # cell-sized coords, or the shoelace sums cancel
        nxt = np.roll(local, -1, axis=0)
        cross = local[:, 0] * nxt[:, 1] - nxt[:, 0] * local[:, 1]
        tot += (cross * (local[:, 0] ** 2 + local[:, 0] * nxt[:, 0] + nxt[:, 0] ** 2
                         + local[:, 1] ** 2 + local[:, 1] * nxt[:, 1] + nxt[:, 1] ** 2)).sum() / 12.0
    return float(tot)
