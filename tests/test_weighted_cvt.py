"""Tests for `weighted_cvt.py`.

Run with pytest, or directly: `python tests/test_weighted_cvt.py`.
"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lloyd_cvt import _area, _cells, _centroid, _ring  # noqa: E402
from weighted_cvt import (  # noqa: E402
    _power_cells,
    best_weighted_cvt,
    cell_areas,
    inset_polygon,
    power_cells,
    validate_weighted,
    weighted_cvt,
)


def square(size: float = 1.0) -> list[tuple[float, float]]:
    return [(0.0, 0.0), (size, 0.0), (size, size), (0.0, size)]


def _raises(exc, fn, *args, **kwargs) -> None:
    try:
        fn(*args, **kwargs)
    except exc:
        return
    raise AssertionError(f"{getattr(fn, '__name__', fn)} did not raise {exc.__name__}")


def _clearance(polygon: list[tuple[float, float]], sites) -> float:
    """Smallest distance from any of `sites` to any edge line of the polygon."""
    ring = np.asarray(_ring(polygon)[0])
    pts = np.asarray(sites, dtype=float)
    edge = np.roll(ring, -1, axis=0) - ring
    rel = pts[:, None, :] - ring[None, :, :]
    cross = edge[None, :, 0] * rel[:, :, 1] - edge[None, :, 1] * rel[:, :, 0]
    return float((cross / np.hypot(edge[:, 0], edge[:, 1])).min())


def test_equal_weights_are_equal_area() -> None:
    n = 7
    result = weighted_cvt(square(), [1.0] * n)
    assert result.converged, result
    assert result.max_area_error <= 1e-3
    for area, target in zip(result.areas, result.targets):
        assert abs(area - 1.0 / n) < 1e-3
    assert validate_weighted(square(), result.sites, result.weights, result.targets) == []


def test_prescribed_areas_random_weights() -> None:
    polygon = [(0.0, 0.0), (3.0, 0.0), (3.0, 2.0), (0.0, 2.0)]
    rng = np.random.default_rng(0)
    ratios = rng.uniform(0.2, 3.0, size=9)
    result = weighted_cvt(polygon, ratios)
    assert result.converged, result
    assert result.max_area_error <= 1e-3
    assert validate_weighted(polygon, result.sites, result.weights, result.targets) == []
    assert abs(sum(result.areas) - 6.0) < 1e-9


def test_weights_scale_invariance() -> None:
    ratios = np.array([1.0, 2.0, 3.0, 4.0])
    first = weighted_cvt(square(), ratios, seed=1)
    second = weighted_cvt(square(), ratios * 10.0, seed=1)
    assert np.allclose(np.asarray(first.sites), np.asarray(second.sites), atol=1e-6)
    assert np.allclose(first.areas, second.areas, atol=1e-9)


def test_areas_are_monotone_in_weight() -> None:
    ratios = np.array([1.0, 1.0, 1.0, 5.0])
    result = weighted_cvt(square(), ratios)
    assert result.converged, result
    assert abs(max(result.areas) - 5.0 / 8.0) < 1e-3
    assert abs(min(result.areas) - 1.0 / 8.0) < 1e-3


def test_equal_weights_power_cells_match_unweighted() -> None:
    polygon = square(2.0)
    ring = np.asarray(_ring(polygon)[0])
    sites = np.array(
        [[0.5, 0.5], [1.5, 0.6], [1.0, 1.6], [0.4, 1.4], [1.6, 1.5]]
    )
    weighted = power_cells(polygon, sites)
    plain = _cells(ring, sites)
    assert len(weighted) == len(plain)
    for a, b in zip(weighted, plain):
        assert abs(_area(a) - _area(b)) < 1e-9
        assert np.allclose(_centroid(a, a[0]), _centroid(b, b[0]), atol=1e-9)


def test_neighbour_modes_agree() -> None:
    polygon = square(2.0)
    ring = np.asarray(_ring(polygon)[0])
    rng = np.random.default_rng(3)
    sites = rng.uniform(0.2, 1.8, size=(12, 2))
    lam = rng.normal(0.0, 0.3, size=12)
    hull = _power_cells(ring, sites, lam, "hull")
    allp = _power_cells(ring, sites, lam, "all")
    assert len(hull) == len(allp) == 12
    # a site squeezed out of the diagram must come back empty from both routes
    assert any(len(cell) < 3 for cell in allp)
    for a, b in zip(hull, allp):
        assert abs(_area(a) - _area(b)) < 1e-9


def test_wide_weight_spread_converges() -> None:
    polygon = [(0.0, 0.0), (3.0, 0.0), (3.0, 2.0), (0.0, 2.0)]
    ratios = np.random.default_rng(0).uniform(0.2, 3.0, size=9)
    result = weighted_cvt(polygon, ratios)
    assert result.converged, result
    assert result.max_area_error <= 1e-3
    assert validate_weighted(polygon, result.sites, result.weights, result.targets) == []


def test_lopsided_weights_stay_clean() -> None:
    polygon = square()
    result = weighted_cvt(polygon, [0.05, 1.0, 1.0, 6.0])
    assert result.converged, result
    assert abs(min(result.areas) - 0.05 / 8.05) < 1e-3
    assert abs(max(result.areas) - 6.0 / 8.05) < 1e-3
    assert validate_weighted(polygon, result.sites, result.weights, result.targets) == []


def test_grid_classification_matches_cells() -> None:
    polygon = square()
    result = weighted_cvt(polygon, [1.0, 1.0, 1.0, 3.0])
    sites = np.asarray(result.sites)
    lam = np.asarray(result.weights)
    n = 400
    axis = np.linspace(0.0, 1.0, n)
    grid = np.column_stack((np.repeat(axis, n), np.tile(axis, n)))
    power = ((grid[:, None, :] - sites[None, :, :]) ** 2).sum(axis=2) - lam[None, :]
    counts = np.bincount(power.argmin(axis=1), minlength=len(sites)) / len(grid)
    for share, area in zip(counts, result.areas):
        assert abs(share - area) < 0.02, (share, area)


def test_cell_areas_matches_result() -> None:
    result = weighted_cvt(square(), [1.0, 2.0, 3.0])
    areas = cell_areas(square(), result.sites, result.weights)
    assert np.allclose(areas, result.areas, atol=1e-9)


def test_validate_weighted_flags_nudged_site() -> None:
    polygon = square()
    result = weighted_cvt(polygon, [1.0, 2.0, 3.0, 4.0])
    assert validate_weighted(polygon, result.sites, result.weights, result.targets) == []
    nudged = [list(site) for site in result.sites]
    nudged[0][0] += 0.1
    assert validate_weighted(polygon, nudged, result.weights, result.targets)


def test_margin_keeps_sites_off_the_edges() -> None:
    polygon = square(2.0)
    margin = 0.25
    for ratios in ([1.0] * 9, [0.2, 1.0, 4.0, 1.0, 1.0, 1.0, 3.0, 0.5, 2.0]):
        result = weighted_cvt(polygon, ratios, margin=margin)
        assert _clearance(polygon, result.sites) >= margin - 1e-9
        assert abs(sum(result.areas) - 4.0) < 1e-9  # the cells still cover the square
        bad = validate_weighted(
            polygon, result.sites, result.weights, result.targets, margin=margin
        )
        # a site held at the margin is no longer its cell's centroid, so only that shows
        assert all("centroid" in message for message in bad), bad


def test_margin_holds_sites_exactly_at_the_limit() -> None:
    polygon = square(2.0)
    margin = 0.6
    result = weighted_cvt(polygon, [1.0] * 9, margin=margin)
    assert result.converged, result
    # the constraint is active, not merely slack: 9 sites cannot all clear 0.6 freely
    assert abs(_clearance(polygon, result.sites) - margin) <= 1e-9


def test_best_weighted_cvt_passes_the_margin_on() -> None:
    polygon = square(2.0)
    margin = 0.3
    result = best_weighted_cvt(polygon, [1.0, 1.0, 2.0, 2.0, 3.0, 3.0], k=2, margin=margin)
    assert _clearance(polygon, result.sites) >= margin - 1e-9
    assert all(
        "centroid" in message
        for message in validate_weighted(
            polygon, result.sites, result.weights, result.targets, margin=margin
        )
    )


def test_margin_zero_leaves_the_default_run_alone() -> None:
    ratios = [1.0, 2.0, 3.0, 4.0]
    assert weighted_cvt(square(), ratios, seed=1) == weighted_cvt(
        square(), ratios, seed=1, margin=0.0
    )


def test_inset_polygon_offsets_every_edge() -> None:
    ring = np.asarray(inset_polygon(square(2.0), 0.25))
    expected = np.array([[0.25, 0.25], [1.75, 0.25], [1.75, 1.75], [0.25, 1.75]])
    assert ring.shape == (4, 2)
    assert np.allclose(np.sort(ring, axis=0), np.sort(expected, axis=0))
    assert abs(float(_area(ring)) - 1.5**2) < 1e-12
    assert inset_polygon(square(2.0), 0.0) == square(2.0)


def test_bad_margin_raises() -> None:
    for bad in (-0.1, float("nan"), 0.5, 1.0):  # 0.5 and 1.0 leave no area in a unit square
        _raises(ValueError, weighted_cvt, square(), [1.0, 1.0], margin=bad)
    _raises(ValueError, best_weighted_cvt, square(), [1.0, 1.0], k=2, margin=0.5)
    _raises(ValueError, inset_polygon, square(), 0.5)


def test_best_weighted_cvt_converges() -> None:
    polygon = square(2.0)
    result = best_weighted_cvt(polygon, [1.0, 1.0, 2.0, 2.0, 3.0], k=3)
    assert result.converged, result
    assert validate_weighted(polygon, result.sites, result.weights, result.targets) == []


def test_bad_inputs_raise() -> None:
    for bad in ([], [0.0, 1.0], [-1.0, 1.0], [1.0, -1.0], [1.0, float("nan")]):
        _raises(ValueError, weighted_cvt, square(), bad)
    nonconvex = [(0.0, 0.0), (2.0, 0.0), (2.0, 2.0), (1.0, 0.5), (0.0, 2.0)]
    _raises(ValueError, weighted_cvt, nonconvex, [1.0] * 5)
    _raises(ValueError, best_weighted_cvt, square(), [1.0, 1.0], 0)


if __name__ == "__main__":
    failures = 0
    for name in sorted(n for n in dir() if n.startswith("test_")):
        try:
            globals()[name]()
            print(f"PASS {name}")
        except Exception:  # noqa: BLE001 - a plain runner, show everything
            failures += 1
            print(f"FAIL {name}")
            traceback.print_exc()
    print(f"\n{len([n for n in dir() if n.startswith('test_')]) - failures} passed, {failures} failed")
    sys.exit(1 if failures else 0)
