"""Tests for `weighted_cvt.py`.

Run with pytest, or directly: `python tests/test_weighted_cvt.py`.
"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lloyd_cvt import _area, _cells, _centroid, _inside, _ring  # noqa: E402
from weighted_cvt import (  # noqa: E402
    _power_cells,
    _triangulate,
    best_weighted_cvt,
    cell_areas,
    cell_masses,
    inset_polygon,
    power_cells,
    validate_weighted,
    weighted_cvt,
)


def square(size: float = 1.0) -> list[tuple[float, float]]:
    return [(0.0, 0.0), (size, 0.0), (size, size), (0.0, size)]


# The right half of a unit square. Density 3 on top of a unit default leaves the left
# half at 1 and the right half at 4, so the square's total mass is 2.5.
DENSE_HALF = [(0.5, 0.0), (1.0, 0.0), (1.0, 1.0), (0.5, 1.0)]


def _raises(exc, fn, *args, **kwargs) -> None:
    try:
        fn(*args, **kwargs)
    except exc:
        return
    raise AssertionError(f"{getattr(fn, '__name__', fn)} did not raise {exc.__name__}")


def _in_ring(ring, pts: np.ndarray) -> np.ndarray:
    """Even-odd (ray crossing) test of the rows of `pts` against a convex or not ring."""
    ring = np.asarray(ring, dtype=float)
    x, y = pts[:, 0], pts[:, 1]
    inside = np.zeros(len(pts), dtype=bool)
    for k in range(len(ring)):
        x1, y1 = ring[k]
        x2, y2 = ring[(k + 1) % len(ring)]
        if y1 == y2:
            continue
        hit = ((y1 > y) != (y2 > y)) & (x < x1 + (y - y1) * (x2 - x1) / (y2 - y1))
        inside ^= hit
    return inside


def _grid_masses(polygon, sites, weights, density_zones=None, default_density=1.0,
                 n: int = 300) -> np.ndarray:
    """Mass per cell from a fine grid over the polygon's bounding box.

    An independent measure, used as a cross-check on the library's own integrals: the
    grid only knows how to classify a point by the power metric and how to add up the
    densities that cover it.
    """
    poly = np.asarray(_ring(polygon)[0])
    pts = np.asarray(sites, dtype=float)
    lam = np.asarray(weights, dtype=float)
    lo, hi = poly.min(axis=0), poly.max(axis=0)
    step_x, step_y = (hi[0] - lo[0]) / n, (hi[1] - lo[1]) / n
    xs = lo[0] + (np.arange(n) + 0.5) * step_x
    ys = lo[1] + (np.arange(n) + 0.5) * step_y
    grid = np.column_stack((np.repeat(xs, n), np.tile(ys, n)))
    grid = grid[_inside(poly, grid)]
    rho = np.full(len(grid), float(default_density))
    for zone, density in density_zones or []:
        rho[_in_ring(zone, grid)] += density
    power = ((grid[:, None, :] - pts[None, :, :]) ** 2).sum(axis=2) - lam[None, :]
    owner = power.argmin(axis=1)
    return np.bincount(owner, weights=rho * step_x * step_y, minlength=len(pts))


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


def test_mass_is_inert_without_a_field() -> None:
    ratios = [1.0, 2.0, 3.0, 4.0]
    plain = weighted_cvt(square(), ratios, seed=1)
    explicit = weighted_cvt(square(), ratios, seed=1, density_zones=None, default_density=1.0,
                            relax_to="area", mass_tol=1e-3)
    assert plain == explicit
    assert explicit.masses == explicit.areas
    assert explicit.mass_targets == explicit.targets
    assert explicit.max_mass_error == explicit.max_area_error


def test_default_density_scale_invariance() -> None:
    ratios = [1.0, 2.0, 3.0, 4.0]
    unit = weighted_cvt(square(), ratios, seed=1)
    doubled = weighted_cvt(square(), ratios, seed=1, default_density=2.0)
    assert doubled.converged, doubled
    assert np.allclose(np.asarray(unit.sites), np.asarray(doubled.sites), atol=1e-9)
    assert np.allclose(doubled.masses, 2.0 * np.asarray(unit.areas), atol=1e-9)
    assert doubled.max_mass_error <= 1e-3


def test_mass_shares_follow_the_weights_across_two_zones() -> None:
    polygon = square()
    result = weighted_cvt(polygon, [1.0] * 4, density_zones=[(DENSE_HALF, 3.0)],
                          relax_to="mass")
    assert result.converged, result
    assert result.max_mass_error <= 1e-3
    assert abs(sum(result.masses) - 2.5) < 1e-9  # 0.5 of density 1 plus 0.5 of density 4
    for mass, target in zip(result.masses, result.mass_targets):
        assert abs(mass - target) < 1e-3 * 2.5
    # the dense half needs less area for the same mass, so the cells differ noticeably
    assert max(result.areas) > 2.0 * min(result.areas)
    assert validate_weighted(polygon, result.sites, result.weights, result.mass_targets,
                             density_zones=[(DENSE_HALF, 3.0)], relax_to="mass") == []
    grid = _grid_masses(polygon, result.sites, result.weights,
                        density_zones=[(DENSE_HALF, 3.0)])
    assert np.allclose(grid / grid.sum(), np.asarray(result.masses) / 2.5, atol=5e-3)


def test_zero_default_density_confines_the_mass_to_the_patch() -> None:
    polygon = square(2.0)
    patch = [(0.5, 0.5), (1.5, 0.5), (1.5, 1.5), (0.5, 1.5)]  # area 1, density 4
    result = weighted_cvt(polygon, [1.0, 1.0], default_density=0.0,
                          density_zones=[(patch, 4.0)])
    assert result.converged, result
    assert abs(sum(result.masses) - 4.0) < 1e-9
    assert abs(min(result.masses) - 2.0) < 1e-3 * 4.0
    for area in result.areas:  # the massless surround is free for the taking
        assert area > 0.0


def test_overlapping_density_zones_add_up() -> None:
    polygon = square()
    left = [(0.0, 0.0), (0.6, 0.0), (0.6, 1.0), (0.0, 1.0)]
    right = [(0.4, 0.0), (1.0, 0.0), (1.0, 1.0), (0.4, 1.0)]
    zones = [(left, 2.0), (right, 5.0)]
    result = weighted_cvt(polygon, [1.0, 1.0], default_density=0.0, density_zones=zones)
    assert abs(sum(result.masses) - 4.2) < 1e-9  # 0.6 * 2 + 0.6 * 5, the overlap counts twice
    assert validate_weighted(polygon, result.sites, result.weights, result.mass_targets,
                             default_density=0.0, density_zones=zones) == []
    grid = _grid_masses(polygon, result.sites, result.weights, density_zones=zones,
                        default_density=0.0)
    assert abs(grid.sum() - 4.2) < 5e-3


def test_non_convex_density_zone_is_triangulated() -> None:
    polygon = square()
    l_shape = [(0.1, 0.1), (0.9, 0.1), (0.9, 0.4), (0.4, 0.4), (0.4, 0.9), (0.1, 0.9)]
    triangles = _triangulate(l_shape, "test zone")
    assert len(triangles) == 4  # two arms of two triangles each
    assert all(_area(tri) > 0.0 for tri in triangles)  # no signed triangle, no cancellation
    assert abs(sum(_area(tri) for tri in triangles) - _area(np.asarray(l_shape))) < 1e-12

    result = weighted_cvt(polygon, [1.0] * 3, density_zones=[(l_shape, 2.0)])
    assert result.converged, result
    assert abs(sum(result.masses) - (1.0 + 2.0 * _area(np.asarray(l_shape)))) < 1e-9
    grid = _grid_masses(polygon, result.sites, result.weights, density_zones=[(l_shape, 2.0)])
    assert np.allclose(grid / grid.sum(), np.asarray(result.masses) / sum(result.masses),
                       atol=5e-3)


def test_negative_density_thins_a_zone() -> None:
    polygon = square(2.0)
    left = [(0.0, 0.0), (1.0, 0.0), (1.0, 2.0), (0.0, 2.0)]   # left half, density 3 - 2 = 1
    right = [(1.0, 0.0), (2.0, 0.0), (2.0, 2.0), (1.0, 2.0)]  # right half, density 3 + 40
    zones = [(left, -2.0), (right, 40.0)]
    result = weighted_cvt(polygon, [1.0, 1.0, 2.0, 2.0], default_density=3.0,
                          density_zones=zones)
    assert result.converged, result
    assert result.max_mass_error <= 1e-3
    assert abs(sum(result.masses) - 88.0) < 1e-9  # 3 * 4 - 2 * 2 + 40 * 2
    assert min(result.masses) > 0.0
    assert validate_weighted(polygon, result.sites, result.weights, result.mass_targets,
                             default_density=3.0, density_zones=zones) == []


def test_unreachable_mass_target_reports_not_converged() -> None:
    polygon = square(2.0)
    left = [(0.0, 0.0), (1.0, 0.0), (1.0, 2.0), (0.0, 2.0)]
    right = [(1.0, 0.0), (2.0, 0.0), (2.0, 2.0), (1.0, 2.0)]
    zones = [(left, -20.0), (right, 40.0)]  # left half -19, right half 41, total 44
    # A single cut can only give the left cell 44b - 82 of the mass, so it cannot hold
    # half of 44: the run reports that instead of settling for a cell that is all hole.
    result = weighted_cvt(polygon, [1.0, 1.0], density_zones=zones, relax_to="mass")
    assert not result.converged
    assert abs(sum(result.masses) - 44.0) < 1e-9  # the total is a property of the field
    assert min(result.masses) <= 0.0
    problems = validate_weighted(polygon, result.sites, result.weights, result.mass_targets,
                                 density_zones=zones, relax_to="mass")
    assert any("masses off target" in message for message in problems), problems
    for x, y in result.sites:  # a site is never dragged out of the polygon
        assert 0.0 <= x <= 2.0 and 0.0 <= y <= 2.0


def test_relax_to_area_and_mass_differ_only_with_density() -> None:
    polygon = square()
    by_area = weighted_cvt(polygon, [1.0] * 4, density_zones=[(DENSE_HALF, 3.0)],
                           relax_to="area")
    by_mass = weighted_cvt(polygon, [1.0] * 4, density_zones=[(DENSE_HALF, 3.0)],
                           relax_to="mass")
    assert not np.allclose(np.asarray(by_area.sites), np.asarray(by_mass.sites), atol=1e-6)
    assert by_area.converged and by_mass.converged
    # a uniform field leaves the centre of mass where the area centroid is
    uniform = weighted_cvt(polygon, [1.0] * 4, default_density=2.0, relax_to="mass")
    plain = weighted_cvt(polygon, [1.0] * 4, default_density=2.0, relax_to="area")
    assert np.allclose(np.asarray(uniform.sites), np.asarray(plain.sites), atol=1e-6)


def test_mass_ratio_follows_unequal_weights() -> None:
    polygon = square()
    result = weighted_cvt(polygon, [1.0, 3.0], density_zones=[(DENSE_HALF, 3.0)])
    assert result.converged, result
    assert abs(result.masses[1] / result.masses[0] - 3.0) < 1e-2
    grid = _grid_masses(polygon, result.sites, result.weights,
                        density_zones=[(DENSE_HALF, 3.0)])
    assert abs(grid[1] / grid[0] - 3.0) < 5e-2


def test_cell_masses_matches_result() -> None:
    polygon = square()
    result = weighted_cvt(polygon, [1.0, 2.0, 3.0], density_zones=[(DENSE_HALF, 3.0)])
    masses = cell_masses(polygon, result.sites, result.weights,
                         density_zones=[(DENSE_HALF, 3.0)])
    assert np.allclose(masses, result.masses, atol=1e-12)
    assert np.allclose(cell_masses(polygon, result.sites, result.weights),
                       cell_areas(polygon, result.sites, result.weights), atol=1e-12)


def test_validate_weighted_flags_a_wrong_mass_ratio() -> None:
    polygon = square()
    result = weighted_cvt(polygon, [1.0] * 4, density_zones=[(DENSE_HALF, 3.0)],
                          relax_to="mass")
    assert result.converged, result
    assert validate_weighted(polygon, result.sites, result.weights, result.mass_targets,
                             density_zones=[(DENSE_HALF, 3.0)], relax_to="mass") == []
    skewed = validate_weighted(polygon, result.sites, result.weights, [1.0, 1.0, 3.0, 3.0],
                               density_zones=[(DENSE_HALF, 3.0)], relax_to="mass")
    assert any("masses off target" in message for message in skewed), skewed
    # the same layout cannot satisfy the area ratio as well: the masses are what was asked
    by_area = validate_weighted(polygon, result.sites, result.weights, result.targets)
    assert any("areas off target" in message for message in by_area), by_area


def test_mass_run_with_margin() -> None:
    polygon = square(2.0)
    margin = 0.3
    result = weighted_cvt(polygon, [1.0] * 4, density_zones=[(DENSE_HALF, 3.0)],
                          relax_to="mass", margin=margin)
    assert result.converged, result
    assert result.max_mass_error <= 1e-3
    assert _clearance(polygon, result.sites) >= margin - 1e-9


def test_best_weighted_cvt_mass_criterion() -> None:
    polygon = square()
    result = best_weighted_cvt(polygon, [1.0] * 4, k=2, criterion="mass_error",
                               density_zones=[(DENSE_HALF, 3.0)], relax_to="mass")
    assert result.converged, result
    assert result.max_mass_error <= 1e-3
    centroid = best_weighted_cvt(polygon, [1.0] * 4, k=1, criterion="centroid",
                                 density_zones=[(DENSE_HALF, 3.0)], relax_to="mass")
    assert abs(sum(centroid.masses) - 2.5) < 1e-9


def test_bad_mass_inputs_raise() -> None:
    ring = square()
    bowtie = [(0.0, 0.0), (1.0, 1.0), (1.0, 0.0), (0.0, 1.0)]     # crosses itself
    sliver = [(0.0, 0.0), (1.0, 1.0), (0.0, 0.0)]                 # two distinct corners
    for bad in ([(bowtie, 1.0)], [(sliver, 1.0)], [(ring, float("nan"))],
                [(ring, 1.0), "not a pair"], [((0.0, 0.0), 1.0)]):
        _raises(ValueError, weighted_cvt, square(), [1.0, 1.0], density_zones=bad)
    # a field with no mass at all
    _raises(ValueError, weighted_cvt, square(), [1.0, 1.0], default_density=-1.0)
    _raises(ValueError, weighted_cvt, square(), [1.0, 1.0], default_density=0.0)
    _raises(ValueError, weighted_cvt, square(), [1.0, 1.0], default_density=float("inf"))
    # a relaxation point that does not exist
    for bad in ("centre", "Area", ""):
        _raises(ValueError, weighted_cvt, square(), [1.0, 1.0], relax_to=bad)
    assert validate_weighted(square(), [(0.3, 0.3), (0.7, 0.7)], [0.0, 0.0], [1.0, 1.0],
                             relax_to="centre")[0].startswith("relax_to")
    problem = validate_weighted(square(), [(0.3, 0.3), (0.7, 0.7)], [0.0, 0.0], [1.0, 1.0],
                                density_zones=[(bowtie, 1.0)])
    assert len(problem) == 1 and problem[0].startswith("mass density field:"), problem


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
