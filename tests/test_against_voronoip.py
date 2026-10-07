"""Cross-checks `weighted_cvt`'s power cells against the third-party `voronoip`.

`voronoip` is optional: when it is not installed every test here early-returns, and
under pytest the module is skipped.

Both engines solve the same power (Laguerre) metric, written differently:

    mine      the cell of i is argmin over j of  |x - p_j|^2 - lam_j
    voronoip  the cell of i is argmin over j of  |x - c_j|^2 - w_j^2
              (`Voronoi(mode="power", radii=...)`, `PowerVoronoi`)

so ``lam_i == w_i ** 2`` and their weight is a radius. Only the *differences* of the
weights matter, so a solve result is handed over after the shift
``lam -> lam - min(lam)``, which makes every weight a real radius and leaves the
diagram unchanged.

`voronoip` also ships `PowerDiagram`, which uses the mirrored ``|x - c|^2 + w^2``
convention (a bigger weight gives a *smaller* cell). That is not the convention used
here; `metrics.power_load_distance` documents it.
"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lloyd_cvt import _area, _ring  # noqa: E402
from weighted_cvt import _neighbours, _power_cells, weighted_cvt  # noqa: E402

try:
    from voronoip import PowerVoronoi, Voronoi as VoronoipVoronoi  # noqa: F401

    HAVE_VORONOIP = True
except ImportError:  # optional cross-check dependency
    HAVE_VORONOIP = False

try:
    import pytest

    pytestmark = pytest.mark.skipif(not HAVE_VORONOIP, reason="voronoip is not installed")
except ImportError:  # the file is also runnable on its own
    pytestmark = None

RECT = [(0.0, 0.0), (3.0, 0.0), (3.0, 2.0), (0.0, 2.0)]
BBOX = [0.0, 0.0, 3.0, 2.0]
TOTAL = 6.0


def _voronoip():
    """The optional dependency, imported lazily so the names are never unbound."""
    from voronoip import PowerVoronoi as pv
    from voronoip import Voronoi as vo

    return pv, vo


def _sites(seed: int, n: int) -> np.ndarray:
    rng = np.random.default_rng(seed)
    return np.column_stack((rng.uniform(0.15, 2.85, n), rng.uniform(0.15, 1.85, n)))


def _mine(sites: np.ndarray, lam: np.ndarray) -> np.ndarray:
    return np.array(
        [_area(c) for c in _power_cells(np.asarray(_ring(RECT)[0]), sites, lam, "hull")]
    )


def _theirs(sites: np.ndarray, radii: np.ndarray) -> np.ndarray:
    _, vo = _voronoip()
    cells = vo(sites, mode="power", radii=radii, bbox=BBOX).polygons
    return np.array([_area(np.asarray(c, dtype=float)) if len(c) >= 3 else 0.0 for c in cells])


def test_cell_areas_match_voronoip() -> None:
    if not HAVE_VORONOIP:
        return
    for seed, n, hi in ((1, 6, 0.4), (2, 12, 0.5), (3, 20, 0.8)):
        sites = _sites(seed, n)
        radii = np.random.default_rng(seed + 50).uniform(0.0, hi, n)
        mine, theirs = _mine(sites, radii**2), _theirs(sites, radii)
        assert np.abs(mine - theirs).max() < 1e-9, (seed, mine, theirs)
        assert abs(mine.sum() - TOTAL) < 1e-9


def test_unweighted_cells_match_voronoip() -> None:
    if not HAVE_VORONOIP:
        return
    sites = _sites(4, 12)
    _, vo = _voronoip()
    mine = _mine(sites, np.zeros(len(sites)))
    theirs = np.array(
        [_area(np.asarray(c, dtype=float)) for c in vo(sites, bbox=BBOX).polygons]
    )
    assert np.abs(mine - theirs).max() < 1e-9


def test_dominated_cells_agree_with_voronoip() -> None:
    if not HAVE_VORONOIP:
        return
    sites = _sites(5, 12)
    radii = np.r_[np.zeros(11), 3.0]  # one site swallows the whole rectangle
    mine, theirs = _mine(sites, radii**2), _theirs(sites, radii)
    assert np.abs(mine - theirs).max() < 1e-9
    assert int((mine < 1e-12).sum()) == 11
    assert int((theirs < 1e-12).sum()) == 11


def test_voronoip_adjacency_is_contained_in_mine() -> None:
    """Every neighbour voronoip gives a cell is one of mine as well.

    Containment, not equality: voronoip also lists sites that meet a cell at a single
    point (a degenerate touch), which the triangulation-based neighbours here leave
    out. Those pairings never cut area, so the cells still match exactly.
    """
    if not HAVE_VORONOIP:
        return
    sites = _sites(6, 14)
    radii = np.random.default_rng(9).uniform(0.0, 0.4, len(sites))
    mine = _neighbours(sites, radii**2, "hull")[0]
    pv, _ = _voronoip()
    query = pv(sites, radii, bbox=BBOX)
    query.compute()
    for i in range(len(sites)):
        extra = query.neighbors_of(i) - mine[i]
        assert not extra, (i, sorted(extra))


def test_solved_layout_is_reproduced_by_voronoip() -> None:
    """The solve's own output, recomputed independently, keeps the requested areas."""
    if not HAVE_VORONOIP:
        return
    for seed in (None, 1, 2):
        result = weighted_cvt(RECT, [1.0, 1.0, 2.0, 2.0, 3.0, 5.0], seed=seed)
        lam = np.asarray(result.weights)
        radii = np.sqrt(lam - lam.min())  # gauge shift: same diagram, real radii
        theirs = _theirs(np.asarray(result.sites), radii)
        assert np.abs(theirs - np.asarray(result.areas)).max() < 1e-9, seed
        assert np.abs(theirs - np.asarray(result.targets)).max() / TOTAL <= 1e-3, seed


if __name__ == "__main__":
    if not HAVE_VORONOIP:
        print("SKIPPED: voronoip is not installed")
        sys.exit(0)
    failures = 0
    names = sorted(n for n in dir() if n.startswith("test_"))
    for name in names:
        try:
            globals()[name]()
            print(f"PASS {name}")
        except Exception:  # noqa: BLE001 - a plain runner, show everything
            failures += 1
            print(f"FAIL {name}")
            traceback.print_exc()
    print(f"\n{len(names) - failures} passed, {failures} failed")
    sys.exit(1 if failures else 0)
