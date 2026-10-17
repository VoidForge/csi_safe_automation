# Mass-weighted (prescribed-mass) CVT - implementation plan

Working document for the `weighted_cvt.py` extension agreed on 2026-10-15. Read `refactor_plan.md`
§1 for the general layering rules; this feature stays inside `weighted_cvt.py` plus its tests, so the
§6.3 import matrix is unchanged (`lloyd_cvt`, `weighted_cvt`, `record_store`, `trivial_optimizer` stay
pure; no new module).

## 1. Goal

Today `weighted_cvt(polygon, area_weights)` solves Laguerre weights so that every cell's **area**
share matches the requested ratio:

    A_i / A_polygon = area_weights[i] / sum(area_weights)

The new feature lets the caller prescribe a mass density over the polygon and then matches the
**mass** share instead:

    M_i / M_polygon = cell_weights[i] / sum(cell_weights)

where `M_i` is the mass covered by cell `i`, `M_polygon` the total mass of the domain, and
`cell_weights` is the renamed former `area_weights`.

## 2. Frozen decisions (user, 2026-10-15)

| # | Decision | Choice |
| - | -------- | ------ |
| D1 | Overlapping mass polygons | Density field is **additive**: overlaps sum. No disjointness requirement, no "which layer wins" logic. |
| D2 | Shape of a density zone | Any simple polygon, convex or **non-convex**. Polygons with holes are not supported (one polygon in, a triangle list out). |
| D3 | Triangulation | Convex polygon -> plain fan from vertex 0 (already a true triangulation). Non-convex polygon -> **ear clipping** into a true triangulation. No new vertices are introduced, and no triangle carries a negative sign. |
| D4 | Relaxation point | New keyword `relax_to`: `"area"` (**default**, today's area centroid) or `"mass"` (the cell's centre of mass). |
| D5 | Negative densities | **Allowed** (a subtractive / lighter zone). Only `default_density` and each density must be finite; the total mass `M > 0` is the guard. |
| D6 | Parameter shape | `density_zones=[(polygon, density), ...] | None = None` plus `default_density: float = 1.0`. |
| D7 | Name of the target weights | `area_weights` -> **`cell_weights`** in `weighted_cvt`, `best_weighted_cvt`, `validate_weighted`. |
| D8 | Scope | `weighted_cvt.py`, `tests/test_weighted_cvt.py`, `auto_analysis.init_pile_layout` pass-through, `cvt_visualization.ipynb` demo cells, repo memory notes. |

## 3. Density field

Definition (additive, no layering):

    rho(x) = default_density + sum_k density_k * 1[x in polygon_k]

Consequences used throughout:

- Total mass of a region `R` is `default_density * area(R) + sum_t density_t * area(R ∩ tri_t)`.
  Because the field is additive, no polygonal union or exclusion is ever needed: each triangle
  contributes its own overlap area with its own signed density coefficient, overlaps included.
- Every area term is positive; only the density coefficient of a triangle can be negative. That is
  why D3 matters: ear clipping keeps the *areas* additive and non-cancelling, so a negative density
  is the only source of a negative total, and it is a visible, intended one.
- A density zone may stick out of the domain. Only `R ∩ tri_t` is measured and `R` is always a cell or
  the domain, so the outside part contributes nothing and never has to be trimmed.

Inert case: `density_zones` empty/None **and** `default_density == 1.0` builds no field at all
(`field is None`). Every public function then takes today's code path unchanged, so existing output
is bit-identical and the feature costs nothing until it is asked for.

## 4. New private helpers in `weighted_cvt.py`

    _DensityField(NamedTuple)
        triangles: list[np.ndarray]   # each (3, 2), CCW, non-degenerate
        densities: np.ndarray         # one float per triangle, may be negative
        default: float                # background density
        total: float                  # mass of the whole domain polygon

    _ring_is_simple(ring) -> bool
        O(m^2) proper-intersection test over the ring's edges, adjacent edges excluded.

    _triangulate(polygon) -> list[np.ndarray]
        Drops repeated/collinear vertices, requires non-zero signed area, orients CCW, then fans a
        convex polygon and ear-clips a non-convex one. Raises ValueError on a self-intersecting
        polygon or when no ear is found (rather than returning a bad decomposition).

    _density_field(poly, density_zones, default_density) -> _DensityField | None
        Validates every input, triangulates every zone polygon, computes `total`. Raises ValueError
        for a non-finite density, a bad polygon, an empty polygon, a self-intersection, or
        `total <= 0`. Returns None for the inert case (§3).

    _cell_mass(cell, field) -> float
        `field.default * _area(cell) + sum_t field.densities[t] * _area(_clip(cell, tri half-planes))`

    _cell_mass_moment(cell, field) -> tuple[float, np.ndarray]
        Same sum, but each term also weighs in `_centroid(piece, origin)`, so the caller gets the mass
        and the first moment. Only computed when `relax_to == "mass"`.

    _face_density_integral(a, b, field) -> float
        `field.default * |ab| + sum_t density_t * length(segment ab ∩ tri_t)`. The segment∩triangle
        clip is the same half-plane clip applied to a 2-point ring, followed by a projection of the
        clipped points onto the chord direction (max - min) to get the length. Positive lengths,
        signed coefficients, no union needed (§3).

Reuse, do not reimplement: `_ring`, `_area`, `_centroid`, `_clip`, `_inside` from `lloyd_cvt`.

## 5. Solver changes

`_jacobian(cells, sites, lam, diag, field=None)`

- Face weight becomes `_face_density_integral(a, b, field) / (2 * dist)` when a field is present, and
  stays `|edge| / (2 * dist)` when it is `None`.
- The diagonal stays minus the row sum. The argument is unchanged: total mass is invariant under a
  shift of the Laguerre weights, so the columns of `dM_i/dlam_j` sum to zero.
- Pair detection (edge midpoint power tie), the `seen` de-duplication and the symmetry are untouched.

`_solve_weights(poly, sites, lam, targets, tol, max_iter, mode, field=None)`

- `targets` are masses when a field is present (`M * w_i / sum(w)`), areas otherwise.
- The residual, the Levenberg escalation and the damping schedule are unchanged; only the measure
  inside the residual changes.
- The backtracking acceptance test keeps its single guard for the area path (no collapsed cell), and
  under a field it refuses a collapsed cell only: a cell measure of zero or less is **accepted** while
  the weights grow it back towards its positive target. Refusing every non-positive measure instead
  locks the solve out of the sign-changing regions a subtracting zone creates, and stalls at the first
  layout whose cells straddle one; the area rule is still what protects the centroid chase.

Outer loop in `weighted_cvt`

- `relax_to="area"`: `shifted[i] = _centroid(cell, sites[i])` (today).
- `relax_to="mass"`: `shifted[i] = moment / mass` for cells with a positive mass, skipped otherwise
  (mirrors the existing collapsed-cell skip). A subtracting zone can put that centre of mass outside
  the cell and even outside the polygon, so the move is only made when the point lies in the polygon:
  otherwise the site stays where it is and the weights are left to move the boundary.
- `margin`/`_project_to_ring` and the damping are applied exactly as now, in the same order.
- Stop test: `max|M_i - target_i| / M <= mass_tol` when a field is present, else the existing area
  test with `area_tol`. A mass run does not use `area_tol` for convergence.

`best_weighted_cvt`

- `criterion="mass_error"` added; `"area_error"` keeps testing `max_area_error`, `"centroid"` measures
  against the active relaxation point (`_centroid_error` gains a `field`/`relax_to` switch).

## 6. Public API

    weighted_cvt(polygon, cell_weights, *, seed=None, tol=1e-9, area_tol=1e-3, max_iter=200,
                 weight_iter=50, damping=0.5, neighbour_mode="auto", margin=0.0,
                 density_zones=None, default_density=1.0, relax_to="area", mass_tol=1e-3)

    best_weighted_cvt(polygon, cell_weights, k=8, *, criterion="area_error", seed=None,
                      max_iter=200, weight_iter=50, damping=0.5, neighbour_mode="auto",
                      area_tol=1e-3, margin=0.0, density_zones=None, default_density=1.0,
                      relax_to="area", mass_tol=1e-3)

    validate_weighted(polygon, sites, weights, cell_weights, tol=1e-3, area_tol=1e-2, margin=0.0,
                      density_zones=None, default_density=1.0, relax_to="area", mass_tol=1e-2)

    cell_masses(polygon, sites, weights=None, *, density_zones=None, default_density=1.0)

- `WeightedCVT` gains `masses`, `mass_targets`, `max_mass_error` appended last, so every existing
  positional construction and field read keeps working. With no field they mirror
  `areas` / `targets` / `max_area_error`.
- `converged` refers to the active objective (mass error when a field is present).
- `areas` and `targets` keep their meaning and are still reported; under a non-uniform field `targets`
  is documented as the area ratio that was asked for, i.e. a by-product rather than an objective.
- `validate_weighted` with a field: checks the mass ratios against `mass_tol`, checks the centroidal
  property against `relax_to`, reports a non-positive total mass and any cell whose mass is <= 0, and
  **skips** the area-ratio check (areas are no longer the objective).
- `area_weights` -> `cell_weights` everywhere (D7), including docstrings, exception messages, the
  notebook and `refactor_plan.md` §5.3.

`auto_analysis.init_pile_layout`

    init_pile_layout(polygon, weights=None, *, n_piles=None, seed=None, best_of=0, area_tol=1e-3,
                     margin=0.0, density_zones=None, default_density=1.0, relax_to="area",
                     mass_tol=1e-3, allow_nonconverged=False, apply=True)

All four new keywords are forwarded to `weighted_cvt` / `best_weighted_cvt`. Zone polygons are given
in the cap polygon's own coordinates. `weights` keeps its name and its meaning is updated in the
docstring: it sets the target share, which is the mass share once densities are supplied. `mass_tol`
is forwarded as well, because a mass run is judged by it and `area_tol` cannot stand in for it. The
summary line prints the mass figures too when a field is given, and the non-convergence
`RuntimeError` has a mass variant saying that a density field can be what puts a target out of reach.

## 7. Tests (`tests/test_weighted_cvt.py`, self-running, no pytest)

1. Inert regression: no mass arguments reproduces the current sites, weights and areas
   (`masses == areas`, `max_mass_error == max_area_error`, `converged` unchanged).
2. Scale invariance: `default_density=2.0` with no polygons gives the same sites as `1.0`.
3. Two-zone square (left density 3, right density 1), `n=4`, equal `cell_weights`: mass shares are
   equal, the dense-side cells are smaller in area, and an independent fine-grid quadrature of the
   returned cells agrees with `WeightedCVT.masses` within `mass_tol`.
4. `default_density=0.0` plus one dense patch: mass lives only in the patch; the total equals
   `density * patch area`.
5. Overlapping polygons sum: total mass equals the analytic `rho_1 + rho_2` value over the overlap,
   and a cell entirely inside both counts both densities.
6. L-shaped (non-convex) region: ear-clipped mass matches grid quadrature, and matches the signed-fan
   value on the same polygon (an independent consistency check of the decomposition).
7. Negative density: a subtractive zone behaves as `default * A + rho * |intersection|`; a case that
   drives a cell's mass to <= 0 terminates with `converged=False` and no frozen site.
8. `relax_to`: `"mass"` and `"area"` give different sites under a non-uniform field, and identical
   sites under a uniform one.
9. Unequal targets: `cell_weights=[1, 3]` gives a mass ratio of 1:3 under a non-uniform field.
10. `validate_weighted` accepts a good mass run and flags a bad one; `cell_masses` agrees with
    `WeightedCVT.masses`.
11. One mass run with `margin > 0` still converges (or reports the documented limit).
12. Rejections: self-intersecting polygon, polygon with < 3 usable vertices, non-finite density,
    non-positive total mass, bad `relax_to` value.
13. All 19 existing tests updated for the rename and still passing; `tests/test_against_voronoip.py`
    5/5 (untouched engine, no mass case there).

Verified with `c:\Python314\python.exe` (numpy 2.3.2, scipy 1.18.0), Pylance clean on every edited
file. `init_pile_layout` verified with Excel stubbed through a throwaway snippet, as was done for the
`margin` work.

## 8. Notebook

Two cells in `cvt_visualization.ipynb`: one drawing a two-density field over the cap polygon, one
showing the same `n` and `cell_weights` solved twice, `relax_to="area"` vs `"mass"`, with the mass
shares annotated.

## 9. Documented limits (not defects)

- A subtracting zone can leave a cell holding no mass, or drag a centre of mass out of its own cell.
  The first shows up as `converged=False` with the layout the weights reached; the second simply
  holds the site where it is. `auto_analysis` raises on the first unless `allow_nonconverged=True`.
- A target share concentrated in a near-zero-density area may simply be unreachable. The 2-cell case
  on a 2x2 cap with a -19 left half and a 41 right half is one: no single cut can give either cell
  half of the 44 total, and the run reports that instead of settling for a cell that is all hole.
- `relax_to="area"` with a non-uniform field deliberately leaves the sites off the centres of mass;
  `validate_weighted` then checks the area centroid, not the mass centroid.
- Polygons with holes are out of scope; a zone polygon must be simple (checked) and non-degenerate.
- `margin` still keeps the cells on the full polygon; mass is integrated over those full cells.

## 10. Deviations from the first draft

1. The weight-solve acceptance test refuses only a collapsed cell under a field, not every
   non-positive cell measure (§5). Refusing non-positive measures stalls any field whose cells
   straddle a subtracting region.
2. `relax_to="mass"` moves a site only to a centre of mass that lies inside the polygon (§5), so a
   negated region cannot drag a site out of the cap and `validate_weighted` does not report it as
   outside. It mirrors that in its own reference-point check.
3. `init_pile_layout` also forwards `mass_tol` (§6), since `area_tol` cannot stand in for it.
4. Naming (2026-10-15, user request): the parameter is **`density_zones`**, not `mass_polygons`, and
   the private `_MassField` / `_mass_field` became `_DensityField` / `_density_field`. Prose says
   *polygon* for what the caller hands in and *zone* for the region it defines; *ring* is kept only
   for a boundary vertex list (`_ring`, `_ring_is_simple`, `_inset_ring`, `_intersection`), which is
   what it means in ISO/OGC terms.
