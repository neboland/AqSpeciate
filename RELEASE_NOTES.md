# AqSpeciate — Release Notes

---

## v2.1b  *(current release)*

**Release type:** Performance and usability update

### Summary

v2.1b introduces computational efficiency improvements to the Newton solver, resolves a
line-continuation compiler error, and adds two new worksheet UDFs (`AqSpeciateLabels` and
an extended `AqSpeciateOne`) along with verbose diagnostic output for `AqSpeciateAll`.
Computed values are identical to v2.1a for all inputs.

### New features

- **`AqSpeciateLabels` UDF** — accepts the same arguments as `AqSpeciateAll` and returns a
  matching label row for use as column headers directly above an `AqSpeciateAll` formula.
  Labels follow the pattern `p[ID]`, `SI[ID]`, `SI[ID](fixed)`, `Calc I (mol/L)`, `Errors`,
  and optional verbose labels when `Verbose=1`.

- **`Verbose` parameter for `AqSpeciateAll`** — optional final argument (default 0). When
  set to 1, eight additional diagnostic columns are appended to the output row:
  - Comma-separated IDs of any precipitated considered solids
  - Newton iteration count at convergence
  - Convergence precision (final residual L2-norm, mol/L)
  - log₁₀(γ) activity coefficients for integer charge magnitudes |z| = 1 through 5

- **`SuppressErrors` parameter for `AqSpeciateOne`** — optional final argument (default 0).
  When set to 1, prevents `#VALUE!` return on solver errors, returning the best available
  result instead. Default behaviour (return `#VALUE!` on error) is unchanged.

- **IS column moved** in `AqSpeciateAll` output — the computed ionic strength column now
  appears after all solid SI columns, immediately before the error string, grouping scalar
  summary values together.

- **Fixed solid SI in output** — `AqSpeciateAll` now returns a SI value for each fixed solid
  (Type 3) in addition to considered solids (Type 5). At convergence the Ksp constraint is
  satisfied, so fixed solid SI values are zero (or a small numerical residual).

### Performance improvements

Six optimisations were applied to the Newton solver. None affect computed values.

| ID | Description | Mechanism |
|---|---|---|
| OPT-A | Fused Residual + Jacobian (highest impact) | `BuildResiduals` and `BuildJacobian` replaced by a single `BuildResidualsAndJacobian` that computes species concentrations once per iteration instead of twice, halving the dominant per-iteration cost |
| OPT-B | Trial array pre-allocated outside Armijo loop | Eliminates up to 17 array allocations per line-search call |
| OPT-C | Gamma arrays pre-allocated in `SolveCore` | `BuildGammaTables` fills pre-allocated arrays rather than allocating on every call inside the Newton loop |
| OPT-D | Pin lookup arrays pre-allocated in `SolveCore` | `compPin` and `compSecPin` allocated once, cleared and refilled each iteration without dynamic allocation |
| OPT-E | `JacSq` copy eliminated | Jacobian passed directly to `SolveLinear` (in-place by design); removes nC² element copy and two allocations per iteration |
| OPT-H | Bulk range reads | `.Value2` array reads replace cell-by-cell access in `PackInputsNew` and `PackSolids`, reducing COM round-trips from O(nC·nS) to 2 |

### Bug fixes

- Fixed "Else without If" compile error in `BuildGammaTables` introduced during OPT-C
  implementation (missing `If zS(i) = 0#` line was accidentally deleted).
- Fixed `#VALUE!` error caused by `UBound()` call on uninitialized arrays in
  `BuildGammaTables` — VBA raises error 9 on `UBound()` of an uninitialized array rather
  than returning a sentinel; reverted to unconditional `ReDim`.
- Resolved "too many line continuations" compiler error — all Sub/Function signatures
  consolidated to fewer than 10 continuation lines.

### UDF registration (`RegisterAqSpeciate`)

A new public sub `RegisterAqSpeciate` registers function and argument descriptions for all
three UDFs so they appear in the Excel Insert Function dialog (Shift+F3) and in formula
autocomplete tooltips.

**Platform support:** `Application.MacroOptions` is supported on Excel for Windows
(2010 and later) only. On Excel for Mac the call is wrapped in `On Error Resume Next`
and silently skipped — the UDFs work normally, argument hints simply do not appear. There
is no cross-platform VBA alternative.

**Persistence:** MacroOptions descriptions are stored in the workbook session and cleared
when the workbook is closed. To restore them automatically on every open, add to the
`ThisWorkbook` module:

```vba
Private Sub Workbook_Open()
    RegisterAqSpeciate
End Sub
```

`RegisterAqSpeciate` can also be run manually at any time from Developer → Macros.


---

## v2.1a  *(previous release)*

**Release type:** Feature update — solid phases, verbose output, label UDF (pre-publication)

### Summary

v2.1a is the first release to include solid phase support (Type-3 fixed solids and Type-5
considered solids with iterative precipitation), extended output diagnostics, and the
`AqSpeciateLabels` helper UDF. This version was not publicly released; it was superseded by
the bug fixes in v2.1b.

### Key additions relative to v2.0g

- Fixed solids (Type 3): Ksp constraint replaces mass balance of the primary component;
  greedy assignment ensures non-conflicting primary components across simultaneously active
  solids; mole-ratio constraint handles components supplied entirely by a solid (compT = 0).
- Considered solids (Type 5): outer precipitation loop with simultaneous state updates,
  oscillation guard (flip-flop detection), and raised iteration limit (30).
- SI output follows the dissolution convention: SI > 0 = supersaturated (matching HYDRAQL).
- Ionic strength column moved to after solid SI columns.
- `Verbose`, `SuppressErrors`, and `AqSpeciateLabels` added.

---

## v2.0g  *(internal — not publicly released)*

Solid SI sign corrected to dissolution convention; fixed solid SI values added to output
array; E15 false-positive fixed for components supplied entirely by a fixed solid.

## v2.0f  *(internal — not publicly released)*

Mole-ratio constraint introduced for components with compT = 0 that are supplied by a
fixed solid; extended BuildPinArrays with secondary pin arrays.

## v2.0e  *(internal — not publicly released)*

Critical fix: Ksp constraint sign corrected from dissolution to formation convention
(logIAP + logK_formation = 0 at equilibrium). This was the root cause of all #VALUE errors
in v2.0a–d when any solid was present.

## v2.0d  *(internal — not publicly released)*

Dynamic greedy primary-component assignment in `BuildPinArrays`, eliminating silent
constraint conflicts when multiple solids competed for the same primary component.

## v2.0c  *(internal — not publicly released)*

E15 false-positive fixed for H⁺ when pH is fixed; `PrimaryComponents` updated to exclude
pH-pinned H⁺ from solid primary selection.

## v2.0b  *(internal — not publicly released)*

Fixed zero-output bug for fixed solids; simultaneous solid state update (oscillation fix);
oscillation guard for considered solids.

## v2.0a  *(internal — not publicly released)*

Initial solid-phase prototype: Type-3 fixed solids and Type-5 considered solids added to
architecture; outer precipitation loop introduced.

---

## v2.1a (aqueous-only line) → v2.0g notes

The v2.0x series was an internal development line focusing on solid-phase support built
on top of the validated v1.3c aqueous solver. The v2.1a release merged these two lines.

---

## v1.3c  *(internal — not publicly released)*

Charge-balance check suppressed when IFixed supplied (E14 fix). Final validated aqueous-
only version before solid-phase development began.

## v1.2b  *(internal — not publicly released)*

IS seeded from component totals instead of 1×10⁻¹²; explicit convergence Boolean flag.

## v1.1f  *(internal — not publicly released)*

Convergence robustness: `MAX_ITER` 200 → 1000, Newton step capping, overflow guards.
Validated against Ni–NTA–en system (4 components, 18 species).

## v1.1e  *(internal — not publicly released)*

Fixed component input orientation (horizontal → vertical). Fixed convergence detection.

## v1.1d  *(internal — not publicly released)*

Fixed `GoTo` inside loop, `On Error` inside loop, `Application.Max`. First working
multi-component version.

## v1.1c  *(internal — not publicly released)*

Removed all `Dim` inside loop/conditional blocks.

## v1.1b  *(internal — not publicly released)*

Renamed ionic-strength parameter `I` → `IonicStrength` to resolve VBA case-collision.

## v1.1a  *(internal — not publicly released)*

Major redesign: H⁺ promoted to full Type-1 component; species charges auto-computed;
thermodynamically correct activity model matching HYDRAQL convention.

## v1.0d  *(internal — not publicly released)*

Bug fixes: Log10 helper, Jacobian assertion, overdetermined system, convergence comparison,
Armijo condition.

## v1.0c  *(internal — not publicly released)*

Initial prototype: mass balance, mass action, charge balance, Davies activity coefficients,
damped Newton method.
