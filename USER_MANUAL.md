# AqSpeciate v2.1b — User Manual

**Aqueous equilibrium speciation solver for Microsoft Excel (VBA)**

Copyright (c) 2026 Nathan E. Boland  
Licensed under the Apache License, Version 2.0 — see [LICENSE](LICENSE) for full text.

---

## How to cite this work

> Boland, N. E. (2026). *AqSpeciate* (Version 2.1b) [Software].
> https://github.com/neboland/AqSpeciate

See `CITATION.cff` in the repository root for machine-readable metadata.

---

## Table of contents

1. [Introduction](#1-introduction)
2. [Installation](#2-installation)
3. [Chemical background](#3-chemical-background)
4. [UDF reference](#4-udf-reference)
5. [Input layout](#5-input-layout)
6. [Output interpretation](#6-output-interpretation)
7. [Worked examples](#7-worked-examples)
8. [Troubleshooting](#8-troubleshooting)
9. [Limitations](#9-limitations)
10. [References](#10-references)

---

## 1. Introduction

This project began with the desire to perform chemical equilibrium calculations directly
within Excel rather than relying on input and output from external programs (e.g. HYDRAQL
(Papelis et al., 1988), MINEQL (Westall et al., 1976)). The user-defined function developed
here uses a chemical equilibrium solver based on the traditional tableau approach
(e.g. Morel and Hering, 1993) and adheres to the species classification system adopted in
programs like HYDRAQL.

AqSpeciate solves for the equilibrium distribution of dissolved components and their
associated complex species in an aqueous system. Inputs are the total component
concentrations, cumulative formation constants, and a stoichiometry matrix. The solver
returns the free (uncomplexed) equilibrium concentration of every component and species,
corrected for non-ideal solution behaviour using the Davies equation.

The solver uses a damped Newton method with backtracking line search, Davies activity
coefficients, and optional fixed or iterative ionic strength.

The package consists of three Excel worksheet functions:

- **`AqSpeciateAll`** — returns the complete equilibrium solution as a horizontal row
- **`AqSpeciateOne`** — returns a single scalar for one component or species by ID
- **`AqSpeciateLabels`** — returns a matching row of column header labels for `AqSpeciateAll`

---

## 2. Installation

### 2.1 Pasting into a workbook

1. Open your Excel workbook and press **Alt + F11** to open the Visual Basic Editor (VBE).
2. In the VBE menu bar choose **Insert → Module**.
3. Open `AqSpeciateV2_1b.txt` in any text editor, select all (**Ctrl+A**), copy (**Ctrl+C**).
4. Click inside the new Module window in the VBE and paste (**Ctrl+V**).
5. Close the VBE and save the workbook as **Excel Macro-Enabled Workbook (.xlsm)**. 
6. Save the workbook as **Excel Macro-Enabled Workbook (.xlsm)**.
   Or alternatively, add the included **Excel Add-in Workbook (.xlam)**. (see Step 7).
6. Enable macros if prompted: File → Options → Trust Center → Macro Settings →
   Enable VBA macros.
7. Excel Add-in files can be made to open every time Excel is opened (so the UDFs are available
   in other workbooks). File → Options → Add-ins → Manage: Excel Add-ins → Go, then Browse to 
   the location on your computer where you saved the **.xlam** file and select it.
   *Note that for downloaded files, you may need to Unblock the file in File Properties*

### 2.2 Enabling macros

Macro security must allow VBA to run. In Excel: File → Options → Trust Center →
Trust Center Settings → Macro Settings → *Enable VBA macros*.

For shared workbooks you may prefer to sign the VBA project or use a trusted location.

---

## 3. Chemical background

### 3.1 Species classification

Following the classification used by HYDRAQL, AqSpeciate uses five species types:

| Type 	| Name 				| Description 	|
|---|---|---|
| **1** | Component 		| Independent chemical building block. Total concentration is a known input. |
| **2** | Aqueous species 	| Formed from Type-1 components by a mass-action equilibrium. Concentration computed by the solver. |
| **3** | Fixed solid 		| Always present; its Ksp constraint is enforced throughout the calculation, replacing the mass-balance equation of one component. |
| **5** | Considered solid 	| May or may not precipitate; the solver determines which considered solids are present at equilibrium. |

Type 4 (dissolved gas) is not implemented in the current version.

### 3.2 Formation constants and the mass-action law

All equilibrium constants are **cumulative formation constants** at infinite dilution
(ionic strength = 0) with the only reactants being Type-1 components.
The reaction is always written in the direction of *formation*:

```
ν₁ C₁ + ν₂ C₂ + … → Species j        logK_formation
```

The equilibrium concentration of species j is then:

```
log[Sⱼ] = logKf_j
           + Σᵢ νᵢⱼ · log[Cᵢ_free]
           + Σᵢ νᵢⱼ · log(γᵢ)
           − log(γ_Sⱼ)
```

where the γ values are Davies activity coefficients (see §3.4).

**Sign convention for ν:** A positive coefficient means the component is consumed to form
the species. A negative coefficient means the component is released. Examples:

- OH⁻: ν(H⁺) = −1 (one H⁺ is released, not consumed, to form OH⁻ from H₂O)
- H₂CO₃: ν(H⁺) = +2, ν(CO₃²⁻) = +1 (two H⁺ and one CO₃²⁻ combine)
- PbOH₂(s) solid: ν(H⁺) = −2, ν(Pb²⁺) = +1

### 3.3 Mass balance

For each Type-1 component k (except H⁺ when pH is fixed):

```
[Ck_free] + Σⱼ νₖⱼ · [Sⱼ] = T_k
```

where T_k is the user-supplied total molar concentration (from `CompRange`). 
The solver finds the set of free concentrations [C₁_free], [C₂_free], … that 
simultaneously satisfies all nC mass-balance equations (plus the H⁺ pH constraint if supplied).

### 3.4 Activity coefficients — Davies equation

Non-ideal activity is accounted for using the Davies equation:

```
log₁₀(γ) = −A · z² · ( √I / (1 + √I) − 0.3·I )

A = 0.509 · (298.15 / (T + 273.15))^1.5        [temperature correction]
```

where z is the ionic charge, I is the ionic strength (mol/L), and T is temperature in °C.
Neutral species (z = 0) have γ = 1. The Davies equation is reliable for I < ~0.5 mol/L.

**Note that the user does not supply species charges;** they are derived automatically from 
the stoichiometry matrix and the component charges supplied in column 4 of `CompRange`.

### 3.5 Ionic strength

When ionic strength is not fixed by the user, it is computed iteratively:

```
I = 0.5 · Σᵢ zᵢ² · [Cᵢ_free]  +  0.5 · Σⱼ zⱼ² · [Sⱼ]
```

The ionic strength is updated once per Newton iteration (one-step lag), which is sufficient
for convergence without requiring a nested inner loop.

### 3.6 Solid phases

**Fixed solids (Type 3)** are always present in the system. Their solubility-product
constraint replaces the mass-balance equation of their *primary component* (the component
with the highest stoichiometric coefficient in the solid formula, excluding pH-pinned H⁺):

```
Σᵢ νᵢ · (log[Cᵢ_free] + log₁₀(γᵢ)) + logK_formation = 0
```

When the user enters compT = 0 for a component that is supplied entirely by a fixed solid,
a *mole-ratio constraint* replaces that component's mass balance, enforcing the
dissolution stoichiometry.

**Considered solids (Type 5)** may or may not precipitate. An outer loop checks the
saturation index (SI) of each considered solid after every converged inner solve:

- If SI > 0 (supersaturated): solid activates (begins to precipitate)
- If SI < 0 (undersaturated): solid deactivates (dissolves)

The loop repeats until no solid changes state. An oscillation guard prevents indefinite
flip-flop for coupled solid pairs.

### 3.7 Saturation index

For both fixed and considered solids, AqSpeciate reports the saturation index as a p-value:

Note that K_sp = -logK_formation

```
pSI = -(logIAP - -logK_formation)
   = −(Σᵢ νᵢ · (log[Cᵢ_free] + log₁₀(γᵢ)) + logK_formation + )
```

|   pSI   |   SI   | Interpretation |
|---|---|---|
| pSI < 0 | SI > 0 | Supersaturated — solid will tend to precipitate |
| pSI = 0 | SI = 0 | At equilibrium with solid |
| pSI > 0 | SI < 0 | Undersaturated — solid will tend to dissolve |

For a converged considered solid (active), SI ≈ 0. For a non-precipitating considered
solid, SI is the residual supersaturation or undersaturation. Note that because outputs are
p-values, SI ≈ 0 will likely be a small negative integers (e.g. -0.081).

### 3.8 pH handling

When pH is provided as an optional argument:

- H⁺ free concentration is **pinned**: `log[H⁺]_free = −pH − log₁₀(γH⁺)`
- The H⁺ mass-balance equation is replaced by this pin
- H⁺ **must still appear** as the **first** Type-1 component in the component table
- The H⁺ total concentration cell may be left at 0 or any value (it is not used)

When pH is omitted, H⁺ is a free variable and its equilibrium value is determined by the
charge balance.

Note: The input pH is considered to be the -log *activity* of hydrogen ion, while 
the reported p[H⁺] output reflects the -log *concentration* of hydrogen ion, so the two 
values will differ slightly (by log₁₀(γH⁺)).

### 3.9 H⁺ and OH⁻ conventions

H⁺ **must be the first component** in the component table (first row of CompRange). The
solver identifies it as the component with charge +1 in the first row.

OH⁻ is not added automatically. Include it as a Type-2 species with:
- ν(H⁺) = −1, all other ν = 0
- logKf = logKw (typically −13.997 at 25 °C, −13.535 at 15 °C, −14.535 at 5 °C)

---

## 4. UDF reference

### 4.1 `AqSpeciateAll`

Returns the full equilibrium speciation as a **horizontal row array** of p-values. 
Gives all components/species concs sorted by ID, solid saturation indices for solids, 
computed I, and errors. Optionally appends solver diagnostics (Verbose=1).

```
=AqSpeciateAll(
    [OptionalPH],              ' optional — Fixed solution pH. When supplied, H+ free 
											concentration is pinned to 10^(-pH)/gamma(H+) 
											Omit to solve pH from charge balance.
    [OptionalTemp],            ' optional — temperature °C, default 25. Used to compute 
											the Davies activity-coefficient parameter A. 
											Reliable from approximately 0 to 60 C.
    [OptionalIFixed],          ' optional — fixes ionic strength mol/L. Fixed ionic strength 
											for Davies activity corrections. When omitted, 
											ionic strength is computed iteratively. The computed 
											value is always returned in AqSpeciateAll output.
    CompRange,                 ' required — 1 row per component (nC) x 4 columns (see §5.1) 
    SpecRange,                 ' required — 1 row per species x (2+nC) columns (see §5.2) 
    [FixedSolidsRange],        ' optional — 1 row per solid x (2+nC) columns. Same layout as 
											SpecRange. Fixed solids are always present 
											and component concs are contrained by Ksp. (see §5.3) 
    [ConsideredSolidsRange],   ' optional — 1 row per solid x (2+nC) columns. Same layout as 
											SpecRange. (see §5.4) 
    [OverrideCompID],          ' optional — integer ID of component whose total concentration 
											you want to override. Use to vary component 
											concentrations without selecting new Component Range. 
											Must be supplied together with OverrideConc.
    [OverrideConc],            ' optional — override total conc (mol/L). Total concentration 
											(mol/L, linear) of the Override Component identified by
											OverrideCompID, replacing the value in Component Range.
    [Verbose]                  ' optional — 0 for none (default) or 1 for diagnostics output (see below)
)
```

**Output column layout:**

| Columns       | Content |
|---|---|
| 1 … nC        | p[component_free] = −log₁₀([Cᵢ_free]), sorted by ascending ID |
| nC+1 … nC+nS  | p[species] = −log₁₀([Sⱼ]), sorted by ascending ID |
| nC+nS+1 … +nF | SI for each fixed solid, sorted by ascending ID |
| +nV           | SI for each considered solid, sorted by ascending ID |
| nC+nS+nF+nV+1 | Computed ionic strength (mol/L) |
| nC+nS+nF+nV+2 | Error/alert string (`""` if no issues) |

When `Verbose = 1`, eight additional columns are appended:

| Extra col | Content |
|---|---|
|    +1     | IDs of precipitated considered solids, comma-separated string |
|    +2     | Newton iteration count at convergence |
|    +3     | Convergence precision (final residual L2-norm, mol/L) |
|  +4 … +8  | log(γ) for charge magnitudes |z| = 1, 2, 3, 4, 5 |

**How to enter as an array formula (pre-Excel 365):**
Select a row of cells spanning the full output width, type the formula, and press
**Ctrl+Shift+Enter**. In Excel 365 / 2021 with dynamic arrays, press **Enter** only —
the result spills automatically.

### 4.2 `AqSpeciateOne`

Returns a single scalar p-value for one component, species, or solid by integer ID.

```
=AqSpeciateOne(
    [OptionalPH], 			   ' same as AqSpeciateAll()
	[OptionalTemp], 		   ' same as AqSpeciateAll()
	[OptionalIFixed],		   ' same as AqSpeciateAll()
    CompRange, 				   ' same as AqSpeciateAll()
	SpecRange,				   ' same as AqSpeciateAll()
    [FixedSolidsRange], 	   ' same as AqSpeciateAll()
	[ConsideredSolidsRange],   ' same as AqSpeciateAll()
    OutputID,             	   ' required — Integer ID of the component, aqueous 
											species, or solid whose value is returned.
    [OverrideCompID], 		   ' same as AqSpeciateAll()
	[OverrideConc],			   ' same as AqSpeciateAll()
    [SuppressErrors]    	   ' optional — 0 (default) or 1 to suppress #VALUE on error
)
```

Returns `−log([free])` for the requested component or species, or the saturation index
for a solid ID. Returns `#VALUE!` if the ID is not found or the solver fails (unless
`SuppressErrors = 1`).

**Typical use cases:**
- Excel Solver objective function (minimise or fix a target species concentration)
- Sensitivity analysis (vary one input, track one output)
- Obtain output for a specific species without generating a full array

### 4.3 `AqSpeciateLabels`

Returns a label row with the same column layout as `AqSpeciateAll`.

```
=AqSpeciateLabels(
    [OptionalPH], 			   ' same as AqSpeciateAll()
	[OptionalTemp], 		   ' same as AqSpeciateAll()
	[OptionalIFixed],		   ' same as AqSpeciateAll()
    CompRange, 				   ' same as AqSpeciateAll()
	SpecRange,				   ' same as AqSpeciateAll()
    [FixedSolidsRange],		   ' same as AqSpeciateAll() 
	[ConsideredSolidsRange],   ' same as AqSpeciateAll()
    [OverrideCompID], 		   ' same as AqSpeciateAll()
	[OverrideConc],			   ' same as AqSpeciateAll()
    [Verbose]				   ' same as AqSpeciateAll()
)
```

Enter this formula in the row **directly above** an `AqSpeciateAll` formula with identical
arguments to generate self-documenting column headers. Example labels:
`p[50]`, `p[8001]`, `SI[20780](fixed)`, `SI[20790]`, `Calc I (mol/L)`, `Errors`.

### 4.4 Optional argument summary

| Argument         | Default | Notes |
|--- | --- | ---|
| `OptionalPH` 	   | omitted | Pins H⁺ activity. H⁺ still required as first component. |
| `OptionalTemp`   | 25 °C   | Modifies Davies A parameter. |
| `OptionalIFixed` | omitted | Fixed ionic strength. Computed I is still reported and flagged if it exceeds IFixed. |
| `OverrideCompID` 
  / `OverrideConc` | omitted | Replaces the total concentration of one component, useful for sensitivity loops. Both must be supplied together. |
| `Verbose`        |    0    | Set to 1 to append diagnostic columns. |
| `SuppressErrors` |    0    | Set to 1 (AqSpeciateOne only) to return a value even when solver errors are present. |

---

## 5. Input layout

### 5.1 CompRange — component table (required)

A range of **nC rows × 4 columns** containing:

| Column | Content                                         | Example |
|--- | ---  | ---  |
|   1    | Integer Component ID                            | `50`    |
|   2    | Initial log guess for free concentration        | `−7`    |
|   3    | Total concentration (mol/L, **linear not log**) | `1E-4`  |
|   4    | Integer charge                                  | `+1`    |

H⁺ **must be in row 1** (first component) with charge '+1'. The integer IDs are arbitrary 
but must be unique and consistent with the IDs used in SpecRange.

**Initial guess tips:**

- For most components, start with log[free] ≈ log[total]; the solver will converge from there.
- For a strongly complexed metal (e.g. logKf > 12), set the free-concentration guess
  2–5 log units below the log total to avoid large initial residuals.
- When pH is fixed, the H⁺ initial guess is ignored (overridden immediately by the pH pin).

### 5.2 SpecRange — aqueous species table (required)

A range of **nS rows × (2 + nC) columns** containing:

| Column   | Content |
| ---  | --- |
|    1     | Integer species ID |
|    2     | Cumulative formation logK (from Type-1 components) |
| 3 … 2+nC | Stoichiometric coefficients ν for each component, one column per component in the **same order as CompRange** |

Blank cells in the stoichiometry block are read as ν = 0.

### 5.3 FixedSolidsRange — fixed solids (optional)

A range of **nF rows × (2 + nC) columns** with the same layout as SpecRange:
ID | logK_formation | stoich columns (one per component)

### 5.4 ConsideredSolidsRange — considered solids (optional)

Same layout as FixedSolidsRange: ID | logK_formation | stoich columns.

### 5.5 Input checklist

- [ ] H⁺ is row 1 of CompRange with charge +1
- [ ] SpecRange has exactly nC stoichiometry columns (2 + nC total)
- [ ] Stoichiometry column order matches component row order in CompRange
- [ ] All logK values use the cumulative formation convention (toward the solid/species)
- [ ] For hydroxide-type solids, ν(H⁺) is **negative** (H⁺ is a reactant in reverse)
- [ ] Total concentrations in CompRange column 3 are in mol/L (linear, not log)
- [ ] OH⁻ is included as a Type-2 species if needed

---

## 6. Output interpretation

### 6.1 p-values

All component and species outputs are **p-values**: p[X] = −log₁₀([X]), where [X] is the
free molar concentration. This follows standard aqueous chemistry notation (e.g. pH = p[H⁺]).

| p-value | Free concentration |
| --- | ---|
|    3    | 1 mmol/L (1×10⁻³ mol/L)   |
|    4    | 100 µmol/L (1×10⁻⁴ mol/L)  |
|    7    | 100 nmol/L (1×10⁻⁷ mol/L)  |
|   10    | 100 pmol/L (1×10⁻¹⁰ mol/L) |

### 6.2 Output ordering

Components and species are sorted by **ascending integer ID** within each block. The order
is determined by the IDs you assign, not the order of rows in the input ranges. Use
`AqSpeciateLabels` to confirm which column corresponds to which species.

### 6.3 Saturation index

Fixed solids report SI ≈ 0 at convergence (the Ksp is satisfied). Considered solids report:
- SI = 0 if precipitating (active pin)
- SI ≠ 0 if not precipitating (supersaturated if SI > 0, undersaturated if SI < 0)

Note: AqSpeciate computes SI at the *post-convergence* free concentrations (the true
equilibrium state). This differs from some programs that report SI evaluated at
pre-precipitation concentrations (e.g. HYDRAQL).

### 6.4 Error string

The last standard column contains a concatenated string of any error or alert codes
generated during the solve. An empty string means no issues. Multiple alerts are separated
by periods.

| Code   | Meaning |
|---  |---------|
| **E1** | H⁺ is not the first component — it must appear in row 1 of `CompRange` |
| **E2** | No component with charge +1 found — H⁺ is missing entirely |
| **E3** | `SpecRange` stoichiometry column count does not match nC |
| **E4** | Newton solver did not converge within `MAX_ITER` (1 000) iterations |
| **E5** | Jacobian became singular — ill-conditioned system |
| **E6** | Computed ionic strength exceeds `IFixed` by more than 0.1 % |
| **E7** | Non-positive concentration at convergence — numerical instability or bad inputs |
| **E8** | Negative component total concentration supplied |
| **E9** | `IFixed` supplied but is ≤ 0 — physically meaningless |
| **E10** | nC < 1 or nS < 1 — empty input ranges |
| **E11** | `CompRange` does not have exactly 4 columns |
| **E12** | `SpecRange` does not have at least 3 columns |
| **E14** | Charge balance at convergence exceeds 5 % of total ionic charge (only reported when `IFixed` is **not** supplied; suppressed in fixed-I mode) |
| **E15** | Zero total concentration for a non-solid-supplied component |
| **E17** | `OverrideCompID` and `OverrideConc` must both be supplied or both omitted. Override not applied |
| **E18** | A `OverrideCompID` does not match any component in `CompRange`. Override not applied |
| **E19** | Solid precipitation loop did not stabilise within 30 iterations |

If the solver fails before output can be assembled, all p-value cells return 0 and the last
cell contains the accumulated error string.

---

### 6.5 Verifying mass balance

For any unpinned component k, the following should hold to within ~10⁻¹² mol/L:

```
10^(−p[k]) + Σⱼ νₖⱼ · 10^(−p[j]) ≈ T_k
```

A consistent violation indicates an incorrect stoichiometry column (wrong order or wrong
count). Use `AqSpeciateLabels` to confirm column assignments.

---

## 7. Worked examples

### 7.1 Ethylenediamine (en) speciation

**Problem:** Calculate the equilibrium speciation of 0.1 mmol/L ethylenediamine (en) in
water at pH 5 and ionic strength I = 0.01 mol/L. What fraction of the total en is
protonated to H₂en²⁺?

**Background:** Ethylenediamine is a bidentate ligand with stepwise pKₐ values of 9.93 and
6.85 (at I = 0). At pH 5, both amine groups are below their pKₐ, so doubly-protonated
H₂en²⁺ is expected to dominate.

**Components:**

| Row | ID  | Log-guess | Total (mol/L) | Charge | (Note) |
|---|---|---|---|---|--|
|  1  | 50  |    −4     |       0       |   +1   | (H⁺) |
|  2  | 166 |    −4     |     1E-4      |    0   | (en) |

Note: H⁺ total is 0 because pH is fixed (the total is not used).

**Species:**

| ID    | logKf   | ν(H⁺)  | ν(en) | (Note) |
| ---| --- | --- | --- | --- |
| 12900 |  9.928  |   +1  |   +1   | Hen⁺ |
| 12910 | 16.776  |   +2  |   +1   | H₂en²⁺|
| 13595 | −13.997 |   −1  |    0   | OH⁻ |

**Formula:**
```
=AqSpeciateAll(5, 25, 0.01, CompRange(2x4), SpecRange(3x4))
```

**Expected results:**

| Output         | ID    | Value         | Result |
| --- | ---| ---| --- |
| p[H⁺_free]     |   50  |  4.9553        | [H⁺_free] = 10⁻⁴·⁹⁵⁵³ mol/L (slightly differs from pH due to γH⁺ correction) |
| p[en_free]     |  166  | 10.9594       | [en_free] = 1.1×10⁻¹¹ mol/L — essentially all en is complexed |
| p[H₂en²⁺]       | 12910 |  4.0045       | [H₂en²⁺] = 9.9×10⁻⁵ mol/L ≈ 99% of total en |
| p[Hen⁺]        | 12900 |  5.9867       | [Hen⁺] = 1.0×10⁻⁶ mol/L ≈ 1% of total en |
| p[OH⁻]         | 13595 |  8.9507       | [OH⁻] = 1.1×10⁻⁹ mol/L (consistent with pH 5) |
| Ionic strength |  —    | 2.0×10⁻⁴ mol/L | (returned but fixed I = 0.01 was used for γ) |

**Interpretation:** At pH 5, H₂en²⁺ accounts for ≈ 99% of the total en. The free en
concentration is vanishingly small (p = 10.96), confirming that both amine groups are
protonated at this pH. The computed mass balance: 10⁻¹⁰·⁹⁶ + 10⁻⁵·⁹⁹ + 10⁻⁴·⁰⁰ =
≈ 1.00×10⁻⁴ mol/L ✓

---

### 7.2 Nickel–NTA–ethylenediamine system

**Problem:** Calculate the equilibrium speciation of a mixture containing Ni²⁺, NTA³⁻
(nitrilotriacetic acid), and ethylenediamine (en) at pH 5, I = 0.01 mol/L. Determine
which Ni species dominates.

**Background:** This is a competitive complexation system. NTA is a tetradentate ligand
(logK_NiNTA = 12.79) and en is a bidentate ligand (logK_Ni(en) = 7.32, logB_Ni(en)2 = 13.5, and logB_Ni(en)3 = 17.61). At pH 5, proton
competition for both NTA and en is significant.

**Components:**

| ID  | Log-guess | Total (mol/L) | Charge | (Note)|
| --- | --- | --- | ---|---|
|  50 |     -5    | 0 (pH-pinned) |    1   | (H⁺) |
|  13 |    -12    | 5.00E-05      |    2   | (Ni²⁺)| 
| 165 |     -5    | 7.50E-05      |   -3   | (NTA³⁻)|
| 121 |     -8    | 5.00E-04      |    0   | (en) |

**Species:**

| ID    | log Kf | ν(H⁺) | ν(Ni²⁺) | ν(NTA³⁻) | ν(en) | Note|
| ---| ---| ---  | --- | --- |  ---|---|
| 7590  |  -9.9  | -1   |    1    |		    |    	 | (NiOH+)|
| 7491	| -19    | -2   |    1    |		    |    	 | (Ni(OH)2)|
| 7492	| -30    | -3   |    1    |		    |    	 | (Ni(OH)3-)|
| 7494	| -27.7  | -4   |    4    |		    |    	 | (Ni4(OH)4 4+)|
| 7510	| 12.79	 |      |    1    |    1    |    	 | (NiNTA)|
| 7511	| 16.96	 |      |    1    |    2    |    	 | (Ni(NTA)2)|
| 7512	| 1.51   | -1   |    1    |    1    |    	 | (Ni(OH)NTA)|
| 7513	| 19.99	 |      |    1    |    1    |    1	 | (NiNTA(en))|
| 7520	| 10.1	 |  1   |         |    1    |    	 | (H-NTA)|
| 7521	| 13.05  |  2   |         |    1    |    	 | (H2NTA)|
| 7522	| 15.07	 |  3   |         |    1    |    	 | (H3NTA)|
| 7523	| 16	 |  4   |         |    1    |    	 | (H4NTA)|
| 7610	| 7.32	 |      |    1    |         |   1	 | (Ni(en))|
| 7611	| 13.5	 |      |    1    |         |   2    | (Ni(en)2)|
| 7612	| 17.61	 |      |    1    |         |   3    | (Ni(en)3)|
| 7620	| 9.928	 |  1   |         |         |   1    | (H(en))|
| 7621	| 16.78	 |  2   |         |         |   1    | (H2(en))|
| 13595 |-13.997 | -1   |         |         |        | (OH-)|

**Formula:**
```
=AqSpeciateAll(5, 25, 0.01, CompRange(4x4), SpecRange(18x6))
```

**Expected results (selected):**

| Output       | ID   | Value | Result  |
| ---  | ---  | --- | --- |
| p[Ni²⁺_free]  |  13  | 7.076 | [Ni²⁺_free] = 8.4×10⁻⁸ (mol/L)|
| p[NTA³⁻_free] |  165 | 9.480 | [NTA³⁻_free] = 3.3×10⁻¹⁰ (mol/L)|
| p[en_free]   |  121 | 10.265 | [en_free] = 5.4×10⁻¹¹ (mol/L)|
| p[NiNTA⁻]     | 7510 | 4.302 | [NiNTA⁻]  = 5.0×10⁻⁵ (mol/L)|
| p[Ni(NTA)₂⁴⁻]  | 7511 | 9.343 | [Ni(NTA)₂⁴⁻] = 4.5×10⁻¹⁰ (mol/L)|
| p[Hen⁺]       | 7620 | 5.292 | [Hen⁺] = 5.1×10⁻⁶ (mol/L)|
| p[H₂en²⁺]     | 7621 | 6.787  | [H₂en²⁺] = 1.6×10⁻⁷ (mol/L)|
| p[H·NTA²⁻]    | 7520 | 4.603 | [H·NTA²⁻] 2.5×10⁻⁵ (mol/L) |

**Interpretation:** NiNTA⁻ (p = 4.302) is overwhelmingly the dominant Ni species —
essentially all of the Ni²⁺ is complexed by NTA, even though en is present at a 10-fold
higher total concentration. This is because:

1. The NiNTA stability constant (logK = 12.79) is much larger than Ni(en) (logK = 7.32). 
   Higher order Ni-EN complexes are not competitive because of their extra reliance on free en.
2. At pH 5, proton competition ties up most of the en as Hen⁺ and H₂en²⁺, making free
   en scarce.
3. Proton competition for NTA is also significant (H·NTA²⁻ is the dominant NTA species),
   but the NiNTA complex is still thermodynamically preferred.

The solver handles all 18 species and 4 components simultaneously in a single call.

---

### 7.3 Lead–carbonate–oxalate system with solids

**Problem:** Calculate the equilibrium speciation of Pb²⁺ in a carbonate/oxalate system
at fixed pH and ionic strength, in the presence of a fixed PbCO_3(am) solid and
considered solid phases.

**System:** pH = 9, I = 0.5 mol/L (fixed), T = 0.1 mmol/L CO₃²⁻,
and Ox²⁻. PbCO_3(am) solid is present (supplying Pb²⁺ and CO₃²⁻).

**Components:**

| ID  | Log-guess | Total (mol/L) | Charge | (Note)|
| --- | ----| ----| --- |--- |
|  50 |    -9     | 0 (pH-pinned) |   1    | (H⁺) |
|  15 |   -12     |      0        |   2    | (Pb²⁺) |
| 101 |   -10     |      0        |  −2    | (CO₃²⁻)|
| 118 |   -10     |    1E-04      |  −2    | (Ox²⁻)|

**Species:**

| ID    | log Kf  | ν(H⁺) | ν(Pb²⁺) | ν(CO₃²⁻) | ν(Ox²⁻) | Note|
| ---| ---| ---| ---| --- | --- | --- |
|  8001 |   4.2   |       |   1    |         |    1    | (PbOx)|
|  8002 |   6.31  |       |   1    |         |    2    | (PbOx2)|
|  8003 |   5.63  |   1   |   1    |         |    1    | (PbHOx)|
|  8009 |   3.8   |   1   |        |         |    1    | (HOx)|
|  8011	|   5     |   2   |        |         |    1    | (H2Ox)|
|  8013 |   7.4   |       |   1    |    1    |         | (PbCO3)|
|  8014 |  10.8   |       |   1    |    2    |         | (PbCO32)|
|  8520 |  -7.1   |  -1   |   1    |         |         | (PbOH)|
|  8530 | -16.5   |  -2   |   1    |         |         | (PbOH2)|
|  8540 | -27.4   |  -3   |   1    |         |         | (PbOH3)|
|  8550 |  -6.3   |  -1   |   2    |         |         | (Pb2OH)|
|  8551 |   9.9   |   1   |        |    1    |         | (HCO3)|
|  8552 |  16.03  |   2   |        |    1    |         | (H2CO3)
| 13595 | -13.997 |  -1   |        |         |         | (OH-)

**Fixed solid:**

| ID    | log Kf  | ν(H⁺) | ν(Pb²⁺) | ν(CO₃²⁻) | ν(Ox²⁻) | Note
| -----| ------ | -----| -----| ------ | --- |--- |
| 20780 |  13.5   |       |   +1   |   +1    |         | (PbCO_3(am)}

**Considered solids:**

| ID    | log Kf  | ν(H⁺) | ν(Pb²⁺) | ν(CO₃²⁻) | ν(Ox²⁻) | Note |
| -----| ----- | ----| -----| ---- | ------ |--- |
| 20790 |  19.4   |  −2   |    3   |    2    |    0    | (Pb3(OH)2CO3)
| 20910 |  −7.3   |  −2   |    1   |    0    |    0    | (Pb(OH)2(am))
| 20911 |   8.02  |   0   |    1   |    0    |    1    | (PbOx) 

**Formula:**
```
=AqSpeciateAll(9, 25, 0.5, CompRange(4x4), SpecRange(14x6), FixedSolidsRange(1x6), ConsideredSolidsRange(3x6), , , 1)
```

**Expected results (selected):**

| Output                 | ID    | Value | Result  |
| ------| ---   | -----| ---- |
| p[Pb²⁺]                |   15  | 10.16 | Free Pb²⁺ is very low — controlled by PbOH2 Ksp at pH 9 |
| p[CO₃²⁻]                |  101  |  2.26  | Much carbonate remains in solution  |
| SI[PbCO3(am)]          | 20780 | -0.69 |(fixed precipitate), slightly supersaturated |
| SI[PbOH2(am)]          | 20910 | -0.15 | Precipitates, slightly supersaturated |
| SI[Pb3(OH)2CO3]        | 20790 | -1.05 | Precipitates, slightly supersaturated |
| SI[PbOx]               | 20911 |  6.62 |  Undersaturated, does not precipitate in this system |
| Precipitated solid IDs |       | 20790,20910 | Both PbOH2(am) and Pb3(OH)2CO3 precipitate |

---

### 7.4 Using AqSpeciateOne for sensitivity analysis

`AqSpeciateOne` is designed for workflows that need a single output value, such as tracking
how one species responds to a changing input parameter.

**Example:** How does the free Ni²⁺ concentration vary with pH from 5 to 9 in the
Ni–NTA–en system?

Set up a column of pH values (e.g. B1:B9 = 5, 5.5, 6, …, 9) and enter in column C:

```
=AqSpeciateOne(B1, 25, 0.01, CompRange, SpecRange, 13, , , )
```

(OutputID = 13 = Ni²⁺). Copy down the column. Each cell runs a full solve and returns
p[Ni²⁺] at that pH without returning a full array.

To automate a parametric study across many conditions, combine with Excel Solver or
Data Table, using `AqSpeciateOne` as the objective.

### 7.5 Using AqSpeciateLabels for self-documenting spreadsheets

Enter `AqSpeciateLabels` in the row above `AqSpeciateAll` with the same arguments, for example:

```
Row 10: =AqSpeciateLabels(5, 25, 0.01, A6:D7, A11:C13)
Row 11: =AqSpeciateAll(5, 25, 0.01, A6:D7, A11:C13)
```

Row 10 will display labels such as: `p[50]  p[166]  p[12900]  p[12910]  p[13595]  Calc I (mol/L)  Errors`

This makes the workbook self-documenting without hardcoding column labels that can become
misaligned if components or species are added.

---

## 8. Troubleshooting 
(see Error Codes in §6.4) 

### `#VALUE!` error

1. **Macros not enabled.** File → Options → Trust Center → Enable VBA macros.
2. **Module not loaded.** Re-paste the code into a fresh VBE module.
3. **Wrong CompRange column count.** CompRange must have exactly 4 columns.
4. **SpecRange column count mismatch.** SpecRange must have 2 + nC columns. With 4
   components, use exactly 6 columns.
5. **H⁺ not in CompRange.** The solver requires H⁺ as the first component (charge +1,
   row 1). Without it, the charge-balance and pH-pin machinery cannot function.
6. **Negative total concentration.** Total concentrations in CompRange column 3 must be
   non-negative. An empty cell reads as 0 and triggers error E15, unless supplied by a fixed solid.

### `#SPILL!` error

Enter the formula into a range wide enough to show all output columns. If the formula
spills into too few cells, only `#SPILL!` is visible. Size the output range to at least
nC + nS + nF + nV + 2 columns.

### Non-convergence (error E4 in output)

1. **Improve initial guesses.** Set log[free] ≈ log[total] for each component. For metals
   with strong ligands, set the initial free-metal guess 3–5 units below log[total].
2. **Check stoichiometry signs.** For hydroxide complexes, ν(H⁺) must be negative (e.g.
   ν(H⁺) = −1 for OH⁻, −2 for M(OH)₂). A sign error can cause the solver to diverge.
3. **Check logK convention.** All logK must be *cumulative formation constants from Type-1
   components*. Do not mix stepwise constants with cumulative ones.

### Mass balance violated at convergence

- Verify that SpecRange stoichiometry columns are in the **same order** as component rows
  in CompRange. If component 165 is row 3 of CompRange, its stoichiometry must be column 3
  of the stoichiometry block (column 5 of SpecRange overall).
- Ensure SpecRange has no extra blank columns embedded in the stoichiometry block.

### Fixed solid gives zero output / `#VALUE!`

- Ensure the fixed solid is in FixedSolidsRange (the 6th argument), not SpecRange.
- If a component's total concentration is 0 because the solid is its only source, this is
  correct — the solver handles it via the mole-ratio constraint.

### Considered solid does not precipitate when expected

- Recall that the function gives pSI (-log SI)
- Verify the logK sign convention: positive logK_formation means the solid is favoured at
  low IAP. The solid activates when logIAP > −logK_formation (SI > 0 in dissolution terms).
- Check that ν(H⁺) for hydroxide solids is negative.

### p[H⁺] differs from supplied pH

This is expected. The output p[H⁺_free] = pH + log₁₀(γH⁺) reflects the free
*concentration* not the *activity*. At I = 0.01 mol/L, γH⁺ ≈ 0.903, so
p[H⁺_free] ≈ pH + 0.044.

---

## 9. Limitations

| Constraint | Detail |
| ----| -----|
| **Ionic strength range** | Davies equation is reliable for I < ~0.5 mol/L. Supply a fixed I estimated independently for saline systems. |
| **H⁺ identification** | H⁺ must be the first component (row 1 of CompRange) with charge +1. |
| **Temperature range** | The Davies A parameter approximation is reliable from ~0–60 °C. |
| **Solid phase conflicts** | When more than nC solids are simultaneously active, not all Ksp constraints can be enforced. The greedy primary-component assignment silently skips solids that cannot be accommodated. |
| **Type 4 (gas phase)** | Gas-phase equilibria are not supported. Dissolved gas concentrations can be included as fixed components with user-computed Henry's law equilibria. |
| **Redox** | No built-in electron activity (pe / Eh). Redox couples can be incorporated as explicit components with appropriate logKf values. |
| **Recalculation cost** | Each UDF call triggers a full Newton solve (typically 5–15 iterations). Avoid embedding these UDFs in very large arrays or chains that invoke them thousands of times simultaneously. |

---

## 10. References

Morel, F. M. M., & Hering, J. G. (1993). *Principles and Applications of Aquatic
Chemistry*. Wiley, New York.

Papelis, C., Hayes, K. F., & Leckie, J. O. (1988). HYDRAQL: A program for the computation
of chemical equilibrium composition of aqueous batch systems including surface-complexation
modeling of ion adsorption at the oxide/solution interface. Technical Report 306,
Department of Civil Engineering, Stanford University.

Westall, J., Zachary, J. L., & Morel, F. M. M. (1976). MINEQL: A computer program for the
calculation of chemical equilibrium composition of aqueous systems. Technical Note 18,
Department of Civil Engineering, MIT.

---

*Copyright (c) 2026 Nathan E. Boland — Apache License 2.0*
