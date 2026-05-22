# AqSpeciate

**Aqueous equilibrium speciation solver for Microsoft Excel (VBA)**

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Version](https://img.shields.io/badge/version-2.1b-green.svg)](https://github.com/neboland/AqSpeciate/releases)

---

## Overview

AqSpeciate performs aqueous chemical equilibrium speciation calculations directly within
Microsoft Excel as a set of worksheet user-defined functions (UDFs). This project began with
the desire to perform chemical equilibrium calculations directly within Excel rather than
relying on input and output from external programs (e.g. HYDRAQL (Papelis et al., 1988),
MINEQL (Westall et al., 1976)). The user-defined function developed here uses a chemical
equilibrium solver based on the traditional tableau approach (e.g. Morel and Hering, 1993)
and adheres to the species classification system adopted in programs like HYDRAQL.

Given a set of total component concentrations and cumulative formation constants, AqSpeciate
returns the equilibrium free concentrations of all components and species, corrected for
non-ideal activity using the Davies equation. It supports:

- Any number of Type-1 components and Type-2 aqueous species
- Optional fixed pH or ionic strength constraints
- Type-3 fixed solids (Ksp constraint always enforced)
- Type-5 considered solids (precipitation determined iteratively)
- Three worksheet UDFs covering different use cases

---

## How to cite this work

> Boland, N. E. (2026). *AqSpeciate* (Version 2.1b) [Software]. 
> https://github.com/neboland/AqSpeciate doi.org/10.5281/zenodo.20338878

[![DOI](https://zenodo.org/badge/1218415740.svg)](https://doi.org/10.5281/zenodo.20338878)

See [`CITATION.cff`](CITATION.cff) for machine-readable citation metadata, including
references to the foundational programs and methods that AqSpeciate builds upon.

---

## Installation

1. Open your Excel workbook and press **Alt + F11** to open the Visual Basic Editor (VBE).
2. In the VBE menu bar choose **Insert → Module**.
3. Open `AqSpeciateV2_1b.txt` in any text editor, select all (**Ctrl+A**), copy (**Ctrl+C**).
4. Click inside the new Module window in the VBE and paste (**Ctrl+V**).
5. Close the VBE and save the workbook as **Excel Macro-Enabled Workbook (.xlsm)**. 
   Or alternatively, add the included **Excel Add-in Workbook (.xlam)**. (see Step 7).
6. Enable macros if prompted: File → Options → Trust Center → Macro Settings →
   Enable VBA macros.
7. Excel Add-in files can be made to open every time Excel is opened (so the UDFs are available
   in other workbooks). File → Options → Add-ins → Manage: Excel Add-ins → Go, then Browse to 
   the location on your computer where you saved the **.xlam** file and select it.
   *Note that for downloaded files, you may need to Unblock the file in File Properties*

The three public UDFs — `AqSpeciateAll`, `AqSpeciateOne`, and `AqSpeciateLabels` — are
now available as worksheet formulas in that workbook, or if you added an Excel Add-in file, all workbooks.

---

## Quick start

### Minimal formula (aqueous species only, no solids)

```
=AqSpeciateAll(pH, Temp, IFixed, CompRange, SpecRange)
```

where:

| Argument | Content |
|---|---|
| `pH` | Optional — fixes H⁺ activity |
| `Temp` | Optional — temperature in °C (default 25) |
| `IFixed` | Optional — fixes ionic strength in mol/L |
| `CompRange` | nC rows × 4 columns: ID, log-guess conc (mol/L), total conc (mol/L), charge |
| `SpecRange` | nS rows × (2+nC) columns: ID, logKf, stoich matrix |

Returns a horizontal row array. Use `AqSpeciateLabels` with the same arguments in the
row above to generate column headers automatically.

See the [User Manual](USER_MANUAL.md) for complete argument descriptions, worked examples,
and troubleshooting guidance.

---

## Repository contents

| File | Description |
|---|---|
| `AqSpeciateV2_1b.bas` | VBA package — import into Excel module |
| `AqSpeciateV2_1b.txt` | VBA source code — paste into an Excel module |
| `AqSpeciateV2_1b.xlam` | Excel Add-in File containing VBA code |
| `CITATION.cff` | Machine-readable citation metadata |
| `LICENSE.md` | Apache 2.0 License notice |
| 'README.md' | This file |
| `USER_MANUAL.md` | Full user manual with worked examples |
| `RELEASE_NOTES.md` | Version history and change log |

---

## UDF summary

### `AqSpeciateAll`
Returns the full equilibrium solution as a **horizontal row array**. Output columns:
p[comp] values (sorted by ID) | p[species] values (sorted by ID) |
SI fixed solids | SI considered solids | computed ionic strength | error string |
[optional verbose diagnostics when `Verbose=1`]

### `AqSpeciateOne`
Returns a **single scalar** p-value (−log₁₀[free]) for one component, species, or solid
by integer ID. Suitable as an objective function for Excel Solver.

### `AqSpeciateLabels`
Returns a **label row** matching the exact column layout of `AqSpeciateAll`. Enter it in
the row above `AqSpeciateAll` for self-documenting spreadsheets.

---

## Methods

The solver is based on the traditional chemical equilibrium tableau approach described by
Morel and Hering (1993). Key features:

- **Damped Newton method** with Armijo backtracking line search for global convergence
- **Davies activity coefficients** with temperature correction (reliable for I < ~0.5 mol/L)
- **Ksp-pin mechanism** for solid phases: a solid's Ksp constraint replaces the mass-balance
  equation of its primary component; a greedy matching algorithm assigns non-conflicting
  primary components when multiple solids are active simultaneously
- **Outer precipitation loop** for considered solids, with simultaneous state updates and
  oscillation-guard logic to prevent flip-flop between precipitation and dissolution
- Convergence tolerance: residual L2-norm < 10⁻¹² mol/L
- Maximum Newton iterations: 1000

---

## AI Attribution

This project was developed collaboratively between humans and artificial intelligence:
* **Human Contributions:** Initial project concept, core feature design, bug testing, validation, and final review/edits. 
* **AI Assistance:** [Claude Sonnet 4.6](https://www.anthropic.com/news/claude-sonnet-4-6) were utilized to generate code, optimize performance, and draft documentation.

---

## References

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

## License

Copyright (c) 2026 Nathan E. Boland.
Licensed under the [Apache License, Version 2.0](LICENSE).
