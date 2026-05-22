Attribute VB_Name = "AqSpeciate2_1b"
Option Explicit

'=========================================================================
' AqSpeciate v2.1b
'
' Copyright (c) 2026 Nathan E. Boland
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
'=========================================================================
' HOW TO CITE THIS WORK
'
' Boland, N. E. (2026). AqSpeciate (Version 2.1b) [Software].
' https://github.com/neboland/AqSpeciate
' https://doi.org/10.5281/zenodo.20338879
'
' See CITATION.cff in the repository root for machine-readable metadata.
'
'=========================================================================
' CHANGE LOG v2.1a -> v2.1b  (computational efficiency review)
'
' No changes to numerical results or function signatures.
' All optimisations are internal to the Newton solver loop.
'
' OPT-A — FUSED RESIDUAL + JACOBIAN BUILD (highest impact)
'   BuildResiduals and BuildJacobian were called sequentially each iteration,
'   each looping over all nS species and calling SpecConcFromLog (containing
'   a 10^x exponentiation) once per species.  A new fused routine
'   BuildResidualsAndJacobian computes Sj once and immediately accumulates
'   into both arrays, halving the number of SpecConcFromLog calls per
'   iteration from 2*nS to nS.  ArmijoLineSearch retains BuildResiduals
'   alone (residual-only evaluation is correct there).
'
' OPT-B — TRIAL ARRAY PRE-ALLOCATED OUTSIDE ARMIJO HALVING LOOP
'   ReDim trial(1 To nC) was inside the Do While loop in ArmijoLineSearch,
'   causing up to 17 array allocations per call.  Moved outside the loop.
'
' OPT-C — GAMMA ARRAYS PRE-ALLOCATED IN SolveCore
'   BuildGammaTables ReDim'd logGammaC and logGammaS on every call (2x per
'   iteration).  Since nC and nS are fixed for the entire solve, SolveCore
'   now allocates them once before the loop and BuildGammaTables fills them
'   in-place.
'
' OPT-D — PIN LOOKUP ARRAYS PRE-ALLOCATED IN SolveCore
'   compPin() and compSecPin() (nC-element Long arrays) were ReDim'd inside
'   BuildResiduals and BuildJacobian on every call.  They are now allocated
'   once in SolveCore and passed by reference; the routines clear and refill
'   them each call without allocating.
'
' OPT-E — ReDim RHS / JacSq COPY ELIMINATED IN NEWTON LOOP
'   Each iteration copied Jacob into JacSq and negated Resid into RHS,
'   requiring two array allocations and an nC^2 element copy.  Jacob is now
'   passed directly to SolveLinear (which modifies it in-place by design,
'   and Jacob is rebuilt from scratch each iteration so this is safe).
'   RHS is pre-allocated once before the loop.
'
' OPT-H — BULK RANGE READS IN PackInputsNew AND PackSolids
'   Cell-by-cell .Value access in nested loops was replaced with a single
'   .Value2 call per range, loading all data into a Variant array in one
'   COM round-trip.  Eliminates nC*4 + nS*(2+nC) individual cell accesses
'   for components and species (e.g. 100 -> 2 calls for nC=4, nS=14).
'
'=========================================================================
'=========================================================================
' CHANGE LOG v2.0g -> v2.1a
'
' NEW 1 — IONIC STRENGTH COLUMN MOVED IN AqSpeciateAll OUTPUT
'   The computed ionic strength column now appears AFTER all solid SI
'   columns and immediately before the error string, rather than between
'   the aqueous species and the solid SI columns.  New layout:
'     cols 1..nC         : p[comp_free]
'     cols nC+1..nC+nS   : p[aq_species]
'     cols +nF           : SI fixed solids (sorted by ID)
'     cols +nV           : SI considered solids (sorted by ID)
'     col  nC+nS+nF+nV+1 : computed ionic strength (mol/L)  <- MOVED
'     col  nC+nS+nF+nV+2 : error/alert string
'     [cols +13 if Verbose=1 — see NEW 2]
'
' NEW 2 — VERBOSE DIAGNOSTIC OUTPUT FOR AqSpeciateAll (Verbose parameter)
'   A new optional final argument Verbose (default 0) may be set to 1 to
'   append 13 additional diagnostic columns to the AqSpeciateAll output:
'     +1 : IDs of precipitated considered solids, comma-separated string
'          (empty string if no considered solids precipitated)
'     +2 : Newton iteration count at convergence (Long integer)
'     +3 : Convergence precision — final residual L2-norm (mol/L)
'     +4..+8  : log10(gamma) for integer charges |z| = 1, 2, 3, 4, 5
'              (Davies equation depends on z^2, so gamma is identical
'               for +z and -z; one value per magnitude is sufficient)
'   Activity coefficients are computed from the Davies equation at the
'   converged ionic strength and temperature.
'   SolveCore was extended to store and return iter count, final residual,
'   and logGammaC values in its output array (positions nC+nS+2 onward).
'
' NEW 3 — SuppressErrors PARAMETER FOR AqSpeciateOne
'   A new optional final argument SuppressErrors (default 0) may be set
'   to 1 to prevent AqSpeciateOne from returning #VALUE when the solver
'   reports errors or fails to converge.  When SuppressErrors=1, the
'   function attempts to return the best available result.  Default
'   behaviour (SuppressErrors=0 or omitted) is unchanged: any error
'   causes an immediate #VALUE return.
'
' NEW 4 — AqSpeciateLabels UDF
'   A new public function AqSpeciateLabels accepts the identical argument
'   list as AqSpeciateAll and returns a 1-row label array whose columns
'   correspond exactly to the AqSpeciateAll output columns (including
'   verbose columns when Verbose=1).  Use it in the row above AqSpeciateAll
'   to provide column headers.  Label format:
'     Components     : "p[<ID>]"
'     Aqueous species: "p[<ID>]"
'     Fixed solid SI : "SI[<ID>](fixed)"
'     Considered SI  : "SI[<ID>]"
'     Ionic strength : "Calc I (mol/L)"
'     Error column   : "Errors"
'     Verbose fields : "Precipitated solid IDs", "Newton iterations",
'                      "Convergence precision",
'                      "log10(g) |z|=1" ... "log10(g) |z|=5"
'
'=========================================================================
'=========================================================================
' CHANGE LOG v2.0f -> v2.0g
'
' CHANGE 1 — SI SIGN CONVENTION CORRECTED TO MATCH HYDRAQL OUTPUT
'
'   In v2.0a–v2.0f, AqSpeciate computed the saturation index internally
'   using the formation convention:
'     SI_form = logIAP + logK_formation
'   where SI_form > 0 means supersaturated (solid should precipitate).
'
'   HYDRAQL uses the dissolution convention:
'     SI_diss = -logK_formation - logIAP = -SI_form
'   where SI_diss > 0 also means supersaturated.
'
'   The two conventions are numerically equal in magnitude but opposite
'   in sign.  AqSpeciate was returning SI_form in the output columns,
'   so all solid SI outputs had the wrong sign compared to HYDRAQL.
'
'   Fix: the internal activation check continues to use SI_form (formation
'   convention; no change to solver logic).  The OUTPUT columns for all
'   solid species now return -SI_form = SI_diss, matching HYDRAQL.
'
' CHANGE 2 — FIXED SOLID SI VALUES ADDED TO OUTPUT ARRAY
'
'   Previously the output array contained SI only for considered (Type 5)
'   solids.  Fixed (Type 3) solids were absent from the output, making it
'   impossible to retrieve their SI via AqSpeciateOne or map them in the
'   AqSpeciateAll output row.
'
'   Fix: nOut is extended by nF.  After the ionic strength column and
'   before the considered-solid columns, AqSpeciateAll now returns nF
'   SI values (dissolution convention, sorted by fixed solid ID).  At
'   a converged solution the Ksp constraint is satisfied, so these values
'   are zero (or a numerical residual close to zero).
'   AqSpeciateOne now also accepts a fixed solid ID as OutputID and
'   returns its SI.
'
'   New output column layout for AqSpeciateAll:
'     cols 1..nC               : p[comp_free]    sorted by component ID
'     cols nC+1..nC+nS         : p[aq_species]   sorted by species ID
'     col  nC+nS+1             : ionic strength (mol/L)
'     cols nC+nS+2..nC+nS+1+nF: SI fixed solids  sorted by solid ID
'     cols nC+nS+2+nF..end-1  : SI consid solids sorted by solid ID
'     col  end                 : error string
'
'=========================================================================
' CHANGE LOG v2.0f -> v2.0g
'
' FIX 12 — E15 FALSE POSITIVE WHEN ALL COMPONENTS OF A FIXED SOLID ARE ZERO
'
'   Root cause:
'     pinnedByFixed() was built using BestPrimary(), which identifies only
'     the SINGLE primary component for each fixed solid (the one whose
'     mass-balance row is replaced by the Ksp constraint).  All other
'     components participating in the solid (non-zero nu, non-primary) were
'     NOT marked, so the E15 zero-concentration check fired on them even
'     though having compT=0 is correct when the fixed solid is the sole
'     source of that component.
'
'   Fix:
'     pinnedByFixed() now marks EVERY component with nu != 0 in ANY fixed
'     solid (excluding pH-pinned H+).  This correctly exempts all components
'     supplied by the solid from E15, whether they are the primary (Ksp-
'     pinned) or a secondary (mole-ratio-constrained) component.
'
' CHANGE 2 — SI SIGN CONVENTION AND FIXED SOLID SI OUTPUT
'   (previously documented above — included in this release)
'
'=========================================================================
' CHANGE LOG v2.0e -> v2.0f
'
' BUG FIX 11 — #VALUE WHEN FIXED SOLID IS THE SOLE SOURCE OF A COMPONENT
'              (compT = 0 for both components of the solid)
'
'   Root cause:
'     When a fixed solid (Type 3) is the ONLY source of one or more
'     components — i.e. the user sets compT = 0 for those components
'     because their entire dissolved concentration is supplied by the
'     solid — two problems occurred:
'
'     (a) E15 "zero total concentration" fired on secondary components
'         (components with nu != 0 in the solid but not chosen as primary).
'         Only the primary component was exempt via pinnedByFixed; the
'         secondary components were not, so the early-exit branch was taken
'         before SolveCore was ever called.
'
'     (b) Even if E15 were bypassed, the mass-balance residual for a
'         secondary component with compT = 0 is:
'           R[k] = [C_k_free] + SUM_j(nu_jk*[Sj]) - 0
'         This has no positive solution because all terms on the left are
'         >= 0, so the residual can never reach zero.  The solver would
'         diverge, producing #VALUE via the FailSafe trap.
'
'   Correct treatment:
'     When the user sets compT = 0 for a component k that participates in
'     a fixed solid (nu_solid,k != 0, k != primary), this signals that
'     the solid is the SOLE source of that component.  The correct equation
'     is NOT a mass balance but a STOICHIOMETRIC RATIO CONSTRAINT derived
'     from the solid dissolution stoichiometry:
'
'       R[k] = ( [C_k] + SUM_j(nu_jk*[Sj]) ) / nu_solid,k
'            - ( [C_p] + SUM_j(nu_jp*[Sj]) ) / nu_solid,p    =  0
'
'     where p is the primary component.  This states that the molar amount
'     of component k in solution (free + complexed) relative to its solid
'     stoichiometric coefficient equals the same ratio for the primary
'     component — i.e. both came entirely from dissolving the same solid.
'     Combined with the Ksp constraint on the primary row, the system is
'     now fully determined without requiring T_k or T_p.
'
'     This mirrors the approach used by HYDRAQL when all dissolved
'     concentrations of a component originate from a single fixed solid.
'
'   Fix applied:
'     1. BuildPinArrays now identifies secondary zero-T components for each
'        pin and returns three new arrays: pinSecCount(), pinSecComp(,),
'        pinSecNu(,), and pinPrimNu().  A secondary component qualifies
'        when nu_solid,k != 0, k != primary, k != iHplusExclude, and
'        Ctot(k) = 0.
'
'     2. BuildResiduals and BuildJacobian receive these new arrays and apply
'        the stoichiometric-ratio residual / Jacobian row for each
'        secondary zero-T component instead of the mass-balance row.
'
'     3. The pinnedByFixed lookup (used for E15 exemption) is extended to
'        mark ALL components with nu != 0 in any fixed solid AND compT = 0,
'        not just the primary.  This prevents E15 from firing before the
'        solver is called.
'
'     4. A new helper SecondaryZeroT() identifies qualifying secondary
'        components given a solid row and the compT array.
'
'=========================================================================
' CHANGE LOG v2.0d -> v2.0e
'
' BUG FIX 10 — WRONG SIGN IN Ksp CONSTRAINT (root cause of all #VALUE errors
'              and incorrect results since v2.0a)
'
'   Root cause:
'     All solid logK values in the input tables (fixed solids Type 3 and
'     considered solids Type 5) follow the FORMATION convention — the same
'     convention used for aqueous species logK values throughout the program.
'     For a solid, the formation reaction is written toward the solid:
'       nu_1*C1 + nu_2*C2 + ... -> Solid    logK_formation = stored value
'     Equilibrium condition: logIAP + logK_formation = 0
'     where logIAP = SUM_i( nu_i * log{Ci} )
'
'     The code in BuildResiduals had:
'       Resid(k) = logIAP - pinLogKsp    <-- WRONG
'     which is never zero at equilibrium (it equals -2*logK_formation when
'     satisfied, i.e. always off by the full scale of the constant).
'
'     The same wrong sign appeared in the SI calculation:
'       SI(v) = logIAP - logKspV(v)     <-- WRONG
'     which gave SI values that were wrong by 2*logK, making all precipitation
'     decisions (activate/deactivate) incorrect.
'
'     For PbCO3(am) with logK_formation=13.5:
'       Correct equilibrium: logIAP = -13.5 (i.e. {Pb2+}{CO32-} = 10^-13.5)
'       Wrong residual: R = -13.5 - 13.5 = -27  (never converges to zero)
'       Correct residual: R = -13.5 + 13.5 = 0  (satisfies equilibrium)
'
'     This caused SolveCore to fail to converge (E4) for every system with
'     any solid constraint, resulting in #VALUE from the FailSafe trap.
'
'   Fix applied (three locations):
'     1. BuildResiduals:  Resid(k) = logIAP + pinLogKsp(p)
'     2. AqSpeciateAll SI loop:  SI(v) = logIAP + logKspV(v)
'     3. AqSpeciateOne SI loop:  SI(v) = logIAP + logKspV(v)
'
'   Note on SI interpretation (formation convention):
'     SI > 0  =>  logIAP > -logK_formation  =>  solution is supersaturated
'                 => solid should precipitate
'     SI = 0  =>  solution is at equilibrium with the solid
'     SI < 0  =>  solution is undersaturated => solid dissolves / does not form
'
'=========================================================================
' CHANGE LOG v2.0c -> v2.0d
'
' BUG FIX 5 — PRIMARY COMPONENT CONFLICT BETWEEN CO-PRECIPITATING SOLIDS
'             (causes #VALUE / wrong results when multiple solids are active)
'
'   Root cause:
'     In v2.0b and v2.0c, each solid's "primary component" (the component
'     whose mass-balance row is replaced by the Ksp constraint) was assigned
'     INDEPENDENTLY, once, at unpack time.  With multiple simultaneously
'     active solids this strategy fails whenever two solids independently
'     select the same primary component:
'
'     Example (Sheet 3, pH 7, fixed PbCO3(am), considered PbOH2(am)):
'       PbCO3(am) nu=[0,1,1,0] -> best non-H+ component: Pb2+ (|nu|=1)
'       PbOH2(am) nu=[-2,1,0,0] -> best non-H+ component: Pb2+ (|nu|=1)
'       Both pre-assigned primary = Pb2+.
'       In BuildPinArrays the "used" guard silently dropped the considered
'       solid's pin — PbOH2 was flagged vActive=True but contributed no
'       Ksp constraint.  The solver then failed to converge or gave wrong
'       concentrations, resulting in #VALUE or incorrect output.
'
'     The same conflict arose in Sheet 4 between PbOH2(am) [primary=Pb2+]
'     and PbOx [primary=Pb2+ or Ox2-] when both became simultaneously active.
'
'   Root cause of root cause:
'     Assigning primaries independently ignores the constraint that each
'     active solid must use a DISTINCT primary component.  The assignment
'     is inherently a matching problem (solids -> components) that must be
'     solved as a group.
'
'   Fix applied — Dynamic greedy primary assignment in BuildPinArrays:
'     The static fPrimary()/vPrimary() arrays are REMOVED.  Instead,
'     BuildPinArrays now performs the full primary-component assignment at
'     every call, using the following greedy algorithm:
'
'     Step 1: For each solid in the active set (all fixed + currently active
'             considered), build a ranked candidate list: all components
'             sorted descending by |nu_solid,i|, excluding H+ when pH is
'             fixed (iHplusExclude > 0) and excluding components whose
'             |nu| = 0.
'
'     Step 2: Sort solids by candidate-list length ascending (most-
'             constrained solid first).  This is the greedy key that
'             minimises the chance of a later solid being left without
'             a viable primary.
'
'     Step 3: Iterate through sorted solids; for each, assign the first
'             candidate that has not yet been claimed by another solid.
'             Mark that component as "used".
'
'     Step 4: If a solid has exhausted all candidates without finding a
'             free component, it is skipped (its Ksp constraint cannot be
'             enforced while all preferred components are taken).  This
'             matches the physical reality that not all combinations of
'             solids can co-precipitate within an nC-component system.
'
'     Consequence: fixed solid and considered solids now cooperate to
'     choose non-overlapping primaries.  PbCO3(am) with candidates
'     [Pb2+, CO32-] will defer to Pb2+ for PbOH2(am) and instead pin
'     CO32- — which is an equally valid Ksp constraint (the same equation
'     written from a different row of the Jacobian).
'
'     iHplusExclude is passed into BuildPinArrays and propagated through
'     the new candidate-building logic.
'
'   Changes to code structure:
'     - fPrimary() and vPrimary() arrays are REMOVED from all UDFs.
'     - pinnedByFixed() is still built (for E15 exemption), but now uses
'       BestPrimary() (a new helper) rather than PrimaryComponents().
'     - PrimaryComponents() is REPLACED by BestPrimary() (single-solid
'       best-candidate finder) and BuildPinArrays() is rewritten to do
'       the full group assignment internally.
'     - CalcLogIAP() still uses NuV, not the assigned primary, so the SI
'       check is unaffected by the primary assignment.
'
'=========================================================================
' FULL INHERITED CHANGE LIST (v1.2b through v2.0c)
'   UI CHANGE 1  : H+ must be first component (row 1 of CompRange)
'   UI CHANGE 2  : Component totals in mol/L (linear), not log10
'   UI CHANGE 3a : Optional args (pH, Temp, IFixed) first in signature
'   UI CHANGE 3b : CompRange = 4 cols (ID, log-guess, conc, charge);
'                  SpecRange = (2+nC) cols (ID, logK, stoich matrix)
'   UI CHANGE 4  : Rich error reporting in last output cell
'   FIX 1 (v1.2b): IS seeded from component totals (not 1e-12)
'   FIX 2 (v1.2b): Explicit convergence Boolean flag
'   FIX 3 (v1.3c): E14 charge balance check suppressed when IFixed supplied
'   NEW 1 (v2.0a): Fixed solids (Type 3) with Ksp-pin mechanism
'   NEW 2 (v2.0a): Considered solids (Type 5) with outer precipitation loop
'   NEW 3 (v2.0a): Component concentration override for Solver integration
'   FIX 4 (v2.0b): pinnedByFixed exemption for E15 zero-conc check
'   FIX 5 (v2.0b): Simultaneous solid state update (newActive array)
'   FIX 6 (v2.0b): Oscillation guard (vEverActive / vEverInactive)
'   FIX 7 (v2.0c): E15 also exempts H+ when pH is fixed
'   FIX 8 (v2.0c): PrimaryComponents excludes pH-pinned H+
'   FIX 9 (v2.0d): Dynamic greedy primary assignment in BuildPinArrays
'   FIX 10(v2.0e): Ksp constraint sign corrected to formation convention
'   FIX 11(v2.0f): Mole-ratio constraint for zero-T secondary solid components
'   FIX 12(v2.0g): SI sign corrected to HYDRAQL dissolution convention
'   NEW 4 (v2.0g): Fixed solid SIs added to AqSpeciateAll output array
'   NEW 5 (v2.1a): IS column moved after solid SI columns
'   NEW 6 (v2.1a): Verbose diagnostic output in AqSpeciateAll
'   NEW 7 (v2.1a): SuppressErrors parameter in AqSpeciateOne
'   NEW 8 (v2.1a): AqSpeciateLabels UDF added
'   OPT 1 (v2.1b): Fused Residual+Jacobian (OPT-A), pre-alloc arrays (B/C/D/E), bulk reads (H)
'=========================================================================

'-------------------------------------------------------------------------
' MODULE-LEVEL CONSTANTS
'-------------------------------------------------------------------------
Private Const LN10           As Double = 2.30258509299405
Private Const LOG10E         As Double = 0.434294481903252  ' 1/ln(10)
Private Const MAX_ITER       As Long   = 1000   ' Newton iterations per solve
Private Const MAX_SOLID_ITER As Long   = 30     ' outer solid loops (raised from 20 in v2.0b)
Private Const CONV_TOL       As Double = 1E-12  ' Newton convergence threshold
Private Const SOLID_TOL      As Double = 1E-6   ' IAP/Ksp threshold for precipitation check
Private Const MIN_LOG_GAMMA  As Double = -300#
Private Const MIN_I          As Double = 1E-12
Private Const ALPHA_MIN      As Double = 1E-6
Private Const MAX_LOG_STEP   As Double = 5#
Private Const LOG_CLAMP      As Double = 300#

'=========================================================================
' PUBLIC UDF 1: AqSpeciateAll  (v2.0b)
'
' Returns a 1-row array of length nC + nS + 1 + nV + 1:
'   cols 1..nC               : p[comp_free]  = -log10([Ci_free]),  sorted by ID
'   cols nC+1..nC+nS         : p[aq_spec]    = -log10([Sj]),       sorted by ID
'   col  nC+nS+1             : computed ionic strength (mol/L)
'   cols nC+nS+2..nC+nS+1+nV: SI = log10(IAP/Ksp) for each considered solid
'   col  nC+nS+2+nV          : "" or period-separated error/alert messages
'=========================================================================
Public Function AqSpeciateAll( _
    Optional OptionalPH As Variant, Optional OptionalTemp As Double = 25#, _
    Optional OptionalIFixed As Variant, _
    Optional CompRange As Variant, Optional SpecRange As Variant, _
    Optional FixedSolidsRange As Variant, Optional ConsideredSolidsRange As Variant, _
    Optional OverrideCompID As Variant, Optional OverrideConc As Variant, _
    Optional Verbose As Variant _
) As Variant

    On Error GoTo FailSafe

    '----------------------------------------------------------------------
    ' STEP 1: Variable declarations (all hoisted per VBA requirement)
    '----------------------------------------------------------------------
    Dim nC        As Long    ' number of Type-1 components
    Dim nS        As Long    ' number of Type-2 aqueous species
    Dim nF        As Long    ' number of Type-3 fixed solids
    Dim nV        As Long    ' number of Type-5 considered solids
    Dim nSpecCols As Long
    Dim nFCols    As Long
    Dim nVCols    As Long
    Dim errMsg    As String

    ' Component arrays
    Dim compID() As Long
    Dim compT()  As Double
    Dim compG()  As Double
    Dim zC()     As Double

    ' Aqueous species arrays
    Dim specID() As Long
    Dim logK()   As Double
    Dim Nu()     As Double   ' stoichiometry matrix (nS x nC)

    ' Fixed solid arrays (Type 3)
    Dim fSolID()  As Long
    Dim logKspF() As Double
    Dim NuF()     As Double  ' stoichiometry matrix (nF x nC)
    ' Note: fPrimary removed in v2.0d — primary assignment is now done
    ' dynamically inside BuildPinArrays using greedy matching.

    ' BUG FIX 1: which components are pinned by any fixed solid?
    ' Used only for E15 exemption — built via BestPrimary() per solid.
    Dim pinnedByFixed() As Boolean

    ' Considered solid arrays (Type 5)
    Dim vSolID()  As Long
    Dim logKspV() As Double
    Dim NuV()     As Double  ' stoichiometry matrix (nV x nC)
    Dim vActive()   As Boolean  ' current active status for each considered solid
    ' Note: vPrimary removed in v2.0d — assigned dynamically in BuildPinArrays.
    ' BUG FIX 2b: oscillation guard tracking arrays
    Dim vEverActive()   As Boolean
    Dim vEverInactive() As Boolean

    Dim pHFixed   As Boolean
    Dim pHval     As Double
    Dim IisFixed  As Boolean
    Dim Ival      As Double
    Dim iHplus    As Long

    Dim result    As Variant
    Dim nOut      As Long
    Dim out()     As Variant
    Dim cOrder()  As Long
    Dim sOrder()  As Long
    Dim vOrder()  As Long
    Dim i As Long, idx As Long, f As Long, v As Long
    Dim Icomputed    As Double
    Dim zS()         As Double
    Dim logGammaC()  As Double
    Dim logGammaS()  As Double
    Dim chargeBalance As Double
    Dim totalIonConc  As Double
    Dim solverErr     As String
    Dim skipConc      As Boolean   ' BUG FIX 3: used in E8/E15 exemption check
    Dim hElsewhere    As Boolean   ' used in E1/E2 H+ position check
    Dim ovrID         As Long      ' OverrideCompID as Long

    Dim CompRng  As Range
    Dim SpecRng  As Range
    Dim FSolRng  As Range
    Dim VSolRng  As Range

    Dim hasOverride  As Boolean
    Dim overrideIdx  As Long
    Dim earlyOut()   As Variant

    ' Verbose output variables (v2.1a)
    Dim verboseMode  As Boolean   ' True when Verbose=1
    Dim pptIDs       As String    ' comma-separated IDs of precipitated considered solids
    Dim iterCount    As Long      ' Newton iterations at convergence
    Dim finalResid   As Double    ' final residual norm (convergence precision)
    Dim lgCverb()    As Double    ' final logGammaC from SolveCore
    Dim nVerbExtra   As Long      ' number of extra verbose output columns

    ' Outer solid loop variables
    Dim solidIter   As Long
    Dim anyChange   As Boolean
    Dim logIAP      As Double
    Dim SI()        As Double
    Dim newActive() As Boolean  ' BUG FIX 2a: staged update array
    Dim suppressedDeact As Boolean

    ' Combined Ksp-pin arrays for SolveCore
    Dim nPins         As Long
    Dim pinComp()     As Long
    Dim pinLogKsp()   As Double
    Dim pinNu()       As Double
    Dim pinPrimNu()   As Double
    Dim pinSecCount() As Long
    Dim pinSecComp()  As Long
    Dim pinSecNu()    As Double

    ' Workspace for post-solve IAP evaluation
    Dim logCconv2() As Double
    Dim Iconvd      As Double
    Dim iHplusExclude As Long   ' BUG FIX 4: pH-pinned component excluded from PrimaryComponents
    Dim bpF           As Long   ' BestPrimary result for fixed solid E15 exemption
    Dim logCconvCB()  As Double  ' log10 free concentrations for charge balance check
    Dim ci            As Double  ' free concentration of component i
    Dim j             As Long    ' species index (charge balance loop)
    Dim Sj            As Double  ' species concentration (charge balance loop)

    '----------------------------------------------------------------------
    ' STEP 2: Validate required range arguments
    '----------------------------------------------------------------------
    errMsg = ""
    If IsMissing(CompRange) Or IsMissing(SpecRange) Then
        AqSpeciateAll = CVErr(xlErrValue)
        Exit Function
    End If

    Set CompRng = CompRange
    Set SpecRng = SpecRange

    nF = 0: nV = 0
    If Not IsMissing(FixedSolidsRange) Then
        Set FSolRng = FixedSolidsRange
        nF = FSolRng.Rows.Count
    End If
    If Not IsMissing(ConsideredSolidsRange) Then
        Set VSolRng = ConsideredSolidsRange
        nV = VSolRng.Rows.Count
    End If

    '----------------------------------------------------------------------
    ' STEP 3: Dimension checks and pre-flight error codes
    '----------------------------------------------------------------------
    nC        = CompRng.Rows.Count
    nS        = SpecRng.Rows.Count
    nSpecCols = SpecRng.Columns.Count

    If nC < 1 Or nS < 1 Then
        errMsg = AppendError(errMsg, "E10: CompRange or SpecRange is empty")
    End If
    If CompRng.Columns.Count <> 4 Then
        errMsg = AppendError(errMsg, "E11: CompRange must have exactly 4 columns " & _
                             "(ID, log-guess, concentration, charge)")
    End If
    If nSpecCols < 3 Then
        errMsg = AppendError(errMsg, "E12: SpecRange must have at least 3 columns " & _
                             "(ID, logK, stoich columns)")
    End If
    If (nSpecCols - 2) <> nC Then
        errMsg = AppendError(errMsg, "E3: Species stoichiometry column count (" & _
                             CStr(nSpecCols - 2) & ") does not match nC (" & CStr(nC) & ")")
    End If
    If nF > 0 Then
        nFCols = FSolRng.Columns.Count
        If (nFCols - 2) <> nC Then
            errMsg = AppendError(errMsg, "E20: FixedSolidsRange stoichiometry column " & _
                                 "count (" & CStr(nFCols - 2) & ") does not match nC (" & _
                                 CStr(nC) & ")")
        End If
    End If
    If nV > 0 Then
        nVCols = VSolRng.Columns.Count
        If (nVCols - 2) <> nC Then
            errMsg = AppendError(errMsg, "E21: ConsideredSolidsRange stoichiometry " & _
                                 "column count (" & CStr(nVCols - 2) & ") does not " & _
                                 "match nC (" & CStr(nC) & ")")
        End If
    End If
    If Not IsMissing(OptionalIFixed) Then
        If CDbl(OptionalIFixed) <= 0 Then
            errMsg = AppendError(errMsg, "E9: OptionalIFixed must be > 0 mol/L")
        End If
    End If

    Dim hasOverID  As Boolean: hasOverID  = Not IsMissing(OverrideCompID)
    Dim hasOverCon As Boolean: hasOverCon = Not IsMissing(OverrideConc)
    If hasOverID Xor hasOverCon Then
        errMsg = AppendError(errMsg, "E17: OverrideCompID and OverrideConc must " & _
                             "both be supplied or both omitted. Override not applied")
        hasOverride = False
    Else
        hasOverride = hasOverID And hasOverCon
    End If

    ' Fatal early exit on dimension / structural errors
    If Len(errMsg) > 0 Then
        nOut = nC + nS + 2 + nV
        If nOut < 2 Then nOut = 2
        ReDim earlyOut(1 To 1, 1 To nOut)
        earlyOut(1, nOut) = errMsg
        AqSpeciateAll = earlyOut
        Exit Function
    End If

    '----------------------------------------------------------------------
    ' STEP 4: Unpack ranges into typed arrays
    '----------------------------------------------------------------------
    Call PackInputsNew(nC, nS, CompRng, SpecRng, _
                       compID, compT, compG, zC, specID, logK, Nu)

    ' BUG FIX 3: Parse pH pin EARLY (before E15 check) so we know which
    ' component is H+ and can exempt it from the zero-concentration check.
    ' iHplus is 1 when zC(1)=+1 (H+ is first component, as required).
    iHplus  = IIf(zC(1) = 1#, 1, 0)
    pHFixed = Not IsMissing(OptionalPH)
    If pHFixed Then pHval = CDbl(OptionalPH)

    IisFixed = Not IsMissing(OptionalIFixed)
    If IisFixed Then
        Ival = CDbl(OptionalIFixed)
    Else
        Ival = SeedIonicStrength(nC, zC, compT)
    End If

    ' BUG FIX 4: Pass iHplus to PrimaryComponents so it is EXCLUDED from
    ' consideration as a primary component.  Solids with the largest |nu|
    ' on H+ (all hydroxide-type solids) must pin the next-largest component
    ' instead, because H+ is already pH-pinned and its mass-balance row
    ' cannot also be replaced by a Ksp constraint.
    ' Pass 0 when pH is not fixed (no exclusion needed).
    iHplusExclude = IIf(pHFixed, iHplus, 0)

    ' Build pinnedByFixed: marks every component exempt from E15 zero-conc check.
    ' A component is exempt if:
    '   (a) It is ANY component with non-zero nu in ANY fixed solid — the fixed
    '       solid either pins its mass balance directly (primary) or supplies it
    '       via the mole-ratio constraint (secondary with compT=0). Either way
    '       the user legitimately leaves compT=0 because the solid is the source.
    '   (b) It is the pH-pinned H+ component (handled separately below).
    ReDim pinnedByFixed(1 To nC)   ' initialised False
    If nF > 0 Then
        Call PackSolids(nF, nC, FSolRng, fSolID, logKspF, NuF)
        For f = 1 To nF
            For i = 1 To nC
                If i <> iHplusExclude Then
                    If NuF(f, i) <> 0# Then
                        pinnedByFixed(i) = True
                    End If
                End If
            Next i
        Next f
        bpF = 0   ' bpF still declared; assign dummy value to avoid uninitialised warning
    End If

    If nV > 0 Then
        Call PackSolids(nV, nC, VSolRng, vSolID, logKspV, NuV)
        ReDim vActive(1 To nV)        ' all False — no solids initially assumed present
        ReDim vEverActive(1 To nV)    ' oscillation guard (v2.0b)
        ReDim vEverInactive(1 To nV)  ' oscillation guard (v2.0b)
        ReDim SI(1 To nV)
        ' vPrimary removed: primary assignment is now done in BuildPinArrays
    End If

    '----------------------------------------------------------------------
    ' STEP 5: Post-unpack validation
    '
    ' BUG FIX 1 (v2.0b): Skip E8/E15 for solid-pinned components.
    ' BUG FIX 3 (v2.0c): Also skip H+ when pH is fixed — its total
    ' concentration cell is not used; the pH pin controls H+ directly.
    '----------------------------------------------------------------------

    ' E16 — duplicate primary components among fixed solids.
    ' In v2.0d, primary assignment is dynamic (greedy in BuildPinArrays),
    ' so fixed solids no longer conflict statically.  E16 is retired.
    ' The greedy algorithm ensures each solid gets a distinct primary at
    ' solve time, so no warning is needed here.

    ' E8 / E15 — negative or zero total concentrations.
    ' Skip: (a) components pinned by a fixed solid (pinnedByFixed)
    '        (b) H+ when pH is fixed — its free conc is set by pHFixed, not compT
    For i = 1 To nC
        skipConc = pinnedByFixed(i) Or (pHFixed And i = iHplus)
        If Not skipConc Then
            If compT(i) < 0# Then
                errMsg = AppendError(errMsg, "E8: Negative total concentration " & _
                                     "for component ID " & CStr(compID(i)))
            End If
            If compT(i) = 0# Then
                errMsg = AppendError(errMsg, "E15: Zero total concentration for " & _
                                     "component ID " & CStr(compID(i)) & _
                                     " (cell may be empty)")
            End If
        End If
    Next i

    ' E1/E2 — H+ (charge +1) must be the first component
    If zC(1) <> 1# Then
        hElsewhere = False
        For i = 2 To nC
            If zC(i) = 1# Then hElsewhere = True: Exit For
        Next i
        If hElsewhere Then
            errMsg = AppendError(errMsg, "E1: H+ (charge +1) is not the first " & _
                                 "component. H+ must appear in row 1 of CompRange")
        Else
            errMsg = AppendError(errMsg, "E2: No component with charge +1 found. " & _
                                 "H+ must be included as the first component")
        End If
    End If
    ' Note: iHplus already set in STEP 4

    ' E18 — OverrideCompID must match a known component
    If hasOverride Then
        ovrID = CLng(OverrideCompID)
        overrideIdx = 0
        For i = 1 To nC
            If compID(i) = ovrID Then overrideIdx = i: Exit For
        Next i
        If overrideIdx = 0 Then
            errMsg = AppendError(errMsg, "E18: OverrideCompID " & CStr(ovrID) & _
                                 " does not match any component in CompRange. " & _
                                 "Override not applied")
            hasOverride = False
        End If
    End If

    '----------------------------------------------------------------------
    ' STEP 6: Parse Verbose flag
    '----------------------------------------------------------------------
    verboseMode = False
    If Not IsMissing(Verbose) Then
        If CLng(Verbose) = 1 Then verboseMode = True
    End If

    '----------------------------------------------------------------------
    ' STEP 7: Outer solid precipitation loop  (v2.0b)
    '
    ' Each pass: assemble Ksp pins -> Newton solve -> evaluate ALL SI values
    ' simultaneously -> apply state changes -> repeat if any change.
    ' Oscillation guard prevents indefinite flip-flop of individual solids.
    '----------------------------------------------------------------------
    solverErr = ""
    anyChange = True
    solidIter = 0

    Do While anyChange And solidIter < MAX_SOLID_ITER
        solidIter = solidIter + 1
        anyChange = False

        ' Build combined pin array — greedy matching assigns conflict-free primaries
        Call BuildPinArrays(nC, nF, nV, logKspF, NuF, _
                            vActive, logKspV, NuV, iHplusExclude, compT, _
                            nPins, pinComp, pinLogKsp, pinNu, _
                            pinPrimNu, pinSecCount, pinSecComp, pinSecNu)

        ' Run Newton solver with this pin configuration
        result = SolveCore(nC, nS, compG, compT, zC, Nu, logK, _
                           pHFixed, pHval, iHplus, IisFixed, Ival, OptionalTemp, _
                           nPins, pinComp, pinLogKsp, pinNu, _
                           pinPrimNu, pinSecCount, pinSecComp, pinSecNu, _
                           solverErr)

        If IsError(result) Then Exit Do   ' solver failure — report below

        ' Only evaluate considered-solid SI if there are any
        If nV > 0 Then

            ' Extract log10([Ci_free]) from result for IAP calculation
            ReDim logCconv2(1 To nC)
            For i = 1 To nC
                logCconv2(i) = -CDbl(result(i))   ' result(i) = p[Ci] = -log10[Ci]
            Next i
            Iconvd = CDbl(result(nC + nS + 1))

            ' Build activity-coefficient tables at the converged ionic strength
            zS = SpeciesCharges(nC, nS, Nu, zC)
            Call BuildGammaTables(nC, nS, zC, zS, Iconvd, OptionalTemp, _
                                  logGammaC, logGammaS)

            ' BUG FIX 2a: Compute ALL SI values BEFORE changing any vActive flag
            ' SI = logIAP + logK_formation: positive = supersaturated (solid should precipitate)
            For v = 1 To nV
                logIAP = CalcLogIAP(v, nC, logCconv2, NuV, logGammaC)
                SI(v)  = logIAP + logKspV(v)   ' formation convention: SI > 0 => supersaturated
            Next v

            ' BUG FIX 2a: Stage proposed state changes in newActive()
            ReDim newActive(1 To nV)
            suppressedDeact = False

            For v = 1 To nV
                newActive(v) = vActive(v)   ' default: retain current state

                If vActive(v) Then
                    ' Currently active: check whether the solid has dissolved
                    If SI(v) < -SOLID_TOL Then
                        ' BUG FIX 2b: Oscillation guard
                        ' If this solid has previously been both active and inactive,
                        ' it is flip-flopping.  Keep it active to force a self-consistent
                        ' solution rather than allowing indefinite oscillation.
                        If vEverActive(v) And vEverInactive(v) Then
                            suppressedDeact = True  ' keep active despite SI < 0
                        Else
                            ' First deactivation — allow it
                            newActive(v)     = False
                            vEverInactive(v) = True
                            anyChange        = True
                        End If
                    End If
                    ' SI >= -SOLID_TOL => solid remains active (no change)
                Else
                    ' Currently inactive: check whether the solid should precipitate
                    If SI(v) > SOLID_TOL Then
                        newActive(v)   = True
                        vEverActive(v) = True
                        anyChange      = True
                    End If
                End If
            Next v

            ' BUG FIX 2a: Apply ALL state changes simultaneously after full evaluation
            For v = 1 To nV
                vActive(v) = newActive(v)
            Next v

            ' Log a note if the oscillation guard was invoked
            If suppressedDeact Then
                solverErr = AppendError(solverErr, _
                    "NOTE: Deactivation of one or more considered solids suppressed " & _
                    "to prevent oscillation (solid retained as active at SI < 0)")
            End If

        End If

    Loop   ' continue outer loop while any solid changed state

    ' E19 — outer loop failed to stabilise within MAX_SOLID_ITER passes
    If solidIter >= MAX_SOLID_ITER And anyChange Then
        solverErr = AppendError(solverErr, "E19: Solid precipitation loop did not " & _
                                "stabilise within " & CStr(MAX_SOLID_ITER) & _
                                " iterations (oscillating precipitation/dissolution)")
    End If

    If Len(solverErr) > 0 Then errMsg = AppendError(errMsg, solverErr)

    '----------------------------------------------------------------------
    ' STEP 7b: Extract verbose diagnostics from the last SolveCore result
    '----------------------------------------------------------------------
    If verboseMode And Not IsError(result) Then
        iterCount  = CLng(result(nC + nS + 2))
        finalResid = CDbl(result(nC + nS + 3))
        ReDim lgCverb(1 To nC)
        For i = 1 To nC
            lgCverb(i) = CDbl(result(nC + nS + 3 + i))
        Next i
        pptIDs = ""
        If nV > 0 Then
            For v = 1 To nV
                If vActive(v) Then
                    If Len(pptIDs) > 0 Then pptIDs = pptIDs & ","
                    pptIDs = pptIDs & CStr(vSolID(v))
                End If
            Next v
        End If
    End If

    '----------------------------------------------------------------------
    ' STEP 8: Assemble output array  (v2.1a layout)
    '
    ' Standard output (nC + nS + nF + nV + 2 columns):
    '   cols 1..nC               : p[comp_free]   sorted by component ID
    '   cols nC+1..nC+nS         : p[aq_species]  sorted by species ID
    '   cols nC+nS+1..nC+nS+nF  : SI fixed solids sorted by solid ID
    '   cols +nV                 : SI consid solids sorted by solid ID
    '   col  nC+nS+nF+nV+1      : computed ionic strength (mol/L)
    '   col  nC+nS+nF+nV+2      : error/alert string
    '
    ' Verbose extra columns (appended when Verbose=1):
    '   col  +1 : IDs of precipitated considered solids (comma-separated string)
    '   col  +2 : Newton iterations at convergence
    '   col  +3 : convergence precision (final residual norm, mol/L)
    '   cols +4..+8  : log10(gamma) for |z| = 1, 2, 3, 4, 5
'              (Davies equation depends on z^2; gamma same for +z and -z)
    '
    ' SI values follow HYDRAQL dissolution convention:
    '   SI > 0  =>  supersaturated  |  SI = 0  =>  equilibrium  |  SI < 0  =>  undersaturated
    '----------------------------------------------------------------------
    nVerbExtra = 0
    If verboseMode Then nVerbExtra = 8    ' 1 pptIDs + 1 iter + 1 residual + 5 gammas

    nOut = nC + nS + nF + nV + 2 + nVerbExtra
    ReDim out(1 To 1, 1 To nOut)

    ' If Newton solver failed, return zeros for numeric columns + error message
    If IsError(result) Then
        For i = 1 To nOut - 1
            out(1, i) = 0
        Next i
        out(1, nC + nS + nF + nV + 2) = errMsg
        AqSpeciateAll = out
        Exit Function
    End If

    ' Sort outputs by ascending ID for consistent column ordering
    cOrder = SortedIndices(compID, nC)
    sOrder = SortedIndices(specID, nS)

    idx = 1
    For i = 1 To nC
        out(1, idx) = result(cOrder(i))   ' p[Ci_free], sorted by compID
        idx = idx + 1
    Next i
    For i = 1 To nS
        out(1, idx) = result(nC + sOrder(i))   ' p[Sj], sorted by specID
        idx = idx + 1
    Next i

    Icomputed = CDbl(result(nC + nS + 1))   ' ionic strength (stored, placed later)

    ' Saturation indices for FIXED solids, sorted by solid ID.
    ' At convergence the Ksp constraint is satisfied, so SI = 0 exactly.
    If nF > 0 Then
        Dim fOrder() As Long
        fOrder = SortedIndices(fSolID, nF)
        Dim logCconvF() As Double
        ReDim logCconvF(1 To nC)
        For i = 1 To nC
            logCconvF(i) = -CDbl(result(i))
        Next i
        Dim IforF As Double: IforF = CDbl(result(nC + nS + 1))
        Dim lgCF() As Double, lgSF() As Double, zSF() As Double
        zSF = SpeciesCharges(nC, nS, Nu, zC)
        Call BuildGammaTables(nC, nS, zC, zSF, IforF, OptionalTemp, lgCF, lgSF)
        For i = 1 To nF
            Dim fi As Long: fi = fOrder(i)
            Dim logIAPF As Double: logIAPF = 0#
            Dim ki As Long
            For ki = 1 To nC
                If NuF(fi, ki) <> 0# Then
                    logIAPF = logIAPF + NuF(fi, ki) * (logCconvF(ki) + lgCF(ki))
                End If
            Next ki
            out(1, idx) = -(logIAPF + logKspF(fi))   ' dissolution convention
            idx = idx + 1
        Next i
    End If

    ' Saturation indices for CONSIDERED solids, sorted by solid ID.
    If nV > 0 Then
        vOrder = SortedIndices(vSolID, nV)
        For i = 1 To nV
            out(1, idx) = -SI(vOrder(i))   ' negate: formation -> dissolution convention
            idx = idx + 1
        Next i
    End If

    ' Ionic strength — placed AFTER all solid SI columns (v2.1a)
    out(1, idx) = Icomputed
    idx = idx + 1

    ' E6 — computed IS exceeds user-fixed IFixed
    If IisFixed And (Icomputed > Ival * 1.001) Then
        errMsg = AppendError(errMsg, "E6: Computed ionic strength (" & _
                             Format(Icomputed, "0.000E+00") & " mol/L) exceeds " & _
                             "fixed IFixed (" & Format(Ival, "0.000E+00") & " mol/L)")
    End If

    ' E14 — charge balance warning (suppressed when IFixed is supplied)
    If Not IisFixed Then
        ReDim logCconvCB(1 To nC)
        For i = 1 To nC
            logCconvCB(i) = -CDbl(result(i))
        Next i
        zS = SpeciesCharges(nC, nS, Nu, zC)
        Call BuildGammaTables(nC, nS, zC, zS, Icomputed, OptionalTemp, logGammaC, logGammaS)
        chargeBalance = 0#: totalIonConc = 0#
        For i = 1 To nC
            ci = 10# ^ logCconvCB(i)
            chargeBalance = chargeBalance + zC(i) * ci
            totalIonConc  = totalIonConc  + Abs(zC(i)) * ci
        Next i
        For j = 1 To nS
            Sj = SpecConcFromLog(j, nC, logCconvCB, Nu, logK, logGammaC, logGammaS(j))
            chargeBalance = chargeBalance + zS(j) * Sj
            totalIonConc  = totalIonConc  + Abs(zS(j)) * Sj
        Next j
        If totalIonConc > 0# Then
            If Abs(chargeBalance) / totalIonConc > 0.05 Then
                errMsg = AppendError(errMsg, "E14: Charge balance at convergence is " & _
                                     Format(100# * Abs(chargeBalance) / totalIonConc, "0.0") & _
                                     "% of total ionic charge (check stoichiometry)")
            End If
        End If
    End If

    ' Error string column
    out(1, idx) = errMsg
    idx = idx + 1

    ' Verbose extra columns (appended only when Verbose=1)
    If verboseMode Then
        ' Col +1: IDs of precipitated considered solids (comma-separated; "" if none)
        out(1, idx) = pptIDs
        idx = idx + 1
        ' Col +2: Newton iterations at convergence
        out(1, idx) = iterCount
        idx = idx + 1
        ' Col +3: Convergence precision (final residual norm in mol/L)
        out(1, idx) = finalResid
        idx = idx + 1
        ' Cols +4..+8: log10(gamma) for |z| = 1, 2, 3, 4, 5.
        ' Davies equation depends on z^2, so gamma is identical for +z and -z;
        ' reporting one value per magnitude avoids redundancy.
        Dim zVerbose As Long
        For zVerbose = 1 To 5
            out(1, idx) = Log10GammaDavies(CDbl(zVerbose), Icomputed, OptionalTemp)
            idx = idx + 1
        Next zVerbose
    End If

    AqSpeciateAll = out
    Exit Function

FailSafe:
    AqSpeciateAll = CVErr(xlErrValue)
End Function

'=========================================================================
' PUBLIC UDF 2: AqSpeciateOne  (v2.0b — signature unchanged from v2.0a)
'
' Returns a scalar: p[comp/species] for the given OutputID, or SI for a
' considered-solid ID.  Inherits both bug fixes from AqSpeciateAll.
'=========================================================================
Public Function AqSpeciateOne( _
    Optional OptionalPH As Variant, Optional OptionalTemp As Double = 25#, _
    Optional OptionalIFixed As Variant, _
    Optional CompRange As Variant, Optional SpecRange As Variant, _
    Optional FixedSolidsRange As Variant, Optional ConsideredSolidsRange As Variant, _
    Optional OutputID As Variant, _
    Optional OverrideCompID As Variant, Optional OverrideConc As Variant, _
    Optional SuppressErrors As Variant _
) As Variant

    On Error GoTo FailSafe

    Dim nC As Long, nS As Long, nF As Long, nV As Long
    Dim nSpecCols As Long, nFCols As Long, nVCols As Long
    Dim compID()  As Long,  compT()  As Double
    Dim compG()   As Double, zC()    As Double
    Dim specID()  As Long,  logK()   As Double,  Nu()  As Double
    Dim fSolID()  As Long,  logKspF() As Double, NuF() As Double
    ' fPrimary removed in v2.0d (dynamic greedy assignment in BuildPinArrays)
    Dim pinnedByFixed() As Boolean   ' BUG FIX 1: E15 exemption for solid-pinned components
    Dim vSolID()  As Long,  logKspV() As Double, NuV()  As Double
    Dim vActive() As Boolean
    ' vPrimary removed in v2.0d (dynamic greedy assignment in BuildPinArrays)
    Dim vEverActive()   As Boolean   ' BUG FIX 2b
    Dim vEverInactive() As Boolean   ' BUG FIX 2b
    Dim pHFixed As Boolean, pHval As Double
    Dim IisFixed As Boolean, Ival As Double
    Dim iHplus As Long
    Dim result As Variant
    Dim k As Long, v As Long, i As Long, f As Long
    Dim CompRng As Range, SpecRng As Range, FSolRng As Range, VSolRng As Range
    Dim solverErr As String
    Dim hasOverride As Boolean, overrideIdx As Long
    Dim nPins As Long
    Dim pinComp() As Long, pinLogKsp() As Double, pinNu() As Double
    Dim solidIter As Long, anyChange As Boolean
    Dim logIAP As Double, SI() As Double
    Dim logGammaC() As Double, logGammaS() As Double, zS() As Double
    Dim Iconvd As Double, logCconv() As Double
    Dim newActive() As Boolean       ' BUG FIX 2a
    Dim suppressedDeact As Boolean
    Dim iHplusExcludeOne As Long     ' BUG FIX 4: pH-pinned component excluded from PrimaryComponents
    Dim bpFOne           As Long     ' BestPrimary result for fixed solid E15 exemption
    Dim ovrIDOne         As Long     ' OverrideCompID as Long (AqSpeciateOne)
    Dim suppressErrMode  As Boolean  ' v2.1a: when True, return value even if errors exist
    ' v2.0f: secondary mole-ratio pin arrays (mirrors AqSpeciateAll)
    Dim pinPrimNu()   As Double
    Dim pinSecCount() As Long
    Dim pinSecComp()  As Long
    Dim pinSecNu()    As Double

    If IsMissing(CompRange) Or IsMissing(SpecRange) Or IsMissing(OutputID) Then
        AqSpeciateOne = CVErr(xlErrValue)
        Exit Function
    End If

    ' Parse SuppressErrors: default 0 (show errors); 1 = return value ignoring errors
    suppressErrMode = False
    If Not IsMissing(SuppressErrors) Then
        If CLng(SuppressErrors) = 1 Then suppressErrMode = True
    End If

    Set CompRng = CompRange
    Set SpecRng = SpecRange

    nC = CompRng.Rows.Count
    nS = SpecRng.Rows.Count
    nSpecCols = SpecRng.Columns.Count
    nF = 0: nV = 0

    If Not IsMissing(FixedSolidsRange) Then
        Set FSolRng = FixedSolidsRange
        nF = FSolRng.Rows.Count
    End If
    If Not IsMissing(ConsideredSolidsRange) Then
        Set VSolRng = ConsideredSolidsRange
        nV = VSolRng.Rows.Count
    End If

    ' Basic dimension checks (GoTo FailSafe on error — no rich messages)
    If nC < 1 Or nS < 1 Then GoTo FailSafe
    If CompRng.Columns.Count <> 4 Then GoTo FailSafe
    If nSpecCols < 3 Then GoTo FailSafe
    If (nSpecCols - 2) <> nC Then GoTo FailSafe
    If nF > 0 Then
        nFCols = FSolRng.Columns.Count
        If (nFCols - 2) <> nC Then GoTo FailSafe
    End If
    If nV > 0 Then
        nVCols = VSolRng.Columns.Count
        If (nVCols - 2) <> nC Then GoTo FailSafe
    End If

    ' Unpack ranges
    Call PackInputsNew(nC, nS, CompRng, SpecRng, _
                       compID, compT, compG, zC, specID, logK, Nu)

    ' BUG FIX 3: Parse iHplus and pHFixed early (before PrimaryComponents calls)
    iHplus  = IIf(zC(1) = 1#, 1, 0)
    pHFixed = Not IsMissing(OptionalPH)
    If pHFixed Then pHval = CDbl(OptionalPH)

    IisFixed = Not IsMissing(OptionalIFixed)
    If IisFixed Then
        Ival = CDbl(OptionalIFixed)
    Else
        Ival = SeedIonicStrength(nC, zC, compT)
    End If

    ' BUG FIX 4: Compute exclusion index for PrimaryComponents
    iHplusExcludeOne = IIf(pHFixed, iHplus, 0)

    ' Build pinnedByFixed: exempt ALL non-zero-nu components of fixed solids
    ' from E15. The fixed solid is the sole source for any component where
    ' the user sets compT=0, so the zero is correct — not an error.
    ReDim pinnedByFixed(1 To nC)
    If nF > 0 Then
        Call PackSolids(nF, nC, FSolRng, fSolID, logKspF, NuF)
        For f = 1 To nF
            For i = 1 To nC
                If i <> iHplusExcludeOne Then
                    If NuF(f, i) <> 0# Then
                        pinnedByFixed(i) = True
                    End If
                End If
            Next i
        Next f
        bpFOne = 0   ' dummy; BestPrimary no longer called here
    End If

    If nV > 0 Then
        Call PackSolids(nV, nC, VSolRng, vSolID, logKspV, NuV)
        ReDim vActive(1 To nV)
        ReDim vEverActive(1 To nV)
        ReDim vEverInactive(1 To nV)
        ReDim SI(1 To nV)
        ' vPrimary removed: primary assignment done in BuildPinArrays
    End If

    hasOverride = (Not IsMissing(OverrideCompID)) And (Not IsMissing(OverrideConc))
    If hasOverride Then
        ovrIDOne = CLng(OverrideCompID)
        overrideIdx = 0
        For i = 1 To nC
            If compID(i) = ovrIDOne Then overrideIdx = i: Exit For
        Next i
        If overrideIdx > 0 Then compT(overrideIdx) = CDbl(OverrideConc)
    End If

    ' Outer solid precipitation loop (same logic as AqSpeciateAll)
    solverErr = ""
    anyChange = True
    solidIter = 0

    Do While anyChange And solidIter < MAX_SOLID_ITER
        solidIter = solidIter + 1
        anyChange = False

        Call BuildPinArrays(nC, nF, nV, logKspF, NuF, _
                            vActive, logKspV, NuV, iHplusExcludeOne, compT, _
                            nPins, pinComp, pinLogKsp, pinNu, _
                            pinPrimNu, pinSecCount, pinSecComp, pinSecNu)

        result = SolveCore(nC, nS, compG, compT, zC, Nu, logK, _
                           pHFixed, pHval, iHplus, IisFixed, Ival, OptionalTemp, _
                           nPins, pinComp, pinLogKsp, pinNu, _
                           pinPrimNu, pinSecCount, pinSecComp, pinSecNu, _
                           solverErr)

        If IsError(result) Then Exit Do

        If nV > 0 Then
            ReDim logCconv(1 To nC)
            For i = 1 To nC
                logCconv(i) = -CDbl(result(i))
            Next i
            Iconvd = CDbl(result(nC + nS + 1))
            zS = SpeciesCharges(nC, nS, Nu, zC)
            Call BuildGammaTables(nC, nS, zC, zS, Iconvd, OptionalTemp, _
                                  logGammaC, logGammaS)

            ' BUG FIX 2a: Compute all SI values first (formation convention)
            For v = 1 To nV
                logIAP = CalcLogIAP(v, nC, logCconv, NuV, logGammaC)
                SI(v)  = logIAP + logKspV(v)   ' formation convention: SI > 0 => supersaturated
            Next v

            ' BUG FIX 2a/2b: Propose changes then apply simultaneously
            ReDim newActive(1 To nV)
            suppressedDeact = False
            For v = 1 To nV
                newActive(v) = vActive(v)
                If vActive(v) Then
                    If SI(v) < -SOLID_TOL Then
                        If vEverActive(v) And vEverInactive(v) Then
                            suppressedDeact = True   ' oscillation guard — keep active
                        Else
                            newActive(v)     = False
                            vEverInactive(v) = True
                            anyChange        = True
                        End If
                    End If
                Else
                    If SI(v) > SOLID_TOL Then
                        newActive(v)   = True
                        vEverActive(v) = True
                        anyChange      = True
                    End If
                End If
            Next v
            For v = 1 To nV
                vActive(v) = newActive(v)
            Next v
        End If
    Loop

    If IsError(result) Then
        ' If SuppressErrors=1, attempt to return the requested value anyway
        ' (only possible for components/species; for errors return #VALUE)
        If Not suppressErrMode Then
            AqSpeciateOne = CVErr(xlErrValue)
            Exit Function
        End If
        ' suppressErrMode=True: fall through to return whatever partial result is available
        ' For a completely failed solve this will return #VALUE from the lookup below
    End If

    ' Return the requested output value.
    ' SI outputs use dissolution convention: SI_output = -(logIAP + logK_formation)
    Dim outID As Long: outID = CLng(OutputID)
    For k = 1 To nC
        If compID(k) = outID Then AqSpeciateOne = result(k): Exit Function
    Next k
    For k = 1 To nS
        If specID(k) = outID Then AqSpeciateOne = result(nC + k): Exit Function
    Next k
    ' Considered solids: negate SI to match dissolution convention
    If nV > 0 Then
        For k = 1 To nV
            If vSolID(k) = outID Then AqSpeciateOne = -SI(k): Exit Function
        Next k
    End If
    ' Fixed solids: compute SI at convergence (= 0 at equilibrium; dissolution convention)
    If nF > 0 Then
        Dim logCconvOne() As Double
        ReDim logCconvOne(1 To nC)
        For i = 1 To nC
            logCconvOne(i) = -CDbl(result(i))
        Next i
        Dim IforOne As Double: IforOne = CDbl(result(nC + nS + 1))
        Dim lgCOne() As Double, lgSOne() As Double, zSOne() As Double
        zSOne = SpeciesCharges(nC, nS, Nu, zC)
        Call BuildGammaTables(nC, nS, zC, zSOne, IforOne, OptionalTemp, lgCOne, lgSOne)
        For k = 1 To nF
            If fSolID(k) = outID Then
                Dim logIAPOne As Double: logIAPOne = 0#
                Dim ki As Long
                For ki = 1 To nC
                    If NuF(k, ki) <> 0# Then
                        logIAPOne = logIAPOne + NuF(k, ki) * (logCconvOne(ki) + lgCOne(ki))
                    End If
                Next ki
                AqSpeciateOne = -(logIAPOne + logKspF(k))
                Exit Function
            End If
        Next k
    End If

    AqSpeciateOne = CVErr(xlErrValue)   ' OutputID not found
    Exit Function

FailSafe:
    AqSpeciateOne = CVErr(xlErrValue)
End Function

'=========================================================================
' PUBLIC UDF 3: AqSpeciateLabels  (v2.1a)
'
' Returns a 1-row label array matching the output layout of AqSpeciateAll.
' Accepts IDENTICAL arguments to AqSpeciateAll.  Instead of computing
' equilibrium concentrations, it reads the input ranges and constructs
' a human-readable label for every column that AqSpeciateAll would
' return — in exactly the same column order.
'
' Intended use: enter AqSpeciateLabels in the row ABOVE AqSpeciateAll
' over the same column span so each output column has a named header.
'
' Label format:
'   Components     : "p[<ID>]"             e.g. "p[50]"
'   Aqueous species: "p[<ID>]"             e.g. "p[8001]"
'   Fixed solid SI : "SI[<ID>](fixed)"     e.g. "SI[20780](fixed)"
'   Consid solid SI: "SI[<ID>]"            e.g. "SI[20790]"
'   Ionic strength : "Calc I (mol/L)"
'   Error string   : "Errors"
'   Verbose pptIDs : "Precipitated solid IDs"
'   Verbose iters  : "Newton iterations"
'   Verbose resid  : "Convergence precision"
'   Verbose gammas : "log10(g) |z|=1" ... "log10(g) |z|=5"
'=========================================================================
Public Function AqSpeciateLabels( _
    Optional OptionalPH As Variant, Optional OptionalTemp As Double = 25#, _
    Optional OptionalIFixed As Variant, _
    Optional CompRange As Variant, Optional SpecRange As Variant, _
    Optional FixedSolidsRange As Variant, Optional ConsideredSolidsRange As Variant, _
    Optional OverrideCompID As Variant, Optional OverrideConc As Variant, _
    Optional Verbose As Variant _
) As Variant

    On Error GoTo FailSafeLabels

    Dim nC As Long, nS As Long, nF As Long, nV As Long
    Dim compID() As Long, specID() As Long
    Dim fSolID() As Long, vSolID() As Long
    Dim nOut As Long, nVerbExtra As Long
    Dim out() As Variant
    Dim i As Long, idx As Long
    Dim cOrder() As Long, sOrder() As Long, fOrder() As Long, vOrder() As Long
    Dim verboseMode As Boolean

    ' Require at minimum CompRange and SpecRange
    If IsMissing(CompRange) Or IsMissing(SpecRange) Then
        AqSpeciateLabels = CVErr(xlErrValue)
        Exit Function
    End If

    ' Read dimensions
    Dim CompRng As Range: Set CompRng = CompRange
    Dim SpecRng As Range: Set SpecRng = SpecRange
    nC = CompRng.Rows.Count
    nS = SpecRng.Rows.Count
    nF = 0: nV = 0
    If Not IsMissing(FixedSolidsRange) Then
        Dim FSolRng As Range: Set FSolRng = FixedSolidsRange
        nF = FSolRng.Rows.Count
    End If
    If Not IsMissing(ConsideredSolidsRange) Then
        Dim VSolRng As Range: Set VSolRng = ConsideredSolidsRange
        nV = VSolRng.Rows.Count
    End If

    verboseMode = False
    If Not IsMissing(Verbose) Then
        If CLng(Verbose) = 1 Then verboseMode = True
    End If

    nVerbExtra = 0
    If verboseMode Then nVerbExtra = 8    ' 1 pptIDs + 1 iter + 1 residual + 5 gammas

    nOut = nC + nS + nF + nV + 2 + nVerbExtra
    ReDim out(1 To 1, 1 To nOut)

    ' Read IDs from ranges
    ReDim compID(1 To nC)
    Dim k As Long
    For k = 1 To nC
        compID(k) = CLng(CompRng.Cells(k, 1).Value)
    Next k

    ReDim specID(1 To nS)
    For k = 1 To nS
        specID(k) = CLng(SpecRng.Cells(k, 1).Value)
    Next k

    If nF > 0 Then
        ReDim fSolID(1 To nF)
        For k = 1 To nF
            fSolID(k) = CLng(FSolRng.Cells(k, 1).Value)
        Next k
    End If

    If nV > 0 Then
        ReDim vSolID(1 To nV)
        For k = 1 To nV
            vSolID(k) = CLng(VSolRng.Cells(k, 1).Value)
        Next k
    End If

    ' Sort IDs to match AqSpeciateAll output order
    cOrder = SortedIndices(compID, nC)
    sOrder = SortedIndices(specID, nS)

    idx = 1

    ' Component labels
    For i = 1 To nC
        out(1, idx) = "p[" & CStr(compID(cOrder(i))) & "]"
        idx = idx + 1
    Next i

    ' Aqueous species labels
    For i = 1 To nS
        out(1, idx) = "p[" & CStr(specID(sOrder(i))) & "]"
        idx = idx + 1
    Next i

    ' Fixed solid SI labels
    If nF > 0 Then
        fOrder = SortedIndices(fSolID, nF)
        For i = 1 To nF
            out(1, idx) = "SI[" & CStr(fSolID(fOrder(i))) & "](fixed)"
            idx = idx + 1
        Next i
    End If

    ' Considered solid SI labels
    If nV > 0 Then
        vOrder = SortedIndices(vSolID, nV)
        For i = 1 To nV
            out(1, idx) = "SI[" & CStr(vSolID(vOrder(i))) & "]"
            idx = idx + 1
        Next i
    End If

    ' Ionic strength label
    out(1, idx) = "Calc I (mol/L)"
    idx = idx + 1

    ' Error column label
    out(1, idx) = "Errors"
    idx = idx + 1

    ' Verbose extra labels
    If verboseMode Then
        out(1, idx) = "Precipitated solid IDs": idx = idx + 1
        out(1, idx) = "Newton iterations":      idx = idx + 1
        out(1, idx) = "Convergence precision":  idx = idx + 1
        Dim zLab As Long
        For zLab = 1 To 5
            out(1, idx) = "log10(g) |z|=" & CStr(zLab)
            idx = idx + 1
        Next zLab
    End If

    AqSpeciateLabels = out
    Exit Function

FailSafeLabels:
    AqSpeciateLabels = CVErr(xlErrValue)
End Function

'=========================================================================
' CORE SOLVER  (private) — unchanged from v2.0a
'
' Solves the Newton system for nC free component log-concentrations.
' Ksp constraints replace mass-balance rows for pinned components.
' When nPins = 0, behaviour is identical to v1.3c.
'=========================================================================
Private Function SolveCore( _
    ByVal nC As Long, ByVal nS As Long, _
    ByRef logCfree0() As Double, ByRef Ctot() As Double, _
    ByRef zC() As Double, ByRef Nu() As Double, ByRef logK() As Double, _
    ByVal pHFixed As Boolean, ByVal pHval As Double, ByVal iHplus As Long, _
    ByVal IisFixed As Boolean, ByVal Ival As Double, ByVal TempC As Double, _
    ByVal nPins As Long, ByRef pinComp() As Long, ByRef pinLogKsp() As Double, _
    ByRef pinNu() As Double, ByRef pinPrimNu() As Double, _
    ByRef pinSecCount() As Long, ByRef pinSecComp() As Long, ByRef pinSecNu() As Double, _
    ByRef solverErr As String) As Variant

    Dim logCfree()  As Double
    Dim zS()        As Double
    Dim logGammaC() As Double
    Dim logGammaS() As Double
    Dim Resid()     As Double
    Dim Jacob()     As Double
    Dim RHS()       As Double
    ' JacSq removed (OPT-E): Jacob passed directly to SolveLinear
    Dim delta()     As Double
    Dim out()       As Variant
    Dim r0          As Double
    Dim Icomputed   As Double
    Dim conc        As Double
    Dim iter        As Long
    Dim i As Long, j As Long, ii As Long, jj As Long
    Dim maxDelta    As Double
    Dim scaleFactor As Double
    Dim converged   As Boolean

    ' Initialise free-concentration guess from compG (log10 space)
    ReDim logCfree(1 To nC)
    For i = 1 To nC
        logCfree(i) = logCfree0(i)
    Next i

    zS = SpeciesCharges(nC, nS, Nu, zC)

    ' OPT-C: pre-allocate gamma arrays once before the Newton loop.
    ' BuildGammaTables now fills in-place rather than ReDim-ing each call.
    ReDim logGammaC(1 To nC)
    ReDim logGammaS(1 To nS)
    ' OPT-D: pre-allocate pin lookup arrays once; BuildResiduals/Jacobian fill them
    Dim compPin()    As Long: ReDim compPin(1 To nC)
    Dim compSecPin() As Long: ReDim compSecPin(1 To nC)
    ' OPT-E: pre-allocate RHS once; Jacob passed directly to SolveLinear
    ReDim RHS(1 To nC)

    ' Apply pH pin before gamma tables exist (uses IS seed value)
    If pHFixed And iHplus > 0 Then
        Call PinHplus(logCfree, iHplus, pHval, zC(iHplus), Ival, TempC)
    End If

    converged = False

    For iter = 1 To MAX_ITER

        ' OPT 1: precompute Davies activity-coefficient tables once per iteration
        Call BuildGammaTables(nC, nS, zC, zS, Ival, TempC, logGammaC, logGammaS)

        ' Refresh H+ pin with updated gamma value
        If pHFixed And iHplus > 0 Then
            Call PinHplusFromLog(logCfree, iHplus, pHval, logGammaC(iHplus))
        End If

        ' OPT-A: fused Residual + Jacobian build — single species pass.
        ' SpecConcFromLog is now called once per species instead of twice,
        ' halving the most expensive per-iteration computation.
        Call BuildResidualsAndJacobian(nC, nS, logCfree, Ctot, zC, zS, Nu, logK, _
                                       logGammaC, logGammaS, pHFixed, iHplus, pHval, _
                                       nPins, pinComp, pinLogKsp, pinNu, _
                                       pinPrimNu, pinSecCount, pinSecComp, pinSecNu, _
                                       compPin, compSecPin, Resid, Jacob)

        ' Check convergence on mass-balance rows (first nC elements of Resid)
        r0 = Norm2N(Resid, nC)
        If r0 < CONV_TOL Then
            converged = True
            Exit For
        End If

        ' OPT-E: Negate Resid into RHS and pass Jacob directly to SolveLinear.
        ' SolveLinear modifies its inputs in-place (by design); Jacob is
        ' rebuilt from scratch each iteration so this is safe.
        For ii = 1 To nC
            RHS(ii) = -Resid(ii)
        Next ii

        ' Compute Newton step by Gaussian elimination with partial pivoting
        delta = SolveLinear(Jacob, RHS, nC)
        If IsEmpty(delta) Then
            solverErr = AppendError(solverErr, "E5: Jacobian became singular at " & _
                                    "iteration " & CStr(iter))
            SolveCore = CVErr(xlErrValue)
            Exit Function
        End If

        ' Cap step size to MAX_LOG_STEP per component (prevents overshooting)
        maxDelta = 0#
        For ii = 1 To nC
            If Abs(delta(ii)) > maxDelta Then maxDelta = Abs(delta(ii))
        Next ii
        If maxDelta > MAX_LOG_STEP Then
            scaleFactor = MAX_LOG_STEP / maxDelta
            For ii = 1 To nC
                delta(ii) = delta(ii) * scaleFactor
            Next ii
        End If

        ' Armijo backtracking line search ensures residual decreases
        Call ArmijoLineSearch(logCfree, delta, Ctot, zC, zS, Nu, logK, _
                              logGammaC, logGammaS, pHFixed, iHplus, pHval, r0, _
                              Ival, TempC, nPins, pinComp, pinLogKsp, pinNu, _
                              pinPrimNu, pinSecCount, pinSecComp, pinSecNu)

        ' Update ionic strength (one-step lag to avoid inner IS iterations)
        If Not IisFixed Then
            Call BuildGammaTables(nC, nS, zC, zS, Ival, TempC, logGammaC, logGammaS)
            Ival = CalcIonicStrength(nC, nS, logCfree, zC, zS, Nu, logK, _
                                    logGammaC, logGammaS)
        End If

    Next iter

    If Not converged Then
        solverErr = AppendError(solverErr, "E4: Newton solver did not converge " & _
                                "within " & CStr(MAX_ITER) & " iterations")
        SolveCore = CVErr(xlErrValue)
        Exit Function
    End If

    ' Final IS at converged solution
    Call BuildGammaTables(nC, nS, zC, zS, Ival, TempC, logGammaC, logGammaS)
    Icomputed = CalcIonicStrength(nC, nS, logCfree, zC, zS, Nu, logK, _
                                  logGammaC, logGammaS)

    ' Pack output: p[Ci_free] (1..nC), p[Sj] (nC+1..nC+nS), IS,
    ' iterCount (nC+nS+2), finalResidual (nC+nS+3), logGammaC table (nC+nS+4..nC+nS+3+nC)
    ReDim out(1 To nC + nS + 3 + nC)
    For i = 1 To nC
        out(i) = -logCfree(i)   ' p[Ci] = -log10([Ci_free])
    Next i
    For i = 1 To nS
        conc = SpecConcFromLog(i, nC, logCfree, Nu, logK, logGammaC, logGammaS(i))
        If conc <= 0# Then
            solverErr = AppendError(solverErr, "E7: Non-positive concentration " & _
                                    "for species index " & CStr(i) & " at convergence")
            SolveCore = CVErr(xlErrValue)
            Exit Function
        End If
        out(nC + i) = -(Log(conc) * LOG10E)   ' p[Sj] = -log10([Sj])
    Next i
    out(nC + nS + 1) = Icomputed               ' ionic strength
    out(nC + nS + 2) = iter                    ' Newton iteration count at convergence
    out(nC + nS + 3) = Norm2N(Resid, nC)      ' final residual norm (convergence precision)
    ' Store final logGammaC values for verbose output retrieval
    For i = 1 To nC
        out(nC + nS + 3 + i) = logGammaC(i)
    Next i

    SolveCore = out
    Exit Function

FailSafe2:
    SolveCore = CVErr(xlErrValue)
End Function

'=========================================================================
' HELPER: PackSolids  (unchanged from v2.0a)
'
' Reads a solid input range (fixed or considered) into typed arrays.
' Layout: col 1 = integer ID, col 2 = log10(Ksp), cols 3..2+nC = stoich
'=========================================================================
Private Sub PackSolids( _
    ByVal nSol     As Long, _
    ByVal nC       As Long, _
    SolRng         As Range, _
    ByRef solID()  As Long, _
    ByRef logKsp() As Double, _
    ByRef NuSol()  As Double)

    ReDim solID(1 To nSol)
    ReDim logKsp(1 To nSol)
    ReDim NuSol(1 To nSol, 1 To nC)

    ' OPT-H: bulk range read (one COM call instead of nSol*(2+nC))
    Dim s As Long, i As Long
    Dim vSol As Variant
    vSol = SolRng.Value2   ' 2D array (1..nSol, 1..2+nC)
    For s = 1 To nSol
        solID(s)  = CLng(vSol(s, 1))
        logKsp(s) = CDbl(vSol(s, 2))
        For i = 1 To nC
            NuSol(s, i) = CDbl(vSol(s, 2 + i))
        Next i
    Next s
End Sub

'=========================================================================
' HELPER: BuildPinArrays  (v2.0d — completely rewritten)
'
' Assembles pinComp, pinLogKsp, pinNu for SolveCore using a greedy
' conflict-free primary-component assignment (BUG FIX 9).
'
' Each active solid needs its OWN distinct primary component — the component
' whose mass-balance row is replaced by the solid's Ksp constraint.
' Earlier versions assigned primaries independently per solid, silently
' dropping any solid that competed for an already-taken component.
'
' Algorithm:
'   1. Collect all active solids (fixed + vActive considered).
'   2. For each solid, build a candidate list: components ranked by
'      descending |nu|, excluding iHplusExclude (pH-pinned H+) and nu=0.
'   3. Sort solids by candidate-list length ascending (most-constrained first).
'   4. Greedily assign: give each solid its best still-available candidate.
'   5. Solids that run out of candidates are silently skipped.
'=========================================================================
Private Sub BuildPinArrays( _
    ByVal nC As Long, ByVal nF As Long, ByVal nV As Long, _
    ByRef logKspF() As Double, ByRef NuF() As Double, _
    ByRef vActive() As Boolean, _
    ByRef logKspV() As Double, ByRef NuV() As Double, _
    ByVal iHplusExclude As Long, ByRef Ctot() As Double, _
    ByRef nPins As Long, _
    ByRef pinComp() As Long, ByRef pinLogKsp() As Double, ByRef pinNu() As Double, _
    ByRef pinPrimNu() As Double, _
    ByRef pinSecCount() As Long, ByRef pinSecComp() As Long, ByRef pinSecNu() As Double)

    Dim f As Long, v As Long, i As Long, s As Long, j As Long, c As Long
    Dim sj      As Long
    Dim si      As Long
    Dim doSwap  As Boolean
    Dim comp    As Long

    ' Count active solids
    Dim nActive As Long: nActive = nF
    For v = 1 To nV
        If vActive(v) Then nActive = nActive + 1
    Next v

    ' Pre-allocate output arrays (min size 1)
    Dim pSize As Long: pSize = IIf(nActive > 0, nActive, 1)
    ReDim pinComp(1 To pSize)
    ReDim pinLogKsp(1 To pSize)
    ReDim pinNu(1 To pSize, 1 To nC)
    nPins = 0

    If nActive = 0 Then Exit Sub

    ' Build per-solid working arrays
    Dim solidType()   As Long:   ReDim solidType(1 To nActive)    ' 1=fixed, 2=considered
    Dim solidIdx()    As Long:   ReDim solidIdx(1 To nActive)     ' row in NuF or NuV
    Dim solidLogKsp() As Double: ReDim solidLogKsp(1 To nActive)
    Dim candComp()    As Long:   ReDim candComp(1 To nActive, 1 To nC)  ' ranked candidates
    Dim nCands()      As Long:   ReDim nCands(1 To nActive)

    s = 0
    For f = 1 To nF       ' fixed solids first
        s = s + 1
        solidType(s) = 1: solidIdx(s) = f: solidLogKsp(s) = logKspF(f)
        Call BuildCandidateList(nC, NuF, f, iHplusExclude, candComp, s, nCands(s))
    Next f
    For v = 1 To nV       ' then active considered solids
        If vActive(v) Then
            s = s + 1
            solidType(s) = 2: solidIdx(s) = v: solidLogKsp(s) = logKspV(v)
            Call BuildCandidateList(nC, NuV, v, iHplusExclude, candComp, s, nCands(s))
        End If
    Next v

    ' Sort solid indices by nCands ascending (most-constrained first);
    ' within same nCands, fixed solids (type 1) before considered (type 2).
    Dim sortOrder() As Long: ReDim sortOrder(1 To nActive)
    For s = 1 To nActive: sortOrder(s) = s: Next s
    Dim tmp As Long
    For i = 2 To nActive
        tmp = sortOrder(i): j = i - 1
        Do While j >= 1
            sj = sortOrder(j)
            si = tmp
            doSwap = False
            If nCands(sj) > nCands(si) Then
                doSwap = True
            ElseIf nCands(sj) = nCands(si) And solidType(sj) > solidType(si) Then
                doSwap = True
            End If
            If Not doSwap Then Exit Do
            sortOrder(j + 1) = sortOrder(j): j = j - 1
        Loop
        sortOrder(j + 1) = tmp
    Next i

    ' Pre-allocate mole-ratio secondary arrays (max nC secondaries per pin)
    ReDim pinPrimNu(1 To pSize)
    ReDim pinSecCount(1 To pSize)
    ReDim pinSecComp(1 To pSize, 1 To nC)
    ReDim pinSecNu(1 To pSize, 1 To nC)

    ' Greedy assignment
    Dim used() As Boolean: ReDim used(1 To nC)
    Dim nuRow() As Double: ReDim nuRow(1 To nC)
    Dim r As Long
    For i = 1 To nActive
        s = sortOrder(i)
        For c = 1 To nCands(s)
            comp = candComp(s, c)
            If comp >= 1 And comp <= nC Then
                If Not used(comp) Then
                    nPins = nPins + 1
                    pinComp(nPins) = comp
                    pinLogKsp(nPins) = solidLogKsp(s)
                    ' Copy nu vector and record primary nu value
                    If solidType(s) = 1 Then
                        For j = 1 To nC: pinNu(nPins, j) = NuF(solidIdx(s), j): Next j
                    Else
                        For j = 1 To nC: pinNu(nPins, j) = NuV(solidIdx(s), j): Next j
                    End If
                    pinPrimNu(nPins) = pinNu(nPins, comp)
                    ' Identify secondary zero-T components for mole-ratio constraint.
                    ' A secondary component qualifies when:
                    '   (a) nu_solid,k <> 0
                    '   (b) k <> primary comp
                    '   (c) k <> iHplusExclude (pH-pinned H+)
                    '   (d) Ctot(k) = 0  (sole source is this solid)
                    r = 0
                    For j = 1 To nC
                        If j <> comp And j <> iHplusExclude Then
                            If pinNu(nPins, j) <> 0# And Ctot(j) = 0# Then
                                r = r + 1
                                pinSecComp(nPins, r) = j
                                pinSecNu(nPins, r)   = pinNu(nPins, j)
                            End If
                        End If
                    Next j
                    pinSecCount(nPins) = r
                    used(comp) = True
                    Exit For   ' move to next solid
                End If
            End If
        Next c
        ' If no candidate was free, this solid is silently skipped.
    Next i
End Sub

'=========================================================================
' HELPER: BuildCandidateList  (new in v2.0d)
'
' Fills row 'row' of candComp with 1-based component indices for solid s
' (row s of NuSol), sorted descending by |nu|.  Components with nu=0
' or index = iHplusExclude are omitted.  Returns count in nCands.
'=========================================================================
Private Sub BuildCandidateList( _
    ByVal nC            As Long, _
    ByRef NuSol()       As Double, _
    ByVal s             As Long, _
    ByVal iHplusExclude As Long, _
    ByRef candComp()    As Long, _
    ByVal row           As Long, _
    ByRef nCands        As Long)

    Dim tempIdx() As Long:   ReDim tempIdx(1 To nC)
    Dim tempAbs() As Double: ReDim tempAbs(1 To nC)
    Dim k As Long, i As Long, j As Long
    k = 0
    For i = 1 To nC
        If i <> iHplusExclude And NuSol(s, i) <> 0# Then
            k = k + 1
            tempIdx(k) = i
            tempAbs(k) = Abs(NuSol(s, i))
        End If
    Next i
    nCands = k

    ' Insertion sort descending by |nu|
    Dim tmp As Long, tmpA As Double
    For i = 2 To nCands
        tmp = tempIdx(i): tmpA = tempAbs(i): j = i - 1
        Do While j >= 1 And tempAbs(j) < tmpA
            tempIdx(j + 1) = tempIdx(j): tempAbs(j + 1) = tempAbs(j): j = j - 1
        Loop
        tempIdx(j + 1) = tmp: tempAbs(j + 1) = tmpA
    Next i

    For i = 1 To nCands: candComp(row, i) = tempIdx(i): Next i
End Sub

'=========================================================================
' HELPER: BestPrimary  (new in v2.0d, replaces PrimaryComponents)
'
' Returns the single best-candidate primary component index for solid s
' (1-based row of NuSol), excluding iHplusExclude and zero-nu components.
' Used only for the conservative pinnedByFixed() E15 exemption lookup.
'=========================================================================
Private Function BestPrimary( _
    ByVal nC            As Long, _
    ByRef NuSol()       As Double, _
    ByVal s             As Long, _
    ByVal iHplusExclude As Long) As Long

    Dim maxAbs As Double: maxAbs = 0#
    Dim best   As Long:   best   = 0
    Dim i      As Long
    For i = 1 To nC
        If i <> iHplusExclude Then
            If Abs(NuSol(s, i)) > maxAbs Then
                maxAbs = Abs(NuSol(s, i)): best = i
            End If
        End If
    Next i
    If best = 0 Then   ' fallback: first non-excluded component
        For i = 1 To nC
            If i <> iHplusExclude Then best = i: Exit For
        Next i
    End If
    BestPrimary = best
End Function

'=========================================================================
' HELPER: CalcLogIAP  (unchanged from v2.0a)
'
' log10(IAP) = SUM_i( nu_s,i * (logCfree_i + logGamma_i) )
'=========================================================================
Private Function CalcLogIAP( _
    ByVal s           As Long, _
    ByVal nC          As Long, _
    ByRef logCfree()  As Double, _
    ByRef NuSol()     As Double, _
    ByRef logGammaC() As Double) As Double

    Dim logIAP As Double: logIAP = 0#
    Dim i As Long
    For i = 1 To nC
        If NuSol(s, i) <> 0# Then
            logIAP = logIAP + NuSol(s, i) * (logCfree(i) + logGammaC(i))
        End If
    Next i
    CalcLogIAP = logIAP
End Function

'=========================================================================
' OPT-A: BuildResidualsAndJacobian  (new in v2.1b)
'
' Computes both the residual vector and the Jacobian matrix in a SINGLE
' pass over the nS aqueous species.  Previously BuildResiduals and
' BuildJacobian each looped over all species independently, calling
' SpecConcFromLog (which contains a 10^x evaluation) once per species
' per call — i.e. 2*nS evaluations per Newton iteration.  The fused
' version computes Sj once and immediately uses it for both residual
' accumulation and Jacobian update, reducing to nS evaluations.
'
' ArmijoLineSearch still uses BuildResiduals alone (residual-only);
' that is correct since the Jacobian is not needed for the line search.
'=========================================================================
Private Sub BuildResidualsAndJacobian( _
    ByVal nC As Long, ByVal nS As Long, _
    ByRef logCfree() As Double, ByRef Ctot() As Double, _
    ByRef zC() As Double, ByRef zS() As Double, ByRef Nu() As Double, ByRef logK() As Double, _
    ByRef logGammaC() As Double, ByRef logGammaS() As Double, _
    ByVal pHFixed As Boolean, ByVal iHplus As Long, ByVal pHval As Double, _
    ByVal nPins As Long, ByRef pinComp() As Long, ByRef pinLogKsp() As Double, _
    ByRef pinNu() As Double, ByRef pinPrimNu() As Double, _
    ByRef pinSecCount() As Long, ByRef pinSecComp() As Long, ByRef pinSecNu() As Double, _
    ByRef compPin() As Long, ByRef compSecPin() As Long, ByRef Resid() As Double, ByRef Jacob() As Double)

    ReDim Resid(1 To nC + 1)
    ReDim Jacob(1 To nC, 1 To nC)

    Dim k As Long, m As Long, j As Long, p As Long, i As Long, r As Long
    Dim Sj As Double, qSum As Double, logCk As Double, logIAP As Double
    Dim primComp As Long, logCprim As Double
    Dim nu_jk As Double, nu_jp As Double

    ' Rebuild pin lookup tables (O(nPins))
    For k = 1 To nC: compPin(k) = 0: Next k
    For p = 1 To nPins: compPin(pinComp(p)) = p: Next p

    For k = 1 To nC: compSecPin(k) = 0: Next k
    For p = 1 To nPins
        For r = 1 To pinSecCount(p)
            compSecPin(pinSecComp(p, r)) = p
        Next r
    Next p

    ' ── Initialise diagonal / structured rows of Residual and Jacobian ──
    For k = 1 To nC
        If pHFixed And k = iHplus Then
            Resid(k)   = logCfree(k) - (-pHval - logGammaC(iHplus))
            Jacob(k,k) = 1#

        ElseIf compPin(k) > 0 Then
            p = compPin(k)
            logIAP = 0#
            For i = 1 To nC
                If pinNu(p, i) <> 0# Then
                    logIAP = logIAP + pinNu(p, i) * (logCfree(i) + logGammaC(i))
                End If
            Next i
            Resid(k) = logIAP + pinLogKsp(p)
            For m = 1 To nC: Jacob(k, m) = pinNu(p, m): Next m

        ElseIf compSecPin(k) > 0 Then
            p = compSecPin(k)
            primComp = pinComp(p)
            logCk = logCfree(k)
            If logCk >  LOG_CLAMP Then logCk =  LOG_CLAMP
            If logCk < -LOG_CLAMP Then logCk = -LOG_CLAMP
            logCprim = logCfree(primComp)
            If logCprim >  LOG_CLAMP Then logCprim =  LOG_CLAMP
            If logCprim < -LOG_CLAMP Then logCprim = -LOG_CLAMP
            Resid(k)         = (10# ^ logCk)    / pinNu(p, k) _
                             - (10# ^ logCprim) / pinPrimNu(p)
            Jacob(k, k)        = Jacob(k, k)        + LN10 * (10# ^ logCk)     / pinNu(p, k)
            Jacob(k, primComp) = Jacob(k, primComp) - LN10 * (10# ^ logCprim) / pinPrimNu(p)

        Else
            logCk = logCfree(k)
            If logCk >  LOG_CLAMP Then logCk =  LOG_CLAMP
            If logCk < -LOG_CLAMP Then logCk = -LOG_CLAMP
            Resid(k)   = 10# ^ logCk - Ctot(k)
            Jacob(k,k) = LN10 * (10# ^ logCk)
        End If
    Next k

    ' ── Single species pass: accumulate Resid AND Jacobian simultaneously ──
    qSum = 0#
    For j = 1 To nS
        Sj = SpecConcFromLog(j, nC, logCfree, Nu, logK, logGammaC, logGammaS(j))

        For k = 1 To nC
            nu_jk = Nu(j, k)
            If nu_jk = 0# Then GoTo NextK

            If pHFixed And k = iHplus Then
                ' pH-pin row: no accumulation in residual or Jacobian
            ElseIf compPin(k) > 0 Then
                ' Ksp-pin primary: residual and Jacobian row already set
            ElseIf compSecPin(k) > 0 Then
                p = compSecPin(k)
                primComp = pinComp(p)
                nu_jp = Nu(j, primComp)
                ' Accumulate mole-ratio residual
                Resid(k) = Resid(k) + nu_jk * Sj / pinNu(p, k) _
                         - nu_jp    * Sj / pinPrimNu(p)
                ' Accumulate mole-ratio Jacobian row
                For m = 1 To nC
                    If Nu(j, m) <> 0# Then
                        Jacob(k, m) = Jacob(k, m) _
                            + nu_jk * Sj * Nu(j, m) * LN10 / pinNu(p, k) _
                            - nu_jp * Sj * Nu(j, m) * LN10 / pinPrimNu(p)
                    End If
                Next m
            Else
                ' Normal mass-balance row
                Resid(k) = Resid(k) + nu_jk * Sj
                ' Jacobian off-diagonals
                For m = 1 To nC
                    If Nu(j, m) <> 0# Then
                        Jacob(k, m) = Jacob(k, m) + nu_jk * Sj * Nu(j, m) * LN10
                    End If
                Next m
            End If
NextK:
        Next k
        qSum = qSum + zS(j) * Sj
    Next j

    ' Free-component charge balance row
    For k = 1 To nC
        qSum = qSum + zC(k) * (10# ^ logCfree(k))
    Next k
    Resid(nC + 1) = qSum
End Sub

'=========================================================================
' EXTENDED: BuildResiduals  (v2.0a — unchanged in v2.0b)
'
' Residual row for each component:
'   pH-pinned H+  : logCfree(iHplus) - (-pH - logGamma)
'   Ksp-pinned    : SUM_i(nu_i*(logC_i+logGamma_i)) - logKsp
'   Normal MB     : [Ck_free] + SUM_j(nu_j,k * [Sj]) - Ctot_k
' Row nC+1 holds the charge balance (informational; not solved directly).
'=========================================================================
Private Sub BuildResiduals( _
    ByVal nC As Long, ByVal nS As Long, _
    ByRef logCfree() As Double, ByRef Ctot() As Double, _
    ByRef zC() As Double, ByRef zS() As Double, ByRef Nu() As Double, ByRef logK() As Double, _
    ByRef logGammaC() As Double, ByRef logGammaS() As Double, _
    ByVal pHFixed As Boolean, ByVal iHplus As Long, ByVal pHval As Double, _
    ByVal nPins As Long, ByRef pinComp() As Long, ByRef pinLogKsp() As Double, _
    ByRef pinNu() As Double, ByRef pinPrimNu() As Double, _
    ByRef pinSecCount() As Long, ByRef pinSecComp() As Long, ByRef pinSecNu() As Double, _
    ByRef compPin() As Long, ByRef compSecPin() As Long, ByRef Resid() As Double)

    ' OPT-D: compPin and compSecPin are pre-allocated by the caller (SolveCore)
    ' and rebuilt here each call (filling is O(nPins*nC), cheap).
    ' ReDim is avoided because the arrays are already the right size.
    ReDim Resid(1 To nC + 1)
    Dim k As Long, j As Long, p As Long, i As Long, r As Long
    Dim Sj As Double, qSum As Double, logCk As Double, logIAP As Double
    Dim primComp As Long, logCprim As Double

    ' Rebuild component-to-pin lookup (O(nPins))
    For k = 1 To nC: compPin(k) = 0: Next k
    For p = 1 To nPins
        compPin(pinComp(p)) = p
    Next p

    ' Rebuild secondary mole-ratio lookup (O(nPins*nC))
    For k = 1 To nC: compSecPin(k) = 0: Next k
    For p = 1 To nPins
        For r = 1 To pinSecCount(p)
            compSecPin(pinSecComp(p, r)) = p
        Next r
    Next p

    ' Fill residual rows
    For k = 1 To nC
        If pHFixed And k = iHplus Then
            ' pH pin: R = 0 when logCfree[H+] = -pH - logGamma[H+]
            Resid(k) = logCfree(k) - (-pHval - logGammaC(iHplus))

        ElseIf compPin(k) > 0 Then
            ' Ksp constraint (primary component of a solid).
            ' Formation convention: R = logIAP + logK_formation = 0 at equilibrium.
            p = compPin(k)
            logIAP = 0#
            For i = 1 To nC
                If pinNu(p, i) <> 0# Then
                    logIAP = logIAP + pinNu(p, i) * (logCfree(i) + logGammaC(i))
                End If
            Next i
            Resid(k) = logIAP + pinLogKsp(p)

        ElseIf compSecPin(k) > 0 Then
            ' Mole-ratio constraint for a zero-T secondary component of a fixed solid.
            ' The solid is the SOLE source of this component (Ctot(k) = 0).
            ' Stoichiometry requires that free + complexed amounts of k and the
            ' primary component p are in proportion to their solid coefficients:
            '   sumTot_k / nu_k  =  sumTot_primary / nu_primary
            ' i.e. R[k] = sumTot_k/nu_k - sumTot_primary/nu_primary = 0
            ' where sumTot_x = [C_x_free] + SUM_j(nu_jx * [Sj])
            ' Species contributions are accumulated in the species loop below.
            p = compSecPin(k)
            primComp = pinComp(p)
            logCk = logCfree(k)
            If logCk >  LOG_CLAMP Then logCk =  LOG_CLAMP
            If logCk < -LOG_CLAMP Then logCk = -LOG_CLAMP
            logCprim = logCfree(primComp)
            If logCprim >  LOG_CLAMP Then logCprim =  LOG_CLAMP
            If logCprim < -LOG_CLAMP Then logCprim = -LOG_CLAMP
            ' Initialise with free-ion terms (stoich-weighted by solid nu values)
            Resid(k) = (10# ^ logCk)    / pinNu(p, k) _
                     - (10# ^ logCprim) / pinPrimNu(p)

        Else
            ' Standard mass-balance: R[k] = [C_k_free] + species - Ctot_k
            logCk = logCfree(k)
            If logCk >  LOG_CLAMP Then logCk =  LOG_CLAMP
            If logCk < -LOG_CLAMP Then logCk = -LOG_CLAMP
            Resid(k) = 10# ^ logCk - Ctot(k)
        End If
    Next k

    ' Single pass over species: accumulate into MB and mole-ratio rows
    qSum = 0#
    For j = 1 To nS
        Sj = SpecConcFromLog(j, nC, logCfree, Nu, logK, logGammaC, logGammaS(j))
        For k = 1 To nC
            If Nu(j, k) <> 0# Then
                If pHFixed And k = iHplus Then
                    ' pH-pinned row: no accumulation
                ElseIf compPin(k) > 0 Then
                    ' Ksp primary row: no accumulation
                ElseIf compSecPin(k) > 0 Then
                    ' Mole-ratio secondary: accumulate species_k/nu_k
                    '   minus species_primary/nu_primary for species j
                    p = compSecPin(k)
                    Resid(k) = Resid(k) + Nu(j, k) * Sj / pinNu(p, k) _
                             - Nu(j, pinComp(p)) * Sj / pinPrimNu(p)
                Else
                    ' Normal mass-balance: accumulate nu_jk * [Sj]
                    Resid(k) = Resid(k) + Nu(j, k) * Sj
                End If
            End If
        Next k
        qSum = qSum + zS(j) * Sj
    Next j

    ' Free-component contribution to charge balance
    For k = 1 To nC
        qSum = qSum + zC(k) * (10# ^ logCfree(k))
    Next k
    Resid(nC + 1) = qSum   ' charge balance (not solved; used for E14 check)
End Sub

'=========================================================================
' EXTENDED: BuildJacobian  (v2.0a — unchanged in v2.0b)
'
' Jacobian rows for each component:
'   pH-pinned H+  : J(iHplus, iHplus) = 1, all other entries = 0
'   Ksp-pinned k  : J(k, m) = pinNu(p, m) for all m (replaces whole row)
'   Normal MB     : J(k, k) += ln(10)*[Ck_free];
'                   J(k, m) += SUM_j( nu_j,k * [Sj] * nu_j,m * ln(10) )
'=========================================================================
Private Sub BuildJacobian( _
    ByVal nC As Long, ByVal nS As Long, _
    ByRef logCfree() As Double, ByRef zC() As Double, _
    ByRef Nu() As Double, ByRef logK() As Double, _
    ByRef logGammaC() As Double, ByRef logGammaS() As Double, _
    ByVal pHFixed As Boolean, ByVal iHplus As Long, ByVal nPins As Long, _
    ByRef pinComp() As Long, ByRef pinNu() As Double, ByRef pinPrimNu() As Double, _
    ByRef pinSecCount() As Long, ByRef pinSecComp() As Long, ByRef pinSecNu() As Double, _
    ByRef compPin() As Long, ByRef compSecPin() As Long, ByRef Jacob() As Double)

    ReDim Jacob(1 To nC, 1 To nC)
    Dim k As Long, m As Long, j As Long, p As Long, r As Long
    Dim Sj As Double
    Dim primComp As Long

    ' OPT-D: compPin and compSecPin passed in pre-allocated; rebuild lookups here
    For k = 1 To nC: compPin(k) = 0: Next k
    For p = 1 To nPins
        compPin(pinComp(p)) = p
    Next p

    For k = 1 To nC: compSecPin(k) = 0: Next k
    For p = 1 To nPins
        For r = 1 To pinSecCount(p)
            compSecPin(pinSecComp(p, r)) = p
        Next r
    Next p

    ' Diagonal initialisation and structured row population
    For k = 1 To nC
        If pHFixed And k = iHplus Then
            ' pH-pin row: unit diagonal
            Jacob(k, k) = 1#

        ElseIf compPin(k) > 0 Then
            ' Ksp-pin row: J(k,m) = nu_solid,m for all m
            p = compPin(k)
            For m = 1 To nC
                Jacob(k, m) = pinNu(p, m)
            Next m

        ElseIf compSecPin(k) > 0 Then
            ' Mole-ratio row for zero-T secondary component k of pin p.
            ' R[k] = sumTot_k/nu_k - sumTot_primary/nu_primary
            ' d(R[k])/d(logCm):
            '   d(sumTot_k)/d(logCm)       = LN10*[Ck] (m=k) + species terms
            '   d(sumTot_primary)/d(logCm)  = LN10*[Cp] (m=primary) + species terms
            ' Species terms are added in the species loop below.
            p = compSecPin(k)
            primComp = pinComp(p)
            ' Free-ion diagonal contributions
            Jacob(k, k)        = Jacob(k, k)        + LN10 * (10# ^ logCfree(k))     / pinNu(p, k)
            Jacob(k, primComp) = Jacob(k, primComp) - LN10 * (10# ^ logCfree(primComp)) / pinPrimNu(p)

        Else
            ' Standard mass-balance diagonal: d([Ck_free])/d(logCk) = ln(10)*[Ck]
            Jacob(k, k) = LN10 * (10# ^ logCfree(k))
        End If
    Next k

    ' Off-diagonal contributions from aqueous species
    For j = 1 To nS
        Sj = SpecConcFromLog(j, nC, logCfree, Nu, logK, logGammaC, logGammaS(j))
        For k = 1 To nC
            If pHFixed And k = iHplus Then
                ' pH-pin: no species contributions
            ElseIf compPin(k) > 0 Then
                ' Ksp-pin: Jacobian row already fully set above
            ElseIf compSecPin(k) > 0 Then
                ' Mole-ratio row: add species contributions
                ' d(nu_jk*Sj/nu_k)/d(logCm) - d(nu_j_primary*Sj/nu_primary)/d(logCm)
                '   = [nu_jk * Sj * nu_jm * LN10] / nu_k
                '   - [nu_j_primary * Sj * nu_jm * LN10] / nu_primary
                p = compSecPin(k)
                primComp = pinComp(p)
                If Nu(j, k) <> 0# Or Nu(j, primComp) <> 0# Then
                    For m = 1 To nC
                        If Nu(j, m) <> 0# Then
                            Jacob(k, m) = Jacob(k, m) _
                                + Nu(j, k)        * Sj * Nu(j, m) * LN10 / pinNu(p, k) _
                                - Nu(j, primComp) * Sj * Nu(j, m) * LN10 / pinPrimNu(p)
                        End If
                    Next m
                End If
            Else
                ' Normal mass-balance: accumulate species Jacobian contributions
                If Nu(j, k) <> 0# Then
                    For m = 1 To nC
                        If Nu(j, m) <> 0# Then
                            Jacob(k, m) = Jacob(k, m) + Nu(j, k) * Sj * Nu(j, m) * LN10
                        End If
                    Next m
                End If
            End If
        Next k
    Next j
End Sub

'=========================================================================
' HELPER: ArmijoLineSearch  (v2.0a — unchanged in v2.0b)
'
' Finds step size alpha that satisfies sufficient decrease condition.
' Halves alpha until ||Resid(x+alpha*d)|| < ||Resid(x)||, then applies.
' Falls back to ALPHA_MIN step if all halvings fail.
'=========================================================================
Private Sub ArmijoLineSearch( _
    ByRef logCfree() As Double, ByRef delta() As Double, ByRef Ctot() As Double, _
    ByRef zC() As Double, ByRef zS() As Double, _
    ByRef Nu() As Double, ByRef logK() As Double, _
    ByRef logGammaC() As Double, ByRef logGammaS() As Double, _
    ByVal pHFixed As Boolean, ByVal iHplus As Long, ByVal pHval As Double, _
    ByVal r0 As Double, ByVal IonicStrength As Double, ByVal TempC As Double, _
    ByVal nPins As Long, ByRef pinComp() As Long, ByRef pinLogKsp() As Double, _
    ByRef pinNu() As Double, ByRef pinPrimNu() As Double, _
    ByRef pinSecCount() As Long, ByRef pinSecComp() As Long, ByRef pinSecNu() As Double)

    Dim alpha   As Double: alpha = 1#
    Dim trial() As Double
    Dim Rtrial() As Double
    Dim nC      As Long: nC = UBound(logCfree)
    Dim iComp   As Long

    ' OPT-B: allocate trial once outside the halving loop
    ReDim trial(1 To nC)
    ReDim Rtrial(1 To nC + 1)
    ' Local lookup arrays for BuildResiduals (acceptable allocation; called infrequently)
    Dim armCompPin()    As Long: ReDim armCompPin(1 To nC)
    Dim armCompSecPin() As Long: ReDim armCompSecPin(1 To nC)

    Do While alpha > ALPHA_MIN
        For iComp = 1 To nC
            trial(iComp) = logCfree(iComp) + alpha * delta(iComp)
        Next iComp

        If pHFixed And iHplus > 0 Then
            Call PinHplusFromLog(trial, iHplus, pHval, logGammaC(iHplus))
        End If

        ' OPT-D: Armijo has its own local compPin/compSecPin (it's called from
        ' SolveCore which passes its arrays, but ArmijoLineSearch is a separate sub).
        ' We accept the small per-halving-step overhead here since line searches
        ' are infrequent (typically 0-1 per iteration near convergence).
        Call BuildResiduals(nC, UBound(logK), trial, Ctot, zC, zS, Nu, logK, _
                            logGammaC, logGammaS, pHFixed, iHplus, pHval, _
                            nPins, pinComp, pinLogKsp, pinNu, _
                            pinPrimNu, pinSecCount, pinSecComp, pinSecNu, _
                            armCompPin, armCompSecPin, Rtrial)

        If Norm2N(Rtrial, nC) < r0 Then
            For iComp = 1 To nC
                logCfree(iComp) = trial(iComp)
            Next iComp
            Exit Sub
        End If
        alpha = alpha * 0.5#
    Loop

    ' Fallback: apply minimal step
    For iComp = 1 To nC
        logCfree(iComp) = logCfree(iComp) + ALPHA_MIN * delta(iComp)
    Next iComp
End Sub

'=========================================================================
' INHERITED HELPERS (unchanged from v1.3c / v2.0a)
'=========================================================================

Private Function AppendError(ByVal existing As String, _
                              ByVal newMsg   As String) As String
    If Len(existing) = 0 Then
        AppendError = newMsg
    Else
        AppendError = existing & ". " & newMsg
    End If
End Function

'-------------------------------------------------------------------------
' PackInputsNew — reads CompRange (4 cols) and SpecRange (2+nC cols)
'-------------------------------------------------------------------------
Private Sub PackInputsNew( _
    ByVal nC As Long, ByVal nS As Long, _
    CompRng As Range, SpecRng As Range, _
    ByRef compID() As Long, ByRef compT() As Double, _
    ByRef compG() As Double, ByRef zC() As Double, _
    ByRef specID() As Long, ByRef logK() As Double, ByRef Nu() As Double)

    ReDim compID(1 To nC): ReDim compT(1 To nC)
    ReDim compG(1 To nC):  ReDim zC(1 To nC)
    ReDim specID(1 To nS): ReDim logK(1 To nS)
    ReDim Nu(1 To nS, 1 To nC)

    ' OPT-H: read entire ranges in two bulk .Value2 calls instead of
    ' cell-by-cell access, eliminating nC*4 + nS*(2+nC) COM round-trips.
    Dim i As Long, j As Long
    Dim vComp As Variant, vSpec As Variant
    vComp = CompRng.Value2   ' 2D array (1..nC, 1..4)
    vSpec = SpecRng.Value2   ' 2D array (1..nS, 1..2+nC)

    For i = 1 To nC
        compID(i) = CLng(vComp(i, 1))
        compG(i)  = CDbl(vComp(i, 2))
        compT(i)  = CDbl(vComp(i, 3))
        zC(i)     = CDbl(vComp(i, 4))
    Next i

    For j = 1 To nS
        specID(j) = CLng(vSpec(j, 1))
        logK(j)   = CDbl(vSpec(j, 2))
        For i = 1 To nC
            Nu(j, i) = CDbl(vSpec(j, 2 + i))
        Next i
    Next j
End Sub

'-------------------------------------------------------------------------
' SeedIonicStrength — upper-bound IS estimate (I = 0.5 * SUM(z_i^2 * T_i))
'-------------------------------------------------------------------------
Private Function SeedIonicStrength( _
    ByVal nC     As Long, _
    ByRef zC()   As Double, _
    ByRef Ctot() As Double) As Double

    Dim Iseed As Double, i As Long
    For i = 1 To nC
        If zC(i) <> 0# Then Iseed = Iseed + 0.5# * zC(i) ^ 2# * Ctot(i)
    Next i
    SeedIonicStrength = WorksheetFunction.Max(Iseed, MIN_I)
End Function

'-------------------------------------------------------------------------
' BuildGammaTables — Davies activity coefficients for all components/species
'-------------------------------------------------------------------------
Private Sub BuildGammaTables( _
    ByVal nC            As Long, _
    ByVal nS            As Long, _
    ByRef zC()          As Double, _
    ByRef zS()          As Double, _
    ByVal IonicStrength As Double, _
    ByVal TempC         As Double, _
    ByRef logGammaC()   As Double, _
    ByRef logGammaS()   As Double)

    ' Arrays are ReDim'd unconditionally. SolveCore pre-allocates them before
    ' the Newton loop (so these ReDims are cheap no-ops there), and external
    ' callers (post-convergence sections, charge-balance checks) pass
    ' uninitialized arrays that must be allocated here. Using UBound() to
    ' test initialization is unsafe in VBA — it raises error 9 on an
    ' uninitialized array rather than returning a sentinel value.
    Dim i As Long
    Dim A     As Double: A   = 0.509# * ((298.15# / (TempC + 273.15#)) ^ 1.5#)
    Dim sqI   As Double: sqI = Sqr(WorksheetFunction.Max(IonicStrength, 0#))
    Dim dTerm As Double: dTerm = sqI / (1# + sqI) - 0.3# * IonicStrength

    ReDim logGammaC(1 To nC)
    For i = 1 To nC
        If zC(i) = 0# Then
            logGammaC(i) = 0#
        Else
            logGammaC(i) = -A * zC(i) ^ 2# * dTerm
            If logGammaC(i) < MIN_LOG_GAMMA Then logGammaC(i) = MIN_LOG_GAMMA
        End If
    Next i

    ReDim logGammaS(1 To nS)
    For i = 1 To nS
        If zS(i) = 0# Then
            logGammaS(i) = 0#
        Else
            logGammaS(i) = -A * zS(i) ^ 2# * dTerm
            If logGammaS(i) < MIN_LOG_GAMMA Then logGammaS(i) = MIN_LOG_GAMMA
        End If
    Next i
End Sub

'-------------------------------------------------------------------------
' PinHplus / PinHplusFromLog — enforce pH constraint on H+ component
'-------------------------------------------------------------------------
Private Sub PinHplus( _
    ByRef logCfree()    As Double, _
    ByVal iHplus        As Long, _
    ByVal pHval         As Double, _
    ByVal zHplus        As Double, _
    ByVal IonicStrength As Double, _
    ByVal TempC         As Double)
    logCfree(iHplus) = -pHval - Log10GammaDavies(zHplus, IonicStrength, TempC)
End Sub

Private Sub PinHplusFromLog( _
    ByRef logCfree()    As Double, _
    ByVal iHplus        As Long, _
    ByVal pHval         As Double, _
    ByVal logGammaHplus As Double)
    logCfree(iHplus) = -pHval - logGammaHplus
End Sub

'-------------------------------------------------------------------------
' SpecConcFromLog — [Sj] in mol/L from log10-space mass-action equation
'-------------------------------------------------------------------------
Private Function SpecConcFromLog( _
    ByVal j           As Long, _
    ByVal nC          As Long, _
    ByRef logCfree()  As Double, _
    ByRef Nu()        As Double, _
    ByRef logK()      As Double, _
    ByRef logGammaC() As Double, _
    ByVal logGammaSj  As Double) As Double

    Dim logS As Double: logS = logK(j)
    Dim iComp As Long
    For iComp = 1 To nC
        If Nu(j, iComp) <> 0# Then
            logS = logS + Nu(j, iComp) * (logCfree(iComp) + logGammaC(iComp))
        End If
    Next iComp
    logS = logS - logGammaSj
    If logS >  LOG_CLAMP Then logS =  LOG_CLAMP
    If logS < -LOG_CLAMP Then logS = -LOG_CLAMP
    SpecConcFromLog = 10# ^ logS
End Function

'-------------------------------------------------------------------------
' CalcIonicStrength — I = 0.5 * SUM_i(z_i^2 * [Ci]) at current state
'-------------------------------------------------------------------------
Private Function CalcIonicStrength( _
    ByVal nC          As Long, _
    ByVal nS          As Long, _
    ByRef logCfree()  As Double, _
    ByRef zC()        As Double, _
    ByRef zS()        As Double, _
    ByRef Nu()        As Double, _
    ByRef logK()      As Double, _
    ByRef logGammaC() As Double, _
    ByRef logGammaS() As Double) As Double

    Dim Icalc As Double, i As Long, j As Long
    For i = 1 To nC
        If zC(i) <> 0# Then
            Icalc = Icalc + 0.5# * zC(i) ^ 2# * (10# ^ logCfree(i))
        End If
    Next i
    For j = 1 To nS
        If zS(j) <> 0# Then
            Icalc = Icalc + 0.5# * zS(j) ^ 2# * _
                SpecConcFromLog(j, nC, logCfree, Nu, logK, logGammaC, logGammaS(j))
        End If
    Next j
    CalcIonicStrength = WorksheetFunction.Max(Icalc, MIN_I)
End Function

'-------------------------------------------------------------------------
' Log10GammaDavies — Davies equation for a single ion species
'-------------------------------------------------------------------------
Private Function Log10GammaDavies( _
    ByVal z             As Double, _
    ByVal IonicStrength As Double, _
    ByVal TempC         As Double) As Double

    If z = 0# Then Log10GammaDavies = 0#: Exit Function
    Dim A   As Double: A   = 0.509# * ((298.15# / (TempC + 273.15#)) ^ 1.5#)
    Dim sqI As Double: sqI = Sqr(WorksheetFunction.Max(IonicStrength, 0#))
    Log10GammaDavies = -A * z ^ 2# * (sqI / (1# + sqI) - 0.3# * IonicStrength)
End Function

'-------------------------------------------------------------------------
' SpeciesCharges — zS(j) = SUM_i( nu_j,i * zC(i) )
'-------------------------------------------------------------------------
Private Function SpeciesCharges( _
    ByVal nC  As Long, _
    ByVal nS  As Long, _
    ByRef Nu() As Double, _
    ByRef zC() As Double) As Double()

    Dim zS() As Double: ReDim zS(1 To nS)
    Dim j As Long, i As Long, q As Double
    For j = 1 To nS
        q = 0#
        For i = 1 To nC
            q = q + Nu(j, i) * zC(i)
        Next i
        zS(j) = q
    Next j
    SpeciesCharges = zS
End Function

'-------------------------------------------------------------------------
' SortedIndices — insertion sort; returns 1-based rank permutation
'-------------------------------------------------------------------------
Private Function SortedIndices(ByRef ids() As Long, ByVal n As Long) As Long()
    Dim idx() As Long: ReDim idx(1 To n)
    Dim i As Long, j As Long, tmp As Long
    For i = 1 To n: idx(i) = i: Next i
    For i = 2 To n
        tmp = idx(i): j = i - 1
        Do While j >= 1
            If ids(idx(j)) <= ids(tmp) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = tmp
    Next i
    SortedIndices = idx
End Function

'-------------------------------------------------------------------------
' Norm2 / Norm2N — Euclidean norms
'-------------------------------------------------------------------------
Private Function Norm2(ByRef v() As Double) As Double
    Dim s As Double, i As Long
    For i = LBound(v) To UBound(v): s = s + v(i) * v(i): Next i
    Norm2 = Sqr(s)
End Function

Private Function Norm2N(ByRef v() As Double, ByVal n As Long) As Double
    Dim s As Double, i As Long
    For i = 1 To n: s = s + v(i) * v(i): Next i
    Norm2N = Sqr(s)
End Function

'-------------------------------------------------------------------------
' SolveLinear — Gaussian elimination with partial pivoting
' Returns empty array if the matrix is (near-)singular.
'-------------------------------------------------------------------------
Private Function SolveLinear( _
    ByRef A() As Double, _
    ByRef b() As Double, _
    ByVal n   As Long) As Double()

    Dim x()  As Double: ReDim x(1 To n)
    Dim i As Long, j As Long, k As Long
    Dim maxVal As Double, tmp As Double, pRow As Long, factor As Double

    For k = 1 To n
        maxVal = Abs(A(k, k)): pRow = k
        For i = k + 1 To n
            If Abs(A(i, k)) > maxVal Then maxVal = Abs(A(i, k)): pRow = i
        Next i
        If maxVal < 1E-300 Then Exit Function   ' singular

        If pRow <> k Then
            For j = 1 To n
                tmp = A(k, j): A(k, j) = A(pRow, j): A(pRow, j) = tmp
            Next j
            tmp = b(k): b(k) = b(pRow): b(pRow) = tmp
        End If

        For i = k + 1 To n
            factor = A(i, k) / A(k, k)
            For j = k To n: A(i, j) = A(i, j) - factor * A(k, j): Next j
            b(i) = b(i) - factor * b(k)
        Next i
    Next k

    For i = n To 1 Step -1
        x(i) = b(i)
        For j = i + 1 To n: x(i) = x(i) - A(i, j) * x(j): Next j
        If Abs(A(i, i)) < 1E-300 Then Exit Function   ' singular
        x(i) = x(i) / A(i, i)
    Next i

    SolveLinear = x
End Function

'=========================================================================
' AqSpeciate UDF Registration  (v2.1b)
'
' RegisterAqSpeciate registers function and argument descriptions for the
' three public UDFs so they appear in the Excel Insert Function dialog
' (Shift+F3) and in formula autocomplete tooltips on Windows.
'
' PLATFORM NOTE: Application.MacroOptions is supported on Excel for
' Windows (2010 and later) only. On Excel for Mac the call is silently
' skipped via On Error Resume Next — no error is raised and the UDFs
' continue to work normally; argument hints simply will not appear.
'
' PERSISTENCE: MacroOptions descriptions are stored in the workbook
' session and are lost when the workbook is closed. To restore them
' automatically every time the workbook is opened, add the following
' one-line call to the ThisWorkbook module:
'
'   Private Sub Workbook_Open()
'       RegisterAqSpeciate
'   End Sub
'
' You can also run RegisterAqSpeciate manually at any time from the
' Developer tab > Macros dialog, or call it once after pasting the code.
'=========================================================================
Sub RegisterAqSpeciate()

    On Error Resume Next   ' silently skip on Mac or unsupported Excel versions

    '----------------------------------------------------------------------
    ' AqSpeciateAll — full solution row
    '----------------------------------------------------------------------
    Dim argsAll(1 To 10) As Variant
    argsAll(1)  = "Fixed pH (optional): Fixed solution pH. When supplied, H+ free " & _
                  "concentration is pinned to 10^(-pH)/gamma(H+) Omit to " & _
                  "solve pH from charge balance."
    argsAll(2)  = "Temperature in C (optional, default 25): Used to compute " & _
                  "the Davies activity-coefficient parameter A. Reliable " & _
                  "from approximately 0 to 60 C."
    argsAll(3)  = "Fixed Ionic Strength in mol/L (optional): Fixed ionic strength for " & _
                  "Davies activity corrections. When omitted, ionic strength " & _
                  "is computed iteratively. The computed value is always returned in AqSpeciateAll output."
    argsAll(4)  = "Component Range (required): 1 row per component (nC) x 4 columns — " & _
                  "col 1: integer component ID; " & _
                  "col 2: initial log10 guess for free conc; " & _
                  "col 3: total conc in mol/L (linear, not log); " & _
                  "col 4: formal charge. " & _
                  "Note: H+ must be in row 1 with charge +1."
    argsAll(5)  = "Species Range (required): 1 row per species x (2+nC) columns — " & _
                  "col 1: integer species ID; " & _
                  "col 2: cumulative formation logK from components; " & _
                  "cols 3 to 2+nC: stoichiometric coefficients (nu) for each " & _
                  "component in the same order as CompRange rows."
    argsAll(6)  = "Fixed Solids Range (optional): 1 row per solid x (2+nC) columns, " & _
                  "same layout as SpecRange. Fixed solids are always present; " & _
                  "and component concs are contrained by Ksp."
    argsAll(7)  = "Considered Solids Range (optional): 1 row per solid x (2+nC) columns, " & _
                  "same layout as SpecRange."
    argsAll(8)  = "Override Component ID (optional): Integer ID of one component " & _
                  "whose total concentration you want to override." & _
                  "Use to vary component concentrations without selecting new " & _
				  "Component Range. Must be supplied together with OverrideConc."
    argsAll(9)  = "Override Conc in mol/L (optional): Total concentration " & _
                  "(mol/L, linear) of the Override Component identified by " & _
                  "OverrideCompID, replacing the value in Component Range."
    argsAll(10) = "Verbose (optional, default 0 = not verbose): Set to 1 to report" & _
                  "diagnostic values: " & _
                  "(1) IDs of precipitated considered solids; " & _
                  "(2) Iteration count; " & _
                  "(3) convergence precision (final residual norm, mol/L); " & _
                  "(4-8) log(gamma) for |z| = 1 to 5."

    Application.MacroOptions _
        Macro:="AqSpeciateAll", _
        Description:="Returns a row of full aqueous equilibrium speciation as p-values." & _
                     "Gives all components/species concs sorted by ID, solid saturation indices for solids, computed I, and" & _
                     "errors. Optionally appends solver diagnostics (Verbose=1).", _
        Category:="AqSpeciate", _
        ArgumentDescriptions:=argsAll

    '----------------------------------------------------------------------
    ' AqSpeciateOne — single scalar output
    '----------------------------------------------------------------------
    Dim argsOne(1 To 11) As Variant
    argsOne(1)  = argsAll(1)   ' OptionalPH
    argsOne(2)  = argsAll(2)   ' OptionalTemp
    argsOne(3)  = argsAll(3)   ' OptionalIFixed
    argsOne(4)  = argsAll(4)   ' CompRange
    argsOne(5)  = argsAll(5)   ' SpecRange
    argsOne(6)  = argsAll(6)   ' FixedSolidsRange
    argsOne(7)  = argsAll(7)   ' ConsideredSolidsRange
    argsOne(8)  = "Output ID (required): Integer ID of the component, aqueous " & _
                  "species, or solid whose value is returned. For components " & _
                  "and species: returns p[X] = -log10([X_free]). For solid " & _
                  "IDs: returns the saturation index SI = -(logIAP + logK)."
    argsOne(9)  = argsAll(8)   ' OverrideCompID
    argsOne(10) = argsAll(9)   ' OverrideConc
    argsOne(11) = "SuppressErrors (optional, default 0 = do not suprress): " & _
                  "Set to 1 to prevent #VALUE! return when the solver encounters " & _
                  "errors. The best available result is returned instead. " & _
                  "Default (0) returns #VALUE! on any solver failure."

    Application.MacroOptions _
        Macro:="AqSpeciateOne", _
        Description:="Returns a single p-value (-log10[free]) for one " & _
                     "component or species identified by OutputID, or the " & _
                     "saturation index SI for a solid ID. Suitable as an Excel Solver objective " & _
                     "function or for parametric sensitivity analysis.", _
        Category:="AqSpeciate", _
        ArgumentDescriptions:=argsOne

    '----------------------------------------------------------------------
    ' AqSpeciateLabels — column header row
    '----------------------------------------------------------------------
    Dim argsLbl(1 To 10) As Variant
    argsLbl(1)  = argsAll(1)   ' Fixed pH (Optional)
    argsLbl(2)  = argsAll(2)   ' Temp (Optional, Default 25 C)
    argsLbl(3)  = argsAll(3)   ' Fixed Ionic Strength (Optional)
    argsLbl(4)  = argsAll(4)   ' Component Range
    argsLbl(5)  = argsAll(5)   ' SpecRange
    argsLbl(6)  = argsAll(6)   ' FixedSolidsRange
    argsLbl(7)  = argsAll(7)   ' ConsideredSolidsRange
    argsLbl(8)  = argsAll(8)   ' OverrideCompID
    argsLbl(9)  = argsAll(9)   ' OverrideConc
    argsLbl(10) = "Verbose (optional, default 0): Set to 1 to include labels " & _
                  "for the 8 verbose diagnostic columns appended by " & _
                  "AqSpeciateAll when Verbose=1."

    Application.MacroOptions _
        Macro:="AqSpeciateLabels", _
        Description:="Returns a row of column header labels " & _
                     "matching the exact output layout of AqSpeciateAll " & _
                     "for the same input arguments. " & _
                     "Labels follow the format p[ID], SI[ID], SI[ID](fixed), " & _
                     "Calc I (mol/L), Errors, and optional verbose labels.", _
        Category:="AqSpeciate", _
        ArgumentDescriptions:=argsLbl

    On Error GoTo 0   ' restore normal error handling

End Sub
