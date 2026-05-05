Option Explicit

'=========================================================================
' AqSpeciate v1.3c
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
' CHANGES FROM v1.3b  (bug fix release)
'
'  FIX 1 — E14 CHARGE BALANCE CHECK SUPPRESSED WHEN IONIC STRENGTH IS FIXED.
'     When IFixed is supplied (IisFixed = True), the solver uses the caller-
'     supplied IS throughout and never iterates it to self-consistency.  The
'     charge balance at convergence is therefore not meaningful as a diagnostic
'     in this mode — the system has one fewer degree of freedom and the charge
'     imbalance simply reflects the IS constraint rather than a stoichiometry
'     error.  The E14 check is now skipped entirely when IisFixed = True.
'
'  All other behaviour is UNCHANGED from v1.3b.
'
'=========================================================================
' FULL CHANGE LIST FROM v1.2b  — Incremental Update #1: User Interface Changes
'
' UI CHANGE 1 — H+ MUST BE THE FIRST COMPONENT (not merely the first +1
'   charge component).  The code checks that row 1 of CompRange has charge
'   +1.  If H+ is absent OR is not in row 1, the solver reports an error
'   string in the last output cell.
'
' UI CHANGE 2 — COMPONENT TOTAL CONCENTRATIONS ARE NOW IN MOLAR (linear),
'   NOT log10.  CompRange column 3 is read as a direct molar concentration.
'   The log conversion previously applied in PackInputsNew is removed.
'
' UI CHANGE 3a — OPTIONAL ARGUMENTS MOVED TO FIRST POSITIONS.
'   Both AqSpeciateAll and AqSpeciateOne now begin with:
'     OptionalPH, OptionalTemp, OptionalIFixed
'   as the first three positional Optional arguments.  All required range
'   inputs follow.  This makes it easier to specify the most commonly
'   customised options without scrolling past many range references.
'
' UI CHANGE 3b — CONSOLIDATED RANGE INPUTS (corrected in v1.3b).
'   Component data are now supplied as a single range CompRange with layout:
'     col 1 : integer ID
'     col 2 : initial guess log10([Ci_free])   (log — unchanged from solver)
'     col 3 : total concentration in mol/L      (linear — see UI CHANGE 2)
'     col 4 : formal charge (integer)            (was separate CompCharges range)
'   Species data are now supplied as a single range SpecRange with layout:
'     col 1 : integer ID
'     col 2 : log10(Kf) cumulative formation constant
'     cols 3+ : stoichiometric recipe matrix (nu_j,i), one column per component
'   The separate CompCharges range argument is ELIMINATED.
'
' UI CHANGE 4 — RICH ERROR REPORTING IN LAST OUTPUT CELL.
'   The last cell of AqSpeciateAll accumulates ALL detected error and alert
'   conditions, separated by full-stop + space.  Conditions checked are:
'     E1 : H+ is not the first component (row 1 of CompRange).
'     E2 : No component with charge +1 found (H+ missing entirely).
'     E3 : Species stoichiometry matrix column count does not match nC.
'     E4 : Newton solver did not converge within MAX_ITER iterations.
'     E5 : Jacobian became singular (ill-conditioned system).
'     E6 : Computed ionic strength exceeds IFixed by more than 0.1%.
'     E7 : A negative or zero species concentration was encountered at
'          convergence (numerical instability or bad input data).
'     E8 : Negative component total concentration supplied.
'     E9 : IFixed supplied but is <= 0 (physically meaningless).
'    E10 : nC < 1 or nS < 1 (empty input ranges).
'    E11 : CompRange does not have exactly 4 columns.
'    E12 : SpecRange does not have at least 3 columns.
'    E13 : logKw appears unreasonably large (reserved for future use).
'    E14 : Charge balance at convergence exceeds 5% of total ionic charge
'          (only checked when ionic strength is NOT fixed; suppressed when
'           IFixed is supplied because IS is not iterated to self-consistency
'           in that mode and the charge imbalance is not diagnostic).
'    E15 : Any component total concentration is exactly zero.
'
'  If the solver itself fails before output is assembled, the entire output
'  array still returns; the last cell contains the error string and all
'  p-value cells return 0.
'
' ALL OTHER BEHAVIOUR — SolveCore, BuildGammaTables, BuildResiduals,
'   BuildJacobian, ArmijoLineSearch, SpecConcFromLog, CalcIonicStrength,
'   Log10GammaDavies, SeedIonicStrength, PinHplus, PinHplusFromLog,
'   SortedIndices, SolveLinear, Norm2, Norm2N, SpeciesCharges —
'   are UNCHANGED from v1.2b.
'=========================================================================

'-------------------------------------------------------------------------
' MODULE-LEVEL CONSTANTS
'-------------------------------------------------------------------------
Private Const LN10          As Double = 2.30258509299405
Private Const LOG10E        As Double = 0.434294481903252  ' 1/ln(10): multiply instead of /Log(10)
Private Const MAX_ITER      As Long   = 1000
Private Const CONV_TOL      As Double = 1E-12
Private Const MIN_LOG_GAMMA As Double = -300#
Private Const MIN_I         As Double = 1E-12   ' floor for IS (never exactly zero)
Private Const ALPHA_MIN     As Double = 1E-6
Private Const MAX_LOG_STEP  As Double = 5#       ' max |Δlog[C]| per Newton step
Private Const LOG_CLAMP     As Double = 300#     ' clamp logS / logC before 10^x

'=========================================================================
' PUBLIC UDF 1: AqSpeciateAll  (v1.3b signature)
'
' Returns a horizontal row array of length nC + nS + 2:
'   columns 1..nC          : p[comp_free]  = -log10([Ci_free])  sorted by CompID
'   columns nC+1..nC+nS    : p[spec_free]  = -log10([Sj])       sorted by SpeciesID
'   column  nC+nS+1        : computed ionic strength (mol/L)
'   column  nC+nS+2        : "" if no issues; otherwise period-separated error messages
'
' --- ARGUMENT ORDER (UI CHANGE 3a: optional args FIRST) ---
'
' Optional arguments (first 3):
'   OptionalPH      fixes H+ activity (dimensionless); omit to solve H+ freely
'   OptionalTemp    temperature in °C (default 25)
'   OptionalIFixed  fixes ionic strength (mol/L); omit to iterate IS
'
' Required range arguments:
'   CompRange       nC rows × 4 cols  (UI CHANGE 3b, corrected in v1.3b):
'                     col 1 = integer component ID
'                     col 2 = initial guess log10([Ci_free])
'                     col 3 = total concentration in mol/L  (LINEAR, not log)
'                     col 4 = formal charge (integer)
'   SpecRange       nS rows × (2+nC) cols:
'                     col 1   = integer species ID
'                     col 2   = log10(Kf) cumulative formation constant
'                     cols 3+ = stoichiometric coefficients (nu_j,i),
'                               one column per component in CompRange row order
'
' NOTE: H+ MUST be the first row of CompRange (UI CHANGE 1).
' NOTE: The separate CompCharges argument from v1.3a is REMOVED; charges
'       are now col 4 of CompRange.
'=========================================================================
Public Function AqSpeciateAll( _
    Optional OptionalPH      As Variant, _
    Optional OptionalTemp    As Double = 25#, _
    Optional OptionalIFixed  As Variant, _
    Optional CompRange       As Variant, _
    Optional SpecRange       As Variant _
) As Variant

    On Error GoTo FailSafe

    '----------------------------------------------------------------------
    ' STEP 1: Declare all working variables (VBA requires top-of-scope)
    '----------------------------------------------------------------------
    Dim nC         As Long     ' number of components
    Dim nS         As Long     ' number of species
    Dim nSpecCols  As Long     ' total columns in SpecRange (= 2 + nC)
    Dim errMsg     As String   ' accumulates error strings (UI CHANGE 4)
    Dim compID()   As Long     ' component integer IDs
    Dim compT()    As Double   ' total component concentrations (mol/L)
    Dim compG()    As Double   ' initial guess log10([Ci_free])
    Dim zC()       As Double   ' component formal charges
    Dim specID()   As Long     ' species integer IDs
    Dim logK()     As Double   ' log10(Kf) formation constants
    Dim Nu()       As Double   ' stoichiometric recipe matrix (nS x nC)
    Dim pHFixed    As Boolean  ' True when OptionalPH is supplied
    Dim pHval      As Double   ' numeric pH value (if pHFixed)
    Dim IisFixed   As Boolean  ' True when OptionalIFixed is supplied
    Dim Ival       As Double   ' numeric fixed IS value (if IisFixed)
    Dim iHplus     As Long     ' 1-based index of H+ in component list
    Dim result     As Variant  ' raw output from SolveCore
    Dim nOut       As Long     ' total output columns
    Dim out()      As Variant  ' final 2-D row-array for spreadsheet
    Dim cOrder()   As Long     ' sort-index for components by ID
    Dim sOrder()   As Long     ' sort-index for species by ID
    Dim i          As Long
    Dim idx        As Long
    Dim CompRng    As Range    ' typed Range references
    Dim SpecRng    As Range
    Dim Icomputed  As Double   ' ionic strength at convergence
    Dim zS()       As Double   ' species charges (derived from Nu and zC)
    Dim logGammaC() As Double  ' log10(gamma) for components at convergence
    Dim logGammaS() As Double  ' log10(gamma) for species  at convergence
    Dim chargeBalance As Double ' charge balance check at convergence
    Dim totalIonConc  As Double ' total ionic charge for relative CB check

    '----------------------------------------------------------------------
    ' STEP 2: Validate that the required range arguments were actually passed
    '----------------------------------------------------------------------
    errMsg = ""   ' start with no errors

    ' Ranges are declared Optional only so that optional args can precede
    ' them in the signature (VBA restriction: optional args must follow
    ' all required args, so we make ALL args optional and validate here).
    If IsMissing(CompRange) Or IsMissing(SpecRange) Then
        ' Cannot produce a meaningful result without input data
        AqSpeciateAll = CVErr(xlErrValue)
        Exit Function
    End If

    ' Bind to typed Range variables (raises error if caller passed non-Range)
    Set CompRng = CompRange
    Set SpecRng = SpecRange

    '----------------------------------------------------------------------
    ' STEP 3: Determine dimensions and run pre-flight error checks (UI CHANGE 4)
    '----------------------------------------------------------------------
    nC       = CompRng.Rows.Count
    nS       = SpecRng.Rows.Count
    nSpecCols = SpecRng.Columns.Count   ' should be 2 + nC

    ' E10 — empty inputs
    If nC < 1 Or nS < 1 Then
        errMsg = AppendError(errMsg, "E10: CompRange or SpecRange is empty")
    End If

    ' E11 — CompRange must have exactly 4 columns (ID, logCguess, Ctot, charge)
    If CompRng.Columns.Count <> 4 Then
        errMsg = AppendError(errMsg, "E11: CompRange must have exactly 4 columns " & _
                             "(ID, log-guess, concentration, charge)")
    End If

    ' E12 — SpecRange must have at least 3 columns (ID, logK, >=1 stoich col)
    If nSpecCols < 3 Then
        errMsg = AppendError(errMsg, "E12: SpecRange must have at least 3 columns " & _
                             "(ID, logK, stoich columns)")
    End If

    ' E3 — stoichiometry matrix column count must match nC
    ' SpecRange has 2 header cols (ID + logK) and then nC stoich columns
    If (nSpecCols - 2) <> nC Then
        errMsg = AppendError(errMsg, "E3: Species stoichiometry matrix column count (" & _
                             CStr(nSpecCols - 2) & ") does not match number of " & _
                             "components (" & CStr(nC) & ")")
    End If

    ' E9 — IFixed physically meaningful
    If Not IsMissing(OptionalIFixed) Then
        If CDbl(OptionalIFixed) <= 0 Then
            errMsg = AppendError(errMsg, "E9: OptionalIFixed must be > 0 mol/L")
        End If
    End If

    ' If fundamental dimension errors exist, bail out now (cannot proceed)
    If Len(errMsg) > 0 Then
        Dim earlyOut() As Variant
        nOut = nC + nS + 2
        If nOut < 2 Then nOut = 2
        ReDim earlyOut(1 To 1, 1 To nOut)
        earlyOut(1, nOut) = errMsg
        AqSpeciateAll = earlyOut
        Exit Function
    End If

    '----------------------------------------------------------------------
    ' STEP 4: Unpack range inputs into typed arrays (UI CHANGE 3b)
    '----------------------------------------------------------------------
    Call PackInputsNew(nC, nS, CompRng, SpecRng, _
                       compID, compT, compG, zC, specID, logK, Nu)

    '----------------------------------------------------------------------
    ' STEP 5: Post-unpack validation checks
    '----------------------------------------------------------------------

    ' E8 — any negative or zero total concentration
    For i = 1 To nC
        If compT(i) < 0# Then
            errMsg = AppendError(errMsg, "E8: Negative total concentration " & _
                                 "for component ID " & CStr(compID(i)))
        End If
    Next i

    ' E15 — exactly zero total concentration (possible forgotten cell)
    For i = 1 To nC
        If compT(i) = 0# Then
            errMsg = AppendError(errMsg, "E15: Zero total concentration for " & _
                                 "component ID " & CStr(compID(i)) & _
                                 " (cell may be empty)")
        End If
    Next i

    ' E1/E2 — H+ must be row 1 AND must have charge +1 (UI CHANGE 1)
    ' First check charge of row 1
    If zC(1) <> 1# Then
        ' Row 1 component does not have charge +1
        ' Determine whether H+ exists elsewhere
        Dim hFoundElsewhere As Boolean
        hFoundElsewhere = False
        For i = 2 To nC
            If zC(i) = 1# Then
                hFoundElsewhere = True
                Exit For
            End If
        Next i
        If hFoundElsewhere Then
            errMsg = AppendError(errMsg, "E1: H+ (charge +1) is not the first " & _
                                 "component. H+ must appear in row 1 of CompRange")
        Else
            errMsg = AppendError(errMsg, "E2: No component with charge +1 found. " & _
                                 "H+ must be included as the first component")
        End If
    End If

    ' Locate H+ index for solver use (always row 1 if charge = +1 there)
    If zC(1) = 1# Then
        iHplus = 1
    Else
        iHplus = 0   ' solver will proceed without pH pinning
    End If

    '----------------------------------------------------------------------
    ' STEP 6: Parse optional pH and IFixed arguments
    '----------------------------------------------------------------------
    pHFixed = Not IsMissing(OptionalPH)
    If pHFixed Then pHval = CDbl(OptionalPH)

    IisFixed = Not IsMissing(OptionalIFixed)
    If IisFixed Then
        Ival = CDbl(OptionalIFixed)
    Else
        ' Seed IS from component totals (FIX 1 carried from v1.2b)
        Ival = SeedIonicStrength(nC, zC, compT)
    End If

    '----------------------------------------------------------------------
    ' STEP 7: Run Newton solver
    '----------------------------------------------------------------------
    ' Note: errMsg may already contain non-fatal warnings (E8/E15); the
    ' solver still runs so partial output is returned alongside the messages.
    Dim solverErr As String
    solverErr = ""

    result = SolveCore(nC, nS, compG, compT, zC, Nu, logK, _
                       pHFixed, pHval, iHplus, IisFixed, Ival, OptionalTemp, _
                       solverErr)

    ' Append any solver-internal error messages
    If Len(solverErr) > 0 Then
        errMsg = AppendError(errMsg, solverErr)
    End If

    '----------------------------------------------------------------------
    ' STEP 8: Assemble output array
    '----------------------------------------------------------------------
    nOut = nC + nS + 2
    ReDim out(1 To 1, 1 To nOut)

    If IsError(result) Then
        ' Solver returned an error — fill p-values with 0 and report errors
        For i = 1 To nOut - 1
            out(1, i) = 0
        Next i
        out(1, nOut) = errMsg
        AqSpeciateAll = out
        Exit Function
    End If

    ' result is 1-D array(1..nC+nS+2) from SolveCore.
    ' Sort by ascending ID for output.
    cOrder = SortedIndices(compID, nC)
    sOrder = SortedIndices(specID, nS)

    idx = 1
    For i = 1 To nC
        out(1, idx) = result(cOrder(i))       ' p[comp_free] in ID order
        idx = idx + 1
    Next i
    For i = 1 To nS
        out(1, idx) = result(nC + sOrder(i))  ' p[spec_free] in ID order
        idx = idx + 1
    Next i

    ' Ionic strength is in result(nC+nS+1)
    Icomputed = result(nC + nS + 1)
    out(1, idx) = Icomputed
    idx = idx + 1

    ' E6 — computed IS exceeds IFixed
    If IisFixed And (Icomputed > Ival * 1.001) Then
        errMsg = AppendError(errMsg, "E6: Computed ionic strength (" & _
                             Format(Icomputed, "0.000E+00") & " mol/L) exceeds " & _
                             "fixed IFixed (" & Format(Ival, "0.000E+00") & " mol/L)")
    End If

    ' E14 — charge balance check at convergence.
    ' SUPPRESSED when ionic strength is fixed (IisFixed = True): in that mode
    ' the solver uses the caller-supplied IS throughout and never iterates it
    ' to self-consistency, so the charge imbalance at convergence reflects the
    ' IS constraint rather than a stoichiometry error and is not diagnostic.
    If Not IisFixed Then
        ' Rebuild species charges and gamma tables at the converged IS
        zS = SpeciesCharges(nC, nS, Nu, zC)
        Call BuildGammaTables(nC, nS, zC, zS, Icomputed, OptionalTemp, logGammaC, logGammaS)

        ' Reconstruct free log-concentrations from p-values in result array
        Dim logCconv() As Double
        ReDim logCconv(1 To nC)
        For i = 1 To nC
            logCconv(i) = -result(i)   ' result stores p = -log[C]
        Next i

        ' Sum charge contributions from free components and all species
        chargeBalance = 0#
        totalIonConc  = 0#
        For i = 1 To nC
            Dim ci As Double
            ci = 10# ^ logCconv(i)
            chargeBalance = chargeBalance + zC(i) * ci
            totalIonConc  = totalIonConc  + Abs(zC(i)) * ci
        Next i
        Dim j As Long
        For j = 1 To nS
            Dim Sj As Double
            Sj = SpecConcFromLog(j, nC, logCconv, Nu, logK, logGammaC, logGammaS(j))
            chargeBalance = chargeBalance + zS(j) * Sj
            totalIonConc  = totalIonConc  + Abs(zS(j)) * Sj
        Next j

        If totalIonConc > 0# Then
            If Abs(chargeBalance) / totalIonConc > 0.05 Then
                errMsg = AppendError(errMsg, "E14: Charge balance at convergence is " & _
                                     Format(100# * Abs(chargeBalance) / totalIonConc, "0.0") & _
                                     "% of total ionic charge (check stoichiometry matrix)")
            End If
        End If
    End If   ' Not IisFixed

    ' Write final error/alert string (empty string = no issues)
    out(1, nOut) = errMsg

    AqSpeciateAll = out
    Exit Function

FailSafe:
    ' Catch-all: return #VALUE so user knows something went wrong
    AqSpeciateAll = CVErr(xlErrValue)
End Function

'=========================================================================
' PUBLIC UDF 2: AqSpeciateOne  (v1.3b signature)
'
' Returns a scalar: -log10([free]) for a single component or species,
' identified by OutputID.  Returns #VALUE if OutputID is not found or
' the solver fails.
'
' --- ARGUMENT ORDER (UI CHANGE 3a: optional args FIRST) ---
'
' Optional arguments (first 3):
'   OptionalPH      fixes H+ activity; omit to solve H+ freely
'   OptionalTemp    temperature in °C (default 25)
'   OptionalIFixed  fixes ionic strength (mol/L); omit to iterate IS
'
' Required arguments:
'   CompRange   nC rows × 4 cols  (UI CHANGE 3b, corrected in v1.3b):
'                 col 1 = integer component ID
'                 col 2 = initial guess log10([Ci_free])
'                 col 3 = total concentration in mol/L  (LINEAR, not log)
'                 col 4 = formal charge (integer)
'   SpecRange   nS rows × (2+nC) cols  (ID | logK | stoich matrix)
'   OutputID    integer ID of the component or species to return
'
' NOTE: H+ must be the first row of CompRange (UI CHANGE 1).
' NOTE: The separate CompCharges argument from v1.3a is REMOVED; charges
'       are now col 4 of CompRange.
'=========================================================================
Public Function AqSpeciateOne( _
    Optional OptionalPH      As Variant, _
    Optional OptionalTemp    As Double = 25#, _
    Optional OptionalIFixed  As Variant, _
    Optional CompRange       As Variant, _
    Optional SpecRange       As Variant, _
    Optional OutputID        As Variant _
) As Variant

    On Error GoTo FailSafe

    '----------------------------------------------------------------------
    ' Variable declarations
    '----------------------------------------------------------------------
    Dim nC        As Long
    Dim nS        As Long
    Dim nSpecCols As Long
    Dim compID()  As Long
    Dim compT()   As Double
    Dim compG()   As Double
    Dim zC()      As Double
    Dim specID()  As Long
    Dim logK()    As Double
    Dim Nu()      As Double
    Dim pHFixed   As Boolean
    Dim pHval     As Double
    Dim IisFixed  As Boolean
    Dim Ival      As Double
    Dim iHplus    As Long
    Dim result    As Variant
    Dim k         As Long
    Dim CompRng   As Range
    Dim SpecRng   As Range
    Dim solverErr As String

    '----------------------------------------------------------------------
    ' Validate that required arguments are present
    '----------------------------------------------------------------------
    If IsMissing(CompRange) Or IsMissing(SpecRange) Or IsMissing(OutputID) Then
        AqSpeciateOne = CVErr(xlErrValue)
        Exit Function
    End If

    Set CompRng = CompRange
    Set SpecRng = SpecRange

    nC        = CompRng.Rows.Count
    nS        = SpecRng.Rows.Count
    nSpecCols = SpecRng.Columns.Count

    ' Basic dimension checks (return #VALUE on fatal errors — AqSpeciateOne
    ' returns a scalar so there is no "last cell" for error messages)
    If nC < 1 Or nS < 1 Then GoTo FailSafe
    If CompRng.Columns.Count <> 4 Then GoTo FailSafe   ' must have 4 cols: ID, log-guess, Ctot, charge
    If nSpecCols < 3 Then GoTo FailSafe
    If (nSpecCols - 2) <> nC Then GoTo FailSafe

    '----------------------------------------------------------------------
    ' Unpack consolidated range inputs (UI CHANGE 3b)
    '----------------------------------------------------------------------
    Call PackInputsNew(nC, nS, CompRng, SpecRng, _
                       compID, compT, compG, zC, specID, logK, Nu)

    '----------------------------------------------------------------------
    ' Locate H+ (must be first component, UI CHANGE 1)
    '----------------------------------------------------------------------
    If zC(1) = 1# Then
        iHplus = 1
    Else
        iHplus = 0
    End If

    '----------------------------------------------------------------------
    ' Optional argument parsing
    '----------------------------------------------------------------------
    pHFixed  = Not IsMissing(OptionalPH)
    If pHFixed Then pHval = CDbl(OptionalPH)

    IisFixed = Not IsMissing(OptionalIFixed)
    If IisFixed Then
        Ival = CDbl(OptionalIFixed)
    Else
        Ival = SeedIonicStrength(nC, zC, compT)
    End If

    '----------------------------------------------------------------------
    ' Run solver
    '----------------------------------------------------------------------
    solverErr = ""
    result = SolveCore(nC, nS, compG, compT, zC, Nu, logK, _
                       pHFixed, pHval, iHplus, IisFixed, Ival, OptionalTemp, _
                       solverErr)

    If IsError(result) Then
        AqSpeciateOne = CVErr(xlErrValue)
        Exit Function
    End If

    '----------------------------------------------------------------------
    ' Find and return the requested output ID
    '----------------------------------------------------------------------
    Dim outID As Long
    outID = CLng(OutputID)

    For k = 1 To nC
        If compID(k) = outID Then
            AqSpeciateOne = result(k)
            Exit Function
        End If
    Next k
    For k = 1 To nS
        If specID(k) = outID Then
            AqSpeciateOne = result(nC + k)
            Exit Function
        End If
    Next k

    ' OutputID not matched in component or species lists
    AqSpeciateOne = CVErr(xlErrValue)
    Exit Function

FailSafe:
    AqSpeciateOne = CVErr(xlErrValue)
End Function

'=========================================================================
' CORE SOLVER  (private)
'
' Implements a damped Newton method with Armijo backtracking line search.
' Activity coefficients follow the Davies equation, updated each iteration.
' Ionic strength is either fixed (IisFixed = True) or iterated alongside
' the Newton solve (one-step-lagged update after each step).
'
' v1.3a change: receives a ByRef solverErr string that it populates with
' any internal error conditions (non-convergence, singular Jacobian, bad
' species concentration).  Caller appends this to the global errMsg.
'
' Returns a 1-D Variant array (1 To nC+nS+2):
'   [1..nC]       -log10([comp_free])  in original input order
'   [nC+1..nC+nS] -log10([spec_free])  in original input order
'   [nC+nS+1]     computed ionic strength (mol/L)
'   [nC+nS+2]     0 (reserved — alert flags now handled by caller)
'
' Returns CVErr(xlErrValue) if the Newton iteration did not converge,
' and sets solverErr to describe the failure.
'=========================================================================
Private Function SolveCore( _
    ByVal nC          As Long, _
    ByVal nS          As Long, _
    ByRef logCfree0() As Double, _
    ByRef Ctot()      As Double, _
    ByRef zC()        As Double, _
    ByRef Nu()        As Double, _
    ByRef logK()      As Double, _
    ByVal pHFixed     As Boolean, _
    ByVal pHval       As Double, _
    ByVal iHplus      As Long, _
    ByVal IisFixed    As Boolean, _
    ByVal Ival        As Double, _
    ByVal TempC       As Double, _
    ByRef solverErr   As String) As Variant  ' v1.3a: solverErr added

    ' --- All variable declarations hoisted to function scope (VBA requirement) ---
    Dim logCfree()  As Double
    Dim zS()        As Double
    Dim logGammaC() As Double  ' log10(gamma) for each component — rebuilt each iter
    Dim logGammaS() As Double  ' log10(gamma) for each species  — rebuilt each iter
    Dim Resid()     As Double
    Dim Jacob()     As Double
    Dim RHS()       As Double
    Dim JacSq()     As Double
    Dim delta()     As Double
    Dim out()       As Variant
    Dim r0          As Double
    Dim Icomputed   As Double
    Dim conc        As Double
    Dim iter        As Long
    Dim i           As Long
    Dim j           As Long
    Dim ii          As Long
    Dim jj          As Long
    Dim maxDelta    As Double
    Dim scaleFactor As Double
    Dim converged   As Boolean  ' explicit convergence flag (FIX 2 from v1.2b)

    ' --- Copy initial guess into working array ---
    ReDim logCfree(1 To nC)
    For i = 1 To nC
        logCfree(i) = logCfree0(i)
    Next i

    ' --- Compute species charges: z_Sj = SUM_i( nu_j,i * z_i ) ---
    zS = SpeciesCharges(nC, nS, Nu, zC)

    ' --- Initialise H+ pin before gamma tables are built for first iteration ---
    If pHFixed And iHplus > 0 Then
        Call PinHplus(logCfree, iHplus, pHval, zC(iHplus), Ival, TempC)
    End If

    ' --- Newton iteration ---
    converged = False   ' start as not converged

    For iter = 1 To MAX_ITER

        ' OPT 1: build log10(gamma) tables once per iteration.
        ' All calls to SpecConcFromLog / BuildResiduals / BuildJacobian
        ' within this iteration reuse these pre-computed values.
        Call BuildGammaTables(nC, nS, zC, zS, Ival, TempC, logGammaC, logGammaS)

        ' Re-pin H+ using the freshly computed gamma_H+ for this IS
        If pHFixed And iHplus > 0 Then
            Call PinHplusFromLog(logCfree, iHplus, pHval, logGammaC(iHplus))
        End If

        ' Build mass-balance + charge-balance residuals (OPT 2: single species pass)
        Call BuildResiduals(nC, nS, logCfree, Ctot, zC, zS, Nu, logK, _
                            logGammaC, logGammaS, pHFixed, iHplus, pHval, Resid)

        ' Convergence check on mass-balance rows only.
        ' Row nC+1 is the charge balance — excluded from convergence norm to
        ' avoid false non-convergence when charge balance is monitoring-only.
        r0 = Norm2N(Resid, nC)
        If r0 < CONV_TOL Then
            converged = True
            Exit For
        End If

        ' Build the nC × nC Jacobian of the mass-balance system
        Call BuildJacobian(nC, nS, logCfree, zC, Nu, logK, _
                           logGammaC, logGammaS, pHFixed, iHplus, Jacob)

        ' Extract nC × nC block; negate residuals to form RHS = -F(x)
        ReDim RHS(1 To nC)
        ReDim JacSq(1 To nC, 1 To nC)
        For ii = 1 To nC
            RHS(ii) = -Resid(ii)
            For jj = 1 To nC
                JacSq(ii, jj) = Jacob(ii, jj)
            Next jj
        Next ii

        ' Solve nC × nC linear system J·Δ = -F for the Newton step
        delta = SolveLinear(JacSq, RHS, nC)
        If IsEmpty(delta) Then
            ' Singular Jacobian — ill-conditioned system
            solverErr = AppendError(solverErr, "E5: Jacobian became singular at " & _
                                    "iteration " & CStr(iter) & _
                                    " (ill-conditioned system or bad input data)")
            SolveCore = CVErr(xlErrValue)
            Exit Function
        End If

        ' Cap Newton step to MAX_LOG_STEP log-units per component.
        ' Large steps from a poor starting point can send concentrations to ±300
        ' on the first iteration, making the line search work very hard.
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

        ' Backtracking Armijo line search: halve alpha until ||F(x+α·Δ)|| < ||F(x)||
        Call ArmijoLineSearch(logCfree, delta, Ctot, zC, zS, Nu, logK, _
                              logGammaC, logGammaS, pHFixed, iHplus, pHval, r0, _
                              Ival, TempC)

        ' Update ionic strength from current free concentrations (one-step lag).
        ' Convergence of the coupled IS/Newton system is guaranteed provided
        ' the step size is small (ensured by the line search above).
        If Not IisFixed Then
            Call BuildGammaTables(nC, nS, zC, zS, Ival, TempC, logGammaC, logGammaS)
            Ival = CalcIonicStrength(nC, nS, logCfree, zC, zS, Nu, logK, _
                                    logGammaC, logGammaS)
        End If

    Next iter

    ' Use explicit converged flag (unambiguous regardless of For-loop exit mode)
    If Not converged Then
        solverErr = AppendError(solverErr, "E4: Newton solver did not converge " & _
                                "within " & CStr(MAX_ITER) & " iterations. " & _
                                "Check initial guesses and input data")
        SolveCore = CVErr(xlErrValue)
        Exit Function
    End If

    ' --- Final ionic strength at converged concentrations ---
    Call BuildGammaTables(nC, nS, zC, zS, Ival, TempC, logGammaC, logGammaS)
    Icomputed = CalcIonicStrength(nC, nS, logCfree, zC, zS, Nu, logK, _
                                  logGammaC, logGammaS)

    ' --- Pack output array ---
    ReDim out(1 To nC + nS + 2)

    ' Free component p-values in original input order
    For i = 1 To nC
        out(i) = -logCfree(i)
    Next i

    ' Free species p-values in original input order
    For i = 1 To nS
        conc = SpecConcFromLog(i, nC, logCfree, Nu, logK, logGammaC, logGammaS(i))
        If conc <= 0# Then
            ' E7 — negative or zero species concentration at convergence
            solverErr = AppendError(solverErr, "E7: Non-positive concentration " & _
                                    "for species index " & CStr(i) & _
                                    " at convergence (numerical instability or " & _
                                    "bad logK/stoichiometry data)")
            SolveCore = CVErr(xlErrValue)
            Exit Function
        End If
        out(nC + i) = -(Log(conc) * LOG10E)   ' OPT 3: multiply by LOG10E
    Next i

    out(nC + nS + 1) = Icomputed   ' computed ionic strength
    out(nC + nS + 2) = 0           ' reserved (alert flags handled by caller in v1.3a)

    SolveCore = out
    Exit Function

FailSafe2:
    SolveCore = CVErr(xlErrValue)
End Function

'=========================================================================
' v1.3a HELPER: AppendError
'
' Concatenates a new error sentence to an existing error string.
' Sentences are separated by ". " (period + space) for readability.
' If existing is empty, returns newMsg alone (no leading separator).
'=========================================================================
Private Function AppendError(ByVal existing As String, _
                              ByVal newMsg   As String) As String
    If Len(existing) = 0 Then
        AppendError = newMsg
    Else
        AppendError = existing & ". " & newMsg
    End If
End Function

'=========================================================================
' v1.3b HELPER: PackInputsNew
'
' Transfers consolidated range data (UI CHANGE 3b, corrected in v1.3b)
' into typed VBA arrays.
'
' CompRange layout  (nC rows × 4 cols):
'   col 1 : integer component ID
'   col 2 : initial guess log10([Ci_free])
'   col 3 : total concentration in mol/L   (LINEAR — UI CHANGE 2)
'   col 4 : formal charge (integer)         (was separate CompCharges in v1.3a)
'
' SpecRange layout  (nS rows × (2+nC) cols):
'   col 1   : integer species ID
'   col 2   : log10(Kf) cumulative formation constant
'   cols 3+ : stoichiometric coefficients nu_j,i (one col per component)
'             column 3 = component 1 (row 1 of CompRange), etc.
'=========================================================================
Private Sub PackInputsNew( _
    ByVal nC       As Long, _
    ByVal nS       As Long, _
    CompRng        As Range, _
    SpecRng        As Range, _
    ByRef compID() As Long, _
    ByRef compT()  As Double, _
    ByRef compG()  As Double, _
    ByRef zC()     As Double, _
    ByRef specID() As Long, _
    ByRef logK()   As Double, _
    ByRef Nu()     As Double)

    ReDim compID(1 To nC)
    ReDim compT(1 To nC)
    ReDim compG(1 To nC)
    ReDim zC(1 To nC)
    ReDim specID(1 To nS)
    ReDim logK(1 To nS)
    ReDim Nu(1 To nS, 1 To nC)

    Dim i As Long
    Dim j As Long

    ' Read component data row-by-row from the consolidated 4-column CompRange.
    ' col 1 = ID, col 2 = log-guess, col 3 = total concentration (mol/L, LINEAR),
    ' col 4 = formal charge
    For i = 1 To nC
        compID(i) = CLng(CompRng.Cells(i, 1).Value)
        compG(i)  = CDbl(CompRng.Cells(i, 2).Value)  ' log10([Ci_free]) initial guess
        compT(i)  = CDbl(CompRng.Cells(i, 3).Value)  ' mol/L directly (UI CHANGE 2)
        zC(i)     = CDbl(CompRng.Cells(i, 4).Value)  ' formal charge (col 4, v1.3b)
    Next i

    ' Read species data and stoichiometry row-by-row from SpecRange.
    ' col 1 = ID, col 2 = logK, cols 3..(2+nC) = stoich matrix
    For j = 1 To nS
        specID(j) = CLng(SpecRng.Cells(j, 1).Value)
        logK(j)   = CDbl(SpecRng.Cells(j, 2).Value)
        For i = 1 To nC
            Nu(j, i) = CDbl(SpecRng.Cells(j, 2 + i).Value)   ' stoich col offset by 2
        Next i
    Next j
End Sub

'=========================================================================
' FIX 1 (from v1.2b): SEED IONIC STRENGTH FROM COMPONENT TOTALS
'
' Estimates IS before the Newton loop as an upper bound:
'   I_seed = max( 0.5 * SUM_i( z_i^2 * T_i ),  MIN_I )
'
' In v1.3a, compT() already holds the linear molar concentration directly
' (no 10^ conversion needed here — UI CHANGE 2 means PackInputsNew does
' NOT take a log, so compT is already mol/L).
'=========================================================================
Private Function SeedIonicStrength( _
    ByVal nC     As Long, _
    ByRef zC()   As Double, _
    ByRef Ctot() As Double) As Double

    Dim Iseed As Double
    Dim i     As Long

    Iseed = 0#
    For i = 1 To nC
        If zC(i) <> 0# Then
            ' Ctot(i) is already in mol/L (linear), no 10^ needed
            Iseed = Iseed + 0.5# * zC(i) ^ 2# * Ctot(i)
        End If
    Next i

    ' Apply floor so IS is never exactly zero (avoids Sqr(0) and divide issues)
    SeedIonicStrength = WorksheetFunction.Max(Iseed, MIN_I)
End Function

'=========================================================================
' OPT 1: BUILD GAMMA TABLES
'
' Computes log10(gamma) for every component and species in one pass at
' the current ionic strength and temperature, using the Davies equation:
'   log10(gamma) = -A * z^2 * ( sqrt(I)/(1+sqrt(I)) - 0.3*I )
'   A = 0.509 * (298.15 / (T+273.15))^1.5
' Neutral species (z = 0) always have log10(gamma) = 0.
'
' Results stored in logGammaC(1..nC) and logGammaS(1..nS) for reuse
' throughout the Newton iteration.
'=========================================================================
Private Sub BuildGammaTables( _
    ByVal nC            As Long, _
    ByVal nS            As Long, _
    ByRef zC()          As Double, _
    ByRef zS()          As Double, _
    ByVal IonicStrength As Double, _
    ByVal TempC         As Double, _
    ByRef logGammaC()   As Double, _
    ByRef logGammaS()   As Double)

    Dim i     As Long
    Dim A     As Double
    Dim sqI   As Double
    Dim dTerm As Double   ' Davies bracket: sqrt(I)/(1+sqrt(I)) - 0.3*I

    ' OPT 3: compute A, sqI, and dTerm once per call (shared by all species)
    A     = 0.509# * ((298.15# / (TempC + 273.15#)) ^ 1.5#)
    sqI   = Sqr(WorksheetFunction.Max(IonicStrength, 0#))
    dTerm = sqI / (1# + sqI) - 0.3# * IonicStrength

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

'=========================================================================
' H+ PINNING HELPERS
'
' When pH is fixed the H+ component is not solved by Newton; instead its
' free concentration is set to:
'   logCfree(iHplus) = -pH - log10(gamma_H+)
' i.e. the activity of H+ equals 10^(-pH) at every iteration.
'
' PinHplus        : used once before gamma tables are built (calls Davies directly).
' PinHplusFromLog : used inside the Newton loop with pre-computed logGammaC(iHplus).
'=========================================================================
Private Sub PinHplus( _
    ByRef logCfree()    As Double, _
    ByVal iHplus        As Long, _
    ByVal pHval         As Double, _
    ByVal zHplus        As Double, _
    ByVal IonicStrength As Double, _
    ByVal TempC         As Double)

    Dim lgH As Double
    lgH = Log10GammaDavies(zHplus, IonicStrength, TempC)
    logCfree(iHplus) = -pHval - lgH
End Sub

Private Sub PinHplusFromLog( _
    ByRef logCfree()    As Double, _
    ByVal iHplus        As Long, _
    ByVal pHval         As Double, _
    ByVal logGammaHplus As Double)

    ' activity(H+) = 10^(-pH) = gamma_H+ * [H+_free]
    ' => log[H+_free] = -pH - log10(gamma_H+)
    logCfree(iHplus) = -pHval - logGammaHplus
End Sub

'=========================================================================
' OPT 2 + OPT 4: SINGLE-PASS RESIDUALS
'
' Visits every species exactly once, accumulating both:
'   (a) mass-balance contribution  Resid(k) += nu_j,k * [Sj]
'   (b) charge-balance contribution  qSum   += z_S,j  * [Sj]
'
' Mass-balance residual for component k (k ≠ iHplus or pH not fixed):
'   Resid(k) = [Ck_free] + SUM_j( nu_j,k * [Sj] ) - Ctot(k)
'
' Pin equation (k = iHplus, pH fixed):
'   Resid(k) = logCfree(k) - (-pH - logGammaC(iHplus))
'
' Charge balance (row nC+1, monitoring only — not solved by Newton):
'   Resid(nC+1) = SUM_i( z_i * [Ci_free] ) + SUM_j( z_S,j * [Sj] )
'=========================================================================
Private Sub BuildResiduals( _
    ByVal nC          As Long, _
    ByVal nS          As Long, _
    ByRef logCfree()  As Double, _
    ByRef Ctot()      As Double, _
    ByRef zC()        As Double, _
    ByRef zS()        As Double, _
    ByRef Nu()        As Double, _
    ByRef logK()      As Double, _
    ByRef logGammaC() As Double, _
    ByRef logGammaS() As Double, _
    ByVal pHFixed     As Boolean, _
    ByVal iHplus      As Long, _
    ByVal pHval       As Double, _
    ByRef Resid()     As Double)

    ReDim Resid(1 To nC + 1)
    Dim k     As Long
    Dim j     As Long
    Dim Sj    As Double
    Dim qSum  As Double
    Dim logCk As Double

    ' Initialise mass-balance residuals from free-component concentrations
    For k = 1 To nC
        If pHFixed And k = iHplus Then
            ' Pin equation: residual in log space (zero at convergence)
            Resid(k) = logCfree(k) - (-pHval - logGammaC(iHplus))
        Else
            logCk = logCfree(k)
            If logCk >  LOG_CLAMP Then logCk =  LOG_CLAMP
            If logCk < -LOG_CLAMP Then logCk = -LOG_CLAMP
            Resid(k) = 10# ^ logCk - Ctot(k)   ' [Ck_free] - Ctot(k)
        End If
    Next k

    ' Single pass over species: add species contributions and accumulate charge balance
    qSum = 0#
    For j = 1 To nS
        Sj = SpecConcFromLog(j, nC, logCfree, Nu, logK, logGammaC, logGammaS(j))
        For k = 1 To nC
            If Nu(j, k) <> 0# Then
                If Not (pHFixed And k = iHplus) Then
                    Resid(k) = Resid(k) + Nu(j, k) * Sj   ' mass balance
                End If
            End If
        Next k
        qSum = qSum + zS(j) * Sj   ' species charge contribution
    Next j

    ' Add free-component charge contributions to charge balance
    For k = 1 To nC
        qSum = qSum + zC(k) * (10# ^ logCfree(k))
    Next k
    Resid(nC + 1) = qSum   ' total charge balance (monitoring row)
End Sub

'=========================================================================
' JACOBIAN  (nC × nC, mass-balance rows only)
'
' Partial derivatives in log-concentration space:
'   d(Resid_k)/d(logCfree_m)
'
' Pin row (k = iHplus):
'   J(k,k) = 1,  J(k,m≠k) = 0
'
' Normal row (k ≠ iHplus or pH not fixed):
'   J(k,k)  = ln(10) * [Ck_free]                              (diagonal)
'   J(k,m) += nu_j,k * [Sj] * nu_j,m * ln(10)  for each j   (species)
'
' Derivation:
'   [Sj] = f(logCfree), d[Sj]/d(logCfree_m) = [Sj] * nu_j,m * ln(10)
'=========================================================================
Private Sub BuildJacobian( _
    ByVal nC          As Long, _
    ByVal nS          As Long, _
    ByRef logCfree()  As Double, _
    ByRef zC()        As Double, _
    ByRef Nu()        As Double, _
    ByRef logK()      As Double, _
    ByRef logGammaC() As Double, _
    ByRef logGammaS() As Double, _
    ByVal pHFixed     As Boolean, _
    ByVal iHplus      As Long, _
    ByRef Jacob()     As Double)

    ReDim Jacob(1 To nC, 1 To nC)
    Dim k  As Long
    Dim m  As Long
    Dim j  As Long
    Dim Sj As Double

    ' Diagonal: free-component terms
    For k = 1 To nC
        If pHFixed And k = iHplus Then
            Jacob(k, k) = 1#   ' pin equation — gradient = 1 in log space
        Else
            Jacob(k, k) = LN10 * (10# ^ logCfree(k))
        End If
    Next k

    ' Off-diagonal (and additional diagonal) contributions from each species
    For j = 1 To nS
        Sj = SpecConcFromLog(j, nC, logCfree, Nu, logK, logGammaC, logGammaS(j))
        For k = 1 To nC
            If Not (pHFixed And k = iHplus) Then
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
' ARMIJO BACKTRACKING LINE SEARCH
'
' Finds the largest step size alpha = 1, 0.5, 0.25, …, ALPHA_MIN such
' that the residual norm at (logCfree + alpha * delta) is strictly less
' than the norm at logCfree.  If no acceptable step is found down to
' ALPHA_MIN, a fractional step of ALPHA_MIN is taken anyway to prevent
' stagnation.
'
' The pre-computed gamma tables (logGammaC, logGammaS) are reused for all
' trial points; IS is not updated mid-search.
'=========================================================================
Private Sub ArmijoLineSearch( _
    ByRef logCfree()  As Double, _
    ByRef delta()     As Double, _
    ByRef Ctot()      As Double, _
    ByRef zC()        As Double, _
    ByRef zS()        As Double, _
    ByRef Nu()        As Double, _
    ByRef logK()      As Double, _
    ByRef logGammaC() As Double, _
    ByRef logGammaS() As Double, _
    ByVal pHFixed     As Boolean, _
    ByVal iHplus      As Long, _
    ByVal pHval       As Double, _
    ByVal r0          As Double, _
    ByVal IonicStrength As Double, _
    ByVal TempC       As Double)

    Dim alpha    As Double: alpha = 1#
    Dim trial()  As Double
    Dim Rtrial() As Double
    Dim nC       As Long: nC = UBound(logCfree)
    Dim iComp    As Long

    Do While alpha > ALPHA_MIN
        ReDim trial(1 To nC)
        For iComp = 1 To nC
            trial(iComp) = logCfree(iComp) + alpha * delta(iComp)
        Next iComp

        ' Maintain H+ pin at each trial point
        If pHFixed And iHplus > 0 Then
            Call PinHplusFromLog(trial, iHplus, pHval, logGammaC(iHplus))
        End If

        Call BuildResiduals(nC, UBound(logK), trial, Ctot, zC, zS, Nu, logK, _
                            logGammaC, logGammaS, pHFixed, iHplus, pHval, Rtrial)

        If Norm2N(Rtrial, nC) < r0 Then
            ' Accept this step size
            For iComp = 1 To nC
                logCfree(iComp) = trial(iComp)
            Next iComp
            Exit Sub
        End If
        alpha = alpha * 0.5#
    Loop

    ' Fallback: take a minimal step rather than making no progress at all
    For iComp = 1 To nC
        logCfree(iComp) = logCfree(iComp) + ALPHA_MIN * delta(iComp)
    Next iComp
End Sub

'=========================================================================
' OPT 4: SPECIES CONCENTRATION FROM PRE-COMPUTED LOG-GAMMA TABLES
'
' log[Sj] = logKf(j)
'           + SUM_i( nu_j,i * ( logCfree_i + logGammaC_i ) )   [component activity]
'           - logGammaS_j                                         [species activity]
'
' All arithmetic is floating-point addition/multiplication; the only
' transcendental call is the final 10^logS.
'=========================================================================
Private Function SpecConcFromLog( _
    ByVal j           As Long, _
    ByVal nC          As Long, _
    ByRef logCfree()  As Double, _
    ByRef Nu()        As Double, _
    ByRef logK()      As Double, _
    ByRef logGammaC() As Double, _
    ByVal logGammaSj  As Double) As Double   ' scalar: logGammaS for species j

    Dim logS  As Double
    Dim iComp As Long
    logS = logK(j)

    For iComp = 1 To nC
        If Nu(j, iComp) <> 0# Then
            logS = logS + Nu(j, iComp) * (logCfree(iComp) + logGammaC(iComp))
        End If
    Next iComp

    logS = logS - logGammaSj   ' subtract species activity correction

    ' Clamp before 10^logS to prevent VBA Double overflow
    If logS >  LOG_CLAMP Then logS =  LOG_CLAMP
    If logS < -LOG_CLAMP Then logS = -LOG_CLAMP

    SpecConcFromLog = 10# ^ logS
End Function

'=========================================================================
' IONIC STRENGTH
'
' I = 0.5 * SUM_i( z_i^2 * [Ci_free] )  +  0.5 * SUM_j( z_S,j^2 * [Sj] )
'
' Species concentrations use pre-computed logGammaS to avoid extra
' DaviesGamma calls.  The result is floored at MIN_I.
'=========================================================================
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

    Dim Icalc As Double
    Dim i     As Long
    Dim j     As Long

    ' Component contributions
    For i = 1 To nC
        If zC(i) <> 0# Then
            Icalc = Icalc + 0.5# * zC(i) ^ 2# * (10# ^ logCfree(i))
        End If
    Next i

    ' Species contributions
    For j = 1 To nS
        If zS(j) <> 0# Then
            Icalc = Icalc + 0.5# * zS(j) ^ 2# * _
                SpecConcFromLog(j, nC, logCfree, Nu, logK, logGammaC, logGammaS(j))
        End If
    Next j

    CalcIonicStrength = WorksheetFunction.Max(Icalc, MIN_I)
End Function

'=========================================================================
' DAVIES log10(gamma) — direct log-space result (OPT 3)
'
' Used only before gamma tables exist (PinHplus on iteration 0) and
' inside BuildGammaTables itself.  Returns log10(gamma), not gamma.
'=========================================================================
Private Function Log10GammaDavies( _
    ByVal z             As Double, _
    ByVal IonicStrength As Double, _
    ByVal TempC         As Double) As Double

    If z = 0# Then
        Log10GammaDavies = 0#
        Exit Function
    End If

    Dim A   As Double
    Dim sqI As Double
    A   = 0.509# * ((298.15# / (TempC + 273.15#)) ^ 1.5#)
    sqI = Sqr(WorksheetFunction.Max(IonicStrength, 0#))

    Log10GammaDavies = -A * z ^ 2# * (sqI / (1# + sqI) - 0.3# * IonicStrength)
End Function

'=========================================================================
' HELPER: SPECIES CHARGES
'
' Derives species formal charge from stoichiometry:
'   z_S,j = SUM_i( nu_j,i * z_i )
'=========================================================================
Private Function SpeciesCharges( _
    ByVal nC  As Long, _
    ByVal nS  As Long, _
    ByRef Nu() As Double, _
    ByRef zC() As Double) As Double()

    Dim zS() As Double
    ReDim zS(1 To nS)
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

'=========================================================================
' HELPER: SortedIndices
'
' Returns a 1-based index array that sorts ids() in ascending order.
' Uses insertion sort (O(n²)), appropriate for small n (typically < 50).
'=========================================================================
Private Function SortedIndices(ByRef ids() As Long, ByVal n As Long) As Long()
    Dim idx() As Long
    ReDim idx(1 To n)
    Dim i As Long, j As Long, tmp As Long

    For i = 1 To n: idx(i) = i: Next i

    For i = 2 To n
        tmp = idx(i)
        j = i - 1
        Do While j >= 1
            If ids(idx(j)) <= ids(tmp) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = tmp
    Next i
    SortedIndices = idx
End Function

'=========================================================================
' HELPER: Norm2
'
' Euclidean norm of all elements of a 1-D Double array.
'=========================================================================
Private Function Norm2(ByRef v() As Double) As Double
    Dim s As Double, i As Long
    For i = LBound(v) To UBound(v)
        s = s + v(i) * v(i)
    Next i
    Norm2 = Sqr(s)
End Function

'=========================================================================
' HELPER: Norm2N
'
' Euclidean norm of the first n elements of a 1-D Double array.
' Used for convergence check on mass-balance rows only
' (deliberately excludes row nC+1 = charge balance).
'=========================================================================
Private Function Norm2N(ByRef v() As Double, ByVal n As Long) As Double
    Dim s As Double, i As Long
    For i = 1 To n
        s = s + v(i) * v(i)
    Next i
    Norm2N = Sqr(s)
End Function

'=========================================================================
' HELPER: SolveLinear
'
' Gaussian elimination with partial pivoting.
' Solves A·x = b for x, where A is n×n and b is n×1.
' Returns an empty array (IsEmpty = True) if A is singular (pivot < 1E-300).
' The input arrays A and b are overwritten in place.
'=========================================================================
Private Function SolveLinear( _
    ByRef A() As Double, _
    ByRef b() As Double, _
    ByVal n   As Long) As Double()

    Dim x()    As Double
    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim maxVal As Double
    Dim tmp    As Double
    Dim pRow   As Long
    Dim factor As Double

    ReDim x(1 To n)

    ' Forward elimination with partial pivoting
    For k = 1 To n
        ' Find pivot row
        maxVal = Abs(A(k, k))
        pRow = k
        For i = k + 1 To n
            If Abs(A(i, k)) > maxVal Then
                maxVal = Abs(A(i, k))
                pRow = i
            End If
        Next i
        If maxVal < 1E-300 Then Exit Function   ' singular — return empty array

        ' Swap rows k and pRow
        If pRow <> k Then
            For j = 1 To n
                tmp = A(k, j): A(k, j) = A(pRow, j): A(pRow, j) = tmp
            Next j
            tmp = b(k): b(k) = b(pRow): b(pRow) = tmp
        End If

        ' Eliminate below pivot
        For i = k + 1 To n
            factor = A(i, k) / A(k, k)
            For j = k To n
                A(i, j) = A(i, j) - factor * A(k, j)
            Next j
            b(i) = b(i) - factor * b(k)
        Next i
    Next k

    ' Back substitution
    For i = n To 1 Step -1
        x(i) = b(i)
        For j = i + 1 To n
            x(i) = x(i) - A(i, j) * x(j)
        Next j
        If Abs(A(i, i)) < 1E-300 Then Exit Function   ' singular
        x(i) = x(i) / A(i, i)
    Next i

    SolveLinear = x
End Function
