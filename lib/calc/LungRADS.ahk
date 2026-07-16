; ============================================================
; lib/calc/LungRADS.ahk -- ACR Lung-RADS v2022
; ------------------------------------------------------------
; Reference: American College of Radiology Lung-RADS v2022
; Assessment System, Release November 2022.
;
; Category buckets (verified against the v2022 source-of-truth grid
; rendered via pdfplumber):
;
;   Cat 0  -- Incomplete / awaiting comparison / 1-3 mo for indeterminate
;             infectious-inflammatory findings
;   Cat 1  -- No nodules; benign-calcified (complete/central/popcorn/
;             concentric ring); fat-containing
;   Cat 2  -- Solid <6 mm baseline OR new <4 mm
;             Part-solid <6 mm total at baseline
;             GGN <30 mm at baseline/new/growing OR >=30 mm stable/slow
;             Juxtapleural <10 mm + solid + smooth + oval/lentiform/triangular
;             Subsegmental airway nodule (baseline/new/stable)
;             Cat 3 stable at 6-mo follow-up OR Cat 4B proven benign
;   Cat 3  -- Solid 6 to <8 mm baseline OR new 4 to <6 mm
;             Part-solid >=6 mm total with solid <6 mm baseline OR new <6 mm total
;             GGN >=30 mm at baseline or new
;             Cat 4A stable at 3-mo follow-up (excl. airway)
;   Cat 4A -- Solid 8 to <15 mm baseline OR growing <8 mm OR new 6 to <8 mm
;             Part-solid >=6 mm total with solid 6 to <8 mm baseline
;                       OR new/growing <4 mm solid component
;             Airway nodule, segmental or more proximal, at baseline
;   Cat 4B -- Solid >=15 mm baseline OR new or growing >=8 mm
;             Part-solid solid component >=8 mm baseline OR new/growing >=4 mm component
;             Airway nodule, segmental or more proximal, stable or growing
;             Slow-growth over multiple screenings (note 8)
;   Cat 4X -- Cat 3 or 4 with ADDITIONAL features suggesting malignancy
;             (spiculation, lymphadenopathy, frank metastatic disease,
;              GGN that doubles in 1 year, etc. -- note 14).
;             NOTE: growth alone is NOT a 4X trigger -- growth shifts size
;             tier (e.g. growing solid <8 mm -> 4A); slow growth over
;             multiple exams is 4B per note 8.
;   S      -- Modifier for clinically significant findings unrelated to
;             lung cancer (added to 0-4).
;   NC     -- "Not classified in Lung-RADS": thin-walled cyst (note 12a),
;             fluid-containing cyst (note 12g), multiple cysts / LCH-LAM
;             (note 12h), and known lung cancer diagnosis (note 16).
;             The exam category is driven by other classified findings.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

LungRADS_Entry(input) {
    ShowLungRADSDialog(input)
    return ""
}

ShowLungRADSDialog(initialText := "") {
    sz := TextScan.Size(initialText)
    comp := TextScan.Composition(initialText)
    typeIdx := (comp = "ground glass") ? 4
            : (comp = "part solid")   ? 3
            : 1
    ; Cavitary nodule -- note 12d: wall thickening is the dominant feature, so
    ; it is classified on the solid-nodule pathway (total mean diameter).
    ; Detect it explicitly since TextScan.Composition has no cavitary category.
    if RegExMatch(initialText, "i)\bcavitary\b|\bcavitation\b|\bcavitating\b")
        typeIdx := 2

    ; Auto-detect 4X features (note 14) from the highlighted sentence.
    ; These are PER-NODULE descriptors -- "8 mm spiculated nodule" is
    ; clearly about the current lesion -- so pre-checking the box matches
    ; the radiologist's intent. _LUR_ScanFeatures includes negation
    ; handling ("no spiculation" -> not flagged) so it's safer than the
    ; previous naive TextScan.ContainsAny pre-detect.
    ;
    ; Infectious features (segmental/lobar consolidation, multiple new
    ; nodules) are intentionally NOT auto-detected -- they're typically
    ; scene-level findings ("multiple new nodules in both lungs") rather
    ; than properties of the specific nodule being scored. Auto-checking
    ; them would silently trigger a Cat 0 override based on report-wide
    ; context. The radiologist explicitly ticks those when they apply.
    preSpicul := preLymph := prePleural := preGGNDouble := false
    if (initialText != "") {
        scan := _LUR_ScanFeatures(initialText)
        featSet := Map()
        for f in scan.features
            featSet[f] := true
        preSpicul    := featSet.Has("spiculation")
        preLymph     := featSet.Has("lymphadenopathy")
        prePleural   := featSet.Has("pleural tethering")
        preGGNDouble := featSet.Has("GGN doubling in 1 year")
    }

    form := RadsForm("Lung-RADS v2022", 620)

    ; Progressive-disclosure layout. Mandatory context first (what is it,
    ; which screening round, how big); everything conditional appears
    ; directly below the field that makes it relevant and is HIDDEN (not
    ; greyed) until then:
    ;   * SolidMm + component-growth row     -> only for part-solid
    ;   * AirwayLoc                          -> only for airway nodules
    ;   * CystFeat                           -> only for atypical cysts
    ;   * BenignFeat                         -> only for parenchymal types
    ;   * whole "Prior comparison" section   -> only for incident rounds
    ;     of parenchymal lesions (a baseline scan has no prior)
    form.Header("Nodule")
    form.Dropdown("LType", "Lesion type:"
        , ["Solid nodule"
        ,  "Cavitary nodule (wall thickening dominant -- managed as solid, note 12d)"
        ,  "Part-solid nodule"
        ,  "Ground-glass nodule (GGN)"
        ,  "Atypical pulmonary cyst"
        ,  "Airway nodule"], typeIdx)
    form.Dropdown("Round", "Screening round:"
        , ["Baseline (first screen)"
        ,  "Incident (new, increased, or decreased from prior)"], 1)
    form.CheckboxRow2("AwaitingPriors", "Awaiting prior exams (note 9 -> Cat 0)"
                   , "KnownCancer",     "Known lung cancer diagnosis (note 16)", false, false)
    form.NumericRow2(
        "SizeMm",  "Mean diameter (mm):", "SolidMm", "Solid component (mm):"
      , sz.mm > 0 ? Round(sz.mm, 1) : 0, 0)
    form.CheckboxNumericRow(
        "NewGrowSolidComp", "Solid component new/growing (note 5)"
      , "NewSolidMm",       "New/growing comp. (mm):", false, 0)
    form.Dropdown("AirwayLoc", "Airway location (note 11):"
        , ["N/A"
        ,  "Subsegmental (benign -- Cat 2, note 11b)"
        ,  "Subsegmental and/or multiple tubular -- favors infection (Cat 0, note 11b)"
        ,  "Segmental or more proximal -- at baseline"
        ,  "Segmental or more proximal -- contains air, favors secretions (Cat 2, note 11c)"
        ,  "Segmental or more proximal -- stable, growing, or persists at 3-mo follow-up (note 11d)"], 1)
    form.Dropdown("CystFeat", "Atypical cyst features (note 12):"
        , ["N/A"
        ,  "Cat 3: Growing cystic component of a thick-walled cyst"
        ,  "Cat 4A: Thick-walled cyst"
        ,  "Cat 4A: Multilocular cyst at baseline"
        ,  "Cat 4A: Thin- or thick-walled that becomes multilocular"
        ,  "Cat 4B: Thick-walled with growing wall thickness or nodularity"
        ,  "Cat 4B: Growing multilocular cyst"
        ,  "Cat 4B: Multilocular with new/increased loculation or opacity"
        ,  "Not classified: Thin-walled cyst (benign, note 12a)"
        ,  "Not classified: Fluid-containing cyst (possibly infectious, note 12g)"
        ,  "Not classified: Multiple cysts (consider LCH / LAM, note 12h)"], 1)
    form.Checkbox("CystNodule", "Cyst has an associated (endophytic / exophytic) nodule -- manage by the most concerning feature (note 12e)", false)
    form.Dropdown("BenignFeat", "Benign features (Cat 1 / Cat 2 override):"
        , ["None"
        ,  "Benign calcification (complete / central / popcorn / concentric ring)"
        ,  "Fat-containing nodule"
        ,  "Juxtapleural <10 mm, solid, smooth, oval / lentiform / triangular"
        ,  "No nodules"], 1)

    form.Header("Prior comparison")
    form.NumericRow2(
        "PriorMm", "Prior diameter (mm; 0=no prior):", "IntervalMo", "Months since prior (0=no prior):", 0, 0)
    form.Dropdown("PriorCat", "Prior Lung-RADS category (auto-inferred for solid / GGN -- override if known):"
        , ["Unknown / not specified"
        ,  "Cat 1 (negative -- no nodules or benign features)"
        ,  "Cat 2 (benign)"
        ,  "Cat 3 (probably benign)"
        ,  "Cat 4A (suspicious)"
        ,  "Cat 4B (very suspicious)"
        ,  "Cat 4X (suspicious with additional features)"], 1)
    form.Note("Behavior auto-fills + locks once prior fields are filled. Change PriorCat to re-trigger inference; clear a prior field to override manually.")
    ; Single dropdown for the nodule's behavior between prior and current
    ; scan -- these states are mutually exclusive observations and were
    ; previously spread across two checkboxes + a separate dropdown, which
    ; allowed contradictory combinations (Growing + Stable at once). "New"
    ; is documented separately from "Not assessed" so the radiologist can
    ; explicitly assert the nodule did not exist on the prior exam (the
    ; classifier maps both to the incident-not-growing size cuts, but the
    ; selected-inputs summary captures the clinical assessment).
    form.Dropdown("Behavior", "Behavior since prior:"
        , ["Not assessed / no prior comparison"
        ,  "New: nodule not present on prior exam"
        ,  "Growing (>1.5 mm in <=12 mo, note 6)"
        ,  "Slow growth across multiple screenings, below threshold (solid/part-solid: note 8 -> 4B; GGN: note 7 -> 2)"
        ,  "Stable or decreased, no specific downgrade -- reclassify by size (note 5)"
        ,  "Cat 3 stable/decreased at 6-mo follow-up (downgrade to Cat 2)"
        ,  "Cat 4A stable/decreased at 3-mo follow-up (downgrade to Cat 3; excludes airway)"
        ,  "Previously biopsied or proven benign by diagnostic workup (downgrade to Cat 2)"], 1)

    form.Header("Features and modifiers (notes 10 / 14 / 15)")
    form.CheckboxRow2("Spiculated",     "Spiculation"
                   , "Lymphadenop",     "Lymphadenopathy", preSpicul, preLymph)
    form.CheckboxRow2("PleuralTether",  "Pleural tethering / retraction / invasion"
                   , "GGNDoubling",     "GGN doubled in size in 1 year", prePleural, preGGNDouble)
    ; Note 10a triggers as a single dropdown: every option routes to the
    ; same Cat 0 + 1-3 month LDCT outcome, so a single choice loses nothing
    ; -- and it keeps the tallest form configuration inside the work area.
    form.Dropdown("InfectFind", "Infectious / inflammatory findings (note 10):"
        , ["None"
        ,  "Segmental or lobar consolidation (-> Cat 0)"
        ,  "Multiple (>6) new nodules (-> Cat 0)"
        ,  "New large solid nodule(s) >=8 mm in a short interval (-> Cat 0)"
        ,  "New nodule(s) in an immunocompromised patient (-> Cat 0)"], 1)
    form.Checkbox("SMod", "S modifier (clinically significant non-lung-cancer finding, note 15)", false)

    ; Lesion-type dependent fields: SolidMm + NewGrowSolidComp + NewSolidMm
    ; only apply to part-solid. AirwayLoc only applies to Airway. CystFeat
    ; only applies to Atypical cyst. SlowGrowth doesn't apply to airway/cyst.
    ; Size + Prior + Interval changes also re-run the growth auto-calc.
    ; LType change resets PriorCat to "Unknown" so the auto-fill re-runs
    ; for the new lesion type. Without this, an inference made earlier
    ; for one lesion type (e.g. "Cat 4A" from solid + prior 9 mm) stays
    ; stuck when the user switches to GGN where prior 9 mm is Cat 2.
    form.OnChange("LType",      _LUR_OnLTypeChange)
    form.OnChange("Round",      _LUR_UpdateForm)
    form.OnChange("SizeMm",     _LUR_UpdateForm)
    form.OnChange("PriorMm",    _LUR_UpdateForm)
    form.OnChange("IntervalMo", _LUR_UpdateForm)
    ; PriorCat change is special -- force re-inference of Behavior even
    ; if it's already at a non-default value, so the radiologist correcting
    ; the inferred PriorCat sees Behavior update accordingly.
    form.OnChange("PriorCat", (f) => _LUR_UpdateForm(f, true))
    _LUR_UpdateForm(form)

    form.SetSubmit(LungRADS_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_LUR_UpdateForm(frm, priorCatChanged := false) {
    t := _LUR_TypeFromLabel(frm.byName["LType"].ctl.Text)
    isParenchymal := (t = "solid" || t = "cavitary" || t = "part-solid" || t = "ggn")
    isIncident := !InStr(frm.byName["Round"].ctl.Text, "Baseline")

    ; Progressive disclosure: irrelevant questions are HIDDEN (the form
    ; reflows and resizes), not greyed. A hidden control reverts to its
    ; default so the classifier never reads a stale answer.
    frm.SetVisible("SolidMm",          t = "part-solid")
    frm.SetVisible("NewGrowSolidComp", t = "part-solid")
    frm.SetVisible("NewSolidMm",       t = "part-solid")
    frm.SetVisible("AirwayLoc",  t = "airway")
    frm.SetVisible("CystFeat",   t = "cyst")
    frm.SetVisible("CystNodule", t = "cyst")
    frm.SetVisible("BenignFeat", isParenchymal)

    ; Prior + Interval + PriorCat + the unified Behavior dropdown are
    ; only meaningful for incident parenchymal lesions. A baseline (first
    ; screen) has no prior to be growing-from / stable-from / decreased-from.
    autoEnabled := isParenchymal && isIncident
    frm.SetSectionVisible("Prior comparison", autoEnabled)

    if !autoEnabled
        return

    ; ---- Auto-infer PriorCat from prior size + lesion type ----
    ; For SOLID and GGN, prior size alone determines the prior category
    ; (assuming baseline-style cuts). For PART-SOLID we'd need the prior
    ; SOLID-component size (not captured in the form), so we leave it at
    ; "Unknown" for the user to set manually. AIRWAY / CYST don't fit
    ; the size-only inference framework. The inference only fires when
    ; PriorCat is at default ("Unknown") so we never clobber a manual pick.
    prior := _LUR_NumVal(frm, "PriorMm")
    priorCatRec := frm.byName["PriorCat"]
    if (prior > 0 && InStr(priorCatRec.ctl.Text, "Unknown")) {
        inferredIdx := _LUR_InferPriorCat(t, prior)
        if (inferredIdx > 0)
            priorCatRec.ctl.Value := inferredIdx
    }

    ; ---- Auto-fill + lock Behavior when full prior comparison given ----
    size     := _LUR_NumVal(frm, "SizeMm")
    interval := _LUR_NumVal(frm, "IntervalMo")
    if (size > 0 && prior > 0 && interval > 0) {
        rate12mo := (size - prior) * 12 / interval
        rec := frm.byName["Behavior"]
        priorCatText := priorCatRec.ctl.Text

        if (rate12mo > 1.5) {
            ; Growth confirmed mathematically -- set to "Growing" per
            ; v2022 note 6 (rate alone determines this state).
            rec.ctl.Value := 3   ; "Growing (>1.5 mm in <=12 mo, note 6)"
        } else if (priorCatChanged || InStr(rec.ctl.Text, "Not assessed")) {
            ; Sub-threshold change (stable or decreased). Auto-SELECT
            ; the specific downgrade option driven by PriorCat + interval.
            ; Fires when Behavior is at default OR when the user just
            ; changed PriorCat (then we re-infer with the new category
            ; even if Behavior was previously auto-set).
            ;
            ; Source rules (v2022 Cat 2 / Cat 3 rows):
            ;   Prior Cat 3 + interval >=6 mo  -> Cat 2 (option 6)
            ;   Prior Cat 4A + interval >=3 mo -> Cat 3 (option 7)
            ;   Cat 4B / biopsy-proven benign: user-only (option 8)
            ;   Cat 1 / 2 / 4X / Unknown: no specific downgrade rule --
            ;     fall back to generic "Stable or decreased" (option 5),
            ;     which uses baseline-style cuts per note 5 (for solid /
            ;     part-solid) and note 7 (for GGN >=30 mm stable).
            autoOption := 5   ; default: generic "Stable or decreased"
            if (InStr(priorCatText, "Cat 3") && interval >= 6)
                autoOption := 6   ; "Cat 3 stable/decreased at 6-mo follow-up..."
            else if (InStr(priorCatText, "Cat 4A") && interval >= 3)
                autoOption := 7   ; "Cat 4A stable/decreased at 3-mo follow-up..."
            rec.ctl.Value := autoOption
        }

        ; Lock the dropdown whenever the user provides a full prior
        ; comparison (size + prior + interval). Consistent UX with the
        ; Growing case. To override the auto-detected option, the user
        ; clears one of the prior fields, picks manually, and re-enters
        ; the prior info (or leaves it cleared). Or they can adjust
        ; PriorCat -- changes propagate via the priorCatChanged path.
        rec.ctl.Enabled := false
        frm._SetMuted(rec, true)
    }
}

; Infer prior Lung-RADS category from prior size + lesion type, assuming
; baseline-style size cuts. Returns the 1-based dropdown index for the
; PriorCat options, or 0 if the lesion type isn't size-only-inferable
; (part-solid needs prior solid component; airway / cyst use different
; scoring frameworks).
_LUR_InferPriorCat(t, priorMm) {
    ; Each `if (cond) stmt` must be on TWO lines -- AHK v2 inline-if
    ; misparse silently swallows the next statement, see
    ; feedback_ahk_v2_inline_if memory note.
    if (t = "solid") {
        if (priorMm < 6)
            return 3   ; Cat 2 (benign): solid <6 baseline
        if (priorMm < 8)
            return 4   ; Cat 3: solid 6-<8 baseline
        if (priorMm < 15)
            return 5   ; Cat 4A: solid 8-<15 baseline
        return 6       ; Cat 4B: solid >=15 baseline
    }
    if (t = "ggn") {
        if (priorMm < 30)
            return 3   ; Cat 2: GGN <30
        return 4       ; Cat 3: GGN >=30
    }
    return 0   ; part-solid / airway / cyst -- not inferable from size alone
}

_LUR_NumVal(frm, name) {
    if !frm.byName.Has(name)
        return 0
    txt := frm.byName[name].ctl.Value
    return IsNumber(txt) ? txt + 0 : 0
}

; Lesion type changed -- reset PriorCat to "Unknown" so the auto-fill in
; _LUR_UpdateForm re-infers it for the new lesion type. Otherwise a stale
; inference (e.g. Cat 4A from solid + prior 9 mm) survives the switch to
; GGN, where prior 9 mm should infer Cat 2.
_LUR_OnLTypeChange(frm) {
    if frm.byName.Has("PriorCat")
        frm.byName["PriorCat"].ctl.Value := 1   ; "Unknown / not specified"
    _LUR_UpdateForm(frm, true)
}

LungRADS_OnSubmit(v, form := "") {
    global g_LastSelectedText
    ntype := _LUR_TypeFromLabel(v.LType)
    size := v.SizeMm + 0.0
    solidMm := v.SolidMm + 0.0
    isBaseline := InStr(v.Round, "Baseline")
    newGrowSolidComp := !!v.NewGrowSolidComp
    newSolidMm := v.NewSolidMm + 0.0
    benignFeat := _LUR_BenignFromLabel(v.BenignFeat)
    airwayLoc := _LUR_AirwayFromLabel(v.AirwayLoc)
    cystFeat := _LUR_CystFromLabel(v.CystFeat)
    sMod := !!v.SMod
    awaitingPriors := !!v.AwaitingPriors
    knownCancer := !!v.KnownCancer
    cystNodule := !!v.CystNodule

    ; Unified Behavior dropdown -> classifier params. The states are
    ; mutually exclusive, which is enforced by the single dropdown.
    behavior := v.Behavior
    isGrowing         := !!InStr(behavior, "Growing (>1.5")
    slowGrowth        := !!InStr(behavior, "Slow growth")
    isStableDecreased := !!InStr(behavior, "Stable or decreased")
    stableDg := ""
    if InStr(behavior, "Cat 3 stable")
        stableDg := "cat3-stable-6mo"
    else if InStr(behavior, "Cat 4A stable")
        stableDg := "cat4a-stable-3mo"
    else if InStr(behavior, "Previously biopsied or proven benign")
        stableDg := "cat4b-proven-benign"

    ; Build a synthetic "report sentence" from explicit feature checkboxes
    ; so the existing _LUR_ScanFeatures pipeline picks them up exactly as
    ; if the radiologist had typed them. Replaces the old free-text
    ; description textarea with structured input.
    ;
    ; CRITICAL: each `if (cond)` body MUST be on its own line. AHK v2's
    ; inline-if form (`if (cond) stmt`) silently executes the statement
    ; as a side effect of the condition expression regardless of the
    ; condition result -- documented misparse, see
    ; feedback_ahk_v2_inline_if memory. Writing the body inline here
    ; was what caused every keyword to be appended unconditionally and
    ; triggered the spurious Cat 0 infectious override.
    text := ""
    if (v.Spiculated = 1)
        text .= " spiculation"
    if (v.Lymphadenop = 1)
        text .= " lymphadenopathy"
    if (v.PleuralTether = 1)
        text .= " pleural tethering"
    if (v.GGNDoubling = 1)
        text .= " doubled in size 1 year"
    if InStr(v.InfectFind, "consolidation")
        text .= " segmental consolidation"
    else if InStr(v.InfectFind, "Multiple")
        text .= " multiple new nodules"
    else if InStr(v.InfectFind, "large solid")
        text .= " new large solid nodules in short interval"
    else if InStr(v.InfectFind, "immunocompromised")
        text .= " immunocompromised patient"
    text := Trim(text)

    r := _LUR_Classify(ntype, size, isBaseline, solidMm, text, isGrowing
                     , benignFeat, airwayLoc, cystFeat, stableDg
                     , newGrowSolidComp, newSolidMm, slowGrowth, sMod
                     , isStableDecreased, awaitingPriors, knownCancer)

    ; v2022 source-gap advisory: a GGN <30 mm that doubled in <=12 months
    ; stays Cat 2 per the size-driven grid (any-growth-state GGN <30 = 2),
    ; but the doubling rate matches note 14's "GGN doubled in 1 year"
    ; 4X-trigger feature -- which only fires for Cat 3/4 baseline GGNs.
    ; This leaves a clinically concerning small-GGN-doubling case scored
    ; the same as any other Cat 2 GGN. Surface a non-algorithmic warning
    ; so the radiologist can apply clinical judgment about follow-up.
    priorMmIn := v.PriorMm + 0.0
    intervalMoIn := v.IntervalMo + 0.0
    if (ntype = "ggn" && r.category = "2"
        && priorMmIn > 0 && intervalMoIn > 0 && intervalMoIn <= 12
        && size >= 2 * priorMmIn) {
        advisory := "Rapid doubling: " priorMmIn " mm -> " Round(size, 1) " mm in " Round(intervalMoIn, 0) " mo. "
                  . "The v2022 grid keeps a GGN <30 mm at Cat 2 regardless of growth rate, but this rate matches the 'GGN doubled in 1 year' 4X trigger (note 14; only promotes Cat 3/4 GGNs). "
                  . "Clinical judgment may warrant short-interval follow-up or specialist consultation despite the Cat 2 designation."
        r.note := (r.note != "" ? r.note . " " : "") . advisory
    }

    sizePhrase := size > 0 ? Format("{:.0f}", size) " mm " : ""
    typeWord := _LUR_TypeWord(ntype)
    descLc := (r.descriptor != "") ? StrLower(SubStr(r.descriptor, 1, 1)) SubStr(r.descriptor, 2) : ""
    shortRec := _LUR_ShortRec(r.category)
    if (r.category = "NC") {
        ; Not-classified pathways (notes 12a / 12g / 12h / 16) have no
        ; Lung-RADS category -- the impression states why instead of
        ; asserting a numbered score.
        impression := r.reason "; not classified in Lung-RADS per Lung-RADS v2022. "
                    . "Exam category is determined by the most suspicious classified finding, if any."
    } else {
        impression := sizePhrase typeWord " pulmonary nodule, Lung-RADS " r.category
                   . " (" descLc ") per Lung-RADS v2022"
        if (shortRec != "")
            impression .= "; " shortRec
        if (SubStr(impression, -1) != ".")
            impression .= "."
    }

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nCategory: " r.category " -- " r.descriptor
    method .= "`nPopulation prevalence: " r.prevalence
    method .= "`nFull management: " r.management
    if (r.suspiciousFeatures.Length > 0)
        method .= "`nSuspicious features detected: " JoinArr(r.suspiciousFeatures, ", ")
    if (r.infectiousFeatures.Length > 0)
        method .= "`nIndeterminate infectious/inflammatory features: " JoinArr(r.infectiousFeatures, ", ")
    if (r.baseCategory != "" && r.baseCategory != r.category)
        method .= "`nBase size-derived category: " r.baseCategory

    ; Practice audit definitions (note 3), keyed off the final category.
    catBase := RegExReplace(r.category, "S$", "")
    if (catBase = "1" || catBase = "2")
        method .= "`nPractice audit (note 3): NEGATIVE screen (categories 1-2). A negative screen does not mean the individual does not have lung cancer."
    else if (catBase = "3" || catBase = "4A" || catBase = "4B" || catBase = "4X")
        method .= "`nPractice audit (note 3): POSITIVE screen (categories 3-4)."
    method .= "`nCoding (note 1): code the exam 0-4 by the nodule with the highest degree of suspicion."
    method .= "`nMeasurement (note 4): mean diameter = (long + short axis) / 2, each axis to 0.1 mm in any plane; volumes to the nearest whole mm3."

    advisories := []
    if (r.note != "")
        advisories.Push(r.note)
    if (ntype = "cyst" && cystNodule)
        advisories.Push("Cyst with an associated nodule (note 12e): management is based on Lung-RADS criteria for the most concerning feature -- also classify the associated nodule under its own lesion type (solid / part-solid / GGN) and act on the higher category.")

    return MakeResult({
        classification: r.category = "NC" ? "Lung-RADS: not classified" : "Lung-RADS " r.category,
        impression:     impression,
        recommendation: "",
        advisories:     advisories,
        methodology:    method,
        citations:      [{ text: "American College of Radiology Committee on Lung-RADS. "
                                . "Lung CT Screening Reporting and Data System (Lung-RADS) v2022. "
                                . "ACR, Release November 2022.",
                           url:  "https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/Lung-RADS" }],
        echo:           g_LastSelectedText
    })
}

_LUR_TypeWord(t) {
    if (t = "solid")
        return "solid"
    if (t = "cavitary")
        return "cavitary"
    if (t = "part-solid")
        return "part-solid"
    if (t = "ggn")
        return "ground-glass"
    if (t = "airway")
        return "airway"
    if (t = "cyst")
        return "atypical cystic"
    return ""
}

; Short impression-line recommendation per Lung-RADS category. The verbose
; multi-sentence management text (PET/CT options, McWilliams discussion, etc.)
; lives in `r.management` and is preserved in the methodology block; the
; impression keeps just the headline action.
_LUR_ShortRec(cat) {
    catBase := RegExReplace(cat, "S$", "")   ; strip S modifier suffix
    if (catBase = "0")
        return "recommend additional imaging or comparison to prior"
    if (catBase = "1" || catBase = "2")
        return "continue annual screening LDCT"
    if (catBase = "3")
        return "recommend 6-month LDCT"
    if (catBase = "4A")
        return "recommend 3-month LDCT"
    if (catBase = "4B" || catBase = "4X")
        return "recommend diagnostic CT, PET/CT, or tissue sampling"
    return ""
}

_LUR_TypeFromLabel(label) {
    if InStr(label, "Cavitary")
        return "cavitary"
    if InStr(label, "Atypical")
        return "cyst"
    if InStr(label, "Airway")
        return "airway"
    if InStr(label, "Ground-glass") || InStr(label, "GGN")
        return "ggn"
    if InStr(label, "Part-solid")
        return "part-solid"
    return "solid"
}
_LUR_BenignFromLabel(label) {
    if InStr(label, "calcification")
        return "benign-calc"
    if InStr(label, "Fat-containing")
        return "fat"
    if InStr(label, "Juxtapleural")
        return "juxtapleural"
    if InStr(label, "No nodules")
        return "no-nodules"
    return "none"
}
_LUR_AirwayFromLabel(label) {
    ; Check the more specific variants before their broader siblings --
    ; "Subsegmental" is a prefix of both note-11b options, and "Segmental
    ; or more proximal" starts three different rows.
    if InStr(label, "favors infection")
        return "subseg-infectious"
    if InStr(label, "Subsegmental")
        return "subsegmental"
    if InStr(label, "favors secretions")
        return "seg-air-secretions"
    if InStr(label, "Segmental or more proximal -- at baseline")
        return "seg-baseline"
    if InStr(label, "Segmental or more proximal -- stable")
        return "seg-stable-or-growing"
    return ""
}
_LUR_CystFromLabel(label) {
    if InStr(label, "Cat 3: Growing")
        return "cyst-3"
    if InStr(label, "Cat 4A: Thick-walled cyst")
        return "cyst-4a-thick"
    if InStr(label, "Cat 4A: Multilocular cyst at baseline")
        return "cyst-4a-multi-base"
    if InStr(label, "Cat 4A: Thin")
        return "cyst-4a-becomes-multi"
    if InStr(label, "Cat 4B: Thick-walled with growing")
        return "cyst-4b-thick-grow"
    if InStr(label, "Cat 4B: Growing multilocular")
        return "cyst-4b-grow-multi"
    if InStr(label, "Cat 4B: Multilocular with new/increased")
        return "cyst-4b-multi-incr"
    if InStr(label, "Not classified: Thin-walled")
        return "cyst-nc-thin"
    if InStr(label, "Not classified: Fluid-containing")
        return "cyst-nc-fluid"
    if InStr(label, "Not classified: Multiple cysts")
        return "cyst-nc-multiple"
    return ""
}
_LUR_Classify(ntype, size, isBaseline, solidMm, text, isGrowing
            , benignFeat := "none", airwayLoc := "", cystFeat := "", stableDg := ""
            , newGrowSolidComp := false, newSolidMm := 0, slowGrowth := false
            , sMod := false, isStableDecreased := false
            , awaitingPriors := false, knownCancer := false) {
    ; ---- Step 0: exam-level states (notes 9 / 16) ----
    ; A lung cancer diagnosis takes the exam out of screening entirely, so
    ; it outranks every nodule-level rule (text scan skipped -- 4X /
    ; infectious routing is meaningless outside screening).
    if knownCancer
        return _LUR_Finalize(_LUR_Cat("NC", "Known lung cancer diagnosis -- further imaging is for staging / management, no longer screening (note 16)"), "", sMod)
    if awaitingPriors
        return _LUR_Finalize(_LUR_Cat("0", "Awaiting prior exams -- Lung-RADS 0 is temporary until the comparison study is available and a new category is assigned (note 9)"), text, sMod)

    ; ---- Step 1: explicit downgrade overrides (v2022 Cat 2 / Cat 3 rows) ----
    ; These short-circuit the size-based classifier.
    if (stableDg = "cat3-stable-6mo" || stableDg = "cat4b-proven-benign")
        return _LUR_Finalize(_LUR_Cat("2", "Downgrade rule -- Cat 3 stable at 6-mo follow-up OR Cat 4B proven benign in etiology"), text, sMod)
    if (stableDg = "cat4a-stable-3mo")
        return _LUR_Finalize(_LUR_Cat("3", "Downgrade rule -- Cat 4A stable at 3-mo follow-up (excludes airway nodules)"), text, sMod)

    ; ---- Step 2: benign overrides (v2022 Cat 1 / Cat 2 morphology) ----
    if (benignFeat = "no-nodules" || benignFeat = "benign-calc" || benignFeat = "fat") {
        reason := benignFeat = "no-nodules" ? "No lung nodules"
                : benignFeat = "fat"        ? "Fat-containing nodule (benign feature, note table)"
                                            : "Benign calcification pattern (complete / central / popcorn / concentric ring)"
        return _LUR_Finalize(_LUR_Cat("1", reason), text, sMod)
    }
    if (benignFeat = "juxtapleural") {
        ; Juxtapleural <10mm + solid + smooth + oval/lentiform/triangular -> Cat 2
        if (size > 0 && size < 10)
            return _LUR_Finalize(_LUR_Cat("2", "Juxtapleural nodule <10 mm, solid, smooth, oval / lentiform / triangular"), text, sMod)
        ; If the user picked juxtapleural but size >=10 the morphology rule
        ; doesn't apply -- fall through to standard size-based classifier.
    }

    ; ---- Step 3: explicit lesion-type branches ----
    if (ntype = "airway") {
        if (airwayLoc = "subsegmental")
            r := _LUR_Cat("2", "Airway nodule, subsegmental -- at baseline / new / stable (note 11b)")
        else if (airwayLoc = "subseg-infectious")
            r := _LUR_Cat("0", "Airway abnormality, subsegmental and/or multiple tubular -- favors an infectious process, no underlying obstructive nodule (note 11b)")
        else if (airwayLoc = "seg-baseline")
            r := _LUR_Cat("4A", "Airway nodule, segmental or more proximal -- at baseline (note 11a)")
        else if (airwayLoc = "seg-air-secretions")
            r := _LUR_Cat("2", "Airway abnormality, segmental or more proximal, containing air -- favors secretions, no underlying soft-tissue nodule (note 11c)")
        else if (airwayLoc = "seg-stable-or-growing")
            r := _LUR_Cat("4B", "Airway nodule, segmental or more proximal -- stable, growing, or persisting at 3-mo follow-up (notes 11a / 11d)")
        else
            r := _LUR_Cat("0", "Airway nodule selected but airway location not specified; please pick a location")
        return _LUR_Finalize(r, text, sMod, true)
    }
    if (ntype = "cyst") {
        ; Not-classified cyst rows (notes 12a / 12g / 12h): these are not
        ; scored in Lung-RADS at all -- the exam category comes from other
        ; classified findings.
        if (cystFeat = "cyst-nc-thin" || cystFeat = "cyst-nc-fluid" || cystFeat = "cyst-nc-multiple") {
            reason := cystFeat = "cyst-nc-thin"  ? "Thin-walled cyst -- considered benign (note 12a)"
                    : cystFeat = "cyst-nc-fluid" ? "Fluid-containing cyst -- may represent an infectious process (note 12g)"
                                                 : "Multiple cysts -- may indicate an alternative diagnosis such as LCH or LAM (note 12h)"
            r := _LUR_Cat("NC", reason)
            if (cystFeat = "cyst-nc-fluid")
                r.note := "Fluid-containing cysts may represent an infectious process; correlate clinically. Classify in Lung-RADS only if other concerning features are identified (note 12g)."
            else if (cystFeat = "cyst-nc-multiple")
                r.note := "Multiple cysts may indicate Langerhans cell histiocytosis (LCH) or lymphangioleiomyomatosis (LAM); consider dedicated evaluation (note 12h)."
            return _LUR_Finalize(r, text, sMod)
        }
        cystCat := cystFeat = "cyst-3"             ? "3"
                : cystFeat = "cyst-4a-thick"        ? "4A"
                : cystFeat = "cyst-4a-multi-base"   ? "4A"
                : cystFeat = "cyst-4a-becomes-multi" ? "4A"
                : cystFeat = "cyst-4b-thick-grow"   ? "4B"
                : cystFeat = "cyst-4b-grow-multi"   ? "4B"
                : cystFeat = "cyst-4b-multi-incr"   ? "4B"
                                                    : ""
        if (cystCat = "")
            r := _LUR_Cat("0", "Atypical cyst selected but cyst features not specified; pick a category from the cyst-feature dropdown")
        else
            r := _LUR_Cat(cystCat, "Atypical pulmonary cyst per note 12")
        return _LUR_Finalize(r, text, sMod)
    }

    ; ---- Step 4: standard solid / part-solid / GGN size classifier ----
    ; A cavitary nodule is classified on the solid pathway by total mean
    ; diameter (note 12d); the only difference is the advisory we attach.
    if (ntype = "solid" || ntype = "cavitary") {
        r := _LUR_Solid(size, isBaseline, isGrowing, isStableDecreased)
        if (ntype = "cavitary")
            r.note := "Cavitary nodule: wall thickening is the dominant feature; managed as a solid nodule by total mean diameter (Lung-RADS v2022 note 12d)."
    }
    else if (ntype = "part-solid")
        r := _LUR_PartSolid(size, solidMm, isBaseline, newGrowSolidComp, newSolidMm, isStableDecreased)
    else if (ntype = "ggn")
        r := _LUR_GGN(size, isStableDecreased)
    else
        r := _LUR_Cat("0", "Unknown lesion type")

    ; ---- Step 5: slow-growth over multiple screenings -> Cat 4B (note 8) ----
    ; Applies to solid AND part-solid lesions only (not GGN; GGN slow-growth
    ; goes to Cat 2 per note 7, which is already handled by the GGN classifier).
    if (slowGrowth && (ntype = "solid" || ntype = "part-solid" || ntype = "cavitary")) {
        ; Only upgrade if the size-derived category is below 4B already.
        if (r.category = "3" || r.category = "4A") {
            base := r.category
            r := _LUR_Cat("4B", "Slow growth over multiple screenings without meeting >1.5 mm threshold (note 8)")
            r.baseCategory := base
        }
    }
    ; GGN slow growth goes the OTHER way: a GGN growing below the >1.5 mm
    ; 12-month threshold stays Cat 2 per note 7 (only a >=30 mm GGN, which
    ; sizes to Cat 3, needs the downgrade -- <30 mm is already Cat 2).
    if (slowGrowth && ntype = "ggn" && r.category = "3") {
        base := r.category
        r := _LUR_Cat("2", "GGN growing over multiple screenings below the >1.5 mm / 12-mo threshold (note 7)")
        r.baseCategory := base
    }

    return _LUR_Finalize(r, text, sMod)
}

_LUR_Solid(size, isBaseline, isGrowing := false, isStableDecreased := false) {
    ; Baseline cuts also apply when the lesion is on follow-up but stable
    ; or decreased without a specific Cat 3/4A/4B downgrade rule firing.
    ; Per v2022 note 5: "When a nodule crosses a new size threshold... it
    ; should be reclassified based on size and managed accordingly."
    ; Baseline-style cuts are the framework for an existing, non-growing,
    ; non-new nodule.
    if (isBaseline || isStableDecreased) {
        sfx := isStableDecreased
            ? " (stable/decreased on follow-up; reclassified by size per note 5)"
            : " at baseline"
        if (size < 6)
            return _LUR_Cat("2",  "Solid nodule <6 mm" sfx)
        if (size < 8)
            return _LUR_Cat("3",  "Solid nodule 6 to <8 mm" sfx)
        if (size < 15)
            return _LUR_Cat("4A", "Solid nodule 8 to <15 mm" sfx)
        return _LUR_Cat("4B", "Solid nodule >=15 mm" sfx)
    }
    ; v2022 incident path: GROWING and NEW have different size cuts.
    ;   Cat 4A solid: "Growing <8 mm" (any size below 8 once growth documented)
    ;   Cat 4B solid: "New or growing >=8 mm"
    if isGrowing {
        if (size < 8)
            return _LUR_Cat("4A", "Growing solid nodule <8 mm (v2022 Cat 4A: 'Growing <8 mm')")
        return _LUR_Cat("4B", "Growing solid nodule >=8 mm")
    }
    ; New lesion (first detection on follow-up; no prior measurement).
    if (size < 4)
        return _LUR_Cat("2",  "New solid nodule <4 mm")
    if (size < 6)
        return _LUR_Cat("3",  "New solid nodule 4 to <6 mm")
    if (size < 8)
        return _LUR_Cat("4A", "New solid nodule 6 to <8 mm")
    return _LUR_Cat("4B", "New solid nodule >=8 mm")
}

_LUR_PartSolid(size, solidMm, isBaseline, newGrowSolidComp, newSolidMm, isStableDecreased := false) {
    ; Part-solid rules per v2022. New-or-growing-solid-component path overrides
    ; size-based path when the radiologist has documented that the solid
    ; component is new or growing (note 5).
    if newGrowSolidComp {
        if (newSolidMm >= 4)
            return _LUR_Cat("4B", "Part-solid: new or growing >=4 mm solid component")
        return _LUR_Cat("4A", "Part-solid: new or growing <4 mm solid component")
    }
    ; Baseline-style cuts apply for an existing, stable/decreased part-
    ; solid lesion (note 5 reclassification).
    if (isBaseline || isStableDecreased) {
        sfx := isStableDecreased
            ? " (stable/decreased on follow-up; reclassified by size per note 5)"
            : " at baseline"
        if (size < 6)
            return _LUR_Cat("2", "Part-solid <6 mm total" sfx)
        if (solidMm < 6)
            return _LUR_Cat("3", "Part-solid >=6 mm total with solid component <6 mm" sfx)
        if (solidMm < 8)
            return _LUR_Cat("4A", "Part-solid >=6 mm total with solid component 6 to <8 mm" sfx)
        return _LUR_Cat("4B", "Part-solid with solid component >=8 mm" sfx)
    }
    ; Incident path: new part-solid lesion as a whole (the lesion itself is
    ; new, not just the solid component growing). Per v2022 Cat 3 row:
    ; "New <6 mm total mean diameter" -> Cat 3.
    if (size < 6)
        return _LUR_Cat("3", "New part-solid <6 mm total mean diameter")
    ; Larger new part-solid: classify by current solid-component size.
    if (solidMm < 6)
        return _LUR_Cat("3", "New part-solid with solid component <6 mm")
    if (solidMm < 8)
        return _LUR_Cat("4A", "New part-solid with solid component 6 to <8 mm")
    return _LUR_Cat("4B", "New part-solid with solid component >=8 mm")
}

_LUR_GGN(size, isStableDecreased := false) {
    if (size < 30)
        return _LUR_Cat("2", "GGN <30 mm at baseline, new, or growing")
    ; v2022 Cat 2 row also covers "GGN >=30 mm stable or slowly growing"
    ; per note 7 -- explicitly downgrade to 2 here when the user has
    ; indicated stable/decreased behavior.
    if isStableDecreased
        return _LUR_Cat("2", "GGN >=30 mm, stable or slowly growing (note 7)")
    return _LUR_Cat("3", "GGN >=30 mm at baseline or new")
}

; Finalize: applies 4X upgrade (suspicious features from text scan), infection
; routing (with negation + tree-in-bud exclusion per note 10c), and the S
; modifier suffix.
_LUR_Finalize(r, text, sMod, isAirway := false) {
    r.suspiciousFeatures := r.HasProp("suspiciousFeatures") ? r.suspiciousFeatures : []
    r.infectiousFeatures := r.HasProp("infectiousFeatures") ? r.infectiousFeatures : []
    if !r.HasProp("baseCategory")
        r.baseCategory := ""
    if !r.HasProp("note")
        r.note := ""

    if (text != "") {
        check := _LUR_ScanFeatures(text)
        ; Note 10c: tree-in-bud and small new <3 cm GGN don't warrant
        ; short-term follow-up -- those triggers are excluded from the scan.
        if (check.infectious && r.category != "1") {
            r.infectiousFeatures := check.infectiousFeatures
            r.baseCategory := r.category
            r.category := "0"
            r.descriptor := "Incomplete -- indeterminate infectious / inflammatory findings"
            r.prevalence := "~1%"
            r.management := "Indeterminate infectious / inflammatory features ("
                . JoinArr(check.infectiousFeatures, ", ")
                . "). 1-3 month LDCT recommended per note 10a. At follow-up, reassign Lung-RADS category based on the most suspicious finding."
            return _LUR_ApplyS(r, sMod)
        }
        if (check.suspicious && (r.category = "3" || r.category = "4A" || r.category = "4B")) {
            r.suspiciousFeatures := check.features
            r.baseCategory := r.category
            r.category := "4X"
            r.descriptor := "Cat 3 or 4 nodule with additional features suggesting malignancy (note 14)"
            r.prevalence := "<1%"
            r.management := "Suspicious features present (" JoinArr(check.features, ", ")
                . "). Chest CT with or without contrast, PET/CT, and/or tissue sampling depending on probability of malignancy and comorbidities. Use the McWilliams et al. assessment tool to support recommendations (note 13)."
        }
    }
    return _LUR_ApplyS(r, sMod)
}

_LUR_ApplyS(r, sMod) {
    if sMod {
        ; Per note 15, S modifier may be added to categories 0-4 -- a
        ; not-classified result gets the advisory but no "S" suffix.
        if (r.category != "NC")
            r.category := r.category . "S"
        r.note := (r.note != "" ? r.note . " " : "")
            . "S modifier added: clinically significant non-lung-cancer finding. Management per ACR Incidental Findings recommendations."
    }
    return r
}

; _LUR_Cat builds the result struct with category metadata pulled from a
; static lookup. Cancer-risk percentages are NOT included; the v2022 source
; only publishes population prevalence in the main grid, so we surface that.
_LUR_Cat(category, reason) {
    info := _LUR_CatInfo().Has(category) ? _LUR_CatInfo()[category] : _LUR_CatInfo()["0"]
    return { category: category
           , descriptor: info.descriptor
           , prevalence: info.prevalence
           , management: info.management
           , reason: reason
           , suspiciousFeatures: []
           , infectiousFeatures: []
           , baseCategory: ""
           , note: "" }
}

_LUR_CatInfo() {
    static m := _LUR_BuildCatInfo()
    return m
}
_LUR_BuildCatInfo() {
    m := Map()
    m["0"]  := { descriptor: "Incomplete"
               , prevalence: "~1%"
               , management: "Comparison to prior chest CT; additional lung cancer screening CT imaging needed; or 1-3 month LDCT for indeterminate infectious / inflammatory findings (note 10)." }
    m["1"]  := { descriptor: "Negative -- no nodules or benign-feature nodules"
               , prevalence: "39%"
               , management: "Continue annual screening with LDCT in 12 months." }
    m["2"]  := { descriptor: "Benign -- imaging features or indolent behavior"
               , prevalence: "45%"
               , management: "Continue annual screening with LDCT in 12 months." }
    m["3"]  := { descriptor: "Probably benign -- imaging features or behavior"
               , prevalence: "9%"
               , management: "6-month LDCT." }
    m["4A"] := { descriptor: "Suspicious"
               , prevalence: "4%"
               , management: "3-month LDCT; PET/CT may be considered if there is a >=8 mm solid nodule or solid component." }
    m["4B"] := { descriptor: "Very suspicious"
               , prevalence: "2%"
               , management: "Diagnostic chest CT with or without contrast; PET/CT may be considered if there is a >=8 mm solid nodule or solid component; tissue sampling; and/or referral for further clinical evaluation. Management depends on clinical evaluation, patient preference, and probability of malignancy (note 13)." }
    m["4X"] := { descriptor: "Cat 3 or 4 with additional features suggesting malignancy (note 14)"
               , prevalence: "<1%"
               , management: "Chest CT with or without contrast, PET/CT, and/or tissue sampling; consider McWilliams assessment tool." }
    m["NC"] := { descriptor: "Not classified in Lung-RADS"
               , prevalence: "n/a"
               , management: "Not classified or managed in Lung-RADS (v2022 notes 12a / 12g / 12h / 16). The exam's Lung-RADS category is determined by the most suspicious classified finding, if any." }
    return m
}

_LUR_ScanFeatures(text) {
    out := { suspicious: false, features: [], infectious: false, infectiousFeatures: [] }
    low := StrLower(text)

    ; Indeterminate infectious / inflammatory triggers, narrowed to v2022 note 10a
    ; (segmental / lobar consolidation, multiple new nodules, large solid nodules
    ; in short interval, immunocompromised). tree-in-bud and small new GGN are
    ; specifically EXCLUDED per note 10c.
    infectious := [
        ["segmental\s+consolid|lobar\s+consolid", "segmental / lobar consolidation"]
      , ["multiple\s+new\s+nodules|more\s+than\s+six\s+new\s+nodules|>?\s*6\s+new\s+nodules", "multiple new nodules"]
      , ["immunocompromised|immunosuppress", "immunocompromised context"]
      , ["new\s+large\s+solid|large\s+solid\s+nodules?.{0,30}short\s+interval", "large solid nodule(s) >=8 mm appearing in a short interval"]
      , ["mucoid\s+impact", "mucoid impaction"]
      , ["bronchopneumon", "bronchopneumonia pattern"]
    ]
    for entry in infectious {
        if (pos := RegExMatch(low, entry[1])) {
            if !_LUR_IsNegated(low, pos)
                out.infectiousFeatures.Push(entry[2])
        }
    }
    if (out.infectiousFeatures.Length > 0)
        out.infectious := true

    ; 4X suspicious features per note 14.
    suspicious := [
        ["spicul", "spiculation"]
      , ["pleural\s+(tether|retract|attach|invol)", "pleural tethering"]
      , ["chest\s+wall\s+inv", "chest wall invasion"]
      , ["rib\s+(erosion|destruct)", "rib involvement"]
      , ["frank\s+metasta", "frank metastatic disease"]
      , ["lymphadenopath|enlarged.*lymph|hilar.*node|mediast.*node", "lymphadenopathy"]
      , ["doubled?\s+in\s+size.*1\s+year|doubling.*1\s+year|GGN\s+that\s+double", "GGN doubling in 1 year"]
      , ["bronchovasc.*distort|distort.*bronchovasc", "bronchovascular distortion"]
      , ["highly\s+suspi.*malig|suspi.*malignan", "report describes high suspicion for malignancy"]
    ]
    for entry in suspicious {
        if (pos := RegExMatch(low, entry[1])) {
            if !_LUR_IsNegated(low, pos)
                out.features.Push(entry[2])
        }
    }
    if (out.features.Length > 0)
        out.suspicious := true
    return out
}

_LUR_IsNegated(text, matchPos) {
    start := Max(1, matchPos - 30)
    prefix := SubStr(text, start, matchPos - start)
    if RegExMatch(prefix, "i)\b(no|not|without|absent|negative|denies|deny|none|neither|nor)\s+\S*\s*$")
        return true
    if RegExMatch(prefix, "i)\b(no\s+evidence\s+of|no\s+signs?\s+of|no\s+definite)\s+\S*\s*$")
        return true
    return false
}
