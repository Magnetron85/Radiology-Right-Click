; ============================================================
; lib/calc/LIRADS.ahk -- LI-RADS v2018 diagnostic algorithm
; ------------------------------------------------------------
; Reference: American College of Radiology LI-RADS v2018 Core.
;
; Applies to observations in patients at risk for HCC.
; Outputs LR-NC, LR-1 / LR-2 / LR-3 / LR-4 / LR-5, LR-M, LR-TIV,
; or LR-TR Viable / Equivocal / Nonviable / Nonevaluable.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

LIRADS_Entry(input) {
    ShowLIRADSDialog(input)
    return ""
}

ShowLIRADSDialog(text := "") {
    sz := TextScan.Size(text)

    aphe    := TextScan.ContainsAny(text
        , ["\b(?:non[- ]rim\s+)?arterial\s+phase\s+(?:hyper)?enhancement\b","\bAPHE\b","\bhyperenhanc.*arterial"])
    washout := TextScan.ContainsAny(text
        , ["\b(?:non[- ]peripheral\s+)?washout\b"])
    capsule := TextScan.ContainsAny(text
        , ["\benhancing\s+capsule\b","\b(?<!non)capsule\s+appear"])
    growth  := TextScan.ContainsAny(text
        , ["\bthreshold\s+growth\b","\binterval\s+(?:increase|enlarge)","\bgrown\s+(?:from|by)\b"])
    tiv     := TextScan.ContainsAny(text
        , ["\btumor\s+(?:in\s+)?(?:vein|thrombus)\b","\bTIV\b","\bportal\s+vein\s+(?:tumor\s+)?thrombus","\bvascular\s+invasion\b"])
    bp      := TextScan.ContainsAny(text
        , ["\bhemangioma","\bparallels\s+blood\s+pool","\bblood\s+pool\s+enhancement"])

    targetoid := TextScan.ContainsAny(text, ["\btargetoid\b","\btarget[- ]like\b"])
    infiltr   := TextScan.ContainsAny(text, ["\binfiltrative\b"])
    rimAphe   := TextScan.ContainsAny(text, ["\brim\s+(?:arterial|APHE)"])
    perWash   := TextScan.ContainsAny(text, ["\bperipheral\s+washout"])
    delCent   := TextScan.ContainsAny(text, ["\bdelayed\s+central\s+enhancement"])
    markDWI   := TextScan.ContainsAny(text, ["\bmarked\s+(?:diffusion|DWI)\s+restrict"])
    necrosis  := TextScan.ContainsAny(text, ["\bnecros","\bsevere\s+ischemia"])

    form := RadsForm("LI-RADS v2018 (CT / MRI)", 620)

    form.Header("Image quality")
    form.Checkbox("IsNC", "Image degradation / omission prevents categorization (LR-NC)", false)

    form.Header("Major features")
    form.Checkbox("Aphe",   "Non-rim arterial phase hyperenhancement (APHE)", aphe)
    form.Numeric("SizeMm",  "Observation size (mm):", sz.mm > 0 ? Round(sz.mm) : 0)
    form.CheckboxRow2("Washout", "Non-peripheral washout"
                   , "Capsule", "Enhancing capsule", washout, capsule)
    form.Checkbox("Growth", "Threshold growth (>=50% size increase in <=6 months)", growth)
    form.Checkbox("Tiv",    "Tumor in vein (TIV) -- overrides table -> LR-TIV", tiv)
    form.Checkbox("BloodPool", "Hemangioma-type enhancement parallels blood pool", bp)

    form.Header("LR-M criteria (any independent feature -> LR-M)")
    form.CheckboxRow2("LM_Target",    "Targetoid mass (concentric pattern of rim APHE + peripheral washout / delayed central)"
                   , "LM_Infiltr",   "Infiltrative appearance", targetoid, infiltr)
    form.CheckboxRow2("LM_MarkedDWI", "Marked diffusion restriction"
                   , "LM_Necrosis",  "Necrosis or severe ischemia", markDWI, necrosis)

    form.Note("Targetoid components below are descriptive subtypes per LI-RADS 2018 -- if >=2 are present they imply targetoid morphology even without the explicit checkbox above.")
    form.CheckboxRow2("LM_RimAphe",   "Rim arterial phase hyperenhancement"
                   , "LM_PerWashout", "Peripheral washout", rimAphe, perWash)
    form.Checkbox("LM_DelCenter", "Delayed central enhancement", delCent)

    form.Header("Definitely / probably benign overrides")
    form.Checkbox("DefBenign", "Definitely benign entity (cyst, hemangioma, perfusion alteration, focal fat / sparing, hypertrophic pseudomass, confluent fibrosis / focal scar, spontaneous disappearance) -> LR-1", false)
    form.Checkbox("DistinctBenign", "Distinctive benign-appearing solid nodule <20 mm (no major HCC feature, no LR-M feature, no ancillary malignancy) -> LR-2", false)

    form.Header("Ancillary feature counts")
    form.Numeric("AncMalig", "Number of ancillary features favoring malignancy (general):", 0)
    form.Numeric("AncBenign", "Number of ancillary features favoring benignity:", 0)
    form.Numeric("AncHCC",   "Number of HCC-specific ancillary features:", 0)

    ; IsNC = image degradation / omission prevents categorization; when ticked
    ; the algorithm returns LR-NC immediately, so the rest of the form is
    ; irrelevant. DefBenign + DistinctBenign similarly short-circuit but those
    ; still want the user to confirm size, so they're not exclusion-tied.
    form.DisableWhenAnyChecked(
        ["IsNC"]
      , ["Aphe", "SizeMm", "Washout", "Capsule", "Growth", "Tiv", "BloodPool"
       , "LM_Target", "LM_Infiltr", "LM_RimAphe", "LM_PerWashout"
       , "LM_DelCenter", "LM_MarkedDWI", "LM_Necrosis"
       , "DefBenign", "DistinctBenign"
       , "AncMalig", "AncBenign", "AncHCC"])

    form.SetSubmit(LIRADS_OnSubmit)
    form.AddButtons()
    form.Show()
}

LIRADS_OnSubmit(v, form := "") {
    global g_LastSelectedText
    aphe := !!v.Aphe
    size := v.SizeMm + 0.0
    wash := !!v.Washout
    cap  := !!v.Capsule
    grow := !!v.Growth
    tiv  := !!v.Tiv
    bp   := !!v.BloodPool
    isNc := !!v.IsNC

    ; LR-M criteria per LI-RADS v2018 Core (lines 1016-1024): targetoid mass OR
    ; nontargetoid mass with infiltrative appearance, marked diffusion restriction,
    ; or necrosis / severe ischemia. Rim APHE, peripheral washout, and delayed
    ; central enhancement are described as components of TARGETOID morphology
    ; (Core lines 1042-1053), not standalone LR-M qualifiers. Treat >=2 of those
    ; three components as implicit targetoid morphology.
    targetComponentCount := (v.LM_RimAphe ? 1 : 0) + (v.LM_PerWashout ? 1 : 0) + (v.LM_DelCenter ? 1 : 0)
    isTargetoid := v.LM_Target || targetComponentCount >= 2

    lrm := []
    if isTargetoid
        lrm.Push(v.LM_Target ? "targetoid" : "targetoid morphology (" targetComponentCount " components)")
    if (v.LM_Infiltr)
        lrm.Push("infiltrative")
    if (v.LM_MarkedDWI)
        lrm.Push("marked diffusion restriction")
    if (v.LM_Necrosis)
        lrm.Push("necrosis")

    defBenign     := !!v.DefBenign
    distinctBenign := !!v.DistinctBenign

    aM := SafeInt(v.AncMalig, 0)
    aB := SafeInt(v.AncBenign, 0)
    aHcc := SafeInt(v.AncHCC, 0)

    r := _LR_Classify(aphe, size, wash, cap, grow, tiv, lrm, aM, aB, isNc, aHcc, bp
                    , defBenign, distinctBenign)

    sizePhrase := size > 0 ? Format("{:.0f}", size) " mm " : ""
    ; r.mgmt is sentence-cased; lowercase the leading char so it flows after "; ".
    mgmtLc := (r.mgmt != "") ? StrLower(SubStr(r.mgmt, 1, 1)) SubStr(r.mgmt, 2) : ""
    descLc := (r.description != "") ? StrLower(SubStr(r.description, 1, 1)) SubStr(r.description, 2) : ""
    impression := sizePhrase "hepatic observation, " r.category " (" descLc ")"
               . " per LI-RADS v2018"
    if (mgmtLc != "")
        impression .= "; " mgmtLc
    if (SubStr(impression, -1) != ".")
        impression .= "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nCategory: " r.category " -- " r.description
    if (r.adjustmentReason != "")
        method .= "`nReason: " r.adjustmentReason
    method .= "`nFull management: " r.mgmt

    return MakeResult({
        classification: r.category,
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "Chernyak V, Fowler KJ, Kamaya A, et al. "
                                . "Liver Imaging Reporting and Data System (LI-RADS) Version 2018: "
                                . "Imaging of Hepatocellular Carcinoma in At-Risk Patients. "
                                . "Radiology. 2018;289(3):816-830.",
                           url:  "https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/LI-RADS" }],
        echo:           g_LastSelectedText
    })
}

_LR_Classify(aphe, size, wash, cap, grow, tiv, lrm, aM, aB, isNc, aHcc, bp
            , defBenign := false, distinctBenign := false) {
    if isNc
        return _LR_R("LR-NC", "", false)

    hasLrm := lrm.Length > 0

    ; LR-1 (Definitely benign): explicit definitely-benign entity, OR classic
    ; hemangioma-pattern blood-pool enhancement in a mass with no concerning
    ; features. Source: LI-RADS v2018 Core lines 1262-1273.
    if (defBenign && !hasLrm && aM = 0 && aHcc = 0 && !tiv)
        return _LR_R("LR-1", "Definitely benign entity", false)
    if (bp && !aphe && !wash && !cap && !grow && !tiv && !hasLrm && aM = 0 && aHcc = 0)
        return _LR_R("LR-1", "Hemangioma-type blood-pool enhancement, no concerning features", false)

    if tiv
        return _LR_R("LR-TIV", "", false)

    ; LR-2 (Probably benign): distinctive benign-appearing solid nodule <20 mm
    ; with no major HCC feature, no LR-M feature, and no ancillary malignancy.
    ; Source: LI-RADS v2018 Core lines 1275-1290.
    if (distinctBenign && size > 0 && size < 20 && !aphe && !wash && !cap && !grow
        && !tiv && !hasLrm && aM = 0 && aHcc = 0)
        return _LR_R("LR-2", "Distinctive benign-appearing solid nodule <20 mm without HCC or LR-M features", false)

    addl := (wash ? 1 : 0) + (cap ? 1 : 0) + (grow ? 1 : 0)
    tableCat := ""
    if !aphe {
        if (size < 20)
            tableCat := addl < 2 ? "LR-3" : "LR-4"
        else
            tableCat := addl = 0 ? "LR-3" : "LR-4"
    } else {
        if (size < 10) {
            tableCat := addl = 0 ? "LR-3" : "LR-4"
        } else if (size < 20) {
            if (addl = 0)
                tableCat := "LR-3"
            else if (addl = 1)
                tableCat := (wash || grow) ? "LR-5" : "LR-4"
            else
                tableCat := "LR-5"
        } else {
            tableCat := addl = 0 ? "LR-4" : "LR-5"
        }
    }

    if hasLrm {
        reason := "LR-M features: " JoinArr(lrm, ", ")
        if (tableCat = "LR-5") {
            return _LR_R("LR-M", reason ". Also meets LR-5 criteria. If confident this is HCC, may categorize as LR-5.", true)
        }
        return _LR_R("LR-M", reason, false)
    }

    final := tableCat
    adjusted := false
    reason := ""
    totalMal := aM + aHcc

    if (totalMal > 0 && aB = 0) {
        if (tableCat = "LR-3") {
            final := "LR-4"
            adjusted := true
            reason := "Upgraded from LR-3 due to ancillary features favoring malignancy"
        }
    }
    if (aB > 0 && totalMal = 0) {
        if (tableCat = "LR-4") {
            final := "LR-3"
            adjusted := true
            reason := "Downgraded from LR-4 due to ancillary features favoring benignity"
        } else if (tableCat = "LR-3") {
            final := "LR-2"
            adjusted := true
            reason := "Downgraded from LR-3 due to ancillary features favoring benignity"
        }
    }
    return _LR_R(final, reason, adjusted)
}

_LR_R(cat, reason, adjusted) {
    desc := _LR_Desc(cat)
    mgmt := _LR_Mgmt(cat)
    return { category: cat, description: desc, mgmt: mgmt
           , adjustmentReason: reason, adjusted: adjusted }
}

_LR_Desc(cat) {
    m := Map(
        "LR-NC",  "Cannot be categorized due to image degradation",
        "LR-1",   "Definitely benign",
        "LR-2",   "Probably benign",
        "LR-3",   "Intermediate probability of malignancy",
        "LR-4",   "Probably hepatocellular carcinoma",
        "LR-5",   "Definitely hepatocellular carcinoma",
        "LR-M",   "Probably or definitely malignant, not HCC specific",
        "LR-TIV", "Tumor in vein")
    return m.Has(cat) ? m[cat] : cat
}

_LR_Mgmt(cat) {
    m := Map(
        "LR-NC",  "Repeat or alternative diagnostic imaging.",
        "LR-1",   "No additional workup needed. Return to routine surveillance.",
        "LR-2",   "No additional workup needed. Return to routine surveillance.",
        "LR-3",   "Variable management per institutional preference. Consider follow-up in 3-6 months or multidisciplinary discussion.",
        "LR-4",   "Close follow-up or multidisciplinary discussion recommended. Consider biopsy if treatment decision depends on diagnosis.",
        "LR-5",   "Multidisciplinary discussion recommended. May qualify for OPTN 5A/5B if size criteria met.",
        "LR-M",   "Multidisciplinary discussion recommended. Biopsy may be considered for tissue diagnosis. Treatment depends on etiology.",
        "LR-TIV", "Multidisciplinary discussion recommended. Treatment depends on etiology of tumor in vein.")
    return m.Has(cat) ? m[cat] : ""
}
