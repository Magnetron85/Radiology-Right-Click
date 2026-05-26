; ============================================================
; lib/calc/IncidentalAdrenal.ahk -- ACR Incidental Adrenal 2017
; ------------------------------------------------------------
; Reference: Mayo-Smith WW, Song JH, Boland GL, et al. Management
; of Incidental Adrenal Masses: A White Paper of the ACR
; Incidental Findings Committee. J Am Coll Radiol. 2017;14(8):
; 1038-1044.
;
; The form has more fields than other RADS calculators because
; the algorithm branches deeply on known malignancy, washout
; values, MRI chemical shift, prior stability, and morphology.
; If unenhanced + enhanced + delayed HU are all entered, APW and
; RPW are auto-computed (matching the v1 adrenal washout
; calculator).
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

IncidentalAdrenal_Entry(input) {
    ShowIncidentalAdrenalDialog(input)
    return ""
}

ShowIncidentalAdrenalDialog(text := "") {
    sz := TextScan.Size(text)
    unenh := TextScan.HU(text, "unenhanced")
    enh   := TextScan.HU(text, "enhanced")
    delay := TextScan.HU(text, "delayed")
    bilat := TextScan.ContainsAny(text, ["\bbilateral\s+adrenal","\bbilaterally\b"])
    calcified := TextScan.ContainsAny(text, ["\bcalcified\s+adrenal","\bcoarse\s+calcif"])
    fat := TextScan.ContainsAny(text, ["\bmyelolipoma\b","\bmacroscopic\s+fat","\bfat[- ]containing"])
    susp := TextScan.ContainsAny(text, ["\birregular\s+(?:margin|contour)","\bnecrosis\b","\binvasion\b","\bheterogeneous\s+enhancement"])
    knownMalig := TextScan.ContainsAny(text
        , ["\bknown\s+(?:malignancy|cancer)","\bhistory\s+of\s+(?:lung|breast|melanoma|renal|kidney|colon|colorectal)\s+cancer"
        ,  "\bmetasta(?:tic|ses)\b","\bs/p\s+(?:lung|breast|melanoma)"])

    form := RadsForm("Incidental Adrenal Mass (ACR 2017)", 620)

    form.Header("Patient and modality")
    form.Dropdown("Modality", "Modality:", ["CT", "MRI", "PET"], 1)
    form.Numeric("SizeMm", "Mass size (mm):", sz.mm > 0 ? Round(sz.mm) : 0)
    form.Checkbox("KnownMalig", "Patient has known extra-adrenal malignancy", knownMalig)
    form.Numeric("MaligType", "If yes, primary type (e.g. lung, breast, melanoma):", "")

    form.Header("CT attenuation (HU; leave blank if not measured)")
    form.Numeric("UnenhHU",  "Unenhanced HU:", unenh != "" ? Round(unenh) : "")
    form.Numeric("EnhHU",    "Portal-venous (~60-70 s) HU:", enh != "" ? Round(enh) : "")
    form.Numeric("DelayHU",  "15-minute delayed HU:", delay != "" ? Round(delay) : "")
    form.Note("APW = (E-D)/(E-U)x100, RPW = (E-D)/Ex100. >=60% APW or >=40% RPW = adenoma.")

    form.Header("MRI / morphology")
    form.Dropdown("ChemShift", "Chemical shift (MRI):"
        , ["Unknown / not done"
        ,  "Yes -- signal loss out-of-phase (adenoma)"
        ,  "No -- no signal loss"], 1)
    form.CheckboxRow2("Fat", "Macroscopic fat (myelolipoma)"
                   , "Calcified", "Benign calcified mass", fat, calcified)
    form.CheckboxRow2("Suspicious", "Irregular / necrosis / invasion"
                   , "Bilateral", "Bilateral adrenal masses", susp, bilat)

    form.Header("Prior comparison")
    form.Dropdown("Stable", "Stable vs prior:"
        , ["Unknown / no prior"
        ,  "Stable for >=1 year"
        ,  "Growing or new"], 1)
    form.Numeric("StableMonths", "Months of documented stability (0 if unknown):", 0)

    form.Header("Other")
    form.Checkbox("Functional", "Functional symptoms (HTN, hypokalemia, pheo sx, Cushingoid, virilization)", false)
    form.Dropdown("Isolated", "If known malignancy: extent of disease:"
        , ["Unknown"
        ,  "Isolated adrenal mass (no other mets)"
        ,  "Widespread metastatic disease"], 1)

    form.OnChange("Modality", _Adr_UpdateModality)
    form.OnChange("KnownMalig", _Adr_UpdateKnownMalig)
    form.OnChange("Stable", _Adr_UpdateStable)
    _Adr_UpdateModality(form)
    _Adr_UpdateKnownMalig(form)
    _Adr_UpdateStable(form)

    form.SetSubmit(IncidentalAdrenal_OnSubmit)
    form.AddButtons()
    form.Show()
}

_Adr_UpdateModality(frm) {
    mod := frm.byName["Modality"].ctl.Text
    isCT := mod = "CT"
    isMRI := mod = "MRI"
    frm.SetEnabled("UnenhHU", isCT)
    frm.SetEnabled("EnhHU",   isCT)
    frm.SetEnabled("DelayHU", isCT)
    frm.SetEnabled("ChemShift", isMRI)
}
_Adr_UpdateKnownMalig(frm) {
    km := !!frm.GetValue("KnownMalig")
    frm.SetEnabled("MaligType", km)
    frm.SetEnabled("Isolated",  km)
}
_Adr_UpdateStable(frm) {
    txt := frm.byName["Stable"].ctl.Text
    frm.SetEnabled("StableMonths", InStr(txt, "Stable"))
}

IncidentalAdrenal_OnSubmit(v, form := "") {
    global g_LastSelectedText
    modality := v.Modality
    size := SafeInt(v.SizeMm, 0)
    knownMalig := !!v.KnownMalig
    maligType := v.MaligType
    unenh := v.UnenhHU = "" ? "" : v.UnenhHU + 0.0
    enh   := v.EnhHU   = "" ? "" : v.EnhHU + 0.0
    delay := v.DelayHU = "" ? "" : v.DelayHU + 0.0
    chem  := InStr(v.ChemShift, "Yes") ? 1
          : InStr(v.ChemShift, "No ") ? 0 : -1
    fat := !!v.Fat
    calc := !!v.Calcified
    susp := !!v.Suspicious
    bilat := !!v.Bilateral
    isStable := InStr(v.Stable, "Stable") ? 1
             : InStr(v.Stable, "Growing") ? 0 : -1
    months := v.StableMonths + 0.0
    funcSx := !!v.Functional
    isolated := InStr(v.Isolated, "Isolated") ? 1
             : InStr(v.Isolated, "Widespread") ? 0 : -1

    r := _Adr_Classify(size, modality, unenh, enh, delay, chem
                     , knownMalig, maligType, isStable, months
                     , susp, funcSx, bilat, fat, calc, isolated)

    sizeMm := SafeInt(v.SizeMm, 0)

    ; --- Impression sentence ---
    ; Subcentimeter findings get a special impression -- per ACR 2017
    ; (Mayo-Smith p. 1041), the paper explicitly says these may not even
    ; qualify as a discrete mass, so the generic "incidental adrenal mass,
    ; category SUBCM" template overclaims. Use clinical prose and the
    ; paper's hedged wording ("in general").
    if (r.cat = "SUBCM") {
        impression := sizeMm " mm subcentimeter adrenal finding; per ACR Incidental "
                    . "Findings 2017, subcentimeter adrenal nodularity or thickening "
                    . "(<10 mm short axis) may not qualify as a discrete mass and in "
                    . "general does not require further evaluation."
    } else {
        ; Lowercase the leading char of the verbose r.rec so it can be glued in
        ; mid-sentence; keep the original (sentence-cased) version inside
        ; methodology for full audit detail.
        recLc := (r.rec != "") ? StrLower(SubStr(r.rec, 1, 1)) SubStr(r.rec, 2) : ""
        impression := sizeMm " mm incidental adrenal mass, category " r.cat
                   . " (" _Adr_CatDesc(r.cat) ") per ACR Incidental Findings 2017"
        if (recLc != "")
            impression .= "; " recLc
        if (SubStr(impression, -1) != ".")
            impression .= "."
    }

    ; --- Methodology ---
    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    if (r.autoWashout != "")
        method .= "`n" r.autoWashout
    method .= "`nCategory: " r.cat " -- " _Adr_CatDesc(r.cat)
    method .= "`nReason: " r.reason
    method .= "`nFull recommendation: " r.rec
    if (r.workup.Length > 0)
        method .= "`nSuggested workup:`n - " JoinArr(r.workup, "`n - ")
    if (r.functionalNote != "")
        method .= "`n" r.functionalNote
    if (r.pheoNote != "")
        method .= "`n" r.pheoNote

    advisories := []
    if (r.functionalNote != "")
        advisories.Push(r.functionalNote)
    if (r.pheoNote != "")
        advisories.Push(r.pheoNote)

    return MakeResult({
        classification: "ACR Adrenal Cat " r.cat,
        impression:     impression,
        recommendation: "",
        advisories:     advisories,
        methodology:    method,
        citations:      [{ text: "Mayo-Smith WW, Song JH, Boland GL, et al. "
                                . "Management of Incidental Adrenal Masses: A White Paper of "
                                . "the ACR Incidental Findings Committee. "
                                . "J Am Coll Radiol. 2017;14(8):1038-1044.",
                           url:  "https://www.acr.org/Clinical-Resources/Incidental-Findings" }],
        echo:           g_LastSelectedText
    })
}

_Adr_Classify(size, modality, unenh, enh, delay, chem
            , knownMalig, maligType, isStable, months
            , susp, funcSx, bilat, fat, calc, isolated) {
    apw := 0, rpw := 0
    hasWashout := false
    autoWashout := ""
    if (enh != "" && delay != "") {
        if (unenh != "" && enh != unenh) {
            apw := Round((enh - delay) / (enh - unenh) * 100, 1)
        }
        if (enh != 0)
            rpw := Round((enh - delay) / enh * 100, 1)
        if (apw != 0 || rpw != 0) {
            hasWashout := true
            interp := (apw >= 60) ? "Adenoma (APW >=60%)"
                    : (rpw >= 40) ? "Adenoma (RPW >=40%)"
                    : "Indeterminate (APW <60%, RPW <40%)"
            autoWashout := "Auto-computed washout: APW " apw "%, RPW " rpw "% (" interp ")"
        }
    }

    functionalNote := ""
    if funcSx
        functionalNote := "Functional workup recommended: consider plasma metanephrines, 1 mg overnight dexamethasone suppression test, and plasma aldosterone/renin ratio if hypertensive."

    pheoNote := ""
    ; ACR 2017 (Mayo-Smith) line 234-240: "avid enhancement (>110-120 HU)... pheochromocytoma
    ; should be considered". Use the lower bound to avoid missing pheo at 111-120 HU --
    ; biopsy without alpha-blockade can precipitate hypertensive crisis.
    if (enh != "" && enh >= 110)
        pheoNote := "Avid enhancement (" enh " HU, in 110-120 range or higher) raises concern for pheochromocytoma. Consider plasma metanephrines / catecholamines BEFORE any biopsy."

    base := { cat: "", reason: "", rec: "", workup: []
            , functionalNote: functionalNote, pheoNote: pheoNote
            , autoWashout: autoWashout }

    ; ACR Incidental Findings 2017 (Mayo-Smith p. 1041, "Five Common Principles"
    ; item #1): "In general, an incidental adrenal mass that is <1 cm in the
    ; short axis need not be pursued. We provide such guidance to address
    ; circumstances in which radiologists identify subcentimeter 'nodularity'
    ; or adrenal 'thickening' and are uncertain whether such findings should
    ; qualify as adrenal masses."
    ;
    ; This short-circuit precedes BOTH the known-malig and no-malig branches.
    ; A subcentimeter finding doesn't qualify as a "mass" for algorithm
    ; purposes regardless of cancer history -- size is the dominant gate.
    if (size > 0 && size < 10) {
        base.cat := "SUBCM"
        base.reason := "Subcentimeter (<10 mm short axis) -- per ACR 2017, may not qualify as a discrete mass"
        ; Hedged wording matches the paper. Mayo-Smith 2017 p. 1041:
        ; "In general, an incidental adrenal mass that is <1 cm in the short
        ; axis need not be pursued. We provide such guidance to address
        ; circumstances in which radiologists identify subcentimeter
        ; 'nodularity' or adrenal 'thickening' and are uncertain whether
        ; such findings should qualify as adrenal masses."
        base.rec := "Subcentimeter adrenal nodularity or thickening (<10 mm short axis); "
                  . "per ACR Incidental Findings 2017, such findings may not qualify as "
                  . "a discrete adrenal mass and in general do not require further pursuit. "
                  . "Clinical correlation as appropriate."
        return base
    }

    if knownMalig
        return _Adr_KnownMalig(base, size, modality, unenh, hasWashout, apw, rpw
                             , chem, maligType, isStable, months, susp, bilat
                             , fat, calc, enh, delay, isolated, funcSx)

    return _Adr_NoMalig(base, size, modality, unenh, hasWashout, apw, rpw
                       , chem, isStable, months, susp, bilat, fat, calc, enh, funcSx)
}

_Adr_KnownMalig(r, size, modality, unenh, hasWashout, apw, rpw, chem
              , maligType, isStable, months, susp, bilat, fat, calc
              , enh, delay, isolated, funcSx) {
    highRisk := false
    low := StrLower(maligType)
    for primary in ["lung","breast","melanoma","renal","kidney","colon","colorectal"] {
        if InStr(low, primary) {
            highRisk := true
            break
        }
    }

    if fat {
        r.cat := "BENIGN"
        r.reason := "Diagnostic of myelolipoma (macroscopic fat) despite known malignancy"
        r.rec := "No follow-up required. Diagnostic of myelolipoma."
        return r
    }
    if (enh != "" && unenh != "" && (enh - unenh) < 10) {
        r.cat := "BENIGN"
        r.reason := "Non-enhancing mass (enhancement change " Round(enh - unenh) " HU, <10 HU) despite known malignancy"
        r.rec := "Non-enhancing mass consistent with cyst or hemorrhage. No follow-up required."
        return r
    }
    if calc {
        r.cat := "BENIGN"
        r.reason := "Benign calcified mass despite known malignancy"
        r.rec := "Benign calcified mass. No follow-up required."
        return r
    }
    if (unenh != "" && unenh <= 10) {
        r.cat := "BENIGN"
        r.reason := "Lipid-rich adenoma (<=10 HU) despite known malignancy"
        r.rec := "Benign adenoma. Metastasis excluded by lipid content."
        return r
    }
    if (hasWashout && (apw >= 60 || rpw >= 40)) {
        r.cat := "BENIGN"
        wt := (apw >= 60) ? ("APW " Round(apw) "%") : ("RPW " Round(rpw) "%")
        r.reason := "Adenoma confirmed by washout (" wt ") despite known malignancy"
        r.rec := "Benign adenoma. Metastasis excluded by washout characteristics."
        return r
    }
    if (chem = 1) {
        r.cat := "BENIGN"
        r.reason := "Adenoma (chemical shift positive) despite known malignancy"
        r.rec := "Benign adenoma. Metastasis excluded by intracellular lipid."
        return r
    }
    if (isStable = 1 && months >= 12) {
        r.cat := "BENIGN"
        r.reason := "Stable for " months " months (>=12 months) despite known " maligType
        r.rec := "Benign based on stability >=1 year. No further imaging follow-up needed."
        return r
    }
    if (isolated = 0) {
        r.cat := "KNOWN_MALIG"
        r.reason := "Adrenal mass in patient with widespread metastatic " maligType
        r.rec := "Widespread metastatic disease. Further evaluation of adrenal mass unlikely to alter management."
        return r
    }
    if (size >= 40) {
        r.cat := "KNOWN_MALIG"
        r.reason := "Large adrenal mass (>=" Round(size/10, 1) " cm) in patient with " maligType
        r.rec := "Consider biopsy or PET-CT to distinguish metastasis from primary adrenal neoplasm."
        r.workup.Push("Consider PET/CT for metabolic characterization")
        r.workup.Push("Consider image-guided biopsy if would change management")
        if !hasWashout
            r.workup.Push("CT adrenal protocol with washout if not done")
        r.workup.Push("Biochemical workup to exclude functional tumor (especially pheochromocytoma before biopsy)")
        return r
    }
    if (isStable = 0) {
        r.cat := "KNOWN_MALIG"
        r.reason := "New or enlarging adrenal mass in patient with " maligType
        r.rec := "Consider biopsy or PET-CT to evaluate for metastasis."
        r.workup.Push("PET/CT for metabolic characterization")
        r.workup.Push("Consider image-guided biopsy if would change management")
        r.workup.Push("Biochemical workup to exclude functional tumor")
        return r
    }
    if bilat {
        r.cat := "KNOWN_MALIG"
        r.reason := "Bilateral adrenal masses in patient with " maligType
        r.rec := "Bilateral masses raise concern for metastatic disease. Further characterization needed."
        r.workup.Push("PET/CT recommended")
        r.workup.Push("Consider biopsy if would change management")
        return r
    }
    if (isolated = 1) {
        r.cat := "KNOWN_MALIG"
        r.reason := "Isolated adrenal mass in patient with " maligType " (only potential metastatic site)"
        r.rec := "Isolated adrenal mass -- characterization critical for staging. Dedicated adrenal CT protocol recommended."
        r.workup.Push("CT adrenal protocol with washout")
        r.workup.Push("Consider PET/CT")
        r.workup.Push("Consider image-guided biopsy if would change management")
        return r
    }
    if highRisk {
        r.cat := "KNOWN_MALIG"
        r.reason := "Indeterminate adrenal mass in patient with " maligType " (high metastatic risk)"
        r.rec := "Cannot exclude metastasis. Further workup recommended."
        r.workup.Push("PET/CT for metabolic characterization")
        if (modality = "CT" && !hasWashout)
            r.workup.Push("CT adrenal protocol with washout")
        r.workup.Push("Consider image-guided biopsy if would change management")
        return r
    }
    r.cat := "INDETERMINATE"
    r.reason := "Indeterminate adrenal mass in patient with known malignancy (" maligType ")"
    r.rec := "Characterization recommended to exclude metastasis."
    if (modality = "CT" && unenh = "")
        r.workup.Push("Unenhanced CT to assess attenuation")
    if (modality = "CT" && !hasWashout)
        r.workup.Push("CT adrenal protocol with washout")
    r.workup.Push("Alternatively: MRI with chemical shift imaging")
    r.workup.Push("Consider PET/CT")
    return r
}

_Adr_NoMalig(r, size, modality, unenh, hasWashout, apw, rpw, chem
            , isStable, months, susp, bilat, fat, calc, enh, funcSx) {
    if fat {
        r.cat := "BENIGN"
        r.reason := "Diagnostic of myelolipoma (macroscopic fat identified)"
        r.rec := "No follow-up required. Diagnostic of myelolipoma."
        if (size > 10 && !funcSx)
            r.workup.Push("Consider biochemical screening for masses >1 cm")
        return r
    }
    if (enh != "" && unenh != "" && (enh - unenh) < 10) {
        r.cat := "BENIGN"
        r.reason := "Non-enhancing mass (enhancement change " Round(enh - unenh) " HU, <10 HU) consistent with cyst or hemorrhage"
        r.rec := "Non-enhancing mass consistent with cyst or hemorrhage. No follow-up required."
        if (size > 10 && !funcSx)
            r.workup.Push("Consider biochemical screening for masses >1 cm")
        return r
    }
    if calc {
        r.cat := "BENIGN"
        r.reason := "Benign calcified mass (old hematoma or granulomatous infection)"
        r.rec := "Benign calcified mass. No follow-up required."
        if (size > 10 && !funcSx)
            r.workup.Push("Consider biochemical screening for masses >1 cm")
        return r
    }
    if (unenh != "" && unenh <= 10) {
        r.cat := "BENIGN"
        r.reason := "Lipid-rich adenoma (unenhanced CT " unenh " HU <=10 HU)"
        r.rec := size < 40
               ? "No imaging follow-up required. Benign lipid-rich adenoma."
               : "Likely benign lipid-rich adenoma. Consider surgical consultation for size >4 cm despite benign imaging features."
        if (size >= 40)
            r.workup.Push("Consider surgical consultation due to size")
        if (size > 10 && !funcSx)
            r.workup.Push("Consider biochemical screening for masses >1 cm")
        return r
    }
    if hasWashout {
        if (apw >= 60) {
            r.cat := "BENIGN"
            r.reason := "Adenoma confirmed by washout (APW " Round(apw) "% >=60%)"
            r.rec := "No imaging follow-up required. Adenoma confirmed by absolute washout."
            if (size >= 40)
                r.workup.Push("Consider surgical consultation due to size >4 cm")
            if (size > 10 && !funcSx)
                r.workup.Push("Consider biochemical screening for masses >1 cm")
            return r
        }
        if (rpw >= 40) {
            r.cat := "BENIGN"
            r.reason := "Adenoma confirmed by washout (RPW " Round(rpw) "% >=40%)"
            r.rec := "No imaging follow-up required. Adenoma confirmed by relative washout."
            if (size >= 40)
                r.workup.Push("Consider surgical consultation due to size >4 cm")
            if (size > 10 && !funcSx)
                r.workup.Push("Consider biochemical screening for masses >1 cm")
            return r
        }
        r.cat := "INDETERMINATE"
        r.reason := "Negative washout (APW " Round(apw) "%, RPW " Round(rpw) "%)"
        r.rec := "Washout criteria not met. Further characterization or close follow-up recommended."
        r.workup.Push("Consider MRI with chemical shift imaging")
        r.workup.Push("Consider PET/CT if high clinical suspicion")
        if (size >= 40)
            r.workup.Push("Consider surgical consultation / biopsy for size >4 cm")
        return r
    }
    if (chem = 1) {
        r.cat := "BENIGN"
        r.reason := "Lipid-containing adenoma (signal loss on chemical shift MRI)"
        r.rec := "No imaging follow-up required. Adenoma confirmed by chemical shift imaging."
        if (size >= 40)
            r.workup.Push("Consider surgical consultation due to size >4 cm")
        if (size > 10 && !funcSx)
            r.workup.Push("Consider biochemical screening for masses >1 cm")
        return r
    }
    if (isStable = 1 && months >= 12) {
        r.cat := "BENIGN"
        r.reason := "Stable for " months " months (>=12 months)"
        r.rec := size < 40
               ? "Benign based on stability >=1 year. No further imaging follow-up needed."
               : "Stable >=1 year but large (>4 cm). Consider additional characterization or surgical consultation."
        if (size >= 40)
            r.workup.Push("Consider CT washout or MRI chemical shift")
        if (size > 10 && !funcSx)
            r.workup.Push("Consider biochemical screening for masses >1 cm")
        return r
    }
    if susp {
        r.cat := "SUSPICIOUS"
        r.reason := "Suspicious imaging features (irregular margins, necrosis, or invasion)"
        r.rec := "Concerning features. Surgical consultation recommended."
        r.workup.Push("Surgical oncology consultation")
        r.workup.Push("Consider PET/CT for staging")
        r.workup.Push("Biochemical workup to exclude functional tumor")
        return r
    }
    if (isStable = 0) {
        if (size >= 40) {
            r.cat := "SUSPICIOUS"
            r.reason := "Growing mass >4 cm"
            r.rec := "Growing large adrenal mass. Surgical consultation recommended."
            r.workup.Push("Surgical consultation")
        } else {
            r.cat := "INDETERMINATE"
            r.reason := "Interval growth documented (no known malignancy)"
            r.rec := "New or enlarging adrenal mass. Consider follow-up adrenal CT or surgical consultation."
            r.workup.Push("CT adrenal protocol with washout")
            r.workup.Push("Consider MRI with chemical shift")
            r.workup.Push("Consider surgical consultation if continued growth")
        }
        r.workup.Push("Biochemical workup to exclude functional tumor")
        return r
    }
    if (size <= 10) {
        r.cat := "LIKELY_BENIGN"
        r.reason := "Small size (<=1 cm)"
        r.rec := "Small adrenal nodule. No dedicated follow-up typically needed if no clinical concern."
        return r
    }
    if (size <= 20) {
        r.cat := "LIKELY_BENIGN"
        r.reason := "Small indeterminate mass (" Round(size/10, 1) " cm, 1-2 cm) without prior imaging"
        r.rec := "Probably benign. Consider 12-month follow-up adrenal CT."
        if (modality = "CT" && unenh = "")
            r.workup.Push("Unenhanced CT to assess attenuation if not done")
        r.workup.Push("Consider 12-month follow-up adrenal CT")
        if !funcSx
            r.workup.Push("Biochemical screening recommended for masses >1 cm")
        return r
    }
    if (size < 40) {
        r.cat := "INDETERMINATE"
        r.reason := "Indeterminate " Round(size/10, 1) " cm mass (>2-4 cm) without definitive benign features"
        r.rec := "Dedicated adrenal CT protocol recommended."
        if (modality = "CT" && unenh = "")
            r.workup.Push("Unenhanced CT to assess attenuation")
        if (modality = "CT" && (unenh = "" || unenh > 10))
            r.workup.Push("CT adrenal protocol with washout calculation")
        r.workup.Push("Alternatively: MRI with chemical shift imaging")
        if !funcSx
            r.workup.Push("Biochemical screening recommended for masses >1 cm")
        return r
    }
    r.cat := "SUSPICIOUS"
    r.reason := "Large size (" Round(size/10, 1) " cm, >4 cm) -- consider adrenocortical carcinoma"
    r.rec := "Large adrenal mass. Surgical consultation recommended."
    r.workup.Push("Surgical consultation")
    r.workup.Push("CT adrenal protocol with washout if not done")
    r.workup.Push("Biochemical workup mandatory")
    r.workup.Push("Consider PET/CT for staging")
    return r
}

_Adr_CatDesc(cat) {
    m := Map(
        "SUBCM",         "Subcentimeter -- below algorithm threshold",
        "BENIGN",        "Benign -- no follow-up",
        "LIKELY_BENIGN", "Likely benign",
        "INDETERMINATE", "Indeterminate -- further workup",
        "SUSPICIOUS",    "Suspicious -- surgical consult",
        "FUNCTIONAL",    "Functional workup indicated",
        "KNOWN_MALIG",   "Known malignancy -- staging workup")
    return m.Has(cat) ? m[cat] : cat
}
