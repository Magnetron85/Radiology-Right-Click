; ============================================================
; lib/calc/ORADSMRI.ahk -- O-RADS MRI risk stratification
; ------------------------------------------------------------
; Reference: ACR O-RADS MRI Risk Stratification System for
; adnexal lesion characterization on MRI (problem-solving after
; ultrasound).
;
; Lesion-type terminology follows the O-RADS MRI lexicon exactly:
;   - No lesion / physiologic finding (premenopausal follicle,
;     hemorrhagic cyst <=3 cm, corpus luteum)
;   - Unilocular cyst
;   - Multilocular cyst (no solid tissue; lipid content excluded)
;   - Lesion with lipid content (dermoid / teratoma; this IS a
;     primary type in O-RADS, not a modifier on another type)
;   - Lesion with solid tissue (includes any morphology with
;     solid component, whether predominantly solid OR mixed
;     cystic-solid; "mixed cystic-solid" is NOT a separate
;     O-RADS descriptor)
;
; Not yet modeled (future work): "Dilated fallopian tube" and
; "Para-ovarian cyst" as primary lesion types.
;
; Algorithm summary:
;   no lesion                                     -> 1
;   peritoneal / omental deposits                 -> 5
;   unilocular, no wall enhancement, no solid     -> 2
;   lipid lesion, no large-volume enhancing solid -> 2
;   lipid lesion + large volume enhancing solid   -> 4
;   solid tissue dark on T2 AND dark on DWI       -> 2
;   multilocular smooth septae + wall enh, no solid -> 3
;   enhancing solid + TIC type 1                  -> 3
;   enhancing solid + TIC type 2                  -> 4
;   enhancing solid + TIC type 3                  -> 5
;   enhancing solid, no TIC                       -> 4 (default)
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

ORADSMRI_Entry(input) {
    ShowORADSMRIDialog(input)
    return ""
}

ShowORADSMRIDialog(text := "") {
    ; Lipid-content auto-detect routes to the new "Lesion with lipid
    ; content" primary type (option 4) per O-RADS lexicon.
    if TextScan.ContainsAny(text, ["\bno\s+(?:adnexal|ovarian)\s+lesion\b","\bnormal\s+ovari"])
        typeIdx := 1
    else if TextScan.ContainsAny(text, ["\bunilocular\b","\bsimple\s+cyst\b"])
        typeIdx := 2
    else if TextScan.ContainsAny(text, ["\bmultilocular\b"])
        typeIdx := 3
    else if TextScan.ContainsAny(text, ["\bdermoid\b","\bteratoma\b","\bfat[- ]containing\b","\bmacroscopic\s+fat\b","\blipid\b"])
        typeIdx := 4
    else
        typeIdx := 5   ; default: "Lesion with solid tissue"
    ; "mixed cystic-solid" colloquial radiology terms map to the O-RADS
    ; "Lesion with solid tissue" category -- O-RADS does not have a
    ; separate "mixed" descriptor.

    solidEnh := TextScan.ContainsAny(text, ["\benhancing\s+solid\b","\bsolid\s+enhancement","\bavidly\s+enhancing","\benhancing\s+component"])
    wallEnh := TextScan.ContainsAny(text, ["\bwall\s+enhanc","\benhancing\s+wall\b"])
    smoothSept := TextScan.ContainsAny(text, ["\bsmooth\s+(?:thin\s+)?septations?","\bthin\s+septations?\b"])
    perit := TextScan.ContainsAny(text, ["\bperitoneal\s+(?:deposit|nodul|implant)","\bomental\s+caking?","\bcarcinomatos"])
    asc   := TextScan.ContainsAny(text, ["\bascites\b","\bfree\s+pelvic\s+fluid\b"])

    form := RadsForm("O-RADS MRI", 560)

    ; Form is organized as a decision tree mirroring how a radiologist
    ; walks through an adnexal MRI: pick a lesion type, then characterize
    ; the features specific to that type. Progressive disclosure: each
    ; type-specific section is HIDDEN (form reflows + resizes) until the
    ; matching lesion type is selected, so the dialog opens minimal.

    form.Header("Lesion type")
    form.Dropdown("LType", "Lesion type:"
        , ["No lesion (or physiologic finding)"
        ,  "Unilocular cyst"
        ,  "Multilocular cyst (no solid tissue)"
        ,  "Lesion with lipid content (dermoid / teratoma)"
        ,  "Lesion with solid tissue"], typeIdx)

    form.Header("Cystic features (unilocular / multilocular cysts)")
    form.Dropdown("FluidType", "Fluid type (unilocular only):"
        , ["Not specified"
        ,  "Simple or endometriotic fluid"
        ,  "Proteinaceous / hemorrhagic / mucinous fluid"], 1)
    form.Checkbox("WallEnh", "Wall / septae enhance (smooth pattern)", wallEnh)
    form.Checkbox("SmoothSeptae", "Smooth thin septations (multilocular only)", smoothSept)

    form.Header("Lipid lesion features (dermoid / teratoma)")
    form.Checkbox("LargeVolSolid", "Large volume enhancing solid tissue (drives Score 4 per source)", false)

    form.Header("Solid tissue characterization (Lesion with solid tissue)")
    form.Checkbox("SolidEnh", "Solid tissue enhances (uncheck if non-enhancing / fibrotic / dark T2-DWI)", solidEnh)
    form.Dropdown("T2",  "Solid tissue T2 signal (dark = Score 2 if DWI also low):"
        , ["Not assessed","Hypo","Intermediate","Hyper"], 1)
    form.Dropdown("DWI", "Solid tissue DWI signal (low = Score 2 if T2 also hypo):"
        , ["Not assessed","Low","Intermediate","High"], 1)
    form.Dropdown("TIC", "Time-intensity curve (DCE; required when solid enhances):"
        , ["Not assessed"
        ,  "Type 1 -- slow progressive rise (slope < myometrium)"
        ,  "Type 2 -- moderate rise + plateau (slope <= myometrium)"
        ,  "Type 3 -- brisk rise (slope > myometrium)"], 1)

    form.Header("Extra-ovarian findings (peritoneal nodularity overrides to Score 5)")
    form.Checkbox("Perit",   "Peritoneal / mesenteric / omental nodularity or thickening", perit)
    form.Checkbox("Ascites", "Ascites", asc)

    form.OnChange("LType", _OM_UpdateLType)
    form.OnChange("SolidEnh", _OM_UpdateSolid)
    _OM_UpdateLType(form)
    _OM_UpdateSolid(form)

    form.SetSubmit(ORADSMRI_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_OM_UpdateLType(frm) {
    label := frm.byName["LType"].ctl.Text
    isUni   := InStr(label, "Unilocular")
    isMulti := InStr(label, "Multilocular")
    isLipid := InStr(label, "lipid content")
    isSolid := InStr(label, "Lesion with solid")

    ; Progressive disclosure: whole type-specific sections are HIDDEN
    ; (the form reflows and resizes), not greyed. A hidden control
    ; reverts to its default so the classifier never reads a stale
    ; answer from a question the user can no longer see.
    frm.SetSectionVisible("Cystic features (unilocular / multilocular cysts)", isUni || isMulti)
    frm.SetVisible("FluidType",    isUni)
    frm.SetVisible("SmoothSeptae", isMulti)
    ; WallEnh is part of the source rules ONLY for unilocular and
    ; multilocular cysts. Lipid and solid lesion scoring don't use
    ; wall enhancement per the O-RADS grid.
    frm.SetVisible("WallEnh", isUni || isMulti)
    ; LargeVolSolid is the lipid-specific equivalent of TIC: it drives
    ; the lipid-lesion Score 2 vs Score 4 decision per source.
    frm.SetSectionVisible("Lipid lesion features (dermoid / teratoma)", isLipid)
    ; SolidEnh / T2 / DWI / TIC apply only to "Lesion with solid
    ; tissue". Unilocular and multilocular cysts have no solid
    ; component by O-RADS definition; "Lesion with lipid content"
    ; uses its own LargeVolSolid question instead of TIC/T2/DWI per
    ; source. T2 / DWI remain visible (regardless of SolidEnh)
    ; for the solid type because "dark T2 + dark DWI" is the way to
    ; confirm Score 2 for a non-enhancing solid lesion.
    frm.SetSectionVisible("Solid tissue characterization (Lesion with solid tissue)", isSolid)
    _OM_UpdateSolid(frm)   ; cascade -- TIC depends on both lesion type and SolidEnh
}
_OM_UpdateSolid(frm) {
    label := frm.byName["LType"].ctl.Text
    isSolid := InStr(label, "Lesion with solid")
    hasSolidEnh := !!frm.GetValue("SolidEnh")
    frm.SetVisible("TIC", isSolid && hasSolidEnh)
}

ORADSMRI_OnSubmit(v, form := "") {
    global g_LastSelectedText
    lt := _OM_LType(v.LType)
    t2 := _OM_T2(v.T2)
    dwi := _OM_DWI(v.DWI)
    tic := _OM_TIC(v.TIC)
    wallEnh := !!v.WallEnh
    solidEnh := !!v.SolidEnh
    largeVolSolid := !!v.LargeVolSolid
    smooth  := !!v.SmoothSeptae
    perit := !!v.Perit
    asc := !!v.Ascites
    fluidType := _OM_Fluid(v.FluidType)

    score := _OM_Score(lt, wallEnh, solidEnh, t2, dwi, tic, smooth, perit, asc, largeVolSolid, fluidType)
    info := _OM_Cats()[score]

    showRisk := Prefs.Get("display", "showMalignancyRisk", true)
    descLc := (info.desc != "") ? StrLower(SubStr(info.desc, 1, 1)) SubStr(info.desc, 2) : ""
    mgmtLc := (info.mgmt != "") ? StrLower(SubStr(info.mgmt, 1, 1)) SubStr(info.mgmt, 2) : ""
    impression := "Adnexal lesion, O-RADS MRI " score " (" descLc ")"
    if (showRisk && info.risk != "")
        impression .= ", malignancy risk " info.risk
    impression .= " per O-RADS MRI v2020"
    if (mgmtLc != "")
        impression .= "; " mgmtLc
    if (SubStr(impression, -1) != ".")
        impression .= "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nScore: " score " -- " info.desc
    method .= "`nMalignancy risk: " info.risk
    method .= "`nFull management: " info.mgmt

    return MakeResult({
        classification: "O-RADS MRI " score,
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "Thomassin-Naggara I, Poncelet E, Jalaguier-Coudray A, et al. "
                                . "Ovarian-Adnexal Reporting Data System Magnetic Resonance Imaging "
                                . "(O-RADS MRI) Score for Risk Stratification of Sonographically "
                                . "Indeterminate Adnexal Masses. JAMA Netw Open. 2020;3(1):e1919896.",
                           url:  "https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/O-RADS" }],
        echo:           g_LastSelectedText
    })
}

_OM_LType(label) {
    if InStr(label, "No lesion")
        return "none"
    if InStr(label, "Unilocular")
        return "unilocular"
    if InStr(label, "Multilocular")
        return "multilocular"
    if InStr(label, "lipid content")
        return "lipid"
    ; "Lesion with solid tissue" -- the O-RADS umbrella category covering
    ; predominantly-solid AND mixed-cystic-solid morphologies.
    return "solid"
}
_OM_T2(label) {
    if (label = "Hypo")
        return "hypo"
    if (label = "Intermediate")
        return "intermediate"
    if (label = "Hyper")
        return "hyper"
    return ""
}
_OM_DWI(label) {
    if (label = "Low")
        return "low"
    if (label = "Intermediate")
        return "intermediate"
    if (label = "High")
        return "high"
    return ""
}
_OM_TIC(label) {
    if InStr(label, "Type 1")
        return 1
    if InStr(label, "Type 2")
        return 2
    if InStr(label, "Type 3")
        return 3
    return 0
}
_OM_Fluid(label) {
    if InStr(label, "Simple or endometriotic")
        return "simple"
    if InStr(label, "Proteinaceous")
        return "complex"
    return ""
}

_OM_Score(lt, wallEnh, solidEnh, t2, dwi, tic, smooth, perit, asc, largeVolSolid := false, fluidType := "") {
    ; Peritoneal / mesenteric / omental nodularity is a Score 5 finding
    ; independent of any visible ovarian lesion (per ACR O-RADS MRI v2024
    ; grid). Check this BEFORE the lt=none short-circuit so a patient with
    ; peritoneal carcinomatosis but no visible primary ovarian lesion is
    ; not mis-scored as Score 1 (normal).
    if perit
        return 5
    if (lt = "none")
        return 1
    ; Per O-RADS MRI v2024 grid (page 1 source-of-truth table):
    ;   Score 2 unilocular: any-fluid + NO wall enhancement + no solid, OR
    ;                       simple/endometriotic fluid + smooth wall enhancement + no solid
    ;   Score 3 unilocular: proteinaceous/hemorrhagic/mucinous fluid +
    ;                       smooth wall enhancement + no solid
    ; If the radiologist did not specify the fluid type, default the wall-
    ; enhanced case to Score 3 (the more conservative classification).
    if (lt = "unilocular" && !solidEnh) {
        if !wallEnh
            return 2
        if (fluidType = "simple")
            return 2
        ; wall enhancement with proteinaceous/hemorrhagic/mucinous OR fluid type
        ; unspecified: route to Score 3 per source row.
        return 3
    }
    ; Lesion with lipid content (O-RADS primary type). The source has only
    ; two rows for this category:
    ;   - "no enhancing solid tissue"          -> Score 2
    ;   - "large volume enhancing solid tissue" -> Score 4
    ; A small focus of enhancement (LargeVolSolid unchecked) falls back to
    ; the Score 2 row -- it does not meet the "large volume" criterion that
    ; defines the Score 4 row. TIC / T2 / DWI are NOT part of the lipid
    ; decision tree per source.
    if (lt = "lipid") {
        if largeVolSolid
            return 4
        return 2
    }
    if (t2 = "hypo" && dwi = "low")
        return 2
    if (lt = "multilocular" && smooth && wallEnh && !solidEnh)
        return 3
    if solidEnh {
        if (tic = 1)
            return 3
        if (tic = 2)
            return 4
        if (tic = 3)
            return 5
        return 4
    }
    return 3
}

_OM_Cats() {
    static m := _OM_BuildCats()
    return m
}
_OM_BuildCats() {
    m := Map()
    m[0] := { desc: "Incomplete",              risk: "N/A",    mgmt: "Repeat MRI with contrast." }
    ; Source (orads mri.txt:10-11): Score 1 PPV is listed as "N/A" -- the
    ; system does not assign a numeric malignancy risk to a normal ovary.
    m[1] := { desc: "Normal",                  risk: "N/A",    mgmt: "No follow-up." }
    m[2] := { desc: "Almost certainly benign", risk: "<0.5%",  mgmt: "No follow-up." }
    m[3] := { desc: "Low risk",                risk: "~5%",    mgmt: "Surveillance MRI in 6-12 months." }
    m[4] := { desc: "Intermediate risk",       risk: "~50%",   mgmt: "Gynecologic oncology consultation." }
    m[5] := { desc: "High risk",               risk: "~90%",   mgmt: "Gynecologic oncology referral." }
    return m
}
