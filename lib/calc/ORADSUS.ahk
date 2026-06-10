; ============================================================
; lib/calc/ORADSUS.ahk -- O-RADS Ultrasound risk stratification
; ------------------------------------------------------------
; Reference: ACR O-RADS US v2022. Adnexal-lesion risk
; stratification on transvaginal / transabdominal ultrasound.
;
; The algorithm is a ladder of priority rules; ascites and/or
; peritoneal nodularity upgrades any score >=3 to 5.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

ORADSUS_Entry(input) {
    ShowORADSUSDialog(input)
    return ""
}

ShowORADSUSDialog(text := "") {
    sz := TextScan.Size(text)
    menoIdx := TextScan.ContainsAny(text, ["\bpostmenopausal\b","\bpost[- ]menopausal\b"]) ? 2 : 1

    lTypeIdx := 2
    if TextScan.ContainsAny(text, ["\bno\s+(?:adnexal|ovarian)\s+lesion","\bnormal\s+ovari"])
        lTypeIdx := 1
    else if TextScan.ContainsAny(text, ["\bbilocular\b"])
        lTypeIdx := 4
    else if TextScan.ContainsAny(text, ["\bmultilocular\b"])
        lTypeIdx := 5
    else if TextScan.ContainsAny(text, ["\bmixed\s+(?:cystic|solid)\b","\bcystic[- ]solid\b"])
        lTypeIdx := 7
    else if TextScan.ContainsAny(text, ["\bsolid\s+(?:adnexal|ovarian)","\b(?:adnexal|ovarian)\s+(?:mass|lesion).*?\bsolid\b"])
        lTypeIdx := 6
    else if TextScan.ContainsAny(text, ["\bunilocular\b"])
        lTypeIdx := 3
    else if TextScan.ContainsAny(text, ["\bsimple\s+cyst\b"])
        lTypeIdx := 2

    contourIdx := TextScan.ContainsAny(text
        , ["\birregular\s+(?:inner\s+)?(?:wall|contour|margin)"]) ? 2 : 1
    hasSolid := TextScan.ContainsAny(text
        , ["\bsolid\s+component\b","\bsolid\s+nodule\b","\bsolid\s+tissue\b","\bvascular\s+solid\b"])

    classicIdx := 1
    if TextScan.ContainsAny(text, ["\bdermoid\b","\bteratoma\b"])
        classicIdx := 2
    else if TextScan.ContainsAny(text, ["\bendometrioma\b","\bchocolate\s+cyst"])
        classicIdx := 3
    else if TextScan.ContainsAny(text, ["\bhemorrhagic\s+cyst","\bhemorrhagic\s+ovarian\s+cyst"])
        classicIdx := 4
    else if TextScan.ContainsAny(text, ["\bhydrosalpinx\b","\bdilated\s+(?:tube|fallopian)"])
        classicIdx := 5

    pap := 0
    if RegExMatch(text, "i)(\d+)\s+(?:or\s+more\s+)?papillary", &m)
        pap := SafeInt(m[1], 0)
    if (pap > 4)
        pap := 4
    papIdx := pap + 1   ; 0->1, 1->2, ..., 4+->5

    asc   := TextScan.ContainsAny(text, ["\bascites\b","\bfree\s+pelvic\s+fluid\b"])
    perit := TextScan.ContainsAny(text, ["\bperitoneal\s+(?:nodul|deposit|implant)","\bcarcinomatos"])

    form := RadsForm("O-RADS Ultrasound v2022", 580)

    form.Header("Patient")
    form.Dropdown("Meno", "Menstrual status:", ["Premenopausal", "Postmenopausal"], menoIdx)

    form.Header("Lesion")
    form.Dropdown("LType", "Lesion type:"
        , ["No lesion"
        ,  "Simple cyst (unilocular, anechoic, smooth thin wall)"
        ,  "Unilocular (with internal complexity, no solid)"
        ,  "Bilocular"
        ,  "Multilocular"
        ,  "Solid"
        ,  "Mixed cystic-solid"], lTypeIdx)
    form.Numeric("SizeCm", "Largest diameter (cm):", sz.cm > 0 ? Round(sz.cm, 1) : 0)
    form.Dropdown("Contour", "Inner contour:", ["Smooth", "Irregular"], contourIdx)
    form.Checkbox("HasSolid", "Solid component present", hasSolid)
    form.Checkbox("Shadowing", "Posterior acoustic shadowing (solid lesions only; diffuse or broad, not refractive artifact)", false)

    form.Dropdown("Classic", "Classic benign features:"
        , ["None"
        ,  "Dermoid / mature cystic teratoma"
        ,  "Endometrioma"
        ,  "Hemorrhagic cyst"
        ,  "Hydrosalpinx"], classicIdx)

    form.Header("Papillary projections and vascularity")
    form.Dropdown("PapCount", "Number of papillary projections:"
        , ["0", "1", "2", "3", "4 or more"], papIdx)
    form.Dropdown("Color", "Color score (vascular flow):"
        , ["1 -- avascular"
        ,  "2 -- minimal"
        ,  "3 -- moderate"
        ,  "4 -- marked"], 1)

    form.Header("Extra-ovarian findings")
    form.Checkbox("Ascites", "Ascites present", asc)
    form.Checkbox("Perit",   "Peritoneal nodularity present", perit)

    ; Progressive disclosure: lesion type gates the descriptor fields.
    ; Run once during build so the dialog opens showing only the
    ; questions relevant to the pre-detected lesion type.
    form.OnChange("LType", _OU_UpdateForm)
    _OU_UpdateForm(form)

    form.SetSubmit(ORADSUS_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_OU_UpdateForm(frm) {
    lt := _OU_LType(frm.byName["LType"].ctl.Text)
    hasLesion := (lt != "none")

    ; Hidden controls reset to their defaults, and _OU_Score provably
    ; ignores every field below for the lesion types that hide it:
    ;   - "No lesion" short-circuits to Score 1 (the ascites/peritoneal
    ;     upgrade only fires for score >=3, so it can never apply).
    ;   - HasSolid is never consulted on the lt="solid" rule rows.
    ;   - Shadowing only appears in the lt="solid" smooth CS<=3 rows.
    ;   - Contour is never consulted on the lt="mixed" rule rows.
    frm.SetVisible("SizeCm",    hasLesion)
    frm.SetVisible("Contour",   hasLesion && lt != "mixed")
    frm.SetVisible("HasSolid",  hasLesion && lt != "solid")
    frm.SetVisible("Shadowing", lt = "solid")
    frm.SetVisible("Classic",   hasLesion)
    frm.SetSectionVisible("Papillary projections and vascularity", hasLesion)
    frm.SetSectionVisible("Extra-ovarian findings", hasLesion)
}

ORADSUS_OnSubmit(v, form := "") {
    global g_LastSelectedText
    meno  := InStr(v.Meno, "Pre") ? "pre" : "post"
    lt    := _OU_LType(v.LType)
    size  := v.SizeCm + 0.0
    contour := InStr(v.Contour, "Irreg") ? "irregular" : "smooth"
    hasSolid := !!v.HasSolid
    shadowing := !!v.Shadowing
    classic := _OU_Classic(v.Classic)
    pap := _OU_PapCount(v.PapCount)
    color := _OU_Color(v.Color)
    asc := !!v.Ascites
    perit := !!v.Perit

    score := _OU_Score(lt, classic, size, hasSolid, pap, contour, color, asc, perit, meno, shadowing)
    info := _OU_Cats()[score]

    showRisk := Prefs.Get("display", "showMalignancyRisk", true)
    sizePhrase := size > 0 ? Format("{:.1f}", size) " cm " : ""
    descLc := (info.desc != "") ? StrLower(SubStr(info.desc, 1, 1)) SubStr(info.desc, 2) : ""
    mgmtLc := (info.mgmt != "") ? StrLower(SubStr(info.mgmt, 1, 1)) SubStr(info.mgmt, 2) : ""
    impression := sizePhrase "adnexal lesion, O-RADS US " score " (" descLc ")"
    if (showRisk && info.risk != "")
        impression .= ", malignancy risk " info.risk
    impression .= " per O-RADS v2022"
    if (mgmtLc != "")
        impression .= "; " mgmtLc
    if (SubStr(impression, -1) != ".")
        impression .= "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nO-RADS " score " -- " info.desc
    if (info.risk != "")
        method .= "`nMalignancy risk: " info.risk
    method .= "`nFull management: " info.mgmt

    return MakeResult({
        classification: "O-RADS US " score,
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "Andreotti RF, Timmerman D, Strachowski LM, et al. "
                                . "O-RADS US Risk Stratification and Management System: A Consensus "
                                . "Guideline from the ACR Ovarian-Adnexal Reporting and Data System "
                                . "Committee. Radiology. 2020;294(1):168-185.",
                           url:  "https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/O-RADS" }],
        echo:           g_LastSelectedText
    })
}

_OU_LType(label) {
    if InStr(label, "No lesion")
        return "none"
    if InStr(label, "Simple cyst")
        return "simple_cyst"
    if InStr(label, "Unilocular")
        return "unilocular"
    if InStr(label, "Bilocular")
        return "bilocular"
    if InStr(label, "Multilocular")
        return "multilocular"
    if InStr(label, "Mixed")
        return "mixed"
    return "solid"
}
_OU_Classic(label) {
    if InStr(label, "Dermoid")
        return "dermoid"
    if InStr(label, "Endometrioma")
        return "endometrioma"
    if InStr(label, "Hemorrhagic")
        return "hemorrhagic_cyst"
    if InStr(label, "Hydrosalpinx")
        return "hydrosalpinx"
    return "none"
}
_OU_PapCount(label) {
    if (label = "4 or more")
        return 4
    return SafeInt(label, 0)
}
_OU_Color(label) {
    return SafeInt(SubStr(label, 1, 1), 1)
}

_OU_Score(lt, classic, size, hasSolid, pap, contour, color, asc, perit, meno, shadowing := false) {
    score := 0

    if (lt = "none")
        score := 1
    ; Source (O-RADS US v2022 grid, Score 1 row): "Physiologic cyst: follicle
    ; (<=3 cm)" applies to PREmenopausal women only (the Score 2 row shows the
    ; premenopausal simple cyst <=3 cm column as "N/A (see follicle)" because
    ; it gets reclassified to Score 1). Postmenopausal simple cyst <=3 cm
    ; remains Score 2 with management "None".
    else if (lt = "simple_cyst" && size <= 3 && meno = "pre")
        score := 1
    else if (pap >= 4)
        score := 5
    else if (lt = "solid" && contour = "irregular")
        score := 5
    else if (lt = "solid" && contour = "smooth" && color >= 4)
        score := 5
    ; Source line 67: Score 5 includes "Bi- or multilocular cyst with solid
    ; component(s), any size, CS 3-4". Added "bilocular" to the rule.
    else if (hasSolid && (lt = "bilocular" || lt = "multilocular" || lt = "mixed") && color >= 3)
        score := 5
    else if (classic != "none" && size < 10)
        score := 2
    else if (classic != "none" && size >= 10)
        score := 3
    else if ((lt = "simple_cyst" || lt = "unilocular") && !hasSolid && contour != "irregular" && size < 10)
        score := 2
    else if (lt = "bilocular" && !hasSolid && contour = "smooth" && size < 10)
        score := 2
    else if (lt = "bilocular" && !hasSolid && contour = "irregular")
        score := 4
    else if ((lt = "simple_cyst" || lt = "unilocular" || lt = "bilocular") && !hasSolid && size >= 10)
        score := 3
    else if ((lt = "unilocular" || lt = "simple_cyst") && !hasSolid && contour = "irregular")
        score := 3
    else if ((lt = "unilocular" || lt = "bilocular" || lt = "simple_cyst") && hasSolid && pap < 4)
        score := 4
    else if (lt = "multilocular" && !hasSolid && contour != "irregular" && size < 10 && color <= 3)
        score := 3
    else if (lt = "multilocular" && !hasSolid && contour != "irregular" && size < 10 && color >= 4)
        score := 4
    else if (lt = "multilocular" && !hasSolid && contour != "irregular" && size >= 10)
        score := 4
    else if (lt = "multilocular" && !hasSolid && contour = "irregular")
        score := 4
    else if (hasSolid && (lt = "multilocular" || lt = "mixed") && color <= 2)
        score := 4
    else if (pap >= 1 && pap <= 3)
        score := 4
    else if (lt = "solid" && contour = "smooth" && color <= 1)
        score := 3
    ; Source row Score 3: "Solid lesion, shadowing, smooth, any size, CS 2-3"
    ; Source row Score 4: "Solid lesion, non-shadowing, smooth, any size, CS 2-3"
    ; Shadowing implies a benign solid mass (e.g. dermoid, fibroma) and
    ; downgrades the CS 2-3 cell from Score 4 to Score 3.
    else if (lt = "solid" && contour = "smooth" && color <= 3 && shadowing)
        score := 3
    else if (lt = "solid" && contour = "smooth" && color <= 3)
        score := 4
    else
        score := 3

    if ((asc || perit) && score >= 3)
        score := 5
    return score
}

_OU_Cats() {
    static m := _OU_BuildCats()
    return m
}
_OU_BuildCats() {
    m := Map()
    m[0] := { desc: "Incomplete",              risk: "N/A",     mgmt: "Additional imaging required." }
    ; Source: O-RADS US v2022 does not state a numeric malignancy risk for
    ; Score 1 (only Score 2+ have stated risk ranges).
    m[1] := { desc: "Normal",                  risk: "N/A",     mgmt: "No follow-up." }
    m[2] := { desc: "Almost certainly benign", risk: "<1%",     mgmt: "No follow-up if <10 cm." }
    m[3] := { desc: "Low risk",                risk: "1-10%",   mgmt: "Follow-up ultrasound in 6 months or MRI." }
    m[4] := { desc: "Intermediate risk",       risk: "10-50%",  mgmt: "MRI or gynecologic oncology consultation." }
    m[5] := { desc: "High risk",               risk: ">=50%",   mgmt: "Referral to gynecologic oncology." }
    return m
}
