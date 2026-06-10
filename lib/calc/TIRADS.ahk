; ============================================================
; lib/calc/TIRADS.ahk -- ACR TI-RADS point system
; ------------------------------------------------------------
; Reference: Tessler FN et al. ACR Thyroid Imaging, Reporting
; and Data System (TI-RADS): White Paper of the ACR TI-RADS
; Committee. J Am Coll Radiol. 2017;14(5):587-595.
;
; Total points -> TR level (Tessler 2017 Figure 1, p.588):
;   0          -> TR1 (benign)            no FNA / no follow-up
;   2          -> TR2 (not suspicious)    no FNA / no follow-up
;   3          -> TR3 (mildly suspicious) FNA >=2.5 cm; f/u >=1.5 cm
;   4-6        -> TR4 (moderate)          FNA >=1.5 cm; f/u >=1.0 cm
;   >=7        -> TR5 (highly suspicious) FNA >=1.0 cm; f/u >=0.5 cm
; Note: 1 point is impossible per Tessler -- the lowest non-zero score is 2
; because mixed composition (+1) is always paired with an echogenicity point.
; Spongiform / purely cystic -> TR1, with ZERO points (no other features count).
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

TIRADS_Entry(input) {
    ShowTIRADSDialog(input)
    return ""
}

ShowTIRADSDialog(text := "") {
    sz := TextScan.Size(text)
    comp := TextScan.Composition(text)
    echo := TextScan.Echogenicity(text)
    shape := TextScan.Shape(text)
    margin := TextScan.Margin(text)
    compIdx := (comp = "cystic") ? 1 : (comp = "spongiform") ? 2
            : (comp = "mixed")  ? 3 : 4
    echoIdx := (echo = "anechoic") ? 1
            : (echo = "hyperechoic" || echo = "isoechoic") ? 2
            : (echo = "very hypoechoic") ? 4
            : (echo = "hypoechoic") ? 3 : 3
    shapeIdx := (shape = "taller") ? 2 : 1
    marginIdx := (margin = "lobulated") ? 3
              : (margin = "extrathyroidal") ? 4
              : (margin = "ill-defined") ? 2 : 1
    micro := TextScan.ContainsAny(text, ["\bmicrocalcification","\bpunctate\s+echogenic"])
    macro := TextScan.ContainsAny(text, ["\bmacrocalcification","\bcoarse\s+calcif"])
    rim   := TextScan.ContainsAny(text, ["\bperipheral\s+calcif","\brim\s+calcif"])
    comet := TextScan.ContainsAny(text, ["\bcomet[- ]tail\b"])

    form := RadsForm("ACR TI-RADS", 540)

    ; Always-needed inputs first: composition (the gate -- cystic /
    ; spongiform short-circuit to TR1 with zero points) and size (echoed
    ; in the impression and used for FNA / follow-up thresholds).
    form.Header("Composition (0-2)")
    form.Dropdown("Comp", "Composition:"
        , ["Cystic / almost completely cystic (0)"
        ,  "Spongiform (0)"
        ,  "Mixed cystic and solid (1)"
        ,  "Solid or almost completely solid (2)"], compIdx)

    form.Header("Size")
    form.Numeric("SizeCm", "Largest diameter (cm):", sz.cm > 0 ? Round(sz.cm, 1) : 0)

    form.Header("Echogenicity (0-3)")
    form.Dropdown("Echo", "Echogenicity:"
        , ["Anechoic (0)"
        ,  "Hyperechoic or isoechoic (1)"
        ,  "Hypoechoic (2)"
        ,  "Very hypoechoic (3)"], echoIdx)

    form.Header("Shape (0 or 3)")
    form.Dropdown("Shape", "Shape:"
        , ["Wider-than-tall (0)"
        ,  "Taller-than-wide (3)"], shapeIdx)

    form.Header("Margin (0-3)")
    form.Dropdown("Margin", "Margin:"
        , ["Smooth (0)"
        ,  "Ill-defined (0)"
        ,  "Lobulated or irregular (2)"
        ,  "Extra-thyroidal extension (3)"], marginIdx)

    form.Header("Echogenic foci (sum, multiple may apply)")
    form.Checkbox("FociMacro",   "Macrocalcifications (+1)", macro)
    form.Checkbox("FociRim",     "Peripheral (rim) calcifications (+2)", rim)
    form.Checkbox("FociPunct",   "Punctate echogenic foci (+3)", micro)
    form.Checkbox("FociComet",   "Comet-tail artifact only (+0; usually benign)", comet)

    ; Progressive disclosure: composition gates the four point sections.
    ; Run once during build so a cystic / spongiform pre-detect opens the
    ; dialog minimal.
    form.OnChange("Comp", _TIRADS_UpdateForm)
    _TIRADS_UpdateForm(form)

    form.SetSubmit(TIRADS_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_TIRADS_UpdateForm(frm) {
    comp := _TIRADS_Comp(frm.byName["Comp"].ctl.Text)
    ; Per Tessler 2017 (and _TIRADS_Pts / _TIRADS_Level): cystic and
    ; spongiform nodules receive ZERO points total and map to TR1 --
    ; "Do not add further points for other categories". Echogenicity,
    ; shape, margin, and echogenic foci are provably never read in that
    ; branch, so hide those whole sections. Size stays visible: it is
    ; echoed into the impression line regardless of TR level.
    scored := (comp != "cystic" && comp != "spongiform")
    frm.SetSectionVisible("Echogenicity (0-3)", scored)
    frm.SetSectionVisible("Shape (0 or 3)", scored)
    frm.SetSectionVisible("Margin (0-3)", scored)
    frm.SetSectionVisible("Echogenic foci (sum, multiple may apply)", scored)
}

TIRADS_OnSubmit(v, form := "") {
    global g_LastSelectedText
    comp := _TIRADS_Comp(v.Comp)
    pts  := _TIRADS_Pts(v.Comp, v.Echo, v.Shape, v.Margin
                     , !!v.FociMacro, !!v.FociRim, !!v.FociPunct)
    size := v.SizeCm + 0.0

    r := _TIRADS_Level(pts, comp)
    mgmt := _TIRADS_Mgmt(r, size)
    showRisk := Prefs.Get("display", "showMalignancyRisk", true)

    sizePhrase := size > 0 ? Format("{:.1f}", size) " cm " : ""
    descLc := (r.desc != "") ? StrLower(SubStr(r.desc, 1, 1)) SubStr(r.desc, 2) : ""
    mgmtLc := (mgmt != "") ? StrLower(SubStr(mgmt, 1, 1)) SubStr(mgmt, 2) : ""
    impression := sizePhrase "thyroid nodule, " r.level " (" descLc ")"
    if (showRisk && r.risk != "")
        impression .= ", malignancy risk " r.risk
    impression .= " per ACR TI-RADS 2017"
    if (mgmtLc != "")
        impression .= "; " mgmtLc
    if (SubStr(impression, -1) != ".")
        impression .= "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nTotal points: " pts
    method .= "`nLevel: " r.level " -- " r.desc
    if (r.risk != "")
        method .= "`nMalignancy risk: " r.risk " (95% CI " r.ci "%)"
    method .= "`nFull management: " mgmt

    citations := [{ text: "Tessler FN, Middleton WD, Grant EG, et al. "
                        . "ACR Thyroid Imaging, Reporting and Data System (TI-RADS): "
                        . "White Paper of the ACR TI-RADS Committee. "
                        . "J Am Coll Radiol. 2017;14(5):587-595.",
                    url:  "https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/TI-RADS" }]
    if showRisk
        citations.Push({ text: "Middleton WD, Teefey SA, Reading CC, et al. "
                              . "Multiinstitutional analysis of thyroid nodule risk stratification "
                              . "using the ACR TIRADS. AJR. 2017;208(6):1331-1341.",
                         url:  "https://doi.org/10.2214/AJR.16.17613" })

    return MakeResult({
        classification: r.level,
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      citations,
        echo:           g_LastSelectedText
    })
}

_TIRADS_Comp(label) {
    ; Mixed must be checked BEFORE Cystic -- the Mixed label
    ; ("Mixed cystic and solid") contains the substring "cystic"
    ; and AHK v2 InStr is case-insensitive by default.
    if InStr(label, "Mixed")
        return "mixed"
    if InStr(label, "Spongiform")
        return "spongiform"
    if InStr(label, "Cystic")
        return "cystic"
    return "solid"
}

_TIRADS_Pts(compLabel, echoLabel, shapeLabel, marginLabel, macro, rim, punct) {
    ; Source (Tessler 2017, Figure 1 + lines 148-150, 190-192): cystic and
    ; spongiform nodules receive ZERO points total -- "Do not add further
    ; points for other categories". Short-circuit instead of computing and
    ; relying on _TIRADS_Level to override.
    ;
    ; Mixed must be checked BEFORE the cystic/spongiform short-circuit
    ; because the Mixed label ("Mixed cystic and solid") contains the
    ; substring "cystic" and AHK v2 InStr is case-insensitive by default.
    total := 0
    if InStr(compLabel, "Mixed")
        total += 1
    else if InStr(compLabel, "Spongiform") || InStr(compLabel, "Cystic")
        return 0
    else if InStr(compLabel, "Solid")
        total += 2

    if InStr(echoLabel, "Hyperechoic") || InStr(echoLabel, "isoechoic")
        total += 1
    else if InStr(echoLabel, "Very hypoechoic")
        total += 3
    else if InStr(echoLabel, "Hypoechoic")
        total += 2
    ; anechoic -> 0

    if InStr(shapeLabel, "Taller")
        total += 3

    if InStr(marginLabel, "Lobulated") || InStr(marginLabel, "irregular")
        total += 2
    else if InStr(marginLabel, "Extra")
        total += 3

    if macro
        total += 1
    if rim
        total += 2
    if punct
        total += 3
    return total
}

_TIRADS_Level(pts, comp) {
    levels := _TIRADS_LevelTable()
    if (comp = "spongiform" || comp = "cystic")
        lvl := "TR1"
    else if (pts = 0)
        lvl := "TR1"
    else if (pts <= 2)
        lvl := "TR2"
    else if (pts = 3)
        lvl := "TR3"
    else if (pts <= 6)
        lvl := "TR4"
    else
        lvl := "TR5"
    info := levels[lvl]
    return { level: lvl, desc: info.desc, risk: info.risk, ci: info.ci
           , fnaThreshold: info.fnaThreshold, fuThreshold: info.fuThreshold }
}

_TIRADS_Mgmt(r, size) {
    if (r.level = "TR1" || r.level = "TR2")
        return "No FNA or routine follow-up recommended."
    ; Tessler 2017 lines 277-285: 5-9 mm TR5 nodules may warrant FNA via
    ; shared decision-making (papillary microcarcinoma considerations).
    ; This branch must run BEFORE the standard FNA threshold check
    ; (TR5 fnaThreshold=1.0) otherwise it is unreachable.
    if (r.level = "TR5" && size >= 0.5 && size < 1.0)
        return "For 5-9 mm TR5 nodules, FNA may be considered with shared decision-making (Tessler 2017, papillary microcarcinoma footnote)." . _TIRADS_FollowupSchedule(r.level)
    if (r.fnaThreshold > 0 && size >= r.fnaThreshold)
        return "FNA recommended (nodule >= " r.fnaThreshold " cm for " r.level ")." . _TIRADS_FollowupSchedule(r.level)
    if (r.fuThreshold > 0 && size >= r.fuThreshold)
        return "Follow-up ultrasound recommended (nodule >= " r.fuThreshold " cm for " r.level ")." . _TIRADS_FollowupSchedule(r.level)
    return "No FNA or routine follow-up at this size for " r.level ". Consider follow-up if nodule grows."
}

; Per-TR follow-up intervals per Tessler 2017 (lines 320-325):
;   TR3 -> at 1, 3, and 5 years
;   TR4 -> at 1, 2, 3, and 5 years
;   TR5 -> annually for up to 5 years
_TIRADS_FollowupSchedule(level) {
    if (level = "TR3")
        return " Follow-up ultrasound schedule: 1, 3, and 5 years."
    if (level = "TR4")
        return " Follow-up ultrasound schedule: 1, 2, 3, and 5 years."
    if (level = "TR5")
        return " Follow-up ultrasound schedule: annually for up to 5 years."
    return ""
}

_TIRADS_LevelTable() {
    static t := _TIRADS_BuildLevels()
    return t
}
; Risk percentages and 95% CIs are the empirical per-TR malignancy rates
; published in Middleton WD, Teefey SA, Reading CC, et al. Multiinstitutional
; analysis of thyroid nodule risk stratification using the ACR TIRADS. AJR
; 2017;208(6):1331-1341 (Figure 3, n=3422 cytopath/histopath-proven nodules).
; Display of these values is gated on display.showMalignancyRisk in Prefs.
_TIRADS_BuildLevels() {
    m := Map()
    m["TR1"] := { desc: "Benign",                risk: "0.3%",  ci: "0.0-1.8",  fnaThreshold: 0,   fuThreshold: 0 }
    m["TR2"] := { desc: "Not suspicious",        risk: "1.5%",  ci: "0.6-2.9",  fnaThreshold: 0,   fuThreshold: 0 }
    m["TR3"] := { desc: "Mildly suspicious",     risk: "4.8%",  ci: "3.4-6.5",  fnaThreshold: 2.5, fuThreshold: 1.5 }
    m["TR4"] := { desc: "Moderately suspicious", risk: "9.1%",  ci: "7.6-10.8", fnaThreshold: 1.5, fuThreshold: 1.0 }
    m["TR5"] := { desc: "Highly suspicious",     risk: "35.0%", ci: "31.0-39.1", fnaThreshold: 1.0, fuThreshold: 0.5 }
    return m
}
