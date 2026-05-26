; ============================================================
; lib/calc/USLIRADS.ahk -- US LI-RADS surveillance categorization
; ------------------------------------------------------------
; Reference: ACR US LI-RADS v2024 (HCC surveillance in at-risk
; patients). Categories:
;   US-1 negative      -> routine 6-month US surveillance
;   US-2 subthreshold  -> short-term 3-6 month US follow-up x2
;   US-3 positive      -> diagnostic CT or MRI
; Visualization score A / B / C reports US technical quality.
; AFP+ in patients with US-1/US-2 routes to diagnostic CT/MRI
; per v2024 (positive AFP overrides surveillance recommendation).
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

USLIRADS_Entry(input) {
    ShowUSLIRADSDialog(input)
    return ""
}

ShowUSLIRADSDialog(text := "") {
    sz := TextScan.Size(text)
    hasObs := sz.mm > 0
    hasThrombus := TextScan.ContainsAny(text
        , ["\bthrombus\b","\bthrombosis\b","\btumor\s+thrombus\b"])
    hasDistortion := TextScan.ContainsAny(text
        , ["\bparenchymal\s+distortion\b"
        ,  "\bill[- ]defined\s+heterogeneity\b"
        ,  "\brefractive\s+edge\s+shadow"])

    form := RadsForm("US LI-RADS v2024", 520)
    form.Header("Observation")
    form.Dropdown("Obs", "Focal observation present?", ["No","Yes"], hasObs ? 2 : 1)
    form.Dropdown("Benign", "If yes, definitely benign (cyst, hemangioma, focal fat, calcified granuloma)?", ["No","Yes"], 1)
    form.Numeric("Size", "Observation size (mm, 0 if none):", sz.mm > 0 ? Round(sz.mm) : 0)
    form.Checkbox("Distortion", "Parenchymal distortion >=10 mm (ill-defined heterogeneity, refractive edge shadows, loss of architecture)", hasDistortion)
    form.Spacer(6)
    form.Header("Vascular")
    form.Checkbox("Thrombus", "New thrombus in portal/hepatic vein", hasThrombus)
    form.Spacer(6)
    form.Header("AFP (optional)")
    form.Checkbox("AfpPos", "AFP positive (>=20 ng/mL OR doubling on 2 consecutive tests)", false)
    form.Spacer(6)
    form.Header("Visualization score")
    form.Dropdown("Viz", "Image quality:"
        , ["A -- no or minimal limitations"
        ,  "B -- moderate limitations"
        ,  "C -- severe limitations"], 1)
    form.Header("VIS-C risk modifiers (only relevant if Visualization = C)")
    form.Checkbox("VizRisk", "MASH- or EtOH-related cirrhosis, Child-Turcotte-Pugh class B/C, or BMI >=35 kg/m^2", false)
    form.OnChange("Viz", _USLI_UpdateViz)
    form.OnChange("Obs", _USLI_UpdateObs)
    _USLI_UpdateViz(form)
    _USLI_UpdateObs(form)

    form.SetSubmit(USLIRADS_OnSubmit)
    form.AddButtons()
    form.Show()
}

_USLI_UpdateViz(frm) {
    isC := InStr(frm.byName["Viz"].ctl.Text, "C --")
    frm.SetEnabled("VizRisk", isC)
}
_USLI_UpdateObs(frm) {
    hasObs := frm.byName["Obs"].ctl.Text = "Yes"
    frm.SetEnabled("Benign", hasObs)
    frm.SetEnabled("Size",   hasObs)
}

USLIRADS_OnSubmit(v, form := "") {
    global g_LastSelectedText
    hasObs := (v.Obs = "Yes")
    isBen  := (v.Benign = "Yes")
    size   := SafeInt(v.Size, 0)
    thr    := !!v.Thrombus
    distortion := !!v.Distortion
    afpPos := !!v.AfpPos
    viz    := SubStr(v.Viz, 1, 1)
    vizRisk := !!v.VizRisk

    r := _USLI_Classify(hasObs, isBen, size, thr, viz, distortion, afpPos, vizRisk)

    mgmtLc := (r.mgmt != "") ? StrLower(SubStr(r.mgmt, 1, 1)) SubStr(r.mgmt, 2) : ""
    descLc := (r.desc != "") ? StrLower(SubStr(r.desc, 1, 1)) SubStr(r.desc, 2) : ""
    vizDescLc := (r.vizDesc != "") ? StrLower(SubStr(r.vizDesc, 1, 1)) SubStr(r.vizDesc, 2) : ""
    impression := "Hepatic screening US " r.cat " (" descLc "), "
               . "visualization " viz " (" vizDescLc ") per US LI-RADS v2024"
    if (mgmtLc != "")
        impression .= "; " mgmtLc
    if (SubStr(impression, -1) != ".")
        impression .= "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nCategory: " r.cat " -- " r.desc
    method .= "`nFull management: " r.mgmt
    if (r.afpNote != "")
        method .= "`nAFP override: " r.afpNote
    method .= "`nVisualization score: " viz " -- " r.vizDesc
    if (r.vizMgmt != "")
        method .= "`nVisualization note: " r.vizMgmt

    advisories := []
    if (r.afpNote != "")
        advisories.Push("AFP override: " r.afpNote)
    if (r.vizMgmt != "")
        advisories.Push(r.vizMgmt)

    return MakeResult({
        classification: r.cat,
        impression:     impression,
        recommendation: "",
        advisories:     advisories,
        methodology:    method,
        citations:      [{ text: "American College of Radiology. US LI-RADS v2024 "
                                . "-- Ultrasound Surveillance.",
                           url:  "https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/LI-RADS" }],
        echo:           g_LastSelectedText
    })
}

_USLI_Classify(hasObs, isBen, size, thr, viz, distortion := false, afpPos := false, vizRisk := false) {
    cats := _USLI_Cats()
    vizs := _USLI_Viz(vizRisk)
    cat := "US-1"
    ; v2024 categorization rules (Step 1):
    ;   - Thrombus at any size  -> US-3
    ;   - Parenchymal distortion >=10 mm  -> US-3 (source lines 659-668)
    ;   - Observation >=10 mm not definitely benign -> US-3
    ;   - Observation <10 mm not definitely benign -> US-2
    ;   - Otherwise (no obs, or definitely benign) -> US-1
    if (thr)
        cat := "US-3"
    else if (distortion)
        cat := "US-3"
    else if (hasObs && !isBen && size >= 10)
        cat := "US-3"
    else if (hasObs && !isBen)
        cat := "US-2"
    else
        cat := "US-1"
    c := cats[cat]
    vi := vizs.Has(viz) ? vizs[viz] : vizs["A"]
    ; AFP override (v2024 lines 111-113, 183-186): AFP+ with category not US-3
    ; routes to diagnostic MRI or CT regardless of US category.
    afpNote := ""
    if (afpPos && cat != "US-3")
        afpNote := "AFP positive in a patient with " cat ". Diagnostic MRI or CT is recommended -- CEUS is unlikely to be helpful when the AFP rise is not explained by a US observation."
    return { cat: cat, desc: c.desc, mgmt: c.mgmt
           , vizDesc: vi.desc, vizMgmt: vi.mgmt, afpNote: afpNote }
}

_USLI_Cats() {
    static m := _USLI_BuildCats()
    return m
}
_USLI_BuildCats() {
    m := Map()
    m["US-1"] := { desc: "Negative"
                 , mgmt: "Routine surveillance ultrasound in 6 months." }
    m["US-2"] := { desc: "Subthreshold"
                 , mgmt: "Repeat surveillance US in 3-6 months, up to two times. If the observation remains <10 mm or is no longer visualized on two consecutive exams, recategorize as US-1 and resume routine 6-month surveillance." }
    m["US-3"] := { desc: "Positive"
                 , mgmt: "Diagnostic contrast-enhanced multiphase CT or MRI recommended (LI-RADS CT/MRI algorithm). CEUS may also be appropriate per local availability." }
    return m
}
_USLI_Viz(vizRisk := false) {
    m := Map()
    m["A"] := { desc: "No or minimal limitations", mgmt: "" }
    ; VIS-B is managed identically to VIS-A per source (lines 162, 286-290 --
    ; "alternative surveillance is NOT recommended after VIS-A or VIS-B").
    m["B"] := { desc: "Moderate limitations -- may obscure small masses"
              , mgmt: "Managed identically to VIS-A (alternative surveillance is not recommended after VIS-B)." }
    ; VIS-C management per v2024 (lines 169-178): if no risk factors, repeat
    ; US within 3 months first; if any of MASH/EtOH cirrhosis, CTP B/C, or
    ; BMI >=35, consider alternative surveillance now.
    if vizRisk {
        m["C"] := { desc: "Severe limitations -- substantially reduced sensitivity"
                  , mgmt: "Risk factor(s) present (MASH/EtOH cirrhosis, CTP B/C, or BMI >=35). Consider alternative surveillance now: abbreviated MRI or multiphase CT." }
    } else {
        m["C"] := { desc: "Severe limitations -- substantially reduced sensitivity"
                  , mgmt: "No risk factors for repeat VIS-C. Repeat US within 3 months; if still VIS-C, consider alternative surveillance (abbreviated MRI or multiphase CT)." }
    }
    return m
}
