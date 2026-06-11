; ============================================================
; lib/calc/RVLV.ahk -- RV/LV diameter ratio (CT pulmonary embolism)
; ------------------------------------------------------------
; The right-ventricular to left-ventricular short-axis diameter
; ratio on CT is a marker of right-heart strain in acute pulmonary
; embolism. Diameters are measured as the maximum minor-axis
; distance between the ventricular endocardium and the
; interventricular septum, on axial images (or reformatted
; 4-chamber). A ratio >= 1.0 is the commonly cited threshold for
; RV dysfunction and is associated with adverse outcomes; some
; studies use >= 0.9 on axial images.
;
; Reference: Meinel FG, Nance JW Jr, Schoepf UJ, et al. Predictive
; Value of Computed Tomography in Acute Pulmonary Embolism:
; Systematic Review and Meta-analysis. Am J Med. 2015;128(7):
; 747-759.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

RVLV_Entry(input) {
    ShowRVLVDialog(input)
    return ""
}

ShowRVLVDialog(text := "") {
    ; Prefill "RV 42 mm ... LV 35 mm" from the selection. \bRV\b / \bLV\b also
    ; match inside the "RV/LV" label, but those positions have no nearby
    ; number+mm, so RegExMatch advances to the real measurement.
    rv := "", lv := ""
    if RegExMatch(text, "i)\bRV\b[^0-9\r\n]{0,8}(\d+(?:\.\d+)?)\h*mm", &m)
        rv := m[1]
    if RegExMatch(text, "i)\bLV\b[^0-9\r\n]{0,8}(\d+(?:\.\d+)?)\h*mm", &m)
        lv := m[1]

    form := RadsForm("RV/LV Diameter Ratio", 460)
    form.Header("Ventricular short-axis diameters (mm)")
    form.Note("Maximum minor-axis distance from endocardium to interventricular septum, on axial (or reformatted 4-chamber) images.")
    form.Numeric("RV", "RV diameter:", rv)
    form.Numeric("LV", "LV diameter:", lv)
    form.SetSubmit(RVLV_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

RVLV_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (v.RV = "" || v.LV = "")
        return MakeResult({ impression: "Please enter both RV and LV diameters.",
                            error: "Missing diameters" })
    rv := v.RV + 0.0, lv := v.LV + 0.0
    if (rv <= 0 || lv <= 0)
        return MakeResult({ impression: "Diameters must be greater than zero.",
                            error: "Non-positive diameter" })

    ratio := _RVLV_Ratio(rv, lv)
    ratioDisp := Format("{:.2f}", ratio)

    ; Factual statement of the ratio, with the published threshold noted
    ; (matches the house style: report the number, attach the validated
    ; cutoff as context rather than a directive).
    rvD := _RVLV_FmtMm(rv), lvD := _RVLV_FmtMm(lv)
    impression := "RV/LV diameter ratio " ratioDisp " (RV " rvD " mm, LV " lvD " mm). "
    if (ratio >= 1.0)
        impression .= "Ratio >= 1.0, associated with right ventricular strain/dysfunction in acute pulmonary embolism."
    else
        impression .= "Ratio below the 1.0 right-heart-strain threshold."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nRV diameter: " rvD " mm"
    method .= "`nLV diameter: " lvD " mm"
    method .= "`nRatio = RV / LV = " ratioDisp
    method .= "`nThreshold: >= 1.0 is the commonly cited cutoff for RV dysfunction; >= 0.9 has also been used on axial images. Diameters are the maximum minor-axis endocardium-to-septum distance."

    return MakeResult({
        classification: "RV/LV ratio",
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "Meinel FG, Nance JW Jr, Schoepf UJ, et al. "
                                . "Predictive Value of Computed Tomography in Acute Pulmonary "
                                . "Embolism: Systematic Review and Meta-analysis. "
                                . "Am J Med. 2015;128(7):747-759.",
                           url:  "https://doi.org/10.1016/j.amjmed.2015.01.023" }],
        echo:           g_LastSelectedText,
        ; Parenthetical append: "...RV 42 mm, LV 35 mm (RV/LV ratio 1.20)."
        paste:          "RV/LV ratio " ratioDisp,
        pasteMode:      ""
    })
}

_RVLV_Ratio(rv, lv) {
    if (lv = 0)
        return 0
    return rv / lv
}

_RVLV_FmtMm(v) {
    return (v = Round(v)) ? Round(v) : Round(v, 1)
}
