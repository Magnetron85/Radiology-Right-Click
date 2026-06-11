; ============================================================
; lib/calc/ICHVolume.ahk -- Intracerebral hemorrhage volume (ABC/2)
; ------------------------------------------------------------
; Reference: Kothari RU, Brott T, Broderick JP, Barsan WG,
; Sauerbeck LR, Zuccarello M, Khoury J. The ABCs of measuring
; intracerebral hemorrhage volumes. Stroke. 1996;27(8):1304-1305.
;
; Method: A x B x C / 2, where A, B, C are the three orthogonal
; hematoma diameters in cm; result is in mL. This is the ellipsoid
; volume with the 0.5 approximation (pi/6 ~= 0.52) that the ICH
; literature and the ICH Score use. C is the vertical extent,
; conventionally (number of slices the hematoma appears on x slice
; thickness). Per user policy this calculator is purely factual --
; it reports the volume and does not assert a prognostic threshold.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

ICHVolume_Entry(input) {
    ShowICHVolumeDialog(input)
    return ""
}

ShowICHVolumeDialog(text := "") {
    a := 0, b := 0, c := 0, unit := ""
    ; \h (not \s) so non-breaking spaces in "3.2<NBSP>x<NBSP>2.1 cm" match.
    if RegExMatch(text
        , "i)(\d+(?:\.\d+)?)\h*[x*]\h*(\d+(?:\.\d+)?)(?:\h*[x*]\h*(\d+(?:\.\d+)?))?\h*(cm|mm)?"
        , &m) {
        a := m[1] + 0.0
        b := m[2] + 0.0
        c := (m.Count >= 3 && m[3] != "") ? m[3] + 0.0 : 0
        if (m.Count >= 4 && m[4] != "")
            unit := m[4]
    }
    if (unit = "")
        unit := (InStr(a, ".") || InStr(b, ".") || InStr(c, ".")) ? "cm" : "mm"

    form := RadsForm("ICH Volume (ABC/2)", 460)
    form.Header("Hematoma dimensions")
    form.Note("ABC/2 method (Kothari 1996): A, B, C are the three orthogonal hematoma diameters; volume = A x B x C / 2. C is the vertical extent (slice count x thickness).")
    form.Numeric("A", "A -- largest diameter:",            a > 0 ? Round(a, 2) : "")
    form.Numeric("B", "B -- perpendicular to A (same slice):", b > 0 ? Round(b, 2) : "")
    form.Numeric("C", "C -- vertical extent (slices x thickness):", c > 0 ? Round(c, 2) : "")
    form.Dropdown("Unit", "Unit:", ["cm", "mm"], unit = "mm" ? 2 : 1)
    form.SetSubmit(ICHVolume_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

ICHVolume_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (v.A = "" || v.B = "" || v.C = "")
        return MakeResult({ impression: "Please enter all three dimensions.",
                            error: "Missing dimensions" })
    a := v.A + 0.0, b := v.B + 0.0, c := v.C + 0.0
    if (a <= 0 || b <= 0 || c <= 0)
        return MakeResult({ impression: "Dimensions must be greater than zero.",
                            error: "Non-positive dimension" })

    unit := v.Unit
    vol := _ICH_Volume(a, b, c, unit)
    volDisp := (vol < 1) ? Round(vol, 2) : Round(vol, 1)

    ; Dimensions echoed back in cm regardless of entry unit (the ABC/2
    ; convention is cm), so the parenthetical is unambiguous.
    aCm := (unit = "mm") ? a / 10 : a
    bCm := (unit = "mm") ? b / 10 : b
    cCm := (unit = "mm") ? c / 10 : c
    dimStr := _ICH_Fmt(aCm) " x " _ICH_Fmt(bCm) " x " _ICH_Fmt(cCm) " cm"

    impression := "Intracerebral hemorrhage volume " volDisp " mL by the ABC/2 method (" dimStr ")."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nDimensions: " dimStr
    method .= "`nFormula: volume = A x B x C / 2 = " _ICH_Fmt(aCm) " x " _ICH_Fmt(bCm) " x " _ICH_Fmt(cCm) " / 2 = " volDisp " mL"
    method .= "`nNote: ABC/2 is the ellipsoid approximation; it tends to overestimate volume for irregular or multinodular hematomas. C is conventionally (number of slices x slice thickness)."

    return MakeResult({
        classification: "ICH volume",
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "Kothari RU, Brott T, Broderick JP, et al. "
                                . "The ABCs of measuring intracerebral hemorrhage volumes. "
                                . "Stroke. 1996;27(8):1304-1305.",
                           url:  "https://doi.org/10.1161/01.str.27.8.1304" }],
        echo:           g_LastSelectedText,
        ; Parenthetical append: "...measures 4.0 x 3.0 x 2.5 cm (ICH volume
        ; 15.0 mL by ABC/2)." reads naturally inside the findings sentence.
        paste:          "ICH volume " volDisp " mL by ABC/2",
        pasteMode:      ""
    })
}

; mL volume from three diameters by ABC/2. `unit` is "cm" or "mm";
; mm diameters are converted to cm before the cm^3 (= mL) computation.
_ICH_Volume(a, b, c, unit) {
    if (unit = "mm") {
        a /= 10
        b /= 10
        c /= 10
    }
    return a * b * c / 2
}

; Integer when whole (avoid "4.0 x 3.0"), else up to 2 decimals.
_ICH_Fmt(v) {
    return (v = Round(v)) ? Round(v) : Round(v, 2)
}
