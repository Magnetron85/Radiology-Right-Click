; ============================================================
; lib/calc/Volumes.ahk -- Ellipsoid + Bullet volume
; ------------------------------------------------------------
; Now form-driven (matching the v2.1 RADS calculators). The
; CalcEllipsoidVolume / CalcBulletVolume functions remain as
; text-parse helpers (used by PSADensity for inline volume
; computation); the menu entries open a small dialog that
; pre-fills three dimension fields from the highlighted text.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

EllipsoidVolume_Entry(input) {
    ShowVolumeDialog(input, "Ellipsoid Volume", CalcEllipsoidVolume)
    return ""
}

BulletVolume_Entry(input) {
    ShowVolumeDialog(input, "Bullet Volume", CalcBulletVolume)
    return ""
}

ShowVolumeDialog(text, title, calcFn) {
    a := 0, b := 0, c := 0
    unit := ""
    ; Use \h (horizontal whitespace) instead of \s -- \h matches the
    ; non-breaking space (U+00A0) that PowerScribe and other rich-text
    ; editors silently insert in measurement contexts. \s only matches
    ; ASCII whitespace, so "3.2\xA0x\xA02.3\xA0x\xA03.5 cm" (visually
    ; identical to "3.2 x 2.3 x 3.5 cm") would not match without \h.
    if RegExMatch(text, "i)(\d+(?:\.\d+)?)\h*[x×]\h*(\d+(?:\.\d+)?)(?:\h*[x×]\h*(\d+(?:\.\d+)?))?\h*(cm|mm)?", &m) {
        a := m[1] + 0.0
        b := m[2] + 0.0
        c := (m.Count >= 3 && m[3] != "") ? m[3] + 0.0 : 0
        if (m.Count >= 4 && m[4] != "")
            unit := m[4]
    }
    if (unit = "") {
        ; v1 heuristic: any decimal -> cm, else mm
        unit := (InStr(a, ".") || InStr(b, ".") || InStr(c, ".")) ? "cm" : "mm"
    }

    form := RadsForm(title, 420)
    form.Header("Dimensions")
    form.Numeric("A", "A:", a > 0 ? Round(a, 2) : "")
    form.Numeric("B", "B:", b > 0 ? Round(b, 2) : "")
    form.Numeric("C", "C:", c > 0 ? Round(c, 2) : "")
    form.Dropdown("Unit", "Unit:", ["cm", "mm"], unit = "mm" ? 2 : 1)
    form.SetSubmit(Volume_OnSubmit.Bind(calcFn))
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

Volume_OnSubmit(calcFn, v, form := "") {
    global g_LastSelectedText
    a := v.A, b := v.B, c := v.C
    if (a = "" || b = "" || c = "")
        return MakeResult({ impression: "Please enter all three dimensions.",
                            error: "Missing dimensions" })
    synth := a " x " b " x " c " " v.Unit
    body := calcFn.Call(synth)
    formulaName := (calcFn = CalcBulletVolume) ? "bullet" : "ellipsoid"

    ; CalcXxxVolume returns "DxDxD <unit> (NN cc)" -- pull the numeric volume
    ; and re-emit as "Bullet volume NN mL (a x b x c cm)." sentence form per
    ; the impression style guide (mL, not cc; prose, not raw expression).
    volNum := ""
    if RegExMatch(body, "\(\s*([\d.]+)\s*(?:cc)?\s*\)", &m)
        volNum := m[1]

    formulaLabel := (formulaName = "bullet") ? "Bullet" : "Ellipsoid"
    impression := volNum != ""
        ? formulaLabel " volume " volNum " mL (" a " x " b " x " c " " v.Unit ")."
        : body

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nDimensions: " a " x " b " x " c " " v.Unit
    method .= "`nFormula: " (formulaName = "bullet"
        ? "V = a * b * c * (5*pi/24)  (bullet / prolate spheroid cap)"
        : "V = (pi/6) * a * b * c     (ellipsoid)")
    method .= "`nRaw result: " body

    return MakeResult({
        classification: formulaLabel " volume",
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [],
        echo:           g_LastSelectedText,
        ; Parenthetical append: "...measures 3.1 x 2.2 x 2.8 cm (ellipsoid
        ; volume 10.0 mL)." reads naturally inside the findings sentence.
        paste:          volNum != "" ? StrLower(formulaLabel) " volume " volNum " mL" : "",
        pasteMode:      ""
    })
}

CalcEllipsoidVolume(input) {
    if !RegExMatch(input
        , "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
        , &m)
        return "Invalid input format for ellipsoid volume.`nExample: 3 x 2 x 1 cm"

    d := SortDimensions([m[1] + 0, m[2] + 0, m[3] + 0])
    isMm := InStr(input, "mm")
    isCm := InStr(input, "cm")
    if (!isMm && !isCm) {
        ; legacy heuristic: integers without dots -> mm
        isMm := (InStr(m[1], ".") = 0 && InStr(m[2], ".") = 0 && InStr(m[3], ".") = 0)
    }
    if isMm
        d[1] /= 10, d[2] /= 10, d[3] /= 10

    volume  := (1/6) * 3.141592653589793 * d[1] * d[2] * d[3]
    rounded := (volume < 1) ? Round(volume, 3) : Round(volume, 1)
    return input " (" rounded (Prefs.Get("display","units",true) ? " cc" : "") ")"
}

CalcBulletVolume(input) {
    if !RegExMatch(input
        , "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
        , &m)
        return "Invalid input format for bullet volume.`nExample: 3 x 2 x 1 cm"

    d := SortDimensions([m[1] + 0, m[2] + 0, m[3] + 0])
    isMm := InStr(input, "mm")
    isCm := InStr(input, "cm")
    if (!isMm && !isCm) {
        isMm := (InStr(m[1], ".") = 0 && InStr(m[2], ".") = 0 && InStr(m[3], ".") = 0)
    }
    if isMm
        d[1] /= 10, d[2] /= 10, d[3] /= 10

    volume  := d[1] * d[2] * d[3] * (5 * 3.141592653589793 / 24)
    rounded := (volume < 1) ? Round(volume, 3) : Round(volume, 1)
    return input " (" rounded (Prefs.Get("display","units",true) ? " cc" : "") ")"
}
