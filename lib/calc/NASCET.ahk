; ============================================================
; lib/calc/NASCET.ahk -- carotid stenosis percent (NASCET method)
; Citation: N Engl J Med 1991;325:445-53.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

NASCET_Entry(input) {
    ShowNASCETDialog(input)
    return ""
}

ShowNASCETDialog(text := "") {
    distal := "", stenosis := ""
    ; Try VALUE-BEFORE-KEYWORD first ("8 mm distal") then value-after as
    ; fallback ("distal ICA 8 mm"). Without the value-before variant the
    ; engine in "8 mm distal, 3 mm at stenosis" matches "distal" then
    ; grabs the *next* number ("3"), assigning the wrong dimension to
    ; the distal slot.
    if RegExMatch(text, "i)(\d+(?:\.\d+)?)\h*(mm|cm)[^0-9]*distal", &m)
        distal := _NAS_ToMm(m[1], m[2])
    else if RegExMatch(text, "i)distal[^0-9]*(\d+(?:\.\d+)?)\h*(mm|cm)", &m)
        distal := _NAS_ToMm(m[1], m[2])
    if RegExMatch(text, "i)(\d+(?:\.\d+)?)\h*(mm|cm)[^0-9]*stenosis", &m)
        stenosis := _NAS_ToMm(m[1], m[2])
    else if RegExMatch(text, "i)stenosis[^0-9]*(\d+(?:\.\d+)?)\h*(mm|cm)", &m)
        stenosis := _NAS_ToMm(m[1], m[2])

    form := RadsForm("NASCET (carotid stenosis)", 440)
    form.Header("Diameters (mm)")
    form.Numeric("Distal",   "Distal normal ICA:", distal)
    form.Numeric("Stenosis", "Residual lumen at stenosis:", stenosis)
    form.SetSubmit(NASCET_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

NASCET_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (v.Distal = "" || v.Stenosis = "")
        return MakeResult({ impression: "Please enter both diameters.",
                            error: "Missing diameters" })

    distal := v.Distal + 0.0
    stenosis := v.Stenosis + 0.0
    if (distal = 0)
        return MakeResult({ impression: "Distal diameter cannot be zero.",
                            error: "Zero distal" })
    if (stenosis > distal)
        return MakeResult({ impression: "Residual lumen exceeds the distal ICA diameter -- "
                                      . "the two values are likely swapped. NASCET requires "
                                      . "distal normal ICA >= residual lumen.",
                            error: "Stenosis > distal" })

    pct := Round((distal - stenosis) / distal * 100, 1)
    severity := (pct < 50) ? "mild (<50%)"
              : (pct < 70) ? "moderate (50-69%)"
                           : "severe (>=70%)"

    impression := "Carotid stenosis " pct "% by NASCET method, " severity "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nDistal ICA: " distal " mm"
    method .= "`nResidual lumen at stenosis: " stenosis " mm"
    method .= "`nNASCET = (distal - stenosis) / distal * 100 = " pct "%"

    return MakeResult({
        classification: severity,
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "North American Symptomatic Carotid Endarterectomy Trial "
                                . "Collaborators. Beneficial Effect of Carotid Endarterectomy in "
                                . "Symptomatic Patients with High-Grade Carotid Stenosis. "
                                . "N Engl J Med. 1991;325(7):445-453.",
                           url:  "https://www.nejm.org/doi/full/10.1056/NEJM199108153250701" }],
        echo:           g_LastSelectedText,
        ; Parenthetical append: "...residual lumen 2 mm (NASCET 65%
        ; stenosis)." reads naturally inside the findings sentence.
        paste:          (pct != "") ? "NASCET " pct "% stenosis" : "",
        pasteMode:      ""
    })
}

CalcNASCET(input) {
    raw := input
    flat := RegExReplace(input, "`r?`n", " ")

    ; \h* between number and unit -- the old (?:mm|cm) with no gap meant the
    ; module's own documented example "Distal ICA = 6 mm" never matched and
    ; everything fell through to the grab-all-numbers fallback. Units are
    ; captured so cm values normalize to mm.
    distal := "", stenosis := ""
    if RegExMatch(flat
        , "i)(?:distal.*?(\d+(?:\.\d+)?)\h*(mm|cm)).*?(?:stenosis.*?(\d+(?:\.\d+)?)\h*(mm|cm))"
        , &m) {
        distal := _NAS_ToMm(m[1], m[2])
        stenosis := _NAS_ToMm(m[3], m[4])
    } else if RegExMatch(flat
        , "i)(?:stenosis.*?(\d+(?:\.\d+)?)\h*(mm|cm)).*?(?:distal.*?(\d+(?:\.\d+)?)\h*(mm|cm))"
        , &m2) {
        stenosis := _NAS_ToMm(m2[1], m2[2])
        distal := _NAS_ToMm(m2[3], m2[4])
    } else {
        nums := []
        pos := 1
        while (pos := RegExMatch(flat, "(\d+(?:\.\d+)?)\h*(mm|cm)?", &mn, pos)) {
            nums.Push(_NAS_ToMm(mn[1], mn[2]))
            pos += mn.Len[0] ? mn.Len[0] : 1
        }
        if (nums.Length < 2)
            return "Could not find two diameters for NASCET. Example: 'Distal ICA = 6 mm, Stenosis = 2 mm'"
        distal := Max(nums*)
        stenosis := Min(nums*)
    }

    if (distal = 0)
        return "Distal diameter cannot be zero."
    pct := Round((distal - stenosis) / distal * 100, 1)
    out := raw "`n`nNASCET Calculation:`nDistal: " distal " mm`nStenosis: " stenosis " mm`nNASCET: " pct "%"
    if (pct < 50)
        out .= "`nMild (<50%)"
    else if (pct < 70)
        out .= "`nModerate (50-69%)"
    else
        out .= "`nSevere (>=70%)"

    if Prefs.Get("display","showCitations",true)
        out .= "`nCitation: NASCET (N Engl J Med 1991;325:445-53)."
    return out
}

; Normalize a captured diameter to millimeters. Unit may be "" (fallback
; number scan) -- assume mm, matching the form's labeling.
_NAS_ToMm(value, unit) {
    v := value + 0.0
    return (unit = "cm") ? v * 10 : v
}
