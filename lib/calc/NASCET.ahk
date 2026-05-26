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
    if RegExMatch(text, "i)(\d+(?:\.\d+)?)\h*(?:mm|cm)[^0-9]*distal", &m)
        distal := m[1] + 0.0
    else if RegExMatch(text, "i)distal[^0-9]*(\d+(?:\.\d+)?)\h*(?:mm|cm)", &m)
        distal := m[1] + 0.0
    if RegExMatch(text, "i)(\d+(?:\.\d+)?)\h*(?:mm|cm)[^0-9]*stenosis", &m)
        stenosis := m[1] + 0.0
    else if RegExMatch(text, "i)stenosis[^0-9]*(\d+(?:\.\d+)?)\h*(?:mm|cm)", &m)
        stenosis := m[1] + 0.0

    form := RadsForm("NASCET (carotid stenosis)", 440)
    form.Header("Diameters (mm)")
    form.Numeric("Distal",   "Distal normal ICA:", distal)
    form.Numeric("Stenosis", "Residual lumen at stenosis:", stenosis)
    form.SetSubmit(NASCET_OnSubmit)
    form.AddButtons()
    form.Show()
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
        echo:           g_LastSelectedText
    })
}

CalcNASCET(input) {
    raw := input
    flat := RegExReplace(input, "`r?`n", " ")

    distal := "", stenosis := ""
    if RegExMatch(flat
        , "i)(?:distal.*?(\d+(?:\.\d+)?)(?:mm|cm)).*?(?:stenosis.*?(\d+(?:\.\d+)?)(?:mm|cm))"
        , &m) {
        distal := m[1] + 0
        stenosis := m[2] + 0
    } else if RegExMatch(flat
        , "i)(?:stenosis.*?(\d+(?:\.\d+)?)(?:mm|cm)).*?(?:distal.*?(\d+(?:\.\d+)?)(?:mm|cm))"
        , &m2) {
        stenosis := m2[1] + 0
        distal := m2[2] + 0
    } else {
        nums := []
        pos := 1
        while (pos := RegExMatch(flat, "(\d+(?:\.\d+)?)(?:mm|cm)?", &mn, pos)) {
            nums.Push(mn[1] + 0)
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
