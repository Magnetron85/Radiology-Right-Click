; ============================================================
; lib/calc/NASCET.ahk -- carotid stenosis percent (NASCET method)
; Citation: N Engl J Med 1991;325:445-53.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

NASCET_Entry(input) {
    return CalcNASCET(input)
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
