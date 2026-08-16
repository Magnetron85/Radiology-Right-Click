; ============================================================
; lib/calc/NoduleSizes.ahk -- Compare + Sort measurement helpers
; ------------------------------------------------------------
; CompareSizes_Entry:  detects previous vs current measurements
;                      and reports dimension / volume / doubling
;                      time / exponential growth rate.
; SortSizes_Entry:     reorders every 2D / 3D measurement in the
;                      selection so the largest dimension is
;                      listed first; pastes back over selection.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

; ---------- Compare ---------------------------------------------------------

CompareSizes_Entry(input) {
    ShowCompareSizesDialog(input)
    return ""
}

ShowCompareSizesDialog(text := "") {
    form := RadsForm("Compare Measurement Sizes", 580)
    form.Header("Sentence with current + previous measurements")
    form.Note("Use keywords like 'now', 'previously', 'current', 'prior'. Dates in MM/DD/YYYY enable doubling-time + growth-rate computation.")
    form.TextArea("Desc", "Description:", 120, text)
    form.SetSubmit(CompareSizes_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

CompareSizes_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (Trim(v.Desc) = "")
        return MakeResult({ impression: "Please paste a sentence with current and previous measurements.",
                            error: "Missing description" })
    body := CompareSizes_Compute(v.Desc)

    ; Compose a single-sentence impression from the labeled body fields.
    longestChange := ""
    volChange := ""
    doublingDays := ""
    growthRate := ""
    if RegExMatch(body, "i)Longest dimension change[:]?\s*([+\-]?[\d.]+%)", &m)
        longestChange := m[1]
    if RegExMatch(body, "i)Volume change[:]?\s*([+\-]?[\d.]+%)", &m)
        volChange := m[1]
    if RegExMatch(body, "i)Doubling time[:]?\s*([\d.]+\s*days)", &m)
        doublingDays := m[1]
    if RegExMatch(body, "i)Exponential Growth Rate[:]?\s*([+\-]?[\d.]+%)", &m)
        growthRate := m[1]

    if (longestChange != "") {
        impression := "Interval change in longest dimension " longestChange
        if (volChange != "")
            impression .= ", volume change " volChange
        if (doublingDays != "")
            impression .= "; doubling time " doublingDays
        if (growthRate != "")
            impression .= ", growth rate " growthRate "/year"
        impression .= "."
    } else {
        impression := body
    }

    ; Short parenthetical fragment: lowercase start, no trailing period.
    pasteFrag := ""
    if (longestChange != "") {
        pasteFrag := "interval change in longest dimension " longestChange
        if (volChange != "")
            pasteFrag .= ", volume " volChange
        if (doublingDays != "")
            pasteFrag .= ", doubling time " doublingDays
    }

    return MakeResult({
        classification: "Size comparison",
        impression:     impression,
        recommendation: "",
        methodology:    body,
        citations:      [],
        echo:           g_LastSelectedText,
        paste:          pasteFrag,
        pasteMode:      ""
    })
}

CompareSizes_Compute(input) {
    if (input = "")
        return "No text selected. Highlight a sentence with two measurements first."

    p1 := "i)(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?.*?(?:previous(?:ly)?|prior|before|old|initial).*?(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?(?:\s*(?:on|dated?)\s*(\d{1,2}/\d{1,2}/\d{2,4}))?"
    p2 := "i)(?:previous(?:ly)?|prior|before|old|initial).*?(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?(?:\s*(?:on|dated?)\s*(\d{1,2}/\d{1,2}/\d{2,4}))?.*?(?:now|current(?:ly)?|present|new|recent(?:ly)?|follow[- ]?up).*?(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?"

    if RegExMatch(input, p1, &m) {
        current := m[2] " " m[3]
        previous := m[5] " " m[6]
        currentDate  := m[1]
        previousDate := m[4] != "" ? m[4] : m[7]
    } else if RegExMatch(input, p2, &m) {
        ; p2 groups: 1=prev date, 2=prev meas, 3=prev unit, 4=prev "on" date,
        ; 5=cur date, 6=cur meas, 7=cur unit. The old indices here read the
        ; current DATE as the measurement and the current UNIT as its date.
        previous := m[2] " " m[3]
        current  := m[6] " " m[7]
        previousDate := m[1] != "" ? m[1] : m[4]
        currentDate  := m[5]
    } else {
        return "Invalid input format. Please provide both current and previous measurements."
    }
    return _CompareMeasurements(previous, current, previousDate, currentDate, input)
}

_CompareMeasurements(previous, current, prevDate, curDate, input) {
    prev := _ParseMeasurement(previous)
    curr := _ParseMeasurement(current)
    if (prev.dims.Length != curr.dims.Length)
        return "Error: Mismatch in number of dimensions between previous and current measurements."
            . "`nPrevious: " previous "`nCurrent: " current

    out := input "`n`n"
    out .= "Previous Date: " (prevDate != "" ? prevDate : "Not provided") "`n"
    out .= "Current Date: " (curDate != "" ? curDate : "Assumed as today") "`n`n"

    prevLong := 0, curLong := 0
    loop prev.dims.Length {
        i := A_Index
        pd := prev.dims[i], cd := curr.dims[i]
        pCm := (prev.unit = "mm") ? pd/10 : pd
        cCm := (curr.unit = "mm") ? cd/10 : cd
        if (pCm > prevLong)
            prevLong := pCm
        if (cCm > curLong)
            curLong := cCm
        change := (pCm != 0) ? (cCm/pCm - 1) * 100 : 0
        out .= "Dimension " i ": " Round(pd,2) " " prev.unit " -> " Round(cd,2) " " curr.unit
            . " (" (change >= 0 ? "+" : "") Round(change,1) "%)`n"
    }

    longChange := (prevLong != 0) ? (curLong/prevLong - 1) * 100 : 0
    out .= "`nLongest dimension change: " (longChange >= 0 ? "+" : "") Round(longChange,1) "%`n"

    prevVol := _Volume(prev.dims, prev.unit)
    curVol  := _Volume(curr.dims, curr.unit)
    if (prevVol != "" && curVol != "") {
        volChange := (prevVol != 0) ? (curVol/prevVol - 1) * 100 : 0
        out .= "Volume change: " (volChange >= 0 ? "+" : "") Round(volChange,1) "%`n"
        out .= "Previous volume: " _FormatVolume(prevVol) "`n"
        out .= "Current volume: "  _FormatVolume(curVol) "`n"

        if (prevDate != "") {
            pParsed := ParseDate(prevDate)
            if (pParsed = "")
                return out . "`nError: Invalid previous date format."
            if (curDate = "")
                curDate := FormatTime(A_Now, "MM/dd/yyyy")
            cParsed := ParseDate(curDate)
            if (cParsed = "")
                return out . "`nError: Invalid current date format."
            years := DaysBetween(pParsed, cParsed) / 365.25
            out .= "Time difference: " Round(years, 2) " years`n"
            if (years > 0) {
                dt := _DoublingTime(prevVol, curVol, years)
                out .= "Doubling time: " (dt != "N/A" ? Round(dt * 365.25, 0) " days" : dt) "`n"
                g := Ln(curVol/prevVol) / years
                out .= "Exponential Growth Rate: " Round(g * 100, 2) "% per year"
            } else {
                out .= "Note: Doubling time and Growth Rate not calculated due to invalid time difference."
            }
        } else {
            out .= "Note: Doubling time and Growth Rate not calculated due to missing previous date."
        }
    } else {
        out .= "Error: Unable to calculate volume for one or both measurements."
    }
    return out
}

_ParseMeasurement(s) {
    dims := []
    unit := ""
    if RegExMatch(s
        , "i)(\d+(?:\.\d+)?)(?:\s*(?:x|\*)\s*(\d+(?:\.\d+)?))?(?:\s*(?:x|\*)\s*(\d+(?:\.\d+)?))?(?=\s*(cm|mm)?)"
        , &m) {
        dims.Push(m[1] + 0)
        if (m.Count >= 2 && m[2] != "")
            dims.Push(m[2] + 0)
        if (m.Count >= 3 && m[3] != "")
            dims.Push(m[3] + 0)
        if (m.Count >= 4 && m[4] != "")
            unit := m[4]
    }
    if (unit = "") {
        hasDecimal := false
        for v in dims {
            if InStr(String(v), ".") {
                hasDecimal := true
                break
            }
        }
        unit := hasDecimal ? "cm" : "mm"
    }
    return { dims: dims, unit: unit }
}

_Volume(dims, unit) {
    pi := 3.14159265358979
    if (dims.Length = 0)
        return ""
    if (dims.Length = 1)
        v := (4/3) * pi * (dims[1]/2) ** 3
    else if (dims.Length = 2)
        v := (4/3) * pi * (dims[1]/2) * (dims[2]/2) * ((dims[1] + dims[2]) / 4)
    else if (dims.Length = 3) {
        sorted := SortDimensions([dims[1] + 0, dims[2] + 0, dims[3] + 0])
        v := (1/6) * pi * sorted[1] * sorted[2] * sorted[3]
    } else
        return ""
    return (unit = "mm") ? v / 1000 : v
}

_FormatVolume(v) {
    return v < 1 ? Round(v * 1000, 1) " cu-mm" : Round(v, 1) " cc"
}

_DoublingTime(v0, v1, yrs) {
    g := (v1/v0) ** (1/yrs) - 1
    return (g > 0) ? Ln(2) / Ln(1 + g) : "N/A"
}

; ---------- Sort -----------------------------------------------------------

SortSizes_Entry(input) {
    global g_LastSelectedText
    if (input = "")
        return ""

    ; Sort pastes directly over the selection and shows no result window, so --
    ; unlike the GUI calculators, which reset the menu's capture cache via
    ; _ClearLastSelection when their result window closes -- nothing else would
    ; clear g_LastSelectedText. Once this selection has been consumed, drop it:
    ; otherwise a later right-click whose fresh capture returns empty (PowerScribe
    ; drops its selection when the context menu takes focus) would fall back to
    ; THIS stale text and re-sort it over a different, newly highlighted
    ; measurement. Clearing turns that failure into a safe no-op instead.
    g_LastSelectedText := ""

    processed := _SortAllMeasurements(input)
    if (processed = input) {
        ; Not silent: the smart-match menu can suggest Sort for a measurement
        ; that is already largest-first, and a wordless return would be
        ; indistinguishable from a failure. A transient tooltip gives feedback
        ; without opening a result window or touching the clipboard.
        ToolTip("Already sorted (largest first) -- nothing to change.")
        SetTimer((*) => ToolTip(), -1500)
        return ""
    }
    leading  := (SubStr(input, 1, 1) = " ") ? " " : ""
    trailing := (SubStr(input, -1) = " ")   ? " " : ""

    ; Save / restore the user's clipboard around the paste so we don't
    ; clobber whatever they had copied. 100 ms is enough for the paste to
    ; land in PowerScribe / Notepad before we put the original back.
    saved := ClipboardAll()
    A_Clipboard := leading . Trim(processed) . trailing
    if !ClipWait(0.5) {
        A_Clipboard := saved
        return "Sort failed: clipboard did not accept the new value."
    }
    ; SendEvent, not the v2 default SendInput: dictation systems (PowerScribe /
    ; Dragon) run low-level keyboard hooks that SendInput bypasses, so the paste
    ; silently no-ops there. Same reasoning as GetSelectedText and the menu.
    SendEvent "^v"
    Sleep 100
    A_Clipboard := saved
    return ""
}

_SortAllMeasurements(input) {
    ; i) so a dictated uppercase "X" separator sorts too -- the smart-match
    ; dispatcher matches case-insensitively, and whatever it suggests this
    ; parser must be able to reorder. Output normalizes to lowercase "x".
    input := _SortPattern(input, "i)\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*", 3)
    input := _SortPattern(input, "i)\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*", 3)
    ; Bare comma PAIRS are capped at 1-3 integer digits so textual dates
    ; ("May 3, 2026") and other large-number enumerations are never rewritten
    ; into "2026 x 3" -- real cm/mm measurements don't have 4-digit values.
    input := _SortPattern(input, "i)\s*(?<![\d.])(\d{1,3}(?:\.\d+)?)\s*,\s*(\d{1,3}(?:\.\d+)?)(?![\d.])\s*", 2)
    input := _SortPattern(input, "i)\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*",                       2)
    return input
}

_SortPattern(input, needle, dims) {
    pos := 1
    while (pos := RegExMatch(input, needle, &m, pos)) {
        if (dims = 3)
            processed := _SortTriple(m[1], m[2], m[3])
        else
            processed := _SortPair(m[1], m[2])
        if (processed != m[0])
            input := SubStr(input, 1, pos - 1) . processed . SubStr(input, pos + m.Len[0])
        pos += StrLen(processed)
    }
    return input
}

_SortTriple(a, b, c) {
    an := a + 0, bn := b + 0, cn := c + 0
    if (an < bn) {
        t := a, a := b, b := t
        t := an, an := bn, bn := t
    }
    if (bn < cn) {
        t := b, b := c, c := t
        t := bn, bn := cn, cn := t
    }
    if (an < bn) {
        t := a, a := b, b := t
    }
    return " " Trim(a) " x " Trim(b) " x " Trim(c) " "
}

_SortPair(a, b) {
    if ((a + 0) < (b + 0)) {
        t := a, a := b, b := t
    }
    return " " Trim(a) " x " Trim(b) " "
}
