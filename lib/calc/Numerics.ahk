; ============================================================
; lib/calc/Numerics.ahk -- Statistics + Number range
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

Statistics_Entry(input) {
    ShowStatisticsDialog(input)
    return ""
}

Range_Entry(input) {
    ShowRangeDialog(input)
    return ""
}

ShowStatisticsDialog(text := "") {
    form := RadsForm("Statistics", 520)
    form.Header("Numbers")
    form.Note("Paste or type the numbers (comma- or newline-separated). Slice / sample / observation labels are ignored.")
    form.TextArea("Nums", "Numbers:", 120, text)
    form.SetSubmit(Statistics_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

Statistics_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (Trim(v.Nums) = "")
        return MakeResult({ impression: "Please enter at least one number.",
                            error: "Empty input" })

    nums := ExtractNumbers(v.Nums)
    if (nums.Length = 0)
        return MakeResult({ impression: "No numbers found in input.",
                            error: "No numbers parsed" })

    ; Impression is a single prose sentence summarizing the headline stats.
    ; The full label list (Count/Sum/Mean/Median/...) goes in methodology.
    impression := "Statistics (n=" nums.Length "): mean " Round(Mean(nums), 1)
                . ", median " Round(Median(nums), 1)
                . ", range " Round(Min(nums*), 1) "-" Round(Max(nums*), 1)
    if (nums.Length >= 9)
        impression .= ", SD " Round(StdDev(nums), 1)
    impression .= "."

    return MakeResult({
        classification: "Statistics",
        impression:     impression,
        recommendation: "",
        methodology:    CalcStatistics(v.Nums),
        citations:      [],
        echo:           g_LastSelectedText,
        paste:          "n " nums.Length ", mean " Round(Mean(nums), 1)
                        . ", range " Round(Min(nums*), 1) "-" Round(Max(nums*), 1),
        pasteMode:      ""
    })
}

ShowRangeDialog(text := "") {
    form := RadsForm("Number Range", 520)
    form.Header("Numbers")
    form.TextArea("Nums", "Numbers (optionally with units):", 100, text)
    form.SetSubmit(Range_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

Range_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (Trim(v.Nums) = "")
        return MakeResult({ impression: "Please enter at least one number.",
                            error: "Empty input" })
    body := CalcRange(v.Nums)
    return MakeResult({
        classification: "Range",
        impression:     body,
        recommendation: "",
        methodology:    "",
        citations:      [],
        echo:           g_LastSelectedText,
        paste:          InStr(body, "No numbers found") ? "" : "range " StrReplace(body, " - ", "-"),
        pasteMode:      ""
    })
}

CalcStatistics(input) {
    nums := ExtractNumbers(input)
    if (nums.Length = 0)
        return "No numbers found in text."

    out := "Statistics:`n"
    out .= "Count: " nums.Length "`n"
    out .= "Sum: "  Round(Sum(nums), 1) "`n"
    out .= "Mean: " Round(Mean(nums), 1) "`n"
    out .= "Median: " Round(Median(nums), 1) "`n"
    out .= "Min: " Round(Min(nums*), 1) "`n"
    out .= "Max: " Round(Max(nums*), 1) "`n"

    if (nums.Length >= 9) {
        q1 := Round(Quartile(nums, 0.25), 1)
        q3 := Round(Quartile(nums, 0.75), 1)
        med := Round(Median(nums), 1)
        iqr := q3 - q1
        out .= "Q1: " q1 "`n"
        out .= "Q3: " q3 "`n"
        out .= "IQR: " Round(iqr, 1) "`n"
        out .= "IQR/Median: " (med != 0 ? Round(iqr / med, 2) : "N/A") "`n"
        out .= "Standard Deviation: " Round(StdDev(nums), 1) "`n"
    }
    return out
}

CalcRange(input) {
    nums := []
    unit := ""
    ; Alternation is ordered longest-first and followed by a letter
    ; lookahead -- with bare "m" / "s" listed before "ml" / "min" the
    ; regex matched the one-letter prefix ("5 ml" reported unit "m").
    needle := "(-?\d+(?:\.\d+)?)(?:\s*((?:cm/s|mm/s|m/s|km/h|mph|ng/ml|ng/mL|mmol/L|µmol/L|°F|°C|min|hr|days?|weeks?|months?|years?|cm|mm|Hz|mg|ng|ml|mL|cc|g|T|m|s)(?:/(?:day|week|month|year))?)(?![A-Za-z]))?"
    pos := 1
    while (pos := RegExMatch(input, needle, &m, pos)) {
        nums.Push(m[1] + 0)
        if (m.Count >= 2 && m[2] != "" && unit = "")
            unit := m[2]
        pos += m.Len[0] ? m.Len[0] : 1
    }
    if (nums.Length = 0)
        return "No numbers found."
    out := Round(Min(nums*), 1) " - " Round(Max(nums*), 1)
    if (unit != "")
        out .= " " unit
    return out
}

; --- helpers ---------------------------------------------------------------

Sum(arr) {
    t := 0
    for v in arr
        t += v
    return t
}

Mean(arr) {
    return Sum(arr) / arr.Length
}

SortNums(arr) {
    out := []
    for v in arr
        out.Push(v)
    ; insertion sort (arrays in v2 don't have a numeric .Sort)
    n := out.Length
    i := 2
    while (i <= n) {
        cur := out[i]
        j := i - 1
        while (j >= 1 && out[j] > cur) {
            out[j+1] := out[j]
            j--
        }
        out[j+1] := cur
        i++
    }
    return out
}

Median(arr) {
    s := SortNums(arr)
    n := s.Length
    if (n = 0)
        return 0
    if (Mod(n, 2) = 0)
        return (s[n//2] + s[n//2 + 1]) / 2
    return s[Floor(n/2) + 1]
}

Quartile(arr, percentile) {
    s := SortNums(arr)
    n := s.Length
    pos := (n - 1) * percentile + 1
    lo := Floor(pos), hi := Ceil(pos)
    if (lo = hi)
        return s[lo]
    return s[lo] + (pos - lo) * (s[hi] - s[lo])
}

StdDev(arr) {
    if (arr.Length < 2)
        return 0
    m := Mean(arr)
    ssq := 0
    for v in arr {
        diff := v - m
        ssq += diff * diff
    }
    return Sqrt(ssq / (arr.Length - 1))
}
