; ============================================================
; lib/calc/Numerics.ahk -- Statistics + Number range
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

Statistics_Entry(input) {
    return CalcStatistics(input)
}

Range_Entry(input) {
    return CalcRange(input)
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
    needle := "(-?\d+(?:\.\d+)?)(?:\s*((?:cm/s|mm/s|m/s|km/h|mph|cm|mm|Hz|T|mg|m|ml|mL|cc|s|min|hr|days?|weeks?|months?|years?|g|ng|ng/ml|ng/mL|mmol/L|µmol/L|°F|°C)(?:/(?:day|week|month|year))?))?"
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
