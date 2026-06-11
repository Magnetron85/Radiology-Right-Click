; ============================================================
; lib/calc/FollowUpDate.ahk -- Recommended follow-up date
; ------------------------------------------------------------
; Computes a follow-up / recommended-imaging date from a base
; (study) date plus an interval. The base date defaults to today,
; is pre-filled from a MM/DD/YYYY date found in the highlighted
; text when present, and is editable. Interval is a number plus a
; unit (days / weeks / months / years).
;
; Month and year arithmetic is CALENDAR-correct, not 30-day
; approximations: AutoHotkey's DateAdd only supports days, so
; months/years are added by component with end-of-month clamping
; (Jan 31 + 1 month -> Feb 28/29) and leap-year handling
; (Feb 29 + 1 year -> Feb 28).
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

FollowUpDate_Entry(input) {
    ShowFollowUpDateDialog(input)
    return ""
}

ShowFollowUpDateDialog(text := "") {
    ; Pre-fill base date from a date in the selection (study date), else today.
    baseYmd := ""
    if RegExMatch(text, "\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}", &dm)
        baseYmd := ParseDate(dm[0])

    ; Pre-fill interval from phrasing like "6 month follow-up" / "follow up in
    ; 1 year". Default 6 months when nothing is detected (modal radiology
    ; follow-up interval).
    n := 6, unitIdx := 3   ; 1=Days 2=Weeks 3=Months 4=Years
    if RegExMatch(text, "i)(\d+)\h*(day|week|month|year)s?", &im) {
        n := SafeInt(im[1], 6)
        u := StrLower(im[2])
        unitIdx := (u = "day") ? 1 : (u = "week") ? 2 : (u = "year") ? 4 : 3
    }

    form := RadsForm("Follow-up Date", 460)
    form.Header("Base (study) date")
    form.DateField("Base", "Study / base date:", baseYmd)
    form.Header("Interval")
    form.Numeric("N", "Amount:", n)
    form.Dropdown("Unit", "Unit:", ["Days", "Weeks", "Months", "Years"], unitIdx)
    form.SetSubmit(FollowUpDate_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

FollowUpDate_OnSubmit(v, form := "") {
    global g_LastSelectedText
    n := SafeInt(v.N, 0)
    if (n <= 0)
        return MakeResult({ impression: "Please enter an interval greater than zero.",
                            error: "Non-positive interval" })

    baseYmd := SubStr(v.Base, 1, 8)
    unit := _FUD_Unit(v.Unit)
    targetYmd := _FUD_AddInterval(baseYmd, n, unit)

    baseFmt   := FormatTime(baseYmd, "MM/dd/yyyy")
    targetFmt := FormatTime(targetYmd, "MM/dd/yyyy")
    unitWord  := (n = 1) ? RTrim(unit, "s") : unit   ; "1 month" vs "6 months"

    impression := "Recommended follow-up by " targetFmt
                . " (" n " " unitWord " from " baseFmt ")."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nBase date: " baseFmt
    method .= "`nInterval: " n " " unitWord
    method .= "`nFollow-up date: " targetFmt
    if (unit = "months" || unit = "years")
        method .= "`nCalendar arithmetic with end-of-month clamping (e.g. Jan 31 + 1 month = Feb 28/29)."

    return MakeResult({
        classification: "Follow-up date",
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [],
        echo:           g_LastSelectedText,
        ; Newline append: a scheduling line reads as its own sentence under
        ; the recommendation, not as a parenthetical to a measurement.
        paste:          "",
        pasteMode:      "newline"
    })
}

_FUD_Unit(label) {
    if InStr(label, "Day")
        return "days"
    if InStr(label, "Week")
        return "weeks"
    if InStr(label, "Year")
        return "years"
    return "months"
}

; Add n units to a YYYYMMDD date. days/weeks use DateAdd; months/years use
; calendar-component math with end-of-month clamping. Returns YYYYMMDD.
_FUD_AddInterval(ymd, n, unit) {
    base := SubStr(ymd, 1, 8)
    if (unit = "days")
        return SubStr(DateAdd(base "000000", n, "days"), 1, 8)
    if (unit = "weeks")
        return SubStr(DateAdd(base "000000", n * 7, "days"), 1, 8)
    if (unit = "years")
        return _FUD_AddMonths(base, n * 12)
    return _FUD_AddMonths(base, n)
}

_FUD_AddMonths(ymd, months) {
    y  := SubStr(ymd, 1, 4) + 0
    mo := SubStr(ymd, 5, 2) + 0
    d  := SubStr(ymd, 7, 2) + 0
    total := (y * 12 + (mo - 1)) + months
    ny  := total // 12
    nmo := Mod(total, 12) + 1
    maxd := _FUD_DaysInMonth(ny, nmo)
    if (d > maxd)
        d := maxd   ; clamp (e.g. Jan 31 + 1 month -> Feb 28/29)
    return Format("{:04d}{:02d}{:02d}", ny, nmo, d)
}

_FUD_DaysInMonth(y, mo) {
    static dm := [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if (mo = 2 && _FUD_IsLeap(y))
        return 29
    return dm[mo]
}

_FUD_IsLeap(y) {
    return (Mod(y, 4) = 0 && Mod(y, 100) != 0) || (Mod(y, 400) = 0)
}
