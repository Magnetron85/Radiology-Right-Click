; ============================================================
; lib/calc/PregnancyDates.ahk -- Pregnancy + Menstrual phase
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

PregnancyDates_Entry(input) {
    return CalcPregnancyDates(input)
}

MenstrualPhase_Entry(input) {
    return CalcMenstrualPhase(input)
}

CalcPregnancyDates(input) {
    if RegExMatch(input
        , "i)(?:LMP|Last\s*Menstrual\s*Period).*?(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})"
        , &lmp) {
        d := ParseDate(lmp[1])
        if (d = "")
            return "Invalid LMP date. Please use MM/DD/YYYY or DD/MM/YYYY."
        return _DatesFromLMP(d)
    }
    if RegExMatch(input
        , "i)(\d+)(?:\s*(?:weeks?|w))?\s*(?:and|&|,|-)?\s*(\d+)?(?:\s*(?:days?|d))?(?:.*?(?:as of|on)\s+(today|\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}))?"
        , &ga) {
        weeks := ga[1] + 0
        days  := (ga.Count >= 2 && ga[2] != "") ? ga[2] + 0 : 0
        refRaw := ga.Count >= 3 ? ga[3] : ""
        ref := (refRaw = "" || refRaw = "today") ? A_Now : ParseDate(refRaw)
        if (ref = "")
            return "Invalid reference date."
        return _DatesFromGA(weeks, days, ref)
    }
    return "Invalid format for pregnancy date calculation.`n"
        . "Example:`nLMP: 01/15/2023`nor`nGA: 12 weeks and 3 days as of today"
}

CalcMenstrualPhase(input) {
    if !RegExMatch(input
        , "i)(?:LMP|Last\s*Menstrual\s*Period)\s*:?\s*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})"
        , &m)
        return "Invalid format for menstrual phase calculation.`nExample: LMP: 05/01/2023"
    d := ParseDate(m[1])
    if (d = "")
        return "Invalid LMP date. Use MM/DD/YYYY or DD/MM/YYYY."
    return _MenstrualPhase(d)
}

_DatesFromLMP(lmpDate) {
    lmpFmt := FormatTime(lmpDate, "MM/dd/yyyy")
    eddDate := DateAddDays(lmpDate, 280)
    eddFmt := FormatTime(eddDate, "MM/dd/yyyy")
    today := SubStr(A_Now, 1, 8)
    days  := DaysBetween(lmpDate, today)
    return "LMP: " lmpFmt
        . "`nEstimated Delivery Date: " eddFmt
        . "`nCurrent Gestational Age: " _FormatGA(days, eddDate, today)
}

_DatesFromGA(weeks, days, refDate) {
    refStr := SubStr(refDate, 1, 8)
    ageDays := weeks * 7 + days
    lmpDate := DateAddDays(refStr, -ageDays)
    lmpFmt := FormatTime(lmpDate, "MM/dd/yyyy")
    eddDate := DateAddDays(lmpDate, 280)
    eddFmt := FormatTime(eddDate, "MM/dd/yyyy")
    today := SubStr(A_Now, 1, 8)
    curDays := DaysBetween(lmpDate, today)
    refFmt := FormatTime(refStr, "MM/dd/yyyy")
    return "LMP: " lmpFmt
        . "`nEstimated Delivery Date: " eddFmt
        . "`nGestational Age as of " refFmt ": " weeks " weeks " days " days"
        . "`nCurrent Gestational Age: " _FormatGA(curDays, eddDate, today)
}

; A current GA past 42 weeks is clinically meaningless -- the patient
; delivered. Show how far past EDD instead so a stale LMP is obvious.
_FormatGA(daysSinceLMP, eddDate, today) {
    if (daysSinceLMP < 0)
        return "(LMP is in the future)"
    pastEDD := DaysBetween(eddDate, today)
    ga := Floor(daysSinceLMP/7) " weeks " Mod(daysSinceLMP, 7) " days"
    if (pastEDD > 0)
        return ga " -- past EDD by " pastEDD " day" (pastEDD = 1 ? "" : "s") " (LMP likely stale)"
    return ga
}

_MenstrualPhase(lmpDate) {
    today := SubStr(A_Now, 1, 8)
    daysSince := DaysBetween(lmpDate, today)
    cycleDay := Mod(daysSince, 28) + 1
    lmpFmt := FormatTime(lmpDate, "MM/dd/yyyy")
    out := "LMP: " lmpFmt "`nCurrent Cycle Day: " cycleDay "/28`n`n"
    if (cycleDay >= 1 && cycleDay <= 5)
        out .= "Menstrual Phase`nExpected endometrial stripe thickness: 1-4 mm"
    else if (cycleDay >= 6 && cycleDay <= 13)
        out .= "Early Proliferative Phase`nExpected endometrial stripe thickness: 5-7 mm"
    else if (cycleDay = 14)
        out .= "Ovulation`nExpected endometrial appearance: Trilaminar, approximately 11 mm"
    else if (cycleDay >= 15 && cycleDay <= 28)
        out .= "Secretory Phase`nExpected endometrial stripe thickness: 7-16 mm"
    else
        out .= "Error: Invalid cycle day calculated"
    return out
}
