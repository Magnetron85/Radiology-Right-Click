; ============================================================
; lib/calc/PregnancyDates.ahk -- Pregnancy + Menstrual phase
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

PregnancyDates_Entry(input) {
    ShowPregnancyDialog(input)
    return ""
}

MenstrualPhase_Entry(input) {
    ShowMenstrualDialog(input)
    return ""
}

ShowPregnancyDialog(text := "") {
    lmpYmd := ""
    if (lmpStr := TextScan.LMP(text)) {
        lmpYmd := ParseDate(lmpStr)
    }
    weeks := 0, days := 0
    if RegExMatch(text, "i)(\d+)\s*(?:weeks?|w)\s*(?:and|&|,|-)?\s*(\d+)?\s*(?:days?|d)?", &m) {
        weeks := SafeInt(m[1], 0)
        days  := (m.Count >= 2) ? SafeInt(m[2], 0) : 0
    }

    ; Progressive disclosure: the mandatory Mode dropdown gates which input
    ; section applies. Only the section matching the selected mode is shown;
    ; the other is hidden (collapsed), and Pregnancy_OnSubmit reads only the
    ; fields for the selected mode, so the hidden defaults are never used.
    form := RadsForm("Pregnancy Dates", 480)
    form.Header("Input mode")
    form.Dropdown("Mode", "Mode:"
        , ["From LMP date"
        ,  "From GA (weeks + days)"], lmpYmd != "" ? 1 : (weeks > 0 ? 2 : 1))

    form.Header("LMP date")
    form.DateField("LMP", "Last menstrual period:", lmpYmd)

    form.Header("Gestational age")
    form.Numeric("Weeks", "Weeks:", weeks)
    form.Numeric("Days",  "Days:",  days)

    form.OnChange("Mode", _PRG_UpdateForm)
    _PRG_UpdateForm(form)

    form.SetSubmit(Pregnancy_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

; Show the input section matching the selected mode; hide (collapse) the
; other one. The form reflows and the window resizes.
_PRG_UpdateForm(frm) {
    isLMP := !!InStr(frm.byName["Mode"].ctl.Text, "LMP")
    frm.SetSectionVisible("LMP date", isLMP)
    frm.SetSectionVisible("Gestational age", !isLMP)
}

Pregnancy_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if InStr(v.Mode, "LMP") {
        lmpDate := FormatTime(v.LMP, "MM/dd/yyyy")
        body := CalcPregnancyDates("LMP: " lmpDate)
    } else {
        weeks := SafeInt(v.Weeks, 0)
        days  := SafeInt(v.Days, 0)
        if (weeks = 0 && days = 0)
            return MakeResult({ impression: "Please enter gestational age in weeks and/or days.",
                                error: "Missing GA" })
        body := CalcPregnancyDates("GA: " weeks " weeks and " days " days as of today")
    }

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`n" body

    ; Compose a single-line impression from the labeled body
    lmp := "", edd := "", ga := ""
    if RegExMatch(body, "i)LMP[:]?\s*(\S+)", &m)
        lmp := m[1]
    if RegExMatch(body, "i)Estimated Delivery Date[:]?\s*(\S+)", &m)
        edd := m[1]
    if RegExMatch(body, "i)Current Gestational Age[:]?\s*([^\r\n]+)", &m)
        ga := Trim(m[1])

    impression := body
    if (lmp != "" && edd != "" && ga != "")
        impression := "LMP " lmp ", EDD " edd ", current GA " ga "."

    return MakeResult({
        classification: "Pregnancy dates",
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "Naegele's rule for estimated delivery date "
                                . "(LMP + 280 days). Standard obstetric formula.",
                           url:  "" }],
        echo:           g_LastSelectedText
    })
}

ShowMenstrualDialog(text := "") {
    lmpYmd := ""
    if (lmpStr := TextScan.LMP(text))
        lmpYmd := ParseDate(lmpStr)

    form := RadsForm("Menstrual Phase", 420)
    form.Header("Last menstrual period")
    form.DateField("LMP", "LMP date:", lmpYmd)
    form.SetSubmit(Menstrual_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

Menstrual_OnSubmit(v, form := "") {
    global g_LastSelectedText
    lmpDate := FormatTime(v.LMP, "MM/dd/yyyy")
    body := CalcMenstrualPhase("LMP: " lmpDate)

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`n" body

    ; Compose a single-line impression by collapsing the labeled body.
    lmp := "", day := "", phase := "", thick := ""
    if RegExMatch(body, "i)LMP[:]?\s*(\S+)", &m)
        lmp := m[1]
    if RegExMatch(body, "i)Current Cycle Day[:]?\s*(\d+/\d+)", &m)
        day := m[1]
    if RegExMatch(body, "i)`n([A-Z][^`r`n]*Phase|Ovulation)`n", &m)
        phase := m[1]
    if RegExMatch(body, "i)Expected endometrial[^:]*:\s*([^\r\n]+)", &m)
        thick := Trim(m[1])

    impression := body
    if (lmp != "" && day != "" && phase != "") {
        impression := "LMP " lmp ", cycle day " day " (" phase ")"
        if (thick != "")
            impression .= "; expected endometrium " thick
        impression .= "."
    }

    return MakeResult({
        classification: "Menstrual phase",
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [],
        echo:           g_LastSelectedText
    })
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
