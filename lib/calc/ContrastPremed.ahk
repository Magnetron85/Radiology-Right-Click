; ============================================================
; lib/calc/ContrastPremed.ahk -- ACR contrast premedication
; ------------------------------------------------------------
; Citation: ACR Manual on Contrast Media. American College of
; Radiology. https://www.acr.org/Clinical-Resources/Contrast-Manual
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\..\lib\Modern.ahk
#Include ..\..\lib\Util.ahk

ContrastPremed_Entry(input) {
    ShowContrastPremedDialog()
    return ""  ; handler shows its own UI; nothing for ShowResult
}

ShowContrastPremedDialog() {
    dark := Prefs.Get("display","darkMode", false)
    palette := ModernPalette(dark)
    MouseGetPos(&mx, &my)
    work := GetWorkAreaAt(mx, my)
    w := 320, h := 220
    pos := ClampToWorkArea(mx + 12, my + 12, w, h, work)

    defDT := _DefaultPremedDateTime()
    defDate := SubStr(defDT, 1, 8)
    defTime := SubStr(defDT, 9, 2) ":" SubStr(defDT, 11, 2)

    g := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Contrast Premedication")
    g.MarginX := 14, g.MarginY := 12
    g.BackColor := palette["bg"]
    g.SetFont("s10 c" palette["fg"], ModernFont())

    g.Add("Text", "x14 y10 w140", "Scan date:")
    dtScan := g.Add("DateTime", "x14 y28 w140 vScanDate", "yyyyMMdd")
    dtScan.Value := defDate

    g.Add("Text", "x166 y10 w120", "Scan time:")
    times := _BuildTimeList()
    chosen := 0
    for i, t in times
        if (t = defTime)
            chosen := i
    ddTime := g.Add("DropDownList"
        , "x166 y28 w120 vScanTime" (chosen ? " Choose" chosen : "")
        , times)

    g.Add("Text", "x14 y66 w280", "Protocol:")
    g.Add("DropDownList", "x14 y86 w280 vProtocol Choose1"
        , ["Prednisone (13-7-1 hour)","Methylprednisolone (12-2 hour)"])
    cbDiph := g.Add("Checkbox", "x14 y116 w280 vIncludeDiph Checked", "Include diphenhydramine")

    btnCalc := g.Add("Button"
        , "Default x14 y150 w120 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Calculate")
    btnCalc.OnEvent("Click", (*) => _ComputePremed(g))

    btnDose := g.Add("Button"
        , "x144 y150 w140 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Show Dosages")
    btnDose.OnEvent("Click", (*) => _ShowPremedDosages(g))

    g.OnEvent("Close",  (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    ApplyModernChrome(g, dark)
    g.Show("x" pos.x " y" pos.y " w" w " h" h)
}

_DefaultPremedDateTime() {
    base := DateAdd(A_Now, 13, "hours")
    hours := FormatTime(base, "HH") + 0
    mins  := FormatTime(base, "mm") + 0
    mins  := Ceil(mins / 5) * 5
    if (mins = 60) {
        mins := 0
        hours++
    }
    if (hours >= 24) {
        hours -= 24
        base := DateAdd(base, 1, "days")
    }
    return FormatTime(base, "yyyyMMdd") . Format("{:02d}{:02d}00", hours, mins)
}

_BuildTimeList() {
    list := []
    loop 24 {
        h := A_Index - 1
        loop 12 {
            m := (A_Index - 1) * 5
            list.Push(Format("{:02d}:{:02d}", h, m))
        }
    }
    return list
}

_ComputePremed(g) {
    saved := g.Submit(false)
    if !RegExMatch(saved.ScanTime, "^(\d{2}):(\d{2})$", &tm) {
        MsgBox("Invalid scan time.", "Contrast Premedication", "Iconx")
        return
    }
    dt := SubStr(saved.ScanDate, 1, 8) . tm[1] . tm[2] . "00"
    out := "Scan time: " FormatTime(dt, "MM/dd/yyyy hh:mm tt")
        . " Contrast Premedication Schedule:`n`n"

    proto := InStr(saved.Protocol, "Prednisone") ? 1 : 2
    diph  := !!saved.IncludeDiph
    if (proto = 1) {
        out .= _FormatPremedStep(dt, -13, "13", 1, diph)
        out .= _FormatPremedStep(dt, -7,  "7",  1, diph)
        out .= _FormatPremedStep(dt, -1,  "1",  1, diph)
    } else {
        out .= _FormatPremedStep(dt, -12, "12", 2, diph)
        out .= _FormatPremedStep(dt, -2,  "2",  2, diph)
    }

    out .= "`nNote: Premedication regimens less than 4-5 hours in duration (oral or IV) have not been shown to be effective.`n"
    out .= "If a patient is unable to take oral medication, 200 mg hydrocortisone IV may be substituted for each dose of oral prednisone in the 13-7-1 premedication regimen.`n"
    if Prefs.Get("display","showCitations", true)
        out .= "`nCitation: ACR Manual on Contrast Media. American College of Radiology. https://www.acr.org/Clinical-Resources/Contrast-Manual`n"

    ShowResult(out)
}

_FormatPremedStep(dt, hoursOffset, label, protocol, diph) {
    when := DateAdd(dt, hoursOffset, "hours")
    formatted := FormatTime(when, "MM/dd/yyyy hh:mm tt")
    med := _PremedMedication(label, protocol, diph)
    return label " hours before (" formatted "):`n" med "`n`n"
}

_PremedMedication(label, protocol, diph) {
    if (protocol = 1) {
        m := "- Prednisone 50 mg PO"
        if (label = "1" && diph)
            m .= "`n- Diphenhydramine 50 mg IV, IM, or PO"
        return m
    }
    m := "- Methylprednisolone 32 mg PO"
    if (label = "2" && diph)
        m .= "`n- Diphenhydramine 50 mg IV, IM, or PO"
    return m
}

_ShowPremedDosages(g) {
    saved := g.Submit(false)
    proto := InStr(saved.Protocol, "Prednisone") ? 1 : 2
    diph  := !!saved.IncludeDiph
    out := "Contrast Premedication Dosages:`n`n"
    if (proto = 1) {
        out .= "Prednisone-based Protocol (13-7-1 hour):`n`n"
        out .= "13 hours before:`n- Prednisone 50 mg PO`n`n"
        out .= "7 hours before:`n- Prednisone 50 mg PO`n`n"
        out .= "1 hour before:`n- Prednisone 50 mg PO`n"
        if diph
            out .= "- Diphenhydramine 50 mg IV, IM, or PO`n"
    } else {
        out .= "Methylprednisolone-based Protocol (12-2 hour):`n`n"
        out .= "12 hours before:`n- Methylprednisolone 32 mg PO`n`n"
        out .= "2 hours before:`n- Methylprednisolone 32 mg PO`n"
        if diph
            out .= "- Diphenhydramine 50 mg IV, IM, or PO`n"
    }
    out .= "`nNotes:`n"
    out .= "- Premedication regimens less than 4-5 hours in duration (oral or IV) have not been shown to be effective.`n"
    out .= "- If a patient is unable to take oral medication, 200 mg hydrocortisone IV may be substituted for each dose of oral prednisone in the 13-7-1 premedication regimen.`n"
    out .= "- Diphenhydramine is considered optional. If a patient is allergic to diphenhydramine, an alternate anti-histamine without cross-reactivity may be considered, or the anti-histamine may be omitted.`n"
    out .= "- These dosages are based on the ACR Manual on Contrast Media. Please consult with a healthcare professional for patient-specific recommendations.`n"
    if Prefs.Get("display","showCitations", true)
        out .= "`nCitation: ACR Manual on Contrast Media. American College of Radiology. https://www.acr.org/Clinical-Resources/Contrast-Manual`n"
    ShowResult(out)
}
