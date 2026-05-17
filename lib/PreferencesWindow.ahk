; ============================================================
; lib/PreferencesWindow.ahk -- Settings GUI
; ------------------------------------------------------------
; Modern flat layout, two-column calculator grid so the window
; fits on a 1366x768 laptop. Total height fits in <650 px;
; FitWindow() further caps it if the work area is smaller still.
; ============================================================

#Requires AutoHotkey v2.0
#Include Modern.ahk
#Include Util.ahk

global CALC_COL1 := [
    ["ellipsoidVolume",       "Ellipsoid Volume"],
    ["bulletVolume",          "Bullet Volume"],
    ["psaDensity",            "PSA Density"],
    ["pregnancyDates",        "Pregnancy Dates"],
    ["menstrualPhase",        "Menstrual Phase"],
    ["adrenalWashout",        "Adrenal Washout"],
    ["thymusChemicalShift",   "Thymus Chemical Shift"],
    ["hepaticSteatosis",      "Hepatic Steatosis"]
]
global CALC_COL2 := [
    ["mriLiverIron",          "MRI Liver Iron Content"],
    ["statistics",            "Statistics"],
    ["numberRange",           "Number Range"],
    ["calciumScorePercentile","Calcium Score Percentile"],
    ["contrastPremedication", "Contrast Premedication"],
    ["fleischnerCriteria",    "Fleischner Criteria"],
    ["nascetCalculator",      "NASCET"]
]

ShowPreferencesWindow() {
    dark    := Prefs.Get("display", "darkMode", false)
    palette := ModernPalette(dark)

    MouseGetPos(&mx, &my)
    work := GetWorkAreaAt(mx, my)
    desiredW := 580
    desiredH := 560   ; dropped from 640 after removing the custom-order field
    fit := FitWindow(desiredW, desiredH, mx + 12, my + 12, work, 24)

    g := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "RightClick Preferences")
    g.MarginX := 18
    g.MarginY := 16
    g.BackColor := palette["bg"]
    g.SetFont("s10 c" palette["fg"], ModernFont())

    col1X := 24
    col2X := 296
    cbW   := 260
    y     := 16

    ; --- DISPLAY
    _AddHeader(g, palette, col1X, y, "Display")
    y += 26
    cbDark := g.Add("Checkbox", "x" col1X " y" y " w" cbW " vDarkMode", "Dark mode")
    cbDark.Value := dark ? 1 : 0
    cbCit  := g.Add("Checkbox", "x" col2X " y" y " w" cbW " vShowCitations"
        , "Show citations in output")
    cbCit.Value := Prefs.Get("display", "showCitations", true) ? 1 : 0
    y += 24
    cbArt := g.Add("Checkbox", "x" col1X " y" y " w" cbW " vShowArterialAge"
        , "Show arterial age (calcium score)")
    cbArt.Value := Prefs.Get("display", "showArterialAge", true) ? 1 : 0
    y += 32

    ; --- CALCULATORS (two-column grid)
    _AddHeader(g, palette, col1X, y, "Calculators (show in menu)")
    y += 26
    gridStartY := y
    for i, entry in CALC_COL1 {
        cb := g.Add("Checkbox"
            , "x" col1X " y" (gridStartY + (i-1)*24) " w" cbW " vCalc_" entry[1]
            , entry[2])
        cb.Value := Prefs.Get("calculations", entry[1], true) ? 1 : 0
    }
    for i, entry in CALC_COL2 {
        cb := g.Add("Checkbox"
            , "x" col2X " y" (gridStartY + (i-1)*24) " w" cbW " vCalc_" entry[1]
            , entry[2])
        cb.Value := Prefs.Get("calculations", entry[1], true) ? 1 : 0
    }
    rows := Max(CALC_COL1.Length, CALC_COL2.Length)
    y := gridStartY + rows * 24 + 12

    ; --- MENU SORT (left)
    _AddHeader(g, palette, col1X, y, "Menu sorting")
    _AddHeader(g, palette, col2X, y, "Right-click modifier")
    y += 26

    sortChoices := ["alphabetical","frequency","none"]
    sortVal := Prefs.Get("menu", "sortingMethod", "alphabetical")
    sortIdx := 1
    for i, v in sortChoices {
        if (v = sortVal)
            sortIdx := i
    }
    g.Add("DropDownList", "x" col1X " y" y " w200 vSortChoice Choose" sortIdx
        , sortChoices)

    ; --- ACTIVATION MODIFIER (right)
    modLabels := ["none (plain right-click)","Ctrl + right-click"
                , "Alt + right-click","Shift + right-click"]
    modKeys   := ["none","ctrl","alt","shift"]
    modVal := Prefs.Get("activation", "modifier", "none")
    modIdx := 1
    for i, k in modKeys {
        if (k = modVal)
            modIdx := i
    }
    g.Add("DropDownList", "x" col2X " y" y " w240 vModifierChoice Choose" modIdx
        , modLabels)
    y += 48

    ; --- BUTTONS (bottom row)
    btnRefs := g.Add("Button"
        , "x" col1X " y" y " w110 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "References...")
    btnRefs.OnEvent("Click", _RefsHandler)

    btnApps := g.Add("Button"
        , "x" (col1X + 120) " y" y " w110 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Target Apps...")
    btnApps.OnEvent("Click", _AppsHandler)

    btnReset := g.Add("Button"
        , "x" (col1X + 240) " y" y " w120 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Restore Defaults")
    btnReset.OnEvent("Click", _ResetHandler.Bind(g))

    btnSave := g.Add("Button"
        , "Default x" (col2X + 166) " y" y " w90 h32 +Background"
          . palette["btnBg"] " c" palette["btnFg"]
        , "Save")
    btnSave.OnEvent("Click", _SaveHandler.Bind(g))

    g.OnEvent("Close",  _CloseHandler.Bind(g))
    g.OnEvent("Escape", _CloseHandler.Bind(g))

    ApplyModernChrome(g, dark)
    finalH := Min(y + 60, fit.h)  ; cap at fitted height
    g.Show("x" fit.x " y" fit.y " w" fit.w " h" finalH)
}

_AddHeader(g, palette, x, y, label) {
    g.SetFont("s10 Bold c" palette["fg"])
    g.Add("Text", "x" x " y" y " w400", label)
    g.SetFont("s10 Norm c" palette["fg"])
}

_RefsHandler(*) {
    ShowReferencesManager()
}

_AppsHandler(*) {
    ShowTargetAppsEditor()
}

_CloseHandler(g, *) {
    g.Destroy()
}

_SaveHandler(g, *) {
    saved := g.Submit(false)

    Prefs.Set("display", "darkMode",        !!saved.DarkMode)
    Prefs.Set("display", "showCitations",   !!saved.ShowCitations)
    Prefs.Set("display", "showArterialAge", !!saved.ShowArterialAge)

    for entry in CALC_COL1
        Prefs.Set("calculations", entry[1], !!saved.%"Calc_" entry[1]%)
    for entry in CALC_COL2
        Prefs.Set("calculations", entry[1], !!saved.%"Calc_" entry[1]%)

    Prefs.Set("menu", "sortingMethod", saved.SortChoice)

    modLabels := ["none (plain right-click)","Ctrl + right-click"
                , "Alt + right-click","Shift + right-click"]
    modKeys   := ["none","ctrl","alt","shift"]
    for i, lbl in modLabels {
        if (saved.ModifierChoice = lbl) {
            Prefs.Set("activation", "modifier", modKeys[i])
            break
        }
    }

    Prefs.Save()
    ApplyActivationHotkey()
    SetAppDarkMode(Prefs.Get("display", "darkMode", false))
    g.Destroy()
}

_ResetHandler(g, *) {
    if (MsgBox("Restore all preferences to their defaults? Your saved references and frequency data will be kept.", "Restore Defaults", "YesNo Iconi") != "Yes")
        return
    Prefs.RestoreDefaults()
    ApplyActivationHotkey()
    SetAppDarkMode(Prefs.Get("display", "darkMode", false))
    g.Destroy()
    ShowPreferencesWindow()
}

; ---- target apps editor ----------------------------------------------------

ShowTargetAppsEditor() {
    dark := Prefs.Get("display","darkMode", false)
    palette := ModernPalette(dark)
    MouseGetPos(&mx, &my)
    work := GetWorkAreaAt(mx, my)
    fit := FitWindow(500, 400, mx + 12, my + 12, work, 24)

    g := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Manage Target Apps")
    g.MarginX := 16
    g.MarginY := 14
    g.BackColor := palette["bg"]
    g.SetFont("s10 c" palette["fg"], ModernFont())

    g.Add("Text", "x16 y14 w468"
        , "Right-click is intercepted only in windows matching these entries."
        . "`nOne per line. Format: class:WindowClass  OR  exe:Process.exe")

    apps := Prefs.Get("activation","targetApps","")
    if !(apps is Array)
        apps := []
    text := ""
    for app in apps {
        if (app is Map && app.Has("type") && app.Has("value"))
            text .= app["type"] ":" app["value"] "`n"
    }

    g.Add("Edit"
        , "x16 y66 w468 h240 vTargetApps -Wrap +HScroll Background"
          . palette["bgAlt"] " c" palette["fg"]
        , RTrim(text, "`n"))

    btnSave := g.Add("Button"
        , "Default x308 y320 w80 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Save")
    btnSave.OnEvent("Click", _SaveTargetApps.Bind(g))

    btnCancel := g.Add("Button"
        , "x398 y320 w86 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Cancel")
    btnCancel.OnEvent("Click", (*) => g.Destroy())

    g.OnEvent("Close",  (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    ApplyModernChrome(g, dark)
    g.Show("x" fit.x " y" fit.y " w" fit.w " h" fit.h)
}

_SaveTargetApps(g, *) {
    saved := g.Submit(false)
    apps := []
    bad := []
    for _, line in StrSplit(saved.TargetApps, "`n", "`r") {
        line := Trim(line)
        if (line = "")
            continue
        if !RegExMatch(line, "i)^\s*(class|exe)\s*:\s*(.+?)\s*$", &m) {
            bad.Push(line)
            continue
        }
        apps.Push(Map("type", StrLower(m[1]), "value", m[2]))
    }
    if (bad.Length > 0) {
        if (MsgBox("These lines were not in 'class:Name' or 'exe:Name' format and will be dropped:`n`n"
                . JoinArr(bad, "`n") "`n`nSave the rest?"
                , "Manage Target Apps", "YesNo Iconi") != "Yes")
            return
    }
    if (apps.Length = 0) {
        MsgBox("At least one target app is required.", "Manage Target Apps", "Iconx")
        return
    }
    Prefs.Set("activation", "targetApps", apps)
    Prefs.Save()
    g.Destroy()
}
