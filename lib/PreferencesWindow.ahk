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

; Single ordered list of all toggleable calculators (key, menu label). The
; preferences grid lays these out column-major; _SaveHandler iterates the
; same list, so adding a calculator here is the only edit needed.
global CALC_ALL := [
    ["ellipsoidVolume",       "Ellipsoid Volume"],
    ["bulletVolume",          "Bullet Volume"],
    ["psaDensity",            "PSA Density"],
    ["pregnancyDates",        "Pregnancy Dates"],
    ["menstrualPhase",        "Menstrual Phase"],
    ["adrenalWashout",        "Adrenal Washout"],
    ["thymusChemicalShift",   "Thymus Chemical Shift"],
    ["hepaticSteatosis",      "Hepatic Steatosis"],
    ["mriLiverIron",          "MRI Liver Iron"],
    ["statistics",            "Statistics"],
    ["numberRange",           "Number Range"],
    ["calciumScorePercentile","Calcium Score Pctile"],
    ["contrastPremedication", "Contrast Premed"],
    ["fleischnerCriteria",    "Fleischner Criteria"],
    ["rvlvRatio",             "RV/LV Ratio (PE)"],
    ["nascetCalculator",      "NASCET (carotid)"],
    ["ichVolume",             "ICH Volume (ABC/2)"],
    ["followUpDate",          "Follow-up Date"],
    ["bosniak",               "Bosniak"],
    ["gbPolyp",               "Gallbladder Polyp"],
    ["incidentalAdrenal",     "Incidental Adrenal"],
    ["incidentalThyroid",     "Incidental Thyroid"],
    ["kyotoIpmn",             "Kyoto IPMN"],
    ["lirads",                "LI-RADS (CT/MRI)"],
    ["lungRads",              "Lung-RADS"],
    ["oradsMri",              "O-RADS MRI"],
    ["oradsUs",               "O-RADS Ultrasound"],
    ["pirads",                "PI-RADS v2.1"],
    ["tirads",                "TI-RADS"],
    ["usLirads",              "US LI-RADS"]
]

; Window geometry. A wider window (740) lets every label render on one line
; -- the old 580 width forced long captions ("Show malignancy risk % ...")
; to wrap to two lines, and the fixed 24 px rows then overlapped the row
; below. Layout is driven by a running `y` that is advanced past each row's
; MEASURED bottom (see _PrefAdvance), so even a wrapped label can never
; overlap the next section.
ShowPreferencesWindow() {
    dark    := Prefs.Get("display", "darkMode", false)
    palette := ModernPalette(dark)

    WIN_W := 740
    marginX := 20
    leftX   := marginX
    contentW := WIN_W - marginX*2

    MouseGetPos(&mx, &my)
    work := GetWorkAreaAt(mx, my)

    g := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox +DPIScale", "RightClick Preferences")
    g.MarginX := marginX
    g.MarginY := 14
    g.BackColor := palette["bg"]
    g.SetFont("s10 c" palette["fg"], ModernFont())

    ; Shared mutable cursor so the helpers can advance a single `y`.
    st := { g: g, pal: palette, y: 16, leftX: leftX, w: contentW }

    ; ---- DISPLAY (two roomy columns, no wrapping) ----
    _PrefSection(st, "Display")
    dColW := (contentW - 24) // 2          ; ~338 each
    dCol2 := leftX + dColW + 24
    cbDark := _PrefCheck(g, palette, leftX, st.y, dColW, "DarkMode", "Dark mode", dark)
    cbCit  := _PrefCheck(g, palette, dCol2, st.y, dColW, "ShowCitations"
                       , "Show citations in output", Prefs.Get("display","showCitations",true))
    _PrefAdvance(st, [cbDark, cbCit])
    cbArt  := _PrefCheck(g, palette, leftX, st.y, dColW, "ShowArterialAge"
                       , "Show arterial age", Prefs.Get("display","showArterialAge",true))
    cbRisk := _PrefCheck(g, palette, dCol2, st.y, dColW, "ShowMalignancyRisk"
                       , "Show malignancy risk %", Prefs.Get("display","showMalignancyRisk",true))
    _PrefAdvance(st, [cbArt, cbRisk])
    cbMethod := _PrefCheck(g, palette, leftX, st.y, dColW, "ShowMethodology"
                         , "Show methodology in result", Prefs.Get("display","showMethodology",true))
    cbWidget := _PrefCheck(g, palette, dCol2, st.y, dColW, "LauncherWidget"
                         , "Show floating launcher widget", Prefs.Get("widget","enabled",false))
    _PrefAdvance(st, [cbMethod, cbWidget])

    ; ---- MENU + ACTIVATION (label above each dropdown) ----
    _PrefSection(st, "Menu and activation")
    cbSmart := _PrefCheck(g, palette, leftX, st.y, contentW, "SmartMatch"
        , "Suggest the best-matching calculator(s) from the highlighted text"
        , Prefs.Get("menu","smartMatch",true))
    _PrefAdvance(st, [cbSmart])
    tSort := g.Add("Text", "x" leftX " y" st.y " w" dColW " c" palette["fg"], "Menu sorting")
    tMod  := g.Add("Text", "x" dCol2 " y" st.y " w" dColW " c" palette["fg"], "Right-click activation")
    _PrefAdvance(st, [tSort, tMod], 2)

    sortChoices := ["grouped","alphabetical","frequency","none"]
    sortIdx := _PrefIndexOf(sortChoices, Prefs.Get("menu","sortingMethod","grouped"), 1)
    ddSort := g.Add("DropDownList"
        , "x" leftX " y" st.y " w" dColW " vSortChoice Choose" sortIdx
          . " Background" palette["bgAlt"] " c" palette["fg"]
        , sortChoices)
    modLabels := ["none (plain right-click)","Ctrl + right-click"
                , "Alt + right-click","Shift + right-click"]
    modKeys   := ["none","ctrl","alt","shift"]
    modIdx := _PrefIndexOf(modKeys, Prefs.Get("activation","modifier","none"), 1)
    ddMod := g.Add("DropDownList"
        , "x" dCol2 " y" st.y " w" dColW " vModifierChoice Choose" modIdx
          . " Background" palette["bgAlt"] " c" palette["fg"]
        , modLabels)
    _PrefAdvance(st, [ddSort, ddMod])

    ; ---- CALCULATORS (compact 3-column grid) ----
    _PrefSection(st, "Calculators shown in the menu")
    cCols := 3
    cColW := (contentW - (cCols-1)*16) // cCols    ; ~228 each
    cRows := Ceil(CALC_ALL.Length / cCols)
    gridTop := st.y
    rowH := 26
    lastRowCtls := []
    for i, entry in CALC_ALL {
        ci := Mod(i - 1, cCols)              ; 0-based column
        ri := (i - 1) // cCols               ; 0-based row
        cx := leftX + ci * (cColW + 16)
        cy := gridTop + ri * rowH
        cb := _PrefCheck(g, palette, cx, cy, cColW, "Calc_" entry[1], entry[2]
                       , Prefs.Get("calculations", entry[1], true))
        if (ri = cRows - 1)
            lastRowCtls.Push(cb)
    }
    st.y := gridTop + cRows * rowH + 6
    _PrefRule(st)

    ; ---- BUTTONS (left group + right group) ----
    btnY := st.y + 6
    btnH := 32
    btnRefs := _PrefButton(g, palette, leftX, btnY, 120, btnH, "References...")
    btnRefs.OnEvent("Click", _RefsHandler)
    btnApps := _PrefButton(g, palette, leftX + 130, btnY, 120, btnH, "Target Apps...")
    btnApps.OnEvent("Click", _AppsHandler)
    btnReset := _PrefButton(g, palette, leftX + 260, btnY, 130, btnH, "Restore Defaults")
    btnReset.OnEvent("Click", _ResetHandler.Bind(g))

    saveX := leftX + contentW - 94
    cancelX := saveX - 100
    btnCancel := _PrefButton(g, palette, cancelX, btnY, 92, btnH, "Cancel")
    btnCancel.OnEvent("Click", _CloseHandler.Bind(g))
    btnSave := g.Add("Button"
        , "Default x" saveX " y" btnY " w94 h" btnH
          . " +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Save")
    btnSave.OnEvent("Click", _SaveHandler.Bind(g))

    g.OnEvent("Close",  _CloseHandler.Bind(g))
    g.OnEvent("Escape", _CloseHandler.Bind(g))

    winH := btnY + btnH + 16
    ApplyModernChrome(g, dark)
    fit := FitWindow(WIN_W, winH, mx + 12, my + 12, work, 24)
    g.Show("x" fit.x " y" fit.y " w" fit.w " h" fit.h)
    WinActivate("ahk_id " g.Hwnd)
}

; ---- layout helpers --------------------------------------------------------

; Section header: accent caption over a hairline rule, then advance the
; cursor below the rule. Adds a little gap before the section (except first).
_PrefSection(st, label) {
    if (st.y > 20)
        st.y += 14
    st.g.SetFont("s10 Bold c" st.pal["accent"])
    t := st.g.Add("Text", "x" st.leftX " y" st.y " w" st.w " +Wrap", label)
    st.g.SetFont("s10 Norm c" st.pal["fg"])
    t.GetPos(, &ty, , &th)
    st.g.Add("Text", "x" st.leftX " y" (ty + th + 3) " w" st.w
        . " h1 +Background" st.pal["border"])
    st.y := ty + th + 3 + 8
}

_PrefRule(st) {
    st.g.Add("Text", "x" st.leftX " y" st.y " w" st.w
        . " h1 +Background" st.pal["border"])
    st.y += 1
}

_PrefCheck(g, pal, x, y, w, vname, label, on) {
    cb := g.Add("Checkbox"
        , "x" x " y" y " w" w " v" vname " c" pal["fg"]
        , label)
    cb.Value := on ? 1 : 0
    return cb
}

_PrefButton(g, pal, x, y, w, h, label) {
    return g.Add("Button"
        , "x" x " y" y " w" w " h" h " +Background" pal["btnBg"] " c" pal["btnFg"]
        , label)
}

; Advance st.y past the measured bottom of the row's controls + a gap, so a
; control that wrapped to two lines pushes the next row down instead of
; overlapping it.
_PrefAdvance(st, ctls, gap := 8) {
    maxBottom := st.y
    for c in ctls {
        c.GetPos(, &cy, , &ch)
        if (cy + ch > maxBottom)
            maxBottom := cy + ch
    }
    st.y := maxBottom + gap
}

_PrefIndexOf(arr, val, default) {
    for i, v in arr {
        if (v = val)
            return i
    }
    return default
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

    Prefs.Set("display", "darkMode",           !!saved.DarkMode)
    Prefs.Set("display", "showCitations",      !!saved.ShowCitations)
    Prefs.Set("display", "showArterialAge",    !!saved.ShowArterialAge)
    Prefs.Set("display", "showMalignancyRisk", !!saved.ShowMalignancyRisk)
    Prefs.Set("display", "showMethodology",    !!saved.ShowMethodology)

    for entry in CALC_ALL
        Prefs.Set("calculations", entry[1], !!saved.%"Calc_" entry[1]%)

    Prefs.Set("menu", "sortingMethod", saved.SortChoice)
    Prefs.Set("menu", "smartMatch", !!saved.SmartMatch)

    modLabels := ["none (plain right-click)","Ctrl + right-click"
                , "Alt + right-click","Shift + right-click"]
    modKeys   := ["none","ctrl","alt","shift"]
    for i, lbl in modLabels {
        if (saved.ModifierChoice = lbl) {
            Prefs.Set("activation", "modifier", modKeys[i])
            break
        }
    }

    Prefs.Set("widget", "enabled", !!saved.LauncherWidget)

    Prefs.Save()
    ApplyActivationHotkey()
    SetAppDarkMode(Prefs.Get("display", "darkMode", false))
    Launcher.Apply()   ; create/destroy the launcher widget to match the toggle
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
    WinActivate("ahk_id " g.Hwnd)
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
