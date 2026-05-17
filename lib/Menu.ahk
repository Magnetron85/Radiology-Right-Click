; ============================================================
; lib/Menu.ahk -- Right-click context menu
; ------------------------------------------------------------
; Registry of calculators (title, freq-key, pref-key, handler);
; the menu is rebuilt each invocation so toggles and sort
; method changes are picked up immediately.
;
; Sorting:
;   "alphabetical" -> case-insensitive by title (default)
;   "frequency"    -> descending by usage count, alpha as tiebreak
;   "none"         -> registry order
; ============================================================

#Requires AutoHotkey v2.0
#Include Util.ahk

global g_LastSelectedText := ""

; ---- registry -- handlers live in lib/calc/*.ahk ---------------------------

CalculatorRegistry() {
    return [
        ; non-toggleable measurement helpers
        { title: "Compare Measurement Sizes", cmd: "CompareNoduleSizes",
          pref: "", handler: CompareSizes_Entry },
        { title: "Sort Measurement Sizes",    cmd: "SortNoduleSizes",
          pref: "", handler: SortSizes_Entry },

        ; toggleable calculators
        { title: "Calculate Calcium Score Percentile", cmd: "CalculateCalciumScorePercentile",
          pref: "calciumScorePercentile", handler: CalciumScore_Entry },
        { title: "Calculate Ellipsoid Volume", cmd: "CalculateEllipsoidVolume",
          pref: "ellipsoidVolume",        handler: EllipsoidVolume_Entry },
        { title: "Calculate Bullet Volume",    cmd: "CalculateBulletVolume",
          pref: "bulletVolume",           handler: BulletVolume_Entry },
        { title: "Calculate PSA Density",      cmd: "CalculatePSADensity",
          pref: "psaDensity",             handler: PSADensity_Entry },
        { title: "Calculate Pregnancy Dates",  cmd: "CalculatePregnancyDates",
          pref: "pregnancyDates",         handler: PregnancyDates_Entry },
        { title: "Calculate Menstrual Phase",  cmd: "CalculateMenstrualPhase",
          pref: "menstrualPhase",         handler: MenstrualPhase_Entry },
        { title: "Calculate Adrenal Washout",  cmd: "CalculateAdrenalWashout",
          pref: "adrenalWashout",         handler: AdrenalWashout_Entry },
        { title: "Calculate Thymus Chemical Shift", cmd: "CalculateThymusChemicalShift",
          pref: "thymusChemicalShift",    handler: ThymusChemicalShift_Entry },
        { title: "Calculate Hepatic Steatosis", cmd: "CalculateHepaticSteatosis",
          pref: "hepaticSteatosis",       handler: HepaticSteatosis_Entry },
        { title: "MRI Liver Iron Content",     cmd: "CalculateIronContent",
          pref: "mriLiverIron",           handler: LiverIron_Entry },
        { title: "Calculate Statistics",       cmd: "Statistics",
          pref: "statistics",             handler: Statistics_Entry },
        { title: "Calculate Number Range",     cmd: "Range",
          pref: "numberRange",            handler: Range_Entry },
        { title: "Calculate Contrast Premedication", cmd: "CalculateContrastPremedication",
          pref: "contrastPremedication",  handler: ContrastPremed_Entry },
        { title: "Calculate Fleischner Criteria", cmd: "CalculateFleischnerCriteria",
          pref: "fleischnerCriteria",     handler: Fleischner_Entry },
        { title: "Calculate NASCET",           cmd: "CalculateNASCET",
          pref: "nascetCalculator",       handler: NASCET_Entry }
    ]
}

; ---- public entry called from RButton hotkey -------------------------------

ShowContextMenu() {
    global g_LastSelectedText
    g_LastSelectedText := GetSelectedText()
    BuildContextMenu().Show()
}

; ---- menu construction -----------------------------------------------------

BuildContextMenu() {
    m := Menu()
    m.Add("Cut",    (*) => Send("^x"))
    m.Add("Copy",   (*) => Send("^c"))
    m.Add("Paste",  (*) => Send("^v"))
    m.Add("Delete", (*) => Send("{Delete}"))
    m.Add()

    items := CalculatorRegistry()
    items := FilterEnabled(items)
    items := SortMenuItems(items)

    for item in items
        m.Add(item.title, MakeHandler(item))

    m.Add()
    refMenu := BuildReferencesMenu()
    m.Add("References", refMenu)
    m.Add()
    m.Add("Preferences", (*) => ShowPreferencesWindow())
    return m
}

FilterEnabled(items) {
    out := []
    for item in items {
        if (item.pref = "" || Prefs.Get("calculations", item.pref, true))
            out.Push(item)
    }
    return out
}

SortMenuItems(items) {
    method := Prefs.Get("menu", "sortingMethod", "alphabetical")
    if (method = "none")
        return items
    if (method = "frequency")
        return _StableSort(items, _CompareFreqThenTitle)
    ; default: alphabetical
    return _StableSort(items, _CompareTitle)
}

_CompareTitle(a, b) {
    return StrCompare(a.title, b.title, false)
}

_CompareFreqThenTitle(a, b) {
    diff := Freq(b.cmd) - Freq(a.cmd)   ; descending by usage
    if (diff != 0)
        return diff
    return StrCompare(a.title, b.title, false)
}

_StableSort(items, cmpFn) {
    sorted := []
    for item in items {
        inserted := false
        for i, existing in sorted {
            if (cmpFn(item, existing) < 0) {
                sorted.InsertAt(i, item)
                inserted := true
                break
            }
        }
        if !inserted
            sorted.Push(item)
    }
    return sorted
}

Freq(cmd) {
    freq := Prefs.Get("frequency", cmd, 0)
    return Integer(freq)
}

; Build a closure that increments frequency, calls the handler with the
; selected text, and (for handlers returning a string) shows the result.
MakeHandler(item) {
    return (*) => RunCalculator(item)
}

RunCalculator(item) {
    global g_LastSelectedText
    Prefs.IncrementFrequency(item.cmd)
    result := item.handler.Call(g_LastSelectedText)
    if (result != "" && Type(result) = "String")
        ShowResult(result)
}

