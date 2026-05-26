; ============================================================
; lib/Menu.ahk -- Right-click context menu
; ------------------------------------------------------------
; Registry of calculators (title, freq-key, pref-key, handler,
; group); the menu is rebuilt each invocation so toggles, sort
; method, and grouping changes are picked up immediately.
;
; Sorting:
;   "grouped"      -> anatomical submenus (default)
;   "alphabetical" -> flat, case-insensitive by title
;   "frequency"    -> flat, descending by usage, alpha tie-break
;   "none"         -> registry order, flat
; ============================================================

#Requires AutoHotkey v2.0
#Include Util.ahk

global g_LastSelectedText := ""

; Display order of anatomical groups in the grouped menu.
global MENU_GROUP_ORDER := [
    "Liver / biliary",
    "Pancreas",
    "Adrenal",
    "Renal",
    "Lung",
    "Cardiovascular",
    "Adnexal / OB-Gyn",
    "Prostate",
    "Thyroid / neck",
    "Volume + measurement",
    "Numeric",
    "Scheduling"
]

; ---- registry -- handlers live in lib/calc/*.ahk ---------------------------

CalculatorRegistry() {
    return [
        ; --- legacy text-parse calculators ---
        { title: "Ellipsoid Volume",          cmd: "CalculateEllipsoidVolume"
        , pref:  "ellipsoidVolume",           group: "Volume + measurement"
        , handler: EllipsoidVolume_Entry }
,       { title: "Bullet Volume",             cmd: "CalculateBulletVolume"
        , pref:  "bulletVolume",              group: "Volume + measurement"
        , handler: BulletVolume_Entry }
,       { title: "Compare Measurement Sizes", cmd: "CompareNoduleSizes"
        , pref:  "",                          group: "Volume + measurement"
        , handler: CompareSizes_Entry }
,       { title: "Sort Measurement Sizes",    cmd: "SortNoduleSizes"
        , pref:  "",                          group: "Volume + measurement"
        , handler: SortSizes_Entry }
,       { title: "PSA Density",               cmd: "CalculatePSADensity"
        , pref:  "psaDensity",                group: "Prostate"
        , handler: PSADensity_Entry }
,       { title: "Pregnancy Dates",           cmd: "CalculatePregnancyDates"
        , pref:  "pregnancyDates",            group: "Adnexal / OB-Gyn"
        , handler: PregnancyDates_Entry }
,       { title: "Menstrual Phase",           cmd: "CalculateMenstrualPhase"
        , pref:  "menstrualPhase",            group: "Adnexal / OB-Gyn"
        , handler: MenstrualPhase_Entry }
,       { title: "Adrenal Washout",           cmd: "CalculateAdrenalWashout"
        , pref:  "adrenalWashout",            group: "Adrenal"
        , handler: AdrenalWashout_Entry }
,       { title: "Thymus Chemical Shift",     cmd: "CalculateThymusChemicalShift"
        , pref:  "thymusChemicalShift",       group: "Thyroid / neck"
        , handler: ThymusChemicalShift_Entry }
,       { title: "Hepatic Steatosis",         cmd: "CalculateHepaticSteatosis"
        , pref:  "hepaticSteatosis",          group: "Liver / biliary"
        , handler: HepaticSteatosis_Entry }
,       { title: "MRI Liver Iron Content",    cmd: "CalculateIronContent"
        , pref:  "mriLiverIron",              group: "Liver / biliary"
        , handler: LiverIron_Entry }
,       { title: "Statistics",                cmd: "Statistics"
        , pref:  "statistics",                group: "Numeric"
        , handler: Statistics_Entry }
,       { title: "Number Range",              cmd: "Range"
        , pref:  "numberRange",               group: "Numeric"
        , handler: Range_Entry }
,       { title: "Calcium Score Percentile",  cmd: "CalculateCalciumScorePercentile"
        , pref:  "calciumScorePercentile",    group: "Cardiovascular"
        , handler: CalciumScore_Entry }
,       { title: "Contrast Premedication",    cmd: "CalculateContrastPremedication"
        , pref:  "contrastPremedication",     group: "Scheduling"
        , handler: ContrastPremed_Entry }
,       { title: "Fleischner Criteria",       cmd: "CalculateFleischnerCriteria"
        , pref:  "fleischnerCriteria",        group: "Lung"
        , handler: Fleischner_Entry }
,       { title: "NASCET (carotid)",          cmd: "CalculateNASCET"
        , pref:  "nascetCalculator",          group: "Cardiovascular"
        , handler: NASCET_Entry }

        ; --- form-based RADS classifiers ---
,       { title: "Bosniak (renal cyst)",      cmd: "CalculateBosniak"
        , pref:  "bosniak",                   group: "Renal"
        , handler: Bosniak_Entry }
,       { title: "Gallbladder Polyp",         cmd: "CalculateGBPolyp"
        , pref:  "gbPolyp",                   group: "Liver / biliary"
        , handler: GBPolyp_Entry }
,       { title: "Incidental Adrenal Mass",   cmd: "CalculateIncidentalAdrenal"
        , pref:  "incidentalAdrenal",         group: "Adrenal"
        , handler: IncidentalAdrenal_Entry }
,       { title: "Incidental Thyroid Nodule", cmd: "CalculateIncidentalThyroid"
        , pref:  "incidentalThyroid",         group: "Thyroid / neck"
        , handler: IncidentalThyroid_Entry }
,       { title: "Kyoto IPMN",                cmd: "CalculateKyotoIPMN"
        , pref:  "kyotoIpmn",                 group: "Pancreas"
        , handler: KyotoIPMN_Entry }
,       { title: "LI-RADS (CT/MRI)",          cmd: "CalculateLIRADS"
        , pref:  "lirads",                    group: "Liver / biliary"
        , handler: LIRADS_Entry }
,       { title: "Lung-RADS",                 cmd: "CalculateLungRADS"
        , pref:  "lungRads",                  group: "Lung"
        , handler: LungRADS_Entry }
,       { title: "O-RADS MRI",                cmd: "CalculateORADSMRI"
        , pref:  "oradsMri",                  group: "Adnexal / OB-Gyn"
        , handler: ORADSMRI_Entry }
,       { title: "O-RADS Ultrasound",         cmd: "CalculateORADSUS"
        , pref:  "oradsUs",                   group: "Adnexal / OB-Gyn"
        , handler: ORADSUS_Entry }
,       { title: "PI-RADS v2.1",              cmd: "CalculatePIRADS"
        , pref:  "pirads",                    group: "Prostate"
        , handler: PIRADS_Entry }
,       { title: "TI-RADS",                   cmd: "CalculateTIRADS"
        , pref:  "tirads",                    group: "Thyroid / neck"
        , handler: TIRADS_Entry }
,       { title: "US LI-RADS",                cmd: "CalculateUSLIRADS"
        , pref:  "usLirads",                  group: "Liver / biliary"
        , handler: USLIRADS_Entry }
    ]
}

; ---- public entry called from RButton hotkey -------------------------------

ShowContextMenu() {
    global g_LastSelectedText
    ; PowerScribe drops its text selection when an AHK Gui takes focus, so the
    ; second right-click (after closing a calc dialog) finds nothing to copy and
    ; Ctrl+C returns empty. Keep the previous capture in that case instead of
    ; clobbering it with "" -- otherwise every calc after the first one in a
    ; PS session opens with zeroed defaults. Notepad preserves selection across
    ; focus loss so this fallback is invisible there.
    captured := GetSelectedText()
    if (captured != "")
        g_LastSelectedText := captured
    BuildContextMenu().Show()   ; Win32 anchors at cursor and auto-flips at screen edges
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

    method := Prefs.Get("menu", "sortingMethod", "grouped")
    if (method = "grouped")
        _BuildGroupedMenuInto(m, items)
    else
        _BuildFlatMenuInto(m, items, method)

    m.Add()
    refMenu := BuildReferencesMenu()
    m.Add("References", refMenu)
    m.Add()
    m.Add("Preferences", (*) => ShowPreferencesWindow())
    return m
}

_BuildFlatMenuInto(m, items, method) {
    if (method = "frequency")
        items := _StableSort(items, _CompareFreqThenTitle)
    else if (method = "alphabetical")
        items := _StableSort(items, _CompareTitle)
    ; method "none" -> registry order
    for item in items
        m.Add(item.title, MakeHandler(item))
}

_BuildGroupedMenuInto(m, items) {
    global MENU_GROUP_ORDER

    ; bucket items by group name
    buckets := Map()
    for item in items {
        g := item.HasOwnProp("group") && item.group != "" ? item.group : "Other"
        if !buckets.Has(g)
            buckets[g] := []
        buckets[g].Push(item)
    }

    ; emit groups in fixed order, alphabetical within each
    for g in MENU_GROUP_ORDER {
        if !buckets.Has(g)
            continue
        sub := Menu()
        sorted := _StableSort(buckets[g], _CompareTitle)
        for item in sorted
            sub.Add(item.title, MakeHandler(item))
        m.Add(g, sub)
        buckets.Delete(g)
    }

    ; any remaining groups not in MENU_GROUP_ORDER -> append alphabetically
    leftover := []
    for g in buckets
        leftover.Push(g)
    for _, g in _SortStrings(leftover) {
        sub := Menu()
        for item in _StableSort(buckets[g], _CompareTitle)
            sub.Add(item.title, MakeHandler(item))
        m.Add(g, sub)
    }
}

_SortStrings(arr) {
    out := []
    for s in arr {
        inserted := false
        for i, existing in out {
            if (StrCompare(s, existing, false) < 0) {
                out.InsertAt(i, s)
                inserted := true
                break
            }
        }
        if !inserted
            out.Push(s)
    }
    return out
}

FilterEnabled(items) {
    out := []
    for item in items {
        if (item.pref = "" || Prefs.Get("calculations", item.pref, true))
            out.Push(item)
    }
    return out
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
    ; Handler may return "" (it showed its own dialog and the form's submit
    ; handler will route the eventual ShowResult), a legacy flat string, or
    ; a CalcResult struct.
    if IsCalcResult(result)
        ShowResult(result)
    else if (result != "" && Type(result) = "String")
        ShowResult(result)
}
