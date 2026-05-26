; ============================================================
; lib/Prefs.ahk -- User preferences (JSON-backed)
; ------------------------------------------------------------
; Stores all settings in preferences.json next to the script.
; On first run, imports legacy preferences.ini (v1 schema) so
; existing colleagues don't lose their toggles, references, or
; frequency data. Atomic writes via .tmp + rename, plus a
; .backup.json kept after every successful save.
; ============================================================

#Requires AutoHotkey v2.0
#Include Json.ahk
#Include Util.ahk     ; for SafeInt()

class Prefs {
    static Path        := A_ScriptDir "\preferences.json"
    static BackupPath  := A_ScriptDir "\preferences.backup.json"
    static LegacyIni   := A_ScriptDir "\preferences.ini"
    static data        := ""
    static dirty       := false   ; set by IncrementFrequency; flushed on exit

    static Load() {
        Prefs.data := Prefs._Defaults()
        if FileExist(Prefs.Path) {
            try {
                raw := FileRead(Prefs.Path, "UTF-8")
                parsed := Json.Parse(raw)
                Prefs._Merge(parsed)
                return
            } catch as e {
                ; corrupt -- try the backup
                if FileExist(Prefs.BackupPath) {
                    try {
                        raw := FileRead(Prefs.BackupPath, "UTF-8")
                        Prefs._Merge(Json.Parse(raw))
                        MsgBox("Preferences file was unreadable and has been"
                            . " restored from backup.`n`nDetail: " e.Message
                            , "RightClick", "Iconi")
                        return
                    } catch {
                    }
                }
                MsgBox("Preferences file is unreadable. Falling back to defaults."
                    . "`n`nDetail: " e.Message, "RightClick", "Iconx")
            }
            return
        }

        ; No json yet -- attempt one-time import from v1 ini.
        if FileExist(Prefs.LegacyIni) {
            try {
                Prefs._ImportFromV1Ini()
                Prefs.Save()
                ; mark imported so we don't re-import accidentally
                try FileMove(Prefs.LegacyIni, Prefs.LegacyIni ".imported", 1)
            } catch as e {
                MsgBox("Could not import v1 preferences.ini -- using defaults."
                    . "`n`nDetail: " e.Message, "RightClick", "Iconi")
            }
        }
    }

    static Save() {
        ; Atomic write: stage, then move into place; keep a backup of the
        ; previous good file.
        tmp := Prefs.Path ".tmp"
        try {
            if FileExist(tmp)
                FileDelete tmp
            FileAppend(Json.Stringify(Prefs.data, "  "), tmp, "UTF-8")
            if FileExist(Prefs.Path)
                FileCopy(Prefs.Path, Prefs.BackupPath, 1)
            FileMove(tmp, Prefs.Path, 1)
            Prefs.dirty := false
        } catch as e {
            MsgBox("Could not save preferences.`n`n" e.Message, "RightClick", "Iconx")
        }
    }

    ; Save only if there are pending in-memory changes (set by
    ; IncrementFrequency). Wire to OnExit so frequency clicks don't write
    ; the entire JSON file on every menu pick.
    static Flush() {
        if Prefs.dirty
            Prefs.Save()
    }

    ; Reset toggles, modifier, sort, and target apps back to defaults.
    ; References and frequency data are preserved by default since they
    ; represent the user's accumulated state.
    static RestoreDefaults(preserveRefs := true, preserveFreq := true) {
        refs := (preserveRefs && Prefs.data.Has("references"))
                ? Prefs.data["references"] : []
        freq := (preserveFreq && Prefs.data.Has("frequency"))
                ? Prefs.data["frequency"]  : Map()
        Prefs.data := Prefs._Defaults()
        Prefs.data["references"] := refs
        Prefs.data["frequency"]  := freq
        Prefs.Save()
    }

    ; ---- accessors ----------------------------------------------------------

    static Get(section, key, default := "") {
        if (Prefs.data is Map && Prefs.data.Has(section)) {
            s := Prefs.data[section]
            if (s is Map && s.Has(key))
                return s[key]
        }
        return default
    }

    static Set(section, key, value) {
        if !Prefs.data.Has(section)
            Prefs.data[section] := Map()
        Prefs.data[section][key] := value
    }

    static IncrementFrequency(cmd) {
        if !Prefs.data.Has("frequency")
            Prefs.data["frequency"] := Map()
        f := Prefs.data["frequency"]
        f[cmd] := (f.Has(cmd) ? f[cmd] : 0) + 1
        Prefs.dirty := true   ; flushed by Prefs.Flush() on script exit
    }

    static References() {
        if Prefs.data.Has("references") && Prefs.data["references"] is Array
            return Prefs.data["references"]
        Prefs.data["references"] := []
        return Prefs.data["references"]
    }

    static FindReferenceIndex(name) {
        for i, r in Prefs.References() {
            if (r is Map && r.Has("name") && r["name"] = name)
                return i
        }
        return 0
    }

    ; ---- internal -----------------------------------------------------------

    static _Defaults() {
        d := Map()
        d["version"] := 2
        d["display"] := Map(
            "units",          true,
            "allValues",      true,
            "darkMode",       false,
            "showCitations",  true,
            "showArterialAge", true,
            ; Generic toggle for any calculator that displays a validated
            ; malignancy / cancer risk percentage from a published cohort
            ; (e.g., TI-RADS per-TR rates from Middleton 2017). Off-by-design
            ; if a clinician prefers points/category only without the risk %.
            "showMalignancyRisk", true,
            ; When true, the result popup includes a "Methodology" section
            ; (selected inputs + reasoning). When false, only the impression
            ; + citations are shown. The result window itself has a local
            ; checkbox to override this per-result without mutating the pref.
            "showMethodology", true
        )
        d["calculations"] := Map(
            "ellipsoidVolume",       true,
            "bulletVolume",          true,
            "psaDensity",            true,
            "pregnancyDates",        true,
            "menstrualPhase",        true,
            "adrenalWashout",        true,
            "thymusChemicalShift",   true,
            "hepaticSteatosis",      true,
            "mriLiverIron",          true,
            "statistics",            true,
            "numberRange",           true,
            "calciumScorePercentile", true,
            "contrastPremedication", true,
            "fleischnerCriteria",    true,
            "nascetCalculator",      true,
            "bosniak",               true,
            "gbPolyp",               true,
            "incidentalAdrenal",     true,
            "incidentalThyroid",     true,
            "kyotoIpmn",             true,
            "lirads",                true,
            "lungRads",              true,
            "oradsMri",              true,
            "oradsUs",               true,
            "pirads",                true,
            "tirads",                true,
            "usLirads",              true
        )
        d["activation"] := Map(
            "modifier", "ctrl",  ; none | ctrl | alt | shift -- ctrl-by-default so
                                 ; plain right-click stays native in PowerScribe etc.
            "targetApps", [
                Map("type", "class", "value", "Notepad"),
                Map("type", "exe",   "value", "notepad.exe"),
                Map("type", "class", "value", "PowerScribe"),
                Map("type", "exe",   "value", "PowerScribe.exe"),
                Map("type", "class", "value", "PowerScribe360"),
                Map("type", "exe",   "value", "Nuance.PowerScribe360.exe"),
                Map("type", "class", "value", "PowerScribe | Reporting")
            ]
        )
        d["menu"] := Map(
            "sortingMethod", "grouped"  ; grouped (anatomical) | alphabetical | frequency | none
        )
        d["references"] := []
        d["frequency"]  := Map()
        return d
    }

    ; Shallow-deep merge of loaded JSON into defaults so missing keys keep
    ; their default values when new versions add fields.
    static _Merge(loaded) {
        if !(loaded is Map)
            return
        for section, value in loaded {
            if !Prefs.data.Has(section) {
                Prefs.data[section] := value
                continue
            }
            current := Prefs.data[section]
            if (current is Map && value is Map) {
                for k, v in value
                    current[k] := v
            } else {
                Prefs.data[section] := value
            }
        }
        ; Migrate retired "custom" sort method to the new alphabetical default
        ; and drop any stored customOrder array.
        if Prefs.data.Has("menu") && (Prefs.data["menu"] is Map) {
            m := Prefs.data["menu"]
            if (m.Has("sortingMethod") && m["sortingMethod"] = "custom")
                m["sortingMethod"] := "alphabetical"
            if m.Has("customOrder")
                m.Delete("customOrder")
        }
    }

    static _ImportFromV1Ini() {
        ini := Prefs.LegacyIni

        ; --- [Display] ---
        Prefs._IniBool(ini, "Display", "DisplayUnits",     "display", "units",          true)
        Prefs._IniBool(ini, "Display", "DisplayAllValues", "display", "allValues",      true)
        Prefs._IniBool(ini, "Display", "DarkMode",         "display", "darkMode",       false)
        Prefs._IniBool(ini, "Display", "ShowCitations",    "display", "showCitations",  true)
        Prefs._IniBool(ini, "Display", "ShowArterialAge",  "display", "showArterialAge", true)
        Prefs._IniBool(ini, "Display", "ShowMethodology",  "display", "showMethodology", true)

        ; --- [Calculations] ---
        Prefs._IniBool(ini, "Calculations", "ShowEllipsoidVolume",        "calculations", "ellipsoidVolume",        true)
        Prefs._IniBool(ini, "Calculations", "ShowBulletVolume",           "calculations", "bulletVolume",           true)
        Prefs._IniBool(ini, "Calculations", "ShowPSADensity",             "calculations", "psaDensity",             true)
        Prefs._IniBool(ini, "Calculations", "ShowPregnancyDates",         "calculations", "pregnancyDates",         true)
        Prefs._IniBool(ini, "Calculations", "ShowMenstrualPhase",         "calculations", "menstrualPhase",         true)
        Prefs._IniBool(ini, "Calculations", "ShowAdrenalWashout",         "calculations", "adrenalWashout",         true)
        Prefs._IniBool(ini, "Calculations", "ShowThymusChemicalShift",    "calculations", "thymusChemicalShift",    true)
        Prefs._IniBool(ini, "Calculations", "ShowHepaticSteatosis",       "calculations", "hepaticSteatosis",       true)
        Prefs._IniBool(ini, "Calculations", "ShowMRILiverIron",           "calculations", "mriLiverIron",           true)
        Prefs._IniBool(ini, "Calculations", "ShowStatistics",             "calculations", "statistics",             true)
        Prefs._IniBool(ini, "Calculations", "ShowNumberRange",            "calculations", "numberRange",            true)
        Prefs._IniBool(ini, "Calculations", "ShowCalciumScorePercentile", "calculations", "calciumScorePercentile", true)
        Prefs._IniBool(ini, "Calculations", "ShowContrastPremedication",  "calculations", "contrastPremedication",  true)
        Prefs._IniBool(ini, "Calculations", "ShowFleischnerCriteria",     "calculations", "fleischnerCriteria",     true)
        Prefs._IniBool(ini, "Calculations", "ShowNASCETCalculator",       "calculations", "nascetCalculator",       true)

        ; --- [Script] section in v1 only stored PauseDuration; the
        ; Pause feature was removed in v2. Skip the section entirely.

        ; --- [Menu] --- map retired "custom" to alphabetical; drop CustomMenuOrder
        sort := IniRead(ini, "Menu", "SortingMethod", "alphabetical")
        if (sort = "custom" || sort = "")
            sort := "alphabetical"
        Prefs.Set("menu", "sortingMethod", sort)

        ; --- [References] (pipe-delimited "name:::type:::path:::uses") ---
        refList := IniRead(ini, "References", "SavedReferences", "")
        if (refList != "") {
            refs := []
            for _, entry in StrSplit(refList, "|") {
                parts := StrSplit(entry, ":::")
                if (parts.Length >= 3) {
                    refs.Push(Map(
                        "name", parts[1],
                        "type", parts[2],
                        "path", parts[3],
                        "uses", parts.Length >= 4 ? SafeInt(parts[4], 0) : 0
                    ))
                }
            }
            Prefs.data["references"] := refs
        }

        ; --- [Frequency] ---
        freqSection := ""
        try freqSection := IniRead(ini, "Frequency")
        if (freqSection != "") {
            freqMap := Map()
            loop parse freqSection, "`n", "`r" {
                if RegExMatch(A_LoopField, "^(.*?)=(.*)$", &m)
                    freqMap[m[1]] := SafeInt(Trim(m[2]), 0)
            }
            Prefs.data["frequency"] := freqMap
        }
    }

    static _IniBool(ini, iniSection, iniKey, prefSection, prefKey, default) {
        val := IniRead(ini, iniSection, iniKey, default ? "1" : "0")
        Prefs.Set(prefSection, prefKey, val = "1" || val = 1 || val = "true")
    }
}
