; ============================================================
; lib/calc/CalciumScore.ahk -- Agatston score percentile + age
; ------------------------------------------------------------
; Citations:
;   Hoff JA et al. Am J Cardiol 2001;87(12):1335-9.
;   McClelland RL et al. Am J Cardiol 2009;103(1):59-63. (age)
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

CalciumScore_Entry(input) {
    ShowCalciumScoreDialog(input)
    return ""
}

ShowCalciumScoreDialog(text := "") {
    age   := TextScan.Age(text)
    sex   := TextScan.Sex(text)
    race  := TextScan.Race(text)
    score := TextScan.CalciumScore(text)
    sexIdx := (sex = "Female") ? 2 : 1
    raceIdx := (race = "White") ? 2
            : (race = "Black") ? 3
            : (race = "Hispanic") ? 4
            : (race = "Chinese") ? 5
            : 1

    ; Mandatory inputs first (Age / Sex / Score); Race is optional and
    ; lives last under its own clearly-labeled section -- it stays visible
    ; (no disclosure gate) because choosing it merely ADDS the MESA
    ; race-stratified percentile on top of the always-computed Hoff one.
    form := RadsForm("Calcium Score Percentile", 480)
    form.Header("Patient")
    form.Numeric("Age",   "Age (years):", age > 0 ? age : "")
    form.Dropdown("Sex",  "Sex:", ["Male", "Female"], sexIdx)
    form.Header("Coronary artery calcium")
    form.Numeric("Score", "Agatston score:", score != "" ? Round(score, 1) : "")
    form.Header("Optional")
    form.Note("Race adds the race-stratified MESA percentile (valid for ages 45-84).")
    form.Dropdown("Race", "Race (optional, enables MESA):"
        , ["Not specified","White","Black","Hispanic","Chinese"], raceIdx)
    form.SetSubmit(CalciumScore_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

CalciumScore_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (v.Age = "" || v.Score = "")
        return MakeResult({ impression: "Age and calcium score are required.",
                            error: "Age and calcium score are required." })

    age := SafeInt(v.Age, 0)
    sex := v.Sex
    score := v.Score + 0.0
    race := (v.Race = "Not specified" || v.Race = "") ? "" : v.Race

    if (age < 30)
        return MakeResult({ impression: "Coronary calcium score calculator is only valid for ages 30 and above.",
                            error: "Age < 30" })

    ; --- Hoff (age + sex) -- same helpers as the legacy string path ---
    ageGroup := _AgeGroup(age)
    pct := _HoffPercentile(ageGroup, sex, score)
    plaque := _DeterminePlaqueBurden(score)
    comp := _DetermineComparison(pct)

    ; --- Optional MESA (race-stratified) for age 45-84 ---
    mesaLabel := ""
    mesaAgeGroup := ""
    if (race != "" && age >= 45 && age <= 84) {
        mesaAgeGroup := _MesaAge(age)
        mesaPct := _MesaPercentile(mesaAgeGroup, race, sex, score)
        mesaLabel := _MesaCompareLabel(mesaPct)
    }

    ; Hoff/MESA percentile sources don't prescribe management -- this is a
    ; pure-reporting impression. No recommendation line per user policy.
    ;
    ; Score display: prefer the user's raw typing (avoids "78" -> "78.0"
    ; coercion when v.Score + 0.0 produces a Float).
    scoreDisp := v.Score
    ; Plaque-burden text from _DeterminePlaqueBurden is two sentences -- the
    ; first is the lesion descriptor, the second is editorial commentary
    ; ("Mild or minimal narrowings likely"). Strip the second for the
    ; impression; keep both in methodology.
    plaqueImpr := plaque
    if InStr(plaqueImpr, ".") {
        firstDot := InStr(plaqueImpr, ".")
        plaqueImpr := SubStr(plaqueImpr, 1, firstDot - 1)
    }
    plaqueImprLc := (plaqueImpr != "")
        ? StrLower(SubStr(plaqueImpr, 1, 1)) SubStr(plaqueImpr, 2) : ""

    ; If race + age window enables MESA, use only MESA in the impression --
    ; the race-stratified percentile supersedes the age/sex Hoff one when
    ; both are available. Hoff still appears in methodology for audit.
    if (mesaLabel != "") {
        impression := "Coronary calcium score " scoreDisp " Agatston"
        if (plaqueImprLc != "")
            impression .= ": " plaqueImprLc ","
        impression .= " " mesaLabel " percentile by MESA"
                    . " (" race " " StrLower(sex) ", age " mesaAgeGroup ")."
    } else {
        impression := "Coronary calcium score " scoreDisp " Agatston"
        if (plaqueImprLc != "")
            impression .= ": " plaqueImprLc ","
        impression .= " " _CS_HoffLabel(pct) " by Hoff (age, sex)."
    }

    ; --- Methodology ---
    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nHoff comparison (age + sex):"
    method .= "`n  Age group: " ageGroup
    method .= "`n  Plaque burden: " plaque
    method .= "`n  Percentile category: " comp
    if Prefs.Get("display","showArterialAge",true)
        method .= "`n  Arterial age: " _CoronaryAge(score) " years"
    if (mesaLabel != "") {
        method .= "`n`nMESA comparison (race-stratified):"
        method .= "`n  Race: " race
        method .= "`n  Age group: " _MesaAge(age)
        method .= "`n  Percentile: " mesaLabel
    }

    citations := [
        { text: "Hoff JA, et al. Age and gender distributions of coronary artery calcium "
              . "detected by electron beam tomography in 35,246 adults. "
              . "Am J Cardiol. 2001 Jun 15;87(12):1335-9.",
          url:  "https://doi.org/10.1016/s0002-9149(01)01548-x" },
        { text: "McClelland RL, Nasir K, Budoff M, Blumenthal RS, Kronmal RA. "
              . "Arterial age as a function of coronary artery calcium "
              . "(from the Multi-Ethnic Study of Atherosclerosis [MESA]). "
              . "Am J Cardiol. 2009 Jan 1;103(1):59-63.",
          url:  "https://doi.org/10.1016/j.amjcard.2008.08.031" }
    ]
    if (mesaLabel != "")
        citations.Push({ text: "McClelland RL, Chung H, Detrano R, Post W, Kronmal RA. "
                              . "Distribution of coronary artery calcium by race, gender, and age: "
                              . "results from the Multi-Ethnic Study of Atherosclerosis (MESA). "
                              . "Circulation. 2006;113(1):30-37.",
                         url:  "https://doi.org/10.1161/CIRCULATIONAHA.105.580696" })

    ; Short inline fragment for parenthetical paste. Prefer the MESA
    ; percentile when available (same precedence as the impression);
    ; _CS_HoffLabel already includes the word "percentile".
    pasteFrag := (scoreDisp != "")
        ? "Agatston score " scoreDisp ", "
          . (mesaLabel != ""
              ? mesaLabel " percentile (MESA)"
              : _CS_HoffLabel(pct) " (Hoff)")
        : ""

    return MakeResult({
        classification: "Calcium score " scoreDisp,
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      citations,
        echo:           g_LastSelectedText,
        paste:          pasteFrag,
        pasteMode:      ""
    })
}

; Impression-line label for a Hoff percentile bucket. The numeric values
; returned by _HoffPercentile are sentinels (0 = no calcium, 10 = sub-25th
; with calcium, 25/50/75 = at-or-above that cutoff but below the next,
; 90 = at-90th, 99 = above 90th) -- map each to a readable range label.
_CS_HoffLabel(pct) {
    if (pct = 0)
        return "0 percentile (no calcium)"
    if (pct < 25)
        return "below 25th percentile"
    if (pct <= 25)
        return "25th-50th percentile"
    if (pct <= 50)
        return "50th-75th percentile"
    if (pct <= 75)
        return "75th-90th percentile"
    if (pct <= 90)
        return "at 90th percentile"
    return ">90th percentile"
}

CalcCalciumScorePercentile(input) {
    ; \b so "Coverage: 55" (which contains the substring "age:") can't
    ; parse as the patient age.
    if !RegExMatch(input, "i)\bAge:\s*(\d+)", &ageM)
        return input "`n`nError: Age not found or invalid. Please provide age in the format 'Age: 55'."
    if !RegExMatch(input, "i)Sex:\s*(Male|Female)", &sexM)
        return input "`n`nError: Sex not found or invalid. Please specify either Male or Female."
    if !RegExMatch(input
        , "i)(?:your\s+)?(?:coronary\s+artery\s+)?calcium\s+score(?:\s+is)?:?\s*(\d+(?:\.\d+)?)\s*(?:\(?\s*Agatston\s*\)?)?|(?:total\s+)?calcium\s+score:?\s*(\d+(?:\.\d+)?)"
        , &scoreM)
        return input "`n`nError: Calcium score not found or invalid. Please ensure the score is provided in the format 'YOUR CORONARY ARTERY CALCIUM SCORE: 0.0 (Agatston)'."

    age := ageM[1] + 0
    sex := sexM[1]
    score := (scoreM[1] != "") ? scoreM[1] + 0 : (scoreM[2] != "" ? scoreM[2] + 0 : "")
    if (score = "")
        return input "`n`nError: Calcium score not found or invalid."

    if (age < 30)
        return input "`n`nCONTEXT:`nError: The calcium score calculators are only valid for ages 30 and above."

    ; Optional race (White / Black / Hispanic / Chinese). If race AND age 45-84,
    ; report MESA percentile (race-stratified) in addition to Hoff (age + sex).
    race := ""
    if RegExMatch(input, "i)Race:\s*(White|Black|Hispanic|Chinese)", &raceM)
        race := _CapFirst(StrLower(raceM[1]))

    out := input "`n`nCONTEXT:`n" _HoffScore(age, sex, score)
    if (race != "" && age >= 45 && age <= 84)
        out .= "`n" _MesaScore(age, sex, race, score)
    return out
}

_CapFirst(s) {
    return StrUpper(SubStr(s, 1, 1)) StrLower(SubStr(s, 2))
}

_HoffScore(age, sex, score) {
    ageGroup := _AgeGroup(age)
    pct := _HoffPercentile(ageGroup, sex, score)
    out := "Plaque Burden: " _DeterminePlaqueBurden(score) "`n`n"
    out .= "Comparison to people of the same age and sex: " _DetermineComparison(pct) "`n`n"
    if Prefs.Get("display","showArterialAge",true)
        out .= "Arterial Age: " _CoronaryAge(score) " years`n`n"
    if Prefs.Get("display","showCitations",true) {
        out .= "Citation 1: Hoff JA, et al. Age and gender distributions of coronary artery calcium detected by electron beam tomography in 35,246 adults. Am J Cardiol. 2001 Jun 15;87(12):1335-9.`n`n"
        out .= "Citation 2: McClelland RL, Nasir K, Budoff M, Blumenthal RS, Kronmal RA. Arterial age as a function of coronary artery calcium (from the Multi-Ethnic Study of Atherosclerosis [MESA]). Am J Cardiol. 2009 Jan 1;103(1):59-63."
    }
    return out
}

_AgeGroup(age) {
    if (age < 40)
        return "<40"
    if (age < 45)
        return "40-44"
    if (age < 50)
        return "45-49"
    if (age < 55)
        return "50-54"
    if (age < 60)
        return "55-59"
    if (age < 65)
        return "60-64"
    if (age < 70)
        return "65-69"
    if (age <= 74)
        return "70-74"
    return ">74"
}

_HoffPercentile(ageGroup, sex, score) {
    static table := _HoffTable()
    cuts := table[ageGroup][sex]
    ; A true zero score must report as "no calcium" -- otherwise patients in
    ; cells whose 25th-percentile cutoff is 0 (typical for young women and
    ; men <40) would be mislabeled "25th / Low" despite having no calcium.
    if (score = 0)
        return 0
    if (score > cuts[4])
        return 99
    if (score >= cuts[4])
        return 90
    if (score >= cuts[3])
        return 75
    if (score >= cuts[2])
        return 50
    if (score >= cuts[1])
        return 25
    ; Score is non-zero but below the 25th-percentile cutoff for this
    ; age/sex cell (e.g. score 59.9 at age 70-74 male where cuts[1]=64).
    ; Return a sub-25 sentinel so _DetermineComparison reports "Low" rather
    ; than "No calcium". The exact value doesn't matter -- any 0<n<=25
    ; routes to the same "Low (<=25%)" bucket.
    return 10
}

_HoffTable() {
    t := Map()
    t["<40"]   := Map("Male", [0,1,3,14],     "Female", [0,0,1,3])
    t["40-44"] := Map("Male", [0,1,9,59],     "Female", [0,0,1,4])
    t["45-49"] := Map("Male", [0,3,36,154],   "Female", [0,0,2,22])
    t["50-54"] := Map("Male", [1,15,103,332], "Female", [0,0,5,55])
    t["55-59"] := Map("Male", [4,48,215,554], "Female", [0,1,23,121])
    t["60-64"] := Map("Male", [13,113,410,994],   "Female", [0,3,57,193])
    t["65-69"] := Map("Male", [32,180,566,1299],  "Female", [1,24,145,410])
    t["70-74"] := Map("Male", [64,310,892,1774],  "Female", [3,52,210,631])
    t[">74"]   := Map("Male", [166,473,1071,1982],"Female", [9,75,241,709])
    return t
}

_DeterminePlaqueBurden(score) {
    if (score = 0)
        return "None. Risk of coronary artery disease is very low, generally less than 5 percent."
    if (score > 0 && score <= 10)
        return "Minimal identifiable plaque. Risk of coronary artery disease is very unlikely, less than 10 percent."
    if (score > 10 && score <= 100)
        return "At least mild atherosclerotic plaque. Mild or minimal coronary narrowings likely."
    if (score > 100 && score <= 400)
        return "At least moderate atherosclerotic plaque. Mild coronary artery disease highly likely, significant narrowings possible."
    return "Extensive atherosclerotic plaque. High likelihood of at least one significant coronary narrowing."
}

; Maps the _HoffPercentile sentinel values (0 = no calcium, 10 = sub-25th,
; 25/50/75 = at-or-above that cutoff but below the next, 90 = at 90th,
; 99 = above 90th) to band labels. Bands must agree with _CS_HoffLabel --
; the old <= chains here were shifted one band low (sentinel 25, meaning
; "25th-50th", was labeled "Low (<=25%)").
_DetermineComparison(pct) {
    if (pct = 0)
        return "No calcium (0 score)"
    if (pct < 25)
        return "Low (below 25th percentile)"
    if (pct = 25)
        return "Average (25th-50th percentile)"
    if (pct = 50)
        return "Average (50th-75th percentile)"
    if (pct = 75)
        return "High (75th-90th percentile)"
    if (pct = 90)
        return "High (at 90th percentile)"
    return "Very high (>90th percentile)"
}

_CoronaryAge(score) {
    return Round(39.1 + 7.25 * Ln(score + 1), 0)
}

; ---- MESA (McClelland 2006) -- race-stratified percentiles ---------------

_MesaScore(age, sex, race, score) {
    ageGroup := _MesaAge(age)
    pct := _MesaPercentile(ageGroup, race, sex, score)
    out := "MESA comparison (" race " " StrLower(sex) ", " ageGroup "): "
        . _MesaCompareLabel(pct) " percentile`n"
    if Prefs.Get("display","showCitations",true)
        out .= "Citation 3: McClelland RL, Chung H, Detrano R, Post W, Kronmal RA. Distribution of coronary artery calcium by race, gender, and age: results from the Multi-Ethnic Study of Atherosclerosis (MESA). Circulation. 2006;113(1):30-37.`n"
    return out
}

_MesaAge(age) {
    if (age < 45)
        return ""
    if (age <= 54)
        return "45-54"
    if (age <= 64)
        return "55-64"
    if (age <= 74)
        return "65-74"
    if (age <= 84)
        return "75-84"
    return ""
}

_MesaCompareLabel(pct) {
    if (pct = 0)
        return "0 (no calcium)"
    if (pct < 25)
        return "<25th"          ; non-zero score below the 25th cutoff
    if (pct <= 25)
        return "25th"
    if (pct <= 50)
        return "50th"
    if (pct <= 75)
        return "75th"
    if (pct <= 90)
        return "90th"
    if (pct <= 95)
        return "95th"
    return ">95th"
}

_MesaPercentile(ageGroup, race, sex, score) {
    table := _MesaTable()
    key := race . sex
    if !table.Has(key) || !table[key].Has(ageGroup)
        return 0
    ; True zero must report as "no calcium" before any cutoff comparison.
    ; Many MESA cells have cuts[1]=cuts[2]=0 (e.g. WhiteFemale 45-54), so
    ; without this guard a score of 0 would match `score >= cuts[2]` and
    ; mislabel as "50th percentile" despite having no calcium.
    if (score = 0)
        return 0
    cuts := table[key][ageGroup]   ; [25th, 50th, 75th, 90th, 95th]
    if (score > cuts[5])
        return 99
    if (score >= cuts[5])
        return 95
    if (score >= cuts[4])
        return 90
    if (score >= cuts[3])
        return 75
    if (score >= cuts[2])
        return 50
    if (score >= cuts[1])
        return 25
    ; Non-zero score below the 25th-percentile cutoff -- sub-25 sentinel
    ; for "Low (<=25%)" (same reasoning as _HoffPercentile).
    return 10
}

_MesaTable() {
    static t := _BuildMesaTable()
    return t
}

_BuildMesaTable() {
    t := Map()
    t["WhiteMale"] := Map(
        "45-54", [0,0,22,110,207],
        "55-64", [0,28,155,452,743],
        "65-74", [21,145,540,1345,2271],
        "75-84", [103,385,1200,2933,4619])
    t["WhiteFemale"] := Map(
        "45-54", [0,0,0,8,31],
        "55-64", [0,0,16,102,209],
        "65-74", [0,13,119,391,674],
        "75-84", [20,106,370,921,1535])
    t["ChineseMale"] := Map(
        "45-54", [0,0,14,89,184],
        "55-64", [0,5,67,242,429],
        "65-74", [0,34,174,487,803],
        "75-84", [11,81,305,769,1299])
    t["ChineseFemale"] := Map(
        "45-54", [0,0,0,12,44],
        "55-64", [0,0,18,105,213],
        "65-74", [0,5,70,246,436],
        "75-84", [0,32,146,398,656])
    t["BlackMale"] := Map(
        "45-54", [0,0,2,45,105],
        "55-64", [0,0,40,173,318],
        "65-74", [0,32,191,575,945],
        "75-84", [23,141,516,1281,2176])
    t["BlackFemale"] := Map(
        "45-54", [0,0,0,9,38],
        "55-64", [0,0,5,74,173],
        "65-74", [0,0,77,310,561],
        "75-84", [0,47,214,582,953])
    t["HispanicMale"] := Map(
        "45-54", [0,0,9,88,195],
        "55-64", [0,3,75,291,512],
        "65-74", [1,56,247,666,1091],
        "75-84", [36,153,494,1221,1943])
    t["HispanicFemale"] := Map(
        "45-54", [0,0,0,2,18],
        "55-64", [0,0,2,50,118],
        "65-74", [0,1,51,203,361],
        "75-84", [0,45,205,557,917])
    return t
}
