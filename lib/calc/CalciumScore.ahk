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

    ; --- Optional MESA (race-stratified, age-specific) for age 45-84 ---
    ; _MesaPercentile now returns the exact age-specific MESA reference
    ; percentile, not a coarse bracket-midpoint bucket.
    mesaLabel := ""
    mesaPct := -1
    if (race != "") {
        mesaPct := _MesaPercentile(age, race, sex, score)
        if (mesaPct >= 0)
            mesaLabel := _MesaOrdinal(mesaPct)
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
                    . " (" race " " StrLower(sex) ", age " age ")."
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
        method .= "`n`nMESA comparison (race-stratified, age-specific):"
        method .= "`n  Race: " race
        method .= "`n  Age: " age
        method .= "`n  P(non-zero calcium): " _MesaNonzeroProb(age, race, sex) "%"
        method .= "`n  Percentile: " mesaLabel (score = 0 ? " (no calcium)" : "")
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

; ---- MESA (McClelland 2006) -- age-specific, race-stratified percentiles ----
;
; The published Table 2 lists 25/50/75/90/95th cutoffs only at each 10-year
; bracket's MIDPOINT, so bucketing it (the old approach) both quantized the
; answer and scored everyone as the bracket midpoint -- a 46-year-old was
; evaluated as ~49-50, badly understating the percentile (score 45 -> "75th"
; when the true, age-specific value is 89th). Instead we use the MESA
; reference model (McClelland 2006) evaluated per integer age: for every
; integer age 45-84 x sex x race we tabulate the reference percentile at a
; fixed ladder of Agatston scores (_MesaLadder) plus the probability of
; non-zero calcium. Age needs no interpolation (each year is stored); a
; patient's score is mapped by log-linear interpolation between the two
; bracketing ladder anchors, accurate to within ~1 percentile (max 2).

_MesaScore(age, sex, race, score) {
    pct := _MesaPercentile(age, race, sex, score)
    label := (score = 0) ? "0 (no calcium)" : _MesaOrdinal(pct)
    out := "MESA comparison (" race " " StrLower(sex) ", age " age "): "
        . label " percentile`n"
    if Prefs.Get("display","showCitations",true)
        out .= "Citation 3: McClelland RL, Chung H, Detrano R, Post W, Kronmal RA. Distribution of coronary artery calcium by race, gender, and age: results from the Multi-Ethnic Study of Atherosclerosis (MESA). Circulation. 2006;113(1):30-37.`n"
    return out
}

_MesaAnchors() {
    static a := _BuildMesaAnchors()
    return a
}

_MesaPnz() {
    static z := _BuildMesaPnz()
    return z
}

; Probability (%) of a non-zero calcium score for this age/race/sex, or -1 if
; unavailable. Reported in methodology alongside the percentile.
_MesaNonzeroProb(age, race, sex) {
    if (age < 45 || age > 84)
        return -1
    z := _MesaPnz()
    key := sex "|" race
    return z.Has(key) ? z[key][age - 44] : -1
}

; Exact MESA reference percentile (integer 0-99) for age 45-84, via log-linear
; interpolation across the score ladder. Returns -1
; when the age/race/sex cell is unavailable so callers can fall back to Hoff.
_MesaPercentile(age, race, sex, score) {
    if (age < 45 || age > 84)
        return -1
    anchors := _MesaAnchors()
    key := sex "|" race
    if !anchors.Has(key)
        return -1
    a := anchors[key][age - 44]         ; percentiles at each _MesaLadder score
    if (score <= 0)
        return 0
    static L := _MesaLadder()
    n := L.Length
    if (score <= L[1])
        return a[1]
    if (score >= L[n])
        return Min(99, a[n])
    i := 1
    while (i < n - 1 && score > L[i + 1])
        i += 1
    f := (Ln(score) - Ln(L[i])) / (Ln(L[i + 1]) - Ln(L[i]))
    pct := Round(a[i] + f * (a[i + 1] - a[i]))
    if (pct < 0)
        pct := 0
    if (pct > 99)
        pct := 99
    return pct
}

; Ordinal label for an exact percentile: 89 -> "89th", 1 -> "1st", 0 -> "0".
_MesaOrdinal(pct) {
    n := Round(pct)
    if (n = 0)
        return "0"
    ones := Mod(n, 10)
    tens := Mod(n, 100)
    suffix := (tens >= 11 && tens <= 13) ? "th"
            : (ones = 1) ? "st"
            : (ones = 2) ? "nd"
            : (ones = 3) ? "rd"
            : "th"
    return n suffix
}

_MesaLadder() {
    ; Fixed Agatston anchor scores for the MESA percentile interpolation.
    return [1, 2, 5, 12, 30, 80, 200, 500, 1500, 4000]
}

_BuildMesaPnz() {
    ; P(non-zero calcium) percent, per age 45..84 (index age-44), from the
    ; MESA reference model (McClelland 2006). Shown in methodology alongside
    ; the percentile.
    z := Map()
    z["Male|White"] := [
        25, 28, 31, 35, 38, 41, 44, 47, 50, 53,   ; age 45-54
        56, 59, 62, 64, 66, 68, 70, 72, 74, 75,   ; age 55-64
        77, 78, 79, 81, 82, 83, 85, 86, 87, 88,   ; age 65-74
        90, 91, 92, 93, 95, 96, 97, 98, 99, 99   ; age 75-84
    ]
    z["Male|Chinese"] := [
        28, 29, 31, 33, 34, 36, 38, 40, 41, 43,   ; age 45-54
        45, 47, 48, 50, 52, 54, 56, 58, 60, 62,   ; age 55-64
        64, 65, 67, 68, 70, 71, 72, 73, 75, 76,   ; age 65-74
        77, 78, 79, 80, 81, 82, 83, 84, 85, 86   ; age 75-84
    ]
    z["Male|Black"] := [
        19, 21, 22, 24, 26, 28, 29, 31, 33, 35,   ; age 45-54
        37, 39, 41, 43, 44, 46, 48, 50, 52, 54,   ; age 55-64
        56, 58, 60, 62, 64, 66, 68, 70, 72, 73,   ; age 65-74
        75, 77, 79, 81, 82, 84, 86, 88, 90, 92   ; age 75-84
    ]
    z["Male|Hispanic"] := [
        23, 25, 26, 28, 29, 31, 32, 34, 35, 37,   ; age 45-54
        38, 40, 43, 46, 49, 51, 54, 57, 59, 62,   ; age 55-64
        64, 67, 69, 71, 73, 75, 77, 79, 80, 82,   ; age 65-74
        84, 86, 87, 89, 91, 92, 94, 96, 99, 99   ; age 75-84
    ]
    z["Female|White"] := [
        7, 9, 11, 13, 15, 16, 18, 20, 22, 24,   ; age 45-54
        26, 28, 30, 32, 34, 36, 38, 40, 43, 45,   ; age 55-64
        47, 49, 52, 54, 56, 59, 62, 65, 67, 70,   ; age 65-74
        73, 76, 78, 81, 84, 87, 89, 92, 95, 98   ; age 75-84
    ]
    z["Female|Chinese"] := [
        4, 7, 9, 11, 14, 16, 18, 20, 23, 25,   ; age 45-54
        27, 30, 32, 34, 36, 37, 39, 41, 42, 44,   ; age 55-64
        46, 48, 50, 51, 53, 55, 57, 59, 60, 62,   ; age 65-74
        64, 66, 68, 70, 72, 74, 76, 78, 79, 81   ; age 75-84
    ]
    z["Female|Black"] := [
        8, 10, 11, 12, 13, 15, 16, 17, 18, 20,   ; age 45-54
        21, 22, 24, 25, 26, 28, 30, 32, 33, 35,   ; age 55-64
        37, 39, 41, 44, 46, 49, 51, 54, 56, 59,   ; age 65-74
        62, 64, 67, 69, 72, 74, 77, 80, 82, 85   ; age 75-84
    ]
    z["Female|Hispanic"] := [
        4, 5, 7, 8, 10, 11, 13, 15, 16, 18,   ; age 45-54
        19, 21, 22, 24, 25, 27, 29, 31, 34, 36,   ; age 55-64
        38, 41, 43, 45, 48, 50, 53, 55, 57, 60,   ; age 65-74
        62, 65, 67, 69, 72, 74, 76, 79, 81, 83   ; age 75-84
    ]
    return z
}

_BuildMesaAnchors() {
    ; MESA reference percentile at each _MesaLadder() score, per age 45..84
    ; (index = age-44), from the MESA reference model (McClelland 2006). Each
    ; row has 10 non-decreasing percentiles; _MesaPercentile interpolates in ln(score).
    a := Map()
    a["Male|White"] := [
        [75, 77, 79, 83, 89, 94, 98, 99, 99, 99],   ; age 45
        [72, 74, 76, 81, 86, 93, 97, 99, 99, 99],   ; age 46
        [69, 70, 73, 78, 84, 92, 97, 99, 99, 99],   ; age 47
        [65, 67, 70, 75, 82, 90, 96, 99, 99, 99],   ; age 48
        [62, 64, 67, 73, 80, 88, 95, 98, 99, 99],   ; age 49
        [59, 60, 64, 70, 77, 87, 94, 98, 99, 99],   ; age 50
        [56, 57, 61, 67, 75, 85, 93, 98, 99, 99],   ; age 51
        [53, 54, 58, 64, 72, 83, 92, 97, 99, 99],   ; age 52
        [49, 51, 55, 61, 70, 81, 91, 97, 99, 99],   ; age 53
        [46, 48, 51, 58, 67, 79, 89, 96, 99, 99],   ; age 54
        [43, 45, 48, 55, 64, 77, 88, 96, 99, 99],   ; age 55
        [40, 42, 45, 51, 61, 74, 86, 95, 99, 99],   ; age 56
        [38, 39, 42, 48, 58, 72, 85, 94, 99, 99],   ; age 57
        [35, 36, 40, 46, 56, 70, 83, 93, 99, 99],   ; age 58
        [33, 34, 37, 43, 53, 67, 81, 92, 98, 99],   ; age 59
        [31, 32, 35, 41, 50, 64, 79, 91, 98, 99],   ; age 60
        [29, 30, 33, 38, 48, 61, 76, 89, 98, 99],   ; age 61
        [27, 28, 31, 36, 45, 59, 74, 88, 97, 99],   ; age 62
        [25, 26, 28, 33, 42, 56, 72, 86, 97, 99],   ; age 63
        [24, 24, 26, 31, 40, 54, 69, 84, 96, 99],   ; age 64
        [22, 23, 25, 29, 38, 51, 67, 82, 95, 99],   ; age 65
        [21, 21, 23, 28, 35, 49, 65, 81, 94, 99],   ; age 66
        [20, 20, 22, 26, 33, 46, 62, 79, 94, 99],   ; age 67
        [19, 19, 20, 24, 31, 44, 60, 77, 92, 98],   ; age 68
        [17, 18, 19, 22, 30, 42, 58, 75, 92, 98],   ; age 69
        [16, 16, 17, 21, 27, 40, 56, 73, 90, 98],   ; age 70
        [15, 15, 16, 19, 26, 38, 54, 71, 90, 98],   ; age 71
        [13, 13, 14, 17, 24, 36, 51, 70, 89, 97],   ; age 72
        [12, 12, 13, 16, 22, 34, 50, 68, 88, 97],   ; age 73
        [11, 11, 12, 14, 20, 32, 47, 66, 86, 96],   ; age 74
        [9, 9, 10, 13, 19, 30, 46, 64, 85, 96],   ; age 75
        [8, 8, 9, 11, 17, 28, 43, 62, 84, 95],   ; age 76
        [7, 7, 8, 10, 15, 26, 42, 61, 82, 95],   ; age 77
        [6, 6, 6, 8, 14, 24, 40, 59, 81, 94],   ; age 78
        [4, 5, 5, 7, 12, 23, 38, 57, 80, 94],   ; age 79
        [3, 3, 4, 6, 11, 21, 36, 55, 78, 93],   ; age 80
        [2, 2, 3, 4, 9, 19, 34, 53, 77, 92],   ; age 81
        [1, 1, 1, 3, 8, 17, 32, 51, 75, 92],   ; age 82
        [0, 0, 0, 2, 6, 16, 30, 49, 74, 91],   ; age 83
        [0, 0, 0, 1, 5, 15, 28, 48, 73, 90]   ; age 84
    ]
    a["Male|Chinese"] := [
        [72, 74, 76, 81, 86, 93, 97, 99, 99, 99],   ; age 45
        [71, 72, 75, 79, 85, 92, 97, 99, 99, 99],   ; age 46
        [69, 70, 73, 78, 84, 91, 96, 99, 99, 99],   ; age 47
        [67, 69, 72, 77, 83, 90, 96, 99, 99, 99],   ; age 48
        [66, 67, 70, 75, 82, 89, 96, 99, 99, 99],   ; age 49
        [64, 65, 68, 73, 80, 89, 95, 98, 99, 99],   ; age 50
        [62, 63, 67, 72, 79, 88, 95, 98, 99, 99],   ; age 51
        [60, 62, 65, 70, 77, 87, 94, 98, 99, 99],   ; age 52
        [58, 60, 63, 69, 76, 86, 93, 98, 99, 99],   ; age 53
        [57, 58, 62, 67, 75, 84, 93, 98, 99, 99],   ; age 54
        [55, 56, 60, 65, 73, 83, 92, 97, 99, 99],   ; age 55
        [53, 55, 58, 64, 72, 82, 91, 97, 99, 99],   ; age 56
        [52, 53, 56, 62, 70, 81, 91, 97, 99, 99],   ; age 57
        [50, 51, 54, 60, 69, 80, 90, 97, 99, 99],   ; age 58
        [47, 49, 52, 58, 67, 78, 89, 96, 99, 99],   ; age 59
        [45, 46, 50, 56, 65, 77, 87, 96, 99, 99],   ; age 60
        [43, 44, 48, 53, 62, 75, 86, 95, 99, 99],   ; age 61
        [41, 42, 46, 51, 60, 73, 85, 94, 99, 99],   ; age 62
        [39, 40, 44, 49, 59, 71, 84, 94, 99, 99],   ; age 63
        [38, 39, 42, 47, 57, 70, 83, 93, 99, 99],   ; age 64
        [36, 37, 40, 46, 55, 69, 82, 93, 99, 99],   ; age 65
        [34, 35, 38, 44, 54, 67, 81, 92, 98, 99],   ; age 66
        [33, 34, 37, 43, 52, 66, 80, 92, 98, 99],   ; age 67
        [31, 32, 35, 41, 50, 65, 79, 91, 98, 99],   ; age 68
        [30, 31, 34, 39, 49, 63, 78, 90, 98, 99],   ; age 69
        [28, 29, 32, 38, 48, 62, 77, 90, 98, 99],   ; age 70
        [27, 28, 31, 37, 46, 61, 76, 89, 98, 99],   ; age 71
        [26, 27, 30, 35, 45, 60, 75, 89, 98, 99],   ; age 72
        [25, 26, 28, 34, 44, 59, 74, 88, 97, 99],   ; age 73
        [24, 24, 27, 33, 43, 57, 73, 88, 97, 99],   ; age 74
        [22, 23, 26, 32, 41, 56, 72, 87, 97, 99],   ; age 75
        [21, 22, 25, 30, 40, 55, 71, 86, 97, 99],   ; age 76
        [20, 21, 23, 29, 38, 53, 70, 85, 96, 99],   ; age 77
        [19, 20, 22, 28, 37, 52, 69, 84, 96, 99],   ; age 78
        [18, 19, 21, 26, 36, 51, 68, 83, 96, 99],   ; age 79
        [17, 18, 20, 25, 34, 49, 66, 83, 96, 99],   ; age 80
        [16, 17, 18, 24, 33, 48, 65, 82, 95, 99],   ; age 81
        [15, 15, 17, 22, 32, 46, 64, 81, 95, 99],   ; age 82
        [14, 14, 16, 21, 30, 45, 63, 80, 94, 99],   ; age 83
        [13, 13, 15, 20, 29, 44, 61, 79, 94, 99]   ; age 84
    ]
    a["Male|Black"] := [
        [81, 82, 85, 88, 92, 96, 99, 99, 99, 99],   ; age 45
        [79, 81, 83, 87, 91, 96, 98, 99, 99, 99],   ; age 46
        [78, 79, 82, 86, 90, 95, 98, 99, 99, 99],   ; age 47
        [76, 78, 80, 84, 89, 94, 98, 99, 99, 99],   ; age 48
        [74, 76, 79, 83, 88, 94, 98, 99, 99, 99],   ; age 49
        [73, 74, 77, 81, 87, 93, 97, 99, 99, 99],   ; age 50
        [71, 72, 75, 80, 86, 92, 97, 99, 99, 99],   ; age 51
        [69, 71, 74, 78, 84, 92, 97, 99, 99, 99],   ; age 52
        [67, 69, 72, 77, 83, 90, 96, 99, 99, 99],   ; age 53
        [65, 67, 70, 75, 82, 89, 95, 99, 99, 99],   ; age 54
        [63, 65, 68, 73, 80, 88, 95, 98, 99, 99],   ; age 55
        [61, 63, 66, 71, 78, 87, 94, 98, 99, 99],   ; age 56
        [59, 61, 64, 69, 76, 86, 93, 98, 99, 99],   ; age 57
        [57, 59, 62, 67, 75, 84, 93, 98, 99, 99],   ; age 58
        [55, 57, 60, 65, 73, 83, 92, 97, 99, 99],   ; age 59
        [54, 55, 58, 63, 71, 82, 91, 97, 99, 99],   ; age 60
        [51, 53, 56, 61, 69, 80, 90, 96, 99, 99],   ; age 61
        [49, 50, 54, 59, 67, 78, 88, 96, 99, 99],   ; age 62
        [47, 48, 51, 56, 65, 76, 87, 95, 99, 99],   ; age 63
        [45, 46, 49, 54, 63, 74, 86, 94, 99, 99],   ; age 64
        [43, 44, 47, 52, 60, 72, 84, 93, 99, 99],   ; age 65
        [41, 42, 45, 50, 58, 70, 82, 93, 99, 99],   ; age 66
        [39, 40, 43, 47, 56, 68, 81, 92, 98, 99],   ; age 67
        [37, 38, 40, 45, 54, 66, 79, 90, 98, 99],   ; age 68
        [35, 36, 38, 43, 51, 63, 77, 89, 98, 99],   ; age 69
        [33, 34, 36, 41, 49, 61, 75, 88, 97, 99],   ; age 70
        [31, 32, 34, 38, 46, 59, 73, 87, 97, 99],   ; age 71
        [29, 30, 32, 36, 44, 57, 71, 85, 96, 99],   ; age 72
        [28, 28, 30, 34, 42, 54, 69, 84, 96, 99],   ; age 73
        [26, 26, 28, 32, 40, 53, 68, 82, 95, 99],   ; age 74
        [24, 24, 26, 30, 38, 51, 66, 81, 95, 99],   ; age 75
        [22, 23, 24, 28, 36, 49, 64, 80, 94, 99],   ; age 76
        [20, 21, 22, 26, 34, 47, 62, 79, 93, 99],   ; age 77
        [19, 19, 20, 24, 32, 45, 61, 77, 93, 98],   ; age 78
        [17, 17, 18, 22, 29, 42, 58, 75, 92, 98],   ; age 79
        [15, 15, 16, 20, 27, 40, 56, 74, 91, 98],   ; age 80
        [13, 13, 15, 18, 25, 38, 54, 72, 90, 98],   ; age 81
        [11, 11, 13, 16, 23, 35, 52, 70, 89, 98],   ; age 82
        [9, 10, 11, 14, 21, 33, 49, 68, 88, 97],   ; age 83
        [7, 8, 9, 11, 18, 31, 47, 66, 87, 97]   ; age 84
    ]
    a["Male|Hispanic"] := [
        [76, 77, 80, 83, 88, 93, 97, 99, 99, 99],   ; age 45
        [75, 76, 78, 82, 87, 92, 97, 99, 99, 99],   ; age 46
        [73, 74, 77, 80, 85, 91, 96, 99, 99, 99],   ; age 47
        [72, 73, 75, 79, 84, 91, 96, 99, 99, 99],   ; age 48
        [70, 71, 74, 78, 83, 90, 95, 98, 99, 99],   ; age 49
        [69, 70, 72, 76, 82, 89, 95, 98, 99, 99],   ; age 50
        [67, 68, 71, 75, 80, 88, 94, 98, 99, 99],   ; age 51
        [66, 67, 69, 73, 79, 87, 93, 98, 99, 99],   ; age 52
        [64, 65, 68, 72, 78, 86, 93, 97, 99, 99],   ; age 53
        [63, 64, 66, 70, 77, 85, 92, 97, 99, 99],   ; age 54
        [61, 62, 65, 69, 75, 84, 91, 97, 99, 99],   ; age 55
        [59, 60, 62, 67, 73, 82, 90, 97, 99, 99],   ; age 56
        [57, 57, 60, 64, 71, 81, 89, 96, 99, 99],   ; age 57
        [54, 54, 57, 61, 69, 79, 88, 95, 99, 99],   ; age 58
        [51, 52, 54, 59, 66, 77, 87, 95, 99, 99],   ; age 59
        [48, 49, 52, 56, 64, 75, 86, 94, 99, 99],   ; age 60
        [45, 46, 49, 54, 62, 73, 85, 94, 99, 99],   ; age 61
        [43, 44, 46, 52, 60, 72, 84, 93, 99, 99],   ; age 62
        [40, 41, 44, 49, 58, 70, 83, 93, 99, 99],   ; age 63
        [38, 39, 41, 47, 55, 68, 81, 92, 98, 99],   ; age 64
        [35, 36, 39, 44, 53, 66, 80, 91, 98, 99],   ; age 65
        [33, 34, 36, 42, 51, 64, 78, 90, 98, 99],   ; age 66
        [30, 31, 34, 39, 48, 62, 76, 89, 98, 99],   ; age 67
        [28, 29, 32, 37, 46, 59, 74, 88, 97, 99],   ; age 68
        [26, 27, 29, 34, 43, 57, 73, 87, 97, 99],   ; age 69
        [24, 25, 27, 32, 41, 55, 71, 86, 97, 99],   ; age 70
        [22, 23, 25, 30, 39, 53, 69, 84, 96, 99],   ; age 71
        [21, 21, 23, 28, 37, 51, 68, 83, 96, 99],   ; age 72
        [19, 20, 21, 26, 35, 49, 66, 82, 95, 99],   ; age 73
        [17, 18, 20, 24, 33, 48, 65, 81, 95, 99],   ; age 74
        [15, 16, 18, 23, 31, 46, 63, 80, 95, 99],   ; age 75
        [14, 14, 16, 21, 30, 44, 62, 79, 94, 99],   ; age 76
        [12, 12, 14, 19, 28, 43, 60, 78, 94, 99],   ; age 77
        [10, 11, 12, 17, 26, 41, 59, 77, 93, 99],   ; age 78
        [8, 9, 11, 15, 24, 39, 57, 76, 93, 99],   ; age 79
        [7, 7, 9, 13, 22, 37, 56, 75, 92, 98],   ; age 80
        [5, 6, 7, 12, 20, 36, 54, 74, 92, 98],   ; age 81
        [3, 4, 6, 10, 19, 34, 52, 72, 91, 98],   ; age 82
        [0, 0, 2, 6, 15, 31, 51, 71, 90, 98],   ; age 83
        [0, 1, 2, 6, 15, 30, 49, 70, 90, 98]   ; age 84
    ]
    a["Female|White"] := [
        [93, 94, 95, 97, 98, 99, 99, 99, 99, 99],   ; age 45
        [91, 92, 94, 96, 97, 99, 99, 99, 99, 99],   ; age 46
        [89, 90, 92, 95, 97, 99, 99, 99, 99, 99],   ; age 47
        [88, 89, 91, 93, 96, 98, 99, 99, 99, 99],   ; age 48
        [86, 87, 89, 92, 95, 98, 99, 99, 99, 99],   ; age 49
        [84, 85, 88, 91, 94, 98, 99, 99, 99, 99],   ; age 50
        [82, 84, 86, 89, 93, 97, 99, 99, 99, 99],   ; age 51
        [80, 82, 84, 88, 92, 96, 99, 99, 99, 99],   ; age 52
        [78, 80, 83, 86, 91, 96, 98, 99, 99, 99],   ; age 53
        [77, 78, 81, 85, 90, 95, 98, 99, 99, 99],   ; age 54
        [74, 76, 79, 83, 88, 94, 98, 99, 99, 99],   ; age 55
        [72, 74, 77, 81, 87, 93, 97, 99, 99, 99],   ; age 56
        [70, 72, 75, 79, 85, 92, 97, 99, 99, 99],   ; age 57
        [68, 70, 73, 77, 83, 90, 96, 99, 99, 99],   ; age 58
        [66, 67, 70, 75, 81, 89, 95, 98, 99, 99],   ; age 59
        [64, 65, 68, 73, 79, 87, 94, 98, 99, 99],   ; age 60
        [61, 63, 66, 70, 77, 86, 93, 98, 99, 99],   ; age 61
        [59, 60, 63, 68, 75, 84, 92, 97, 99, 99],   ; age 62
        [57, 58, 61, 65, 73, 82, 91, 97, 99, 99],   ; age 63
        [55, 56, 58, 63, 71, 80, 89, 96, 99, 99],   ; age 64
        [52, 53, 56, 61, 68, 79, 88, 96, 99, 99],   ; age 65
        [50, 51, 54, 58, 66, 77, 87, 95, 99, 99],   ; age 66
        [48, 49, 52, 56, 64, 75, 86, 94, 99, 99],   ; age 67
        [45, 46, 49, 54, 62, 73, 85, 94, 99, 99],   ; age 68
        [43, 44, 47, 51, 60, 72, 83, 93, 99, 99],   ; age 69
        [40, 41, 44, 49, 57, 69, 82, 92, 98, 99],   ; age 70
        [37, 38, 41, 46, 54, 67, 80, 91, 98, 99],   ; age 71
        [35, 35, 38, 43, 52, 64, 78, 90, 98, 99],   ; age 72
        [32, 33, 35, 40, 49, 61, 76, 89, 97, 99],   ; age 73
        [29, 30, 32, 37, 46, 59, 74, 87, 97, 99],   ; age 74
        [26, 27, 29, 34, 43, 56, 72, 86, 97, 99],   ; age 75
        [24, 24, 26, 32, 40, 54, 70, 85, 96, 99],   ; age 76
        [21, 22, 24, 29, 37, 52, 68, 84, 96, 99],   ; age 77
        [18, 19, 21, 26, 35, 49, 66, 82, 95, 99],   ; age 78
        [15, 16, 18, 23, 32, 47, 64, 81, 95, 99],   ; age 79
        [13, 13, 15, 20, 29, 45, 62, 80, 94, 99],   ; age 80
        [10, 10, 12, 17, 27, 42, 60, 79, 94, 99],   ; age 81
        [7, 7, 9, 15, 24, 40, 58, 77, 94, 99],   ; age 82
        [4, 5, 7, 12, 21, 37, 56, 76, 93, 99],   ; age 83
        [1, 2, 4, 9, 18, 35, 54, 75, 93, 98]   ; age 84
    ]
    a["Female|Chinese"] := [
        [95, 95, 96, 97, 98, 99, 99, 99, 99, 99],   ; age 45
        [93, 93, 94, 96, 97, 99, 99, 99, 99, 99],   ; age 46
        [91, 91, 93, 94, 96, 98, 99, 99, 99, 99],   ; age 47
        [89, 89, 91, 93, 95, 98, 99, 99, 99, 99],   ; age 48
        [86, 87, 89, 91, 94, 97, 99, 99, 99, 99],   ; age 49
        [84, 85, 87, 90, 93, 96, 99, 99, 99, 99],   ; age 50
        [82, 83, 85, 88, 92, 96, 98, 99, 99, 99],   ; age 51
        [80, 81, 83, 86, 91, 95, 98, 99, 99, 99],   ; age 52
        [77, 79, 81, 85, 89, 94, 98, 99, 99, 99],   ; age 53
        [75, 76, 79, 83, 88, 94, 97, 99, 99, 99],   ; age 54
        [73, 74, 77, 81, 87, 93, 97, 99, 99, 99],   ; age 55
        [70, 72, 75, 79, 85, 92, 97, 99, 99, 99],   ; age 56
        [68, 70, 73, 77, 84, 91, 96, 99, 99, 99],   ; age 57
        [66, 68, 71, 76, 82, 90, 96, 99, 99, 99],   ; age 58
        [64, 66, 69, 74, 80, 88, 95, 98, 99, 99],   ; age 59
        [63, 64, 67, 72, 78, 87, 94, 98, 99, 99],   ; age 60
        [61, 62, 65, 70, 77, 86, 93, 98, 99, 99],   ; age 61
        [59, 60, 63, 68, 75, 84, 92, 97, 99, 99],   ; age 62
        [57, 58, 61, 66, 73, 83, 91, 97, 99, 99],   ; age 63
        [55, 57, 60, 64, 72, 82, 91, 97, 99, 99],   ; age 64
        [54, 55, 58, 63, 70, 81, 90, 96, 99, 99],   ; age 65
        [52, 53, 56, 61, 69, 79, 89, 96, 99, 99],   ; age 66
        [50, 51, 54, 60, 68, 79, 89, 96, 99, 99],   ; age 67
        [48, 49, 53, 58, 67, 78, 88, 96, 99, 99],   ; age 68
        [47, 48, 51, 57, 66, 77, 88, 96, 99, 99],   ; age 69
        [45, 46, 49, 55, 64, 76, 87, 96, 99, 99],   ; age 70
        [43, 44, 48, 53, 63, 75, 86, 95, 99, 99],   ; age 71
        [41, 42, 46, 52, 61, 74, 86, 95, 99, 99],   ; age 72
        [39, 40, 44, 50, 60, 73, 85, 95, 99, 99],   ; age 73
        [37, 39, 42, 48, 59, 72, 85, 94, 99, 99],   ; age 74
        [36, 37, 40, 47, 57, 71, 84, 94, 99, 99],   ; age 75
        [34, 35, 39, 45, 55, 70, 83, 94, 99, 99],   ; age 76
        [32, 33, 37, 43, 54, 68, 83, 93, 99, 99],   ; age 77
        [30, 31, 35, 41, 52, 67, 82, 93, 99, 99],   ; age 78
        [28, 29, 33, 39, 51, 66, 81, 93, 99, 99],   ; age 79
        [26, 27, 31, 38, 49, 64, 80, 92, 98, 99],   ; age 80
        [24, 25, 29, 36, 47, 63, 79, 91, 98, 99],   ; age 81
        [22, 23, 27, 34, 45, 62, 78, 91, 98, 99],   ; age 82
        [20, 21, 25, 32, 43, 60, 77, 90, 98, 99],   ; age 83
        [18, 19, 23, 30, 42, 58, 75, 90, 98, 99]   ; age 84
    ]
    a["Female|Black"] := [
        [91, 92, 93, 95, 97, 98, 99, 99, 99, 99],   ; age 45
        [90, 91, 92, 94, 96, 98, 99, 99, 99, 99],   ; age 46
        [89, 90, 91, 93, 96, 98, 99, 99, 99, 99],   ; age 47
        [88, 89, 90, 92, 95, 98, 99, 99, 99, 99],   ; age 48
        [86, 87, 89, 91, 94, 97, 99, 99, 99, 99],   ; age 49
        [85, 86, 88, 91, 94, 97, 99, 99, 99, 99],   ; age 50
        [84, 85, 87, 89, 93, 96, 99, 99, 99, 99],   ; age 51
        [83, 84, 86, 88, 92, 96, 98, 99, 99, 99],   ; age 52
        [81, 82, 84, 87, 91, 95, 98, 99, 99, 99],   ; age 53
        [80, 81, 83, 86, 90, 95, 98, 99, 99, 99],   ; age 54
        [79, 80, 82, 85, 90, 94, 98, 99, 99, 99],   ; age 55
        [77, 79, 81, 84, 89, 94, 97, 99, 99, 99],   ; age 56
        [76, 77, 79, 83, 88, 93, 97, 99, 99, 99],   ; age 57
        [75, 76, 78, 82, 86, 92, 97, 99, 99, 99],   ; age 58
        [73, 74, 77, 80, 85, 91, 96, 99, 99, 99],   ; age 59
        [72, 73, 75, 79, 84, 90, 95, 98, 99, 99],   ; age 60
        [70, 71, 73, 77, 82, 89, 95, 98, 99, 99],   ; age 61
        [68, 69, 71, 75, 81, 88, 94, 98, 99, 99],   ; age 62
        [66, 67, 69, 73, 79, 86, 93, 97, 99, 99],   ; age 63
        [64, 65, 68, 71, 77, 85, 92, 97, 99, 99],   ; age 64
        [62, 63, 66, 69, 76, 84, 91, 97, 99, 99],   ; age 65
        [60, 61, 63, 67, 74, 82, 90, 96, 99, 99],   ; age 66
        [58, 59, 61, 65, 72, 81, 89, 96, 99, 99],   ; age 67
        [56, 56, 59, 63, 69, 79, 88, 95, 99, 99],   ; age 68
        [53, 54, 56, 60, 67, 77, 86, 94, 99, 99],   ; age 69
        [50, 51, 54, 58, 65, 75, 85, 94, 99, 99],   ; age 70
        [48, 49, 51, 55, 62, 73, 84, 93, 99, 99],   ; age 71
        [45, 46, 48, 53, 60, 72, 83, 93, 98, 99],   ; age 72
        [43, 44, 46, 51, 58, 70, 82, 92, 98, 99],   ; age 73
        [40, 41, 44, 48, 56, 69, 81, 92, 98, 99],   ; age 74
        [38, 39, 41, 46, 54, 67, 80, 91, 98, 99],   ; age 75
        [35, 36, 39, 44, 52, 65, 78, 90, 98, 99],   ; age 76
        [33, 33, 36, 41, 50, 63, 77, 89, 98, 99],   ; age 77
        [30, 31, 33, 39, 48, 61, 76, 89, 98, 99],   ; age 78
        [27, 28, 31, 36, 46, 59, 74, 88, 97, 99],   ; age 79
        [25, 26, 28, 34, 43, 57, 73, 87, 97, 99],   ; age 80
        [22, 23, 26, 31, 41, 56, 72, 87, 97, 99],   ; age 81
        [20, 20, 23, 29, 39, 54, 70, 86, 97, 99],   ; age 82
        [17, 18, 20, 26, 36, 52, 69, 85, 97, 99],   ; age 83
        [14, 15, 18, 24, 34, 50, 68, 84, 96, 99]   ; age 84
    ]
    a["Female|Hispanic"] := [
        [96, 96, 97, 98, 99, 99, 99, 99, 99, 99],   ; age 45
        [95, 95, 96, 97, 98, 99, 99, 99, 99, 99],   ; age 46
        [93, 94, 95, 96, 98, 99, 99, 99, 99, 99],   ; age 47
        [92, 92, 94, 95, 97, 99, 99, 99, 99, 99],   ; age 48
        [90, 91, 92, 94, 97, 98, 99, 99, 99, 99],   ; age 49
        [89, 89, 91, 93, 96, 98, 99, 99, 99, 99],   ; age 50
        [87, 88, 90, 92, 95, 98, 99, 99, 99, 99],   ; age 51
        [86, 87, 89, 91, 94, 97, 99, 99, 99, 99],   ; age 52
        [84, 85, 87, 90, 94, 97, 99, 99, 99, 99],   ; age 53
        [82, 84, 86, 89, 93, 96, 99, 99, 99, 99],   ; age 54
        [81, 82, 84, 88, 92, 96, 98, 99, 99, 99],   ; age 55
        [79, 80, 83, 86, 91, 95, 98, 99, 99, 99],   ; age 56
        [78, 79, 82, 85, 90, 95, 98, 99, 99, 99],   ; age 57
        [76, 78, 80, 84, 89, 94, 98, 99, 99, 99],   ; age 58
        [75, 76, 79, 83, 88, 93, 97, 99, 99, 99],   ; age 59
        [73, 74, 77, 81, 86, 93, 97, 99, 99, 99],   ; age 60
        [71, 72, 75, 79, 85, 91, 96, 99, 99, 99],   ; age 61
        [69, 70, 73, 77, 83, 90, 96, 99, 99, 99],   ; age 62
        [66, 67, 71, 75, 82, 89, 95, 99, 99, 99],   ; age 63
        [64, 65, 68, 73, 80, 88, 95, 98, 99, 99],   ; age 64
        [62, 63, 66, 71, 78, 87, 94, 98, 99, 99],   ; age 65
        [59, 61, 64, 69, 76, 86, 93, 98, 99, 99],   ; age 66
        [57, 58, 62, 67, 74, 84, 92, 98, 99, 99],   ; age 67
        [54, 56, 59, 64, 72, 83, 91, 97, 99, 99],   ; age 68
        [52, 53, 57, 62, 70, 81, 91, 97, 99, 99],   ; age 69
        [50, 51, 54, 60, 68, 79, 89, 96, 99, 99],   ; age 70
        [47, 48, 52, 57, 66, 78, 88, 96, 99, 99],   ; age 71
        [45, 46, 49, 55, 63, 75, 87, 95, 99, 99],   ; age 72
        [42, 43, 47, 52, 61, 73, 85, 94, 99, 99],   ; age 73
        [40, 41, 44, 50, 59, 71, 84, 94, 99, 99],   ; age 74
        [37, 38, 41, 47, 56, 69, 82, 93, 99, 99],   ; age 75
        [35, 36, 39, 45, 54, 67, 81, 92, 98, 99],   ; age 76
        [33, 33, 36, 42, 51, 65, 79, 91, 98, 99],   ; age 77
        [30, 31, 34, 40, 49, 63, 78, 90, 98, 99],   ; age 78
        [28, 29, 31, 37, 47, 61, 76, 89, 98, 99],   ; age 79
        [25, 26, 29, 34, 44, 58, 74, 88, 97, 99],   ; age 80
        [23, 24, 26, 32, 41, 56, 72, 87, 97, 99],   ; age 81
        [20, 21, 24, 29, 39, 54, 70, 85, 97, 99],   ; age 82
        [18, 19, 21, 27, 36, 51, 68, 84, 96, 99],   ; age 83
        [16, 17, 19, 24, 34, 49, 66, 83, 96, 99]   ; age 84
    ]
    return a
}
