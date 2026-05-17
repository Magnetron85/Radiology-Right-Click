; ============================================================
; lib/calc/CalciumScore.ahk -- Agatston score percentile + age
; ------------------------------------------------------------
; Citations:
;   Hoff JA et al. Am J Cardiol 2001;87(12):1335-9.
;   McClelland RL et al. Am J Cardiol 2009;103(1):59-63. (age)
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

CalciumScore_Entry(input) {
    return CalcCalciumScorePercentile(input)
}

CalcCalciumScorePercentile(input) {
    if !RegExMatch(input, "i)Age:\s*(\d+)", &ageM)
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

    return input "`n`nCONTEXT:`n" _HoffScore(age, sex, score)
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
    if (score <= cuts[1])
        return 25
    if (score <= cuts[2])
        return 50
    if (score <= cuts[3])
        return 75
    if (score <= cuts[4])
        return 90
    return 99
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

_DetermineComparison(pct) {
    if (pct <= 25)
        return "Low (<=25%)"
    if (pct <= 50)
        return "Average (25-50%)"
    if (pct <= 75)
        return "Average (50-75%)"
    if (pct <= 90)
        return "High (75-90%)"
    return "Very high (>90%)"
}

_CoronaryAge(score) {
    return Round(39.1 + 7.25 * Ln(score + 1), 0)
}
