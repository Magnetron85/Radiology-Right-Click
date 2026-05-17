; ============================================================
; lib/calc/Fleischner.ahk -- Fleischner 2017 nodule criteria
; ------------------------------------------------------------
; Citation: MacMahon H, Naidich DP, Goo JM, et al. Guidelines
; for Management of Incidental Pulmonary Nodules Detected on
; CT Images: From the Fleischner Society 2017. Radiology.
; 2017;284(1):228-243. doi:10.1148/radiol.2017161659.
;
; The Nodule class parses a sentence to infer multiplicity,
; composition (solid / ground-glass / part-solid), calcification,
; location, morphology, and at least one numeric measurement.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

Fleischner_Entry(input) {
    return ProcessNodules(input)
}

; ---- recommendations table -------------------------------------------------

FleischnerRecommendations() {
    static t := _BuildFleischnerRecs()
    return t
}

_BuildFleischnerRecs() {
    r := Map()
    r["single_solid_small"]   := Map("low","No routine follow-up.", "high","Optional CT at 12 months.")
    r["single_solid_medium"]  := Map("low","CT at 6-12 months, then consider CT at 18-24 months.", "high","CT at 6-12 months, then CT at 18-24 months.")
    r["single_solid_large"]   := Map("low","Consider CT at 3 months, PET/CT, or tissue sampling.", "high","Consider CT at 3 months, PET/CT, or tissue sampling.")
    r["multiple_solid_small"] := Map("low","No routine follow-up.", "high","Optional CT at 12 months.")
    r["multiple_solid_medium"]:= Map("low","CT at 3-6 months, then consider CT at 18-24 months.", "high","CT at 3-6 months, then at 18-24 months.")
    r["multiple_solid_large"] := Map("low","CT at 3-6 months, then consider CT at 18-24 months.", "high","CT at 3-6 months, then at 18-24 months.")
    r["single_gg_small"]      := "No routine follow-up."
    r["single_gg_large"]      := "CT at 6-12 months to confirm persistence, then every 2 years until 5 years."
    r["single_ps_small"]      := "No routine follow-up."
    r["single_ps_large"]      := "CT at 3-6 months to confirm persistence. If unchanged and solid component remains <6mm, annual CT should be performed for 5 years."
    r["multiple_subsolid_small"] := "CT at 3-6 months. If stable, consider CT at 2 and 4 years."
    r["multiple_subsolid_large"] := "CT at 3-6 months. Subsequent management based on the most suspicious nodule(s)."
    return r
}

; ---- Nodule class ----------------------------------------------------------

class Nodule {
    Description := ""
    Composition := ""
    Calcified   := false
    mString     := ""
    Units       := ""
    Measurements := ""
    OriginalMeasurements := ""
    OriginalUnits := ""
    HighRisk    := false
    Location    := ""
    Multiplicity := "single"
    Perifissural := false
    Morphology  := ""
    GlobalHighRisk := false

    __New(nString) {
        this.Description := nString
        this.Measurements := []
        if !this._ContainsNoduleReference(nString)
            throw Error("No nodule reference: " nString)
        this._ParseProperties(nString)
        this._ExtractMeasurements(nString)
        if (InStr(nString, "up to") && InStr(nString, "multiple"))
            this.Multiplicity := "multiple"
        this.OriginalMeasurements := this.Measurements.Clone()
        this.OriginalUnits := this.Units
    }

    _ContainsNoduleReference(s) {
        for _, term in ["nodule","nodules","mass","masses","opacity","opacities"
                       ,"lesion","lesions","micronodule","micronodules"] {
            if InStr(s, term)
                return true
        }
        return false
    }

    _ParseProperties(s) {
        words := StrSplit(s, A_Space)
        for i, word in words {
            if this._IsMultiplicityWord(word)
                this.Multiplicity := "multiple"

            if FuzzyMatch(word, "solid") {
                if (this.Composition = "")
                    this.Composition := "solid"
                else if (this.Composition = "ground glass")
                    this.Composition := "part solid"
            }

            if (FuzzyMatch(word, "ground") && i < words.Length && FuzzyMatch(words[i+1], "glass")) {
                if (this.Composition = "")
                    this.Composition := "ground glass"
                else if (this.Composition = "solid")
                    this.Composition := "part solid"
            }

            if FuzzyMatch(word, "groundglass") {
                if (this.Composition = "")
                    this.Composition := "ground glass"
                else if (this.Composition = "solid")
                    this.Composition := "part solid"
            }

            if ((FuzzyMatch(word, "part") && i < words.Length && FuzzyMatch(words[i+1], "solid"))
              || FuzzyMatch(word, "part-solid") || FuzzyMatch(word, "partsolid"))
                this.Composition := "part solid"

            if (FuzzyMatch(word, "calcified") || FuzzyMatch(word, "calcification") || FuzzyMatch(word, "calcifications"))
                this.Calcified := true
            if (FuzzyMatch(word, "noncalcified") || FuzzyMatch(word, "non-calcified"))
                this.Calcified := false

            if (FuzzyMatch(word, "emphysema") || FuzzyMatch(word, "fibrosis"))
                this.GlobalHighRisk := true

            for _, kw in ["upper","middle","lower","lingula","apical","basal"] {
                if FuzzyMatch(word, kw)
                    this.Location .= " " word
            }
            if (FuzzyMatch(word, "right") || FuzzyMatch(word, "left"))
                this.Location .= " " word

            if (FuzzyMatch(word, "perifissural") || FuzzyMatch(word, "fissure"))
                this.Perifissural := true

            for _, kw in ["spiculated","lobulated","irregular","smooth"] {
                if FuzzyMatch(word, kw) {
                    this.Morphology := kw
                    break
                }
            }
        }
        this.Location := Trim(this.Location)
        if (InStr(s, "multiple") || InStr(s, "several") || InStr(s, "numerous"))
            this.Multiplicity := "multiple"
    }

    _IsMultiplicityWord(word) {
        for _, w in ["nodules","multiple","several","few","numerous","masses"
                   ,"opacities","lesions","micronodules"] {
            if FuzzyMatch(word, w)
                return true
        }
        return false
    }

    _ExtractMeasurements(s) {
        needle3 := "i)(?:up to|approximately|about|~)?\s*((?:\d*\.)?\d+)\s*(?:x\s*((?:\d*\.)?\d+))?\s*(?:x\s*((?:\d*\.)?\d+))?\s*([cm]m)"
        if RegExMatch(s, needle3, &m) {
            this.mString := m[0]
            this.Units   := m[4]
            loop 3 {
                if (m[A_Index] != "")
                    this.Measurements.Push(m[A_Index] + 0)
            }
            return
        }
        needle1 := "i)(?:up to|approximately|about|~)?\s*((?:\d*\.)?\d+)\s*([cm]m)"
        if RegExMatch(s, needle1, &m) {
            this.mString := m[0]
            this.Units   := m[2]
            this.Measurements.Push(m[1] + 0)
            return
        }
        if (InStr(s, "micronodule") || InStr(s, "micronodules")) {
            this.mString := "5 mm"
            this.Units   := "mm"
            this.Measurements.Push(5)
            return
        }
        ; any number + units
        needleAny := "i)((?:\d*\.)?\d+)\s*([cm]m)"
        if RegExMatch(s, needleAny, &m) {
            this.mString := m[0]
            this.Units   := m[2]
            this.Measurements.Push(m[1] + 0)
            return
        }
        throw Error("No measurements found: " s)
    }

    Size() {
        if (this.Measurements.Length = 0)
            return 0
        t := 0
        for v in this.Measurements
            t += v
        avg := t / this.Measurements.Length
        return (this.Units = "cm") ? avg * 10 : avg  ; mm
    }

    UpdateMString() {
        if (this.OriginalMeasurements.Length = 0)
            return
        parts := []
        for v in this.OriginalMeasurements
            parts.Push(Format("{:.1f}", v))
        this.mString := JoinArr(parts, " x ") . " " . this.OriginalUnits
    }

    Category() {
        s := this.Size()  ; mm
        if (this.Composition = "solid" || this.Composition = "") {
            if (this.Multiplicity = "multiple")
                return (s < 6) ? "multiple_solid_small" : (s <= 8 ? "multiple_solid_medium" : "multiple_solid_large")
            return (s < 6) ? "single_solid_small" : (s <= 8 ? "single_solid_medium" : "single_solid_large")
        } else if (this.Composition = "ground glass") {
            if (this.Multiplicity = "multiple")
                return (s < 6) ? "multiple_subsolid_small" : "multiple_subsolid_large"
            return (s < 6) ? "single_gg_small" : "single_gg_large"
        } else if (this.Composition = "part solid") {
            if (this.Multiplicity = "multiple")
                return (s < 6) ? "multiple_subsolid_small" : "multiple_subsolid_large"
            return (s < 6) ? "single_ps_small" : "single_ps_large"
        }
        return "single_solid_small"
    }

    Recommendation() {
        recs := FleischnerRecommendations()
        cat := this.Category()
        risk := this.HighRisk ? "high" : "low"
        if !recs.Has(cat)
            return "Unable to determine recommendation."
        v := recs[cat]
        return (v is Map) ? v[risk] : v
    }
}

; ---- main processor --------------------------------------------------------

ProcessNodules(text) {
    sentences := _PreprocessFleischner(text)
    globalHighRisk := _CheckHighRiskConditions(text)
    isMultiple := false
    nodules := []

    ; Feed the full sentence to Nodule -- the description-only slice used in
    ; v1 (via SplitNoduleDescriptions) was dropping location words and
    ; "up to N mm" qualifiers before the parser saw them. The class is
    ; tolerant of extra context and only fails when the sentence has no
    ; nodule reference at all, which is exactly when we want to skip it.
    for _, sentence in sentences {
        prob := _MultipleNodulesProbability(sentence)
        isMultiple := (prob >= 0.45) || isMultiple
        try {
            n := Nodule(sentence)
            n.GlobalHighRisk := globalHighRisk
            n.Multiplicity := isMultiple ? "multiple" : "single"
            nodules.Push(n)
        } catch {
            ; skip non-nodule sentences
        }
    }

    if (nodules.Length = 0) {
        if (InStr(text, "micronodule") || InStr(text, "micronodules")
          || (InStr(text, "multiple") && InStr(text, "nodule"))) {
            n := Nodule("Multiple nodules measuring up to 5 mm")
            n.GlobalHighRisk := globalHighRisk
            n.Multiplicity := "multiple"
            nodules.Push(n)
        } else {
            return "Error: No valid nodules found."
        }
    }

    most := nodules[1]
    for n in nodules {
        n.UpdateMString()
        if (n.Size() > most.Size() || (n.Size() = most.Size() && _IsMoreSignificant(n, most)))
            most := n
    }
    if (nodules.Length > 1 || isMultiple || InStr(text, "multiple") || InStr(text, "micronodules"))
        most.Multiplicity := "multiple"

    out := text "`n`n"
    out .= "Fleischner 2017 assessment:`n"
    out .= (most.Multiplicity = "multiple")
        ? "Multiple pulmonary nodules are present. The most significant nodule forming the basis of follow-up:`n"
        : "A solitary pulmonary nodule is described forming the basis of follow-up:`n"
    out .= "- Location: " (most.Location != "" ? most.Location : "Not specified") "`n"
    out .= "- Extracted (or inferred) Size: " most.mString "`n"
    out .= "- Fleischner Size (mm): " Format("{:.1f} mm", most.Size()) "`n"
    out .= "- Composition (or inferred): " (most.Composition != "" ? most.Composition : "solid") "`n"
    if most.Calcified
        out .= "- Calcification: Present`n"
    if (most.Morphology != "")
        out .= "- Morphology: " most.Morphology "`n"

    currentDate := A_Now
    out .= "`nFLEISCHNER SOCIETY RECOMMENDATION:`n"
    if ((globalHighRisk || most.Morphology = "spiculated") && !most.Calcified) {
        most.HighRisk := true
        out .= _AddFollowUpDates(most.Recommendation(), currentDate)
    } else {
        if (!most.Calcified && (most.Composition = "solid" || most.Composition = "")) {
            lowRec := most.Recommendation()
            out .= "For low-risk patients: " _AddFollowUpDates(lowRec, currentDate) "`n"
            most.HighRisk := true
            out .= "`nFor high-risk patients: " _AddFollowUpDates(most.Recommendation(), currentDate) "`n"
        } else if (!most.Calcified && (most.Composition = "part solid" || most.Composition = "ground glass")) {
            out .= _AddFollowUpDates(most.Recommendation(), currentDate)
        } else {
            out .= "Incidental calcified nodules do not typically require routine follow up.`n"
        }
    }

    if (globalHighRisk || most.Morphology = "spiculated")
        out .= "`n`nNote: Patient has risk factors or morphologic characteristics that may increase the risk of lung cancer."

    if Prefs.Get("display","showCitations", true)
        out .= "`n`nCitation: MacMahon H, Naidich DP, Goo JM, et al. Guidelines for Management of Incidental Pulmonary Nodules Detected on CT Images: From the Fleischner Society 2017. Radiology. 2017;284(1):228-243. doi:10.1148/radiol.2017161659"

    return out
}

; ---- helpers ---------------------------------------------------------------

_PreprocessFleischner(text) {
    sentences := []
    for _, line in StrSplit(text, "`n", "`r") {
        line := RegExReplace(line, "^\s*[\*\-•]\s*", "")
        parts := StrSplit(line, ".")
        for i, part in parts {
            part := Trim(part)
            if (i > 1 && RegExMatch(parts[i-1], "\d+$") && RegExMatch(part, "^\d+"))
                sentences[sentences.Length] .= "." part
            else if (part != "")
                sentences.Push(part)
        }
    }
    return sentences
}

_CheckHighRiskConditions(text) {
    low := StrLower(text)
    for _, term in ["emphysema","fibrosis","fibrotic","emphysematous"] {
        for _, word in StrSplit(low, A_Space) {
            if FuzzyMatch(word, term)
                return true
        }
    }
    return false
}

_MultipleNodulesProbability(text) {
    p := 0
    if (InStr(text, "nodules") || InStr(text, "multiple") || InStr(text, "several") || InStr(text, "few"))
        p += 0.7
    if (InStr(text, "scattered") && InStr(text, "micronodules"))
        p += 0.6

    count := 0
    pos := 1
    while (pos := RegExMatch(text, "i)(\d+(?:\.\d+)?\s*(mm|cm))", &m, pos)) {
        count++
        pos += m.Len[0] ? m.Len[0] : 1
    }
    if (count > 1)
        p += 0.5

    locs := Map()
    for _, kw in ["upper","middle","lower","lingula","apical","basal","right","left"] {
        if InStr(text, kw)
            locs[kw] := true
    }
    if (locs.Count > 1)
        p += 0.3

    noduleHits := 0
    pos := 1
    while (pos := InStr(text, "nodule", false, pos)) {
        noduleHits++
        pos += 6
    }
    if (noduleHits > 1)
        p += 0.4

    return (p > 1) ? 1 : p
}

_SplitNoduleDescriptions(text) {
    out := []
    needle := "i)(?:(?:\d+(?:\.\d+)?\s*(?:x\s*\d+(?:\.\d+)?)*\s*(?:mm|cm))\s*(?:[a-z\s]+\s+)?(?:pulmonary\s+)?(?:nodule|mass|opacity|lesion)s?|(?:(?:multiple\s+)?(?:nodule|mass|opacity|lesion)s?(?:\s+(?:up to|approximately|about|~)?)?\s*(?:\d+(?:\.\d+)?\s*(?:x\s*\d+(?:\.\d+)?)*\s*(?:mm|cm))?|(?:(?:up to|approximately|about|~)?\s*\d+(?:\.\d+)?\s*(?:x\s*\d+(?:\.\d+)?)*\s*(?:mm|cm))?\s*(?:(?:solid|ground glass|part[- ]solid|calcified|noncalcified|spiculated|lobulated|irregular|smooth)?\s*(?:pulmonary\s+)?(?:nodule|mass|opacity|lesion)s?)))(?:\s*\([^)]+\))?"
    for _, line in StrSplit(text, "`n", "`r") {
        pos := 1
        while (pos := RegExMatch(line, needle, &m, pos)) {
            out.Push(Trim(m[0]))
            pos += m.Len[0] ? m.Len[0] : 1
        }
    }
    if (out.Length = 0) {
        if RegExMatch(text, "i)(?:multiple\s+)?nodules?.*(?:measure|up to).*(\d+(?:\.\d+)?)\s*(?:x\s*\d+(?:\.\d+)?)*\s*(mm|cm)", &m)
            out.Push(m[0])
        else if (InStr(text, "micronodule") || InStr(text, "micronodules"))
            out.Push("Multiple micronodules")
        else if (InStr(text, "multiple") && InStr(text, "nodule"))
            out.Push("Multiple nodules")
    }
    if (out.Length = 0)
        out.Push(text)
    return out
}

_IsMoreSignificant(a, b) {
    order := Map("solid", 3, "part solid", 2, "ground glass", 1, "", 0)
    aOrd := order.Has(a.Composition) ? order[a.Composition] : 0
    bOrd := order.Has(b.Composition) ? order[b.Composition] : 0
    return aOrd > bOrd
}

_AddFollowUpDates(rec, currentDate) {
    periods := []
    if InStr(rec, "3 months")
        periods.Push({ min: 90,  max: 90,  text: "3 months" })
    if InStr(rec, "3-6 months")
        periods.Push({ min: 90,  max: 183, text: "3-6 months" })
    if InStr(rec, "6-12 months")
        periods.Push({ min: 180, max: 365, text: "6-12 months" })
    if InStr(rec, "18-24 months")
        periods.Push({ min: 540, max: 730, text: "18-24 months" })
    if InStr(rec, "2 years until 5 years")
        periods.Push({ min: 1095, max: 1825, text: "2 years until 5 years" })

    if (periods.Length > 0) {
        rec .= "`nFollow-up dates:"
        nowFmt := FormatTime(currentDate, "MMMM yyyy")
        for p in periods {
            minD := DateAdd(currentDate, p.min, "days")
            maxD := DateAdd(currentDate, p.max, "days")
            minFmt := FormatTime(minD, "MMMM yyyy")
            maxFmt := FormatTime(maxD, "MMMM yyyy")
            if (minFmt != maxFmt)
                rec .= " " p.text " is " minFmt " to " maxFmt " from " nowFmt ". "
            else
                rec .= " " p.text " is " minFmt " from " nowFmt ". "
        }
    }
    return rec
}
