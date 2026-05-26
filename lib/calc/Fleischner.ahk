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
#Include ..\FormGui.ahk

Fleischner_Entry(input) {
    ShowFleischnerDialog(input)
    return ""
}

ShowFleischnerDialog(text := "") {
    ; --- pre-fill from highlighted text ---
    sz := TextScan.Size(text)
    comp := TextScan.Composition(text)
    multi := TextScan.ContainsAny(text
        , ["\bmultiple\b","\bseveral\b","\bnumerous\b","\bscattered\b","\bfew\b","\bmany\b","\bnodules\b"])
    solitary := TextScan.ContainsAny(text, ["\bsolitary\b","\bsingle\b","\ba\s+(?:single\s+)?pulmonary\s+nodule\b"])
    multIdx := (multi && !solitary) ? 2 : 1

    compIdx := (comp = "ground glass") ? 2
            : (comp = "part solid")   ? 3
            : 1   ; default solid (most common)

    calcified := TextScan.ContainsAny(text, ["\bcalcified\b","\bcalcification","\bcalcific\b"])
    noncalc   := TextScan.ContainsAny(text, ["\bnoncalcified\b","\bnon-calcified\b"])
    if noncalc
        calcified := false   ; explicit negation wins

    spiculated := TextScan.ContainsAny(text, ["\bspiculat"])
    lobulated  := TextScan.ContainsAny(text, ["\blobulat"])
    morphIdx := spiculated ? 3 : (lobulated ? 2 : 1)

    emphysema := TextScan.ContainsAny(text, ["\bemphysema\b","\bemphysematous\b"])
    fibrosis  := TextScan.ContainsAny(text, ["\bfibrosis\b","\bfibrotic\b","\bUIP\b","\bIPF\b"])
    highRisk  := emphysema || fibrosis || spiculated

    location := ""
    if RegExMatch(text, "i)(right|left)\s+(upper|middle|lower)\s+lobe", &m)
        location := m[0]
    else if RegExMatch(text, "i)(lingula|apical|basal)", &m)
        location := m[0]

    ; --- build the form ---
    form := RadsForm("Fleischner 2017", 580)

    form.Header("Number of nodules")
    form.Dropdown("Mult", "Multiplicity:"
        , ["Solitary (single nodule)"
        ,  "Multiple nodules"], multIdx)

    form.Header("Composition")
    form.Dropdown("Comp", "Composition:"
        , ["Solid"
        ,  "Ground-glass"
        ,  "Part-solid"], compIdx)

    form.Header("Size of dominant nodule")
    form.Numeric("SizeMm", "Largest dimension (mm):", sz.mm > 0 ? Round(sz.mm, 1) : "")

    form.Header("Morphology and additional features")
    form.Dropdown("Morph", "Margin / morphology:"
        , ["Smooth or not specified"
        ,  "Lobulated"
        ,  "Spiculated (high-risk)"
        ,  "Irregular"], morphIdx)
    form.CheckboxRow2("Calcified",  "Calcified"
                   , "Perifissural", "Perifissural", calcified, false)

    form.Header("Patient risk")
    form.Note("Per Fleischner 2017: heavy smoking (>=30 pack-years, or quit within past 15 years), family history of lung cancer, asbestos / radon / uranium exposure, or pulmonary fibrosis / emphysema on the study all contribute to high-risk classification.")
    form.CheckboxRow2("HighRisk",  "High-risk patient (heavy smoker / family hx / occupational exposure)"
                   , "Background", "Emphysema / pulmonary fibrosis on this study"
                   , highRisk, emphysema || fibrosis)

    form.Header("Location (optional)")
    form.Numeric("Loc", "Anatomic location:", location)

    form.SetSubmit(Fleischner_OnSubmit)
    form.AddButtons()
    form.Show()
}

Fleischner_OnSubmit(v, form := "") {
    global g_LastSelectedText
    sizeMm := v.SizeMm + 0
    if (sizeMm <= 0)
        return MakeResult({ impression: "Please enter the nodule size in mm.",
                            error: "Please enter the nodule size in mm." })

    multStr := InStr(v.Mult, "Multiple") ? "Multiple" : "Solitary"
    compStr := InStr(v.Comp, "Ground")    ? "ground glass"
             : InStr(v.Comp, "Part-solid") ? "part-solid"
             : "solid"
    morphPart := InStr(v.Morph, "Spiculated") ? " spiculated"
              : InStr(v.Morph, "Lobulated")  ? " lobulated"
              : InStr(v.Morph, "Irregular")  ? " irregular"
              : ""
    calcPart := v.Calcified ? " calcified" : ""
    locPart  := (v.Loc != "") ? " in the " v.Loc : ""

    sentence := multStr " " sizeMm " mm" morphPart calcPart " " compStr
              . " pulmonary nodule" (multStr = "Multiple" ? "s" : "") locPart "."
    if (v.Background)
        sentence .= " The lungs show emphysema."
    fullText := ProcessNodules(sentence)

    ; Strip the inline citation (now a structured field).
    citPos := InStr(fullText, "`n`nCitation:")
    if (citPos > 0)
        fullText := SubStr(fullText, 1, citPos - 1)
    fullText := RTrim(fullText, " `t`r`n")

    ; Build a single-sentence impression. The ProcessNodules output has a
    ; "FLEISCHNER SOCIETY RECOMMENDATION:" section that contains the
    ; report-ready prose; everything above it is descriptive detail that
    ; belongs in methodology. Parse the rec block, strip the header /
    ; low-risk / high-risk labels, and collapse to a single paragraph.
    impression := ""
    recHdr := "FLEISCHNER SOCIETY RECOMMENDATION:`n"
    recPos := InStr(fullText, recHdr)
    notePos := InStr(fullText, "`n`nNote: Patient has")
    advisories := []
    if (recPos > 0) {
        endPos := notePos > 0 ? notePos : StrLen(fullText) + 1
        recBlock := SubStr(fullText, recPos + StrLen(recHdr)
                         , endPos - recPos - StrLen(recHdr))
        recBlock := Trim(recBlock, " `t`r`n")

        ; Strip the "Follow-up dates: ..." narrative that _AddFollowUpDates
        ; appended. The dates are useful for scheduling and stay in the
        ; methodology block (which renders the full ProcessNodules text);
        ; in the impression they're noise -- duplicated across both risk
        ; branches and verbose ("November 2026 to May 2027 from May 2026").
        recBlock := RegExReplace(recBlock, "\s*Follow-up dates:[^`r`n]*", "")

        ; ProcessNodules emits BOTH "For low-risk patients:" and
        ; "For high-risk patients:" branches for solid uncalcified nodules
        ; because the radiologist often doesn't know the patient's clinical
        ; risk profile (smoking history, occupational exposure, family hx).
        ; Fleischner convention: emit BOTH recs when risk is unknown; the
        ; ordering clinician then picks the relevant one.
        ;
        ; If the user has affirmatively indicated high risk via the form
        ; (HighRisk checkbox -- patient-level indicators; or Background --
        ; emphysema/fibrosis on this study), collapse to the high-risk
        ; branch only. Otherwise keep both labeled.
        affirmHigh := v.HighRisk || v.Background
        lPos := InStr(recBlock, "For low-risk patients:")
        hPos := InStr(recBlock, "For high-risk patients:")
        hasBoth := (lPos > 0 && hPos > 0)

        compStr := InStr(v.Comp, "Ground")    ? "ground-glass"
                 : InStr(v.Comp, "Part-solid") ? "part-solid"
                                               : "solid"
        multStr := InStr(v.Mult, "Multiple") ? "multiple pulmonary nodules"
                                             : "solitary pulmonary nodule"
        ; Display size as integer when whole (avoids "6.0 mm" from
        ; pre-filled Round(sz.mm, 1)).
        sizeDisp := (sizeMm = Floor(sizeMm)) ? Integer(sizeMm) : Round(sizeMm, 1)
        sizeStr := sizeDisp " mm " compStr " " multStr

        ; Build the impression as "<nodule>. <follow-up-preamble>: <rec>."
        ; The "per Fleischner 2017" label attaches to the recommendation,
        ; not the nodule -- Fleischner is a follow-up framework, not a
        ; classification system, so tagging the nodule with it reads off.
        if (hasBoth && affirmHigh) {
            recText := _FLE_CleanWhitespace(
                Trim(SubStr(recBlock, hPos + StrLen("For high-risk patients:"))
                   , " `t`r`n"))
            impression := sizeStr ". Per Fleischner 2017 (high-risk patient): " recText
        } else if (hasBoth) {
            lowRec := _FLE_CleanWhitespace(
                Trim(SubStr(recBlock, lPos + StrLen("For low-risk patients:")
                          , hPos - lPos - StrLen("For low-risk patients:"))
                   , " `t`r`n"))
            highRec := _FLE_CleanWhitespace(
                Trim(SubStr(recBlock, hPos + StrLen("For high-risk patients:"))
                   , " `t`r`n"))
            impression := sizeStr
                       . ". Fleischner 2017 follow-up stratified by patient risk -- "
                       . "low-risk: " RTrim(lowRec, ".")
                       . "; high-risk: " RTrim(highRec, ".") "."
        } else {
            recText := _FLE_CleanWhitespace(recBlock)
            if InStr(recText, "calcified nodules do not") {
                ; Calcified path -- the rec is a self-contained sentence;
                ; just append "per Fleischner 2017" before the period.
                impression := sizeStr ". " RTrim(recText, ".") " per Fleischner 2017."
            } else {
                preamble := affirmHigh
                    ? "Per Fleischner 2017 (high-risk patient): "
                    : "Per Fleischner 2017: "
                impression := sizeStr ". " preamble recText
            }
        }

        ; Collapse any accidental double periods from rec strings already
        ; ending in a period.
        impression := StrReplace(impression, "..", ".")
        impression := RTrim(impression, " `t`r`n")
        if (notePos > 0) {
            noteBlock := Trim(SubStr(fullText, notePos), " `t`r`n")
            noteBlock := RegExReplace(noteBlock, "^Note:\s*", "")
            advisories.Push(noteBlock)
        }
    } else {
        impression := fullText
    }

    return MakeResult({
        classification: "Fleischner 2017",
        impression:     impression,
        recommendation: "",
        advisories:     advisories,
        methodology:    fullText,
        citations:      [{ text: "MacMahon H, Naidich DP, Goo JM, et al. "
                                . "Guidelines for Management of Incidental Pulmonary Nodules "
                                . "Detected on CT Images: From the Fleischner Society 2017. "
                                . "Radiology. 2017;284(1):228-243.",
                           url:  "https://pubs.rsna.org/doi/10.1148/radiol.2017161659" }],
        echo:           g_LastSelectedText
    })
}

; Collapse newlines + repeated whitespace inside a Fleischner rec block so
; it reads as one sentence in the impression line.
_FLE_CleanWhitespace(s) {
    s := RegExReplace(s, "[`r`n]+", " ")
    s := RegExReplace(s, "\s{2,}", " ")
    return Trim(s)
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
