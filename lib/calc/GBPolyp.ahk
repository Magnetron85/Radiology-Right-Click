; ============================================================
; lib/calc/GBPolyp.ahk -- SRU 2022 gallbladder polyp consensus
; ------------------------------------------------------------
; Reference: Kamaya A et al. Society of Radiologists in
; Ultrasound Consensus Conference Statement on the Management
; of Gallbladder Polyps. Radiology. 2022;305(2):277-289.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

GBPolyp_Entry(input) {
    ShowGBPolypDialog(input)
    return ""
}

ShowGBPolypDialog(text := "") {
    sz := TextScan.Size(text)
    morphIdx := 6
    if TextScan.ContainsAny(text, ["\bball[- ]on[- ]the[- ]wall\b"])
        morphIdx := 1
    else if TextScan.ContainsAny(text, ["\bpedunculated\b.*\bthin\s+stalk\b","\bthin\s+stalk\b"])
        morphIdx := 2
    else if TextScan.ContainsAny(text, ["\bpedunculated\b.*\bthick\s+stalk\b","\bthick\s+stalk\b","\bwide\s+stalk\b"])
        morphIdx := 3
    else if TextScan.ContainsAny(text, ["\bsessile\b","\bbroad[- ]based\b"])
        morphIdx := 4
    else if TextScan.ContainsAny(text, ["\bfocal\s+wall\s+thickening\b","\bwall\s+thickening\s+adjacent"])
        morphIdx := 5
    hasPSC := TextScan.ContainsAny(text
        , ["\bprimary\s+sclerosing\s+cholangitis\b","\bPSC\b"])
    poorViz := TextScan.ContainsAny(text
        , ["\btechnically\s+(?:inadequate|limited)","\bpoor\s+visualization","\blimited\s+visualization"])

    form := RadsForm("Gallbladder Polyp (SRU 2022)", 540)

    ; Progressive-disclosure layout. The exclusion gate comes FIRST: when
    ; any exclusion is ticked the SRU algorithm does not apply (the
    ; classifier short-circuits on psc / susp / poor before reading any
    ; other input), so every other section is HIDDEN -- the form collapses
    ; to just the exclusions. Hidden inputs revert to their defaults, so
    ; Gui.Submit() never returns a stale answer and the "Selected inputs:"
    ; echo stays honest.
    form.Header("Exclusions (algorithm does NOT apply)")
    form.Checkbox("PSC", "Primary sclerosing cholangitis", hasPSC)
    form.Checkbox("Susp"
        , "Suspicious features: invasion, liver masses, biliary obstruction, pathologic nodes"
        , false)
    form.Checkbox("PoorViz", "Technically inadequate (poor visualization)", poorViz)

    form.Header("Polyp")
    form.Numeric("SizeMm", "Polyp size (mm):", sz.mm > 0 ? Round(sz.mm) : 0)
    form.Dropdown("Morph", "Morphology:"
        , ["Pedunculated with ball-on-the-wall appearance (extremely low risk)"
        ,  "Pedunculated with thin stalk (extremely low risk)"
        ,  "Pedunculated with thick or wide stalk (low risk)"
        ,  "Sessile / broad-based attachment (low risk)"
        ,  "Focal wall thickening >=4 mm adjacent (indeterminate)"
        ,  "Unknown / cannot characterize"], morphIdx)

    form.Header("Adjacent wall")
    form.Checkbox("WallThick", "Focal wall thickening >=4 mm adjacent to polyp", false)

    form.Header("Patient factors")
    form.Checkbox("HighRiskEth"
        , "High-risk geographic / genetic background (North/South American Indigenous, North Indian, Japanese, or Hispanic American populations -- elevated GBC incidence)"
        , false)

    form.Header("Prior comparison (optional)")
    form.Numeric("PriorMm",     "Prior size (mm, 0 if no prior):", 0)
    form.Numeric("PriorMonths", "Months since prior (0 if no prior):", 0)

    form.OnChange("PSC",     _GBP_UpdateForm)
    form.OnChange("Susp",    _GBP_UpdateForm)
    form.OnChange("PoorViz", _GBP_UpdateForm)
    _GBP_UpdateForm(form)

    form.SetSubmit(GBPolyp_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_GBP_UpdateForm(frm) {
    excluded := !!frm.GetValue("PSC") || !!frm.GetValue("Susp")
             || !!frm.GetValue("PoorViz")
    applies := !excluded
    frm.SetSectionVisible("Polyp", applies)
    frm.SetSectionVisible("Adjacent wall", applies)
    frm.SetSectionVisible("Patient factors", applies)
    frm.SetSectionVisible("Prior comparison (optional)", applies)
}

GBPolyp_OnSubmit(v, form := "") {
    global g_LastSelectedText
    size := SafeInt(v.SizeMm, 0)
    morph := _GBP_MorphKey(v.Morph)
    wallThick := !!v.WallThick
    psc := !!v.PSC
    susp := !!v.Susp
    poor := !!v.PoorViz
    highEth := !!v.HighRiskEth
    priorMm := SafeInt(v.PriorMm, 0)
    priorMonths := v.PriorMonths + 0.0

    r := _GBP_Classify(size, morph, wallThick, psc, susp, poor, highEth, priorMm, priorMonths)

    ; --- Build impression sentence (report-ready) ---
    ; Excluded paths (PSC, suspicious features, poor viz) -> the impression
    ; explains why the SRU algorithm doesn't apply rather than a category.
    impression := _GBP_Impression(size, morph, r)
    ; Option B: impression already ends with the recommendation phrase
    ; ("...recommend follow-up US at 12 months."), so leave `recommendation`
    ; empty to avoid double-pasting it into the clipboard.
    recommendation := ""

    ; --- Methodology block (selected inputs + reasoning) ---
    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nRisk category: " r.riskDesc
    if (r.reason != "")
        method .= "`nReason: " r.reason
    if (r.recommendation != "")
        method .= "`nFull recommendation: " r.recommendation
    if (r.followup != "")
        method .= "`nFollow-up schedule (months): " r.followup

    advisories := []
    ; The classifier embeds the high-risk-ethnicity advisory inside
    ; recommendation text; surface it as an advisory if present.
    ; (No-op for now -- the classifier already inlines it; left here
    ;  as the pattern for advisories so Stage 2 can split them out.)

    return MakeResult({
        classification: r.riskDesc,
        impression:     impression,
        recommendation: recommendation,
        advisories:     advisories,
        methodology:    method,
        citations:      [{ text: "Kamaya A, Fung C, Szpakowski JL, et al. "
                                . "Management of Incidentally Detected Gallbladder Polyps: "
                                . "Society of Radiologists in Ultrasound Consensus Conference "
                                . "Recommendations. Radiology. 2022;305(2):277-289.",
                           url:  "https://pubs.rsna.org/doi/10.1148/radiol.213079" }],
        echo:           g_LastSelectedText
    })
}

_GBP_Impression(size, morph, r) {
    ; Exclusion paths -- the algorithm doesn't apply.
    if (r.excluded) {
        if (r.excludeReason = "PSC")
            return "Gallbladder polyp in the setting of primary sclerosing cholangitis; SRU 2022 algorithm does not apply. Refer to GI for management per AGA/ACG guidelines."
        if (r.excludeReason = "Suspicious features")
            return "Gallbladder polyp with suspicious features (wall invasion, liver mass, biliary obstruction, or pathologic lymphadenopathy); SRU 2022 algorithm does not apply. Recommend oncologic consultation."
        if (r.excludeReason = "Poor visualization")
            return "Gallbladder polyp seen on technically limited examination; SRU 2022 algorithm does not apply. Recommend repeat US in 1-2 months with optimized technique, or alternatively CEUS or MRI."
        return "Gallbladder polyp -- algorithm not applicable. " r.reason
    }

    morphPhrase := _GBP_MorphPhrase(morph)
    riskWord := _GBP_RiskWord(r.riskCategory)

    impr := size " mm " morphPhrase " gallbladder polyp"
    if (riskWord != "")
        impr .= ", " riskWord " per SRU 2022"

    ; Always end the impression with the recommendation (active voice).
    rec := _GBP_RecommendationShort(r)
    if (rec != "")
        impr .= "; " rec
    return impr "."
}

_GBP_MorphPhrase(morph) {
    ; Body on a separate line from `if` -- AHK v2's inline-if form
    ; (`if (cond) return ...`) silently misparses; documented in
    ; feedback_ahk_v2_inline_if.
    if (morph = "ball_on_wall")
        return "pedunculated (ball-on-the-wall morphology)"
    if (morph = "thin_stalk")
        return "pedunculated thin-stalk"
    if (morph = "thick_stalk")
        return "pedunculated thick-stalk"
    if (morph = "sessile")
        return "sessile"
    if (morph = "wall_thickening")
        return "with adjacent focal wall thickening"
    return "unspecified-morphology"
}

_GBP_RiskWord(cat) {
    if (cat = "EXTREMELY_LOW")
        return "extremely low risk"
    if (cat = "LOW")
        return "low risk"
    if (cat = "INDETERMINATE")
        return "indeterminate risk"
    if (cat = "SURGICAL")
        return "surgical-consultation indication"
    return ""
}

; Short recommendation phrase suitable for the impression line. The full
; explanatory text (incl. ethnic-modifier note) goes in `methodology`.
_GBP_RecommendationShort(r) {
    if (r.followup != "")
        return "recommend follow-up US at " r.followup " months"
    if (r.riskCategory = "SURGICAL")
        return "recommend surgical consultation"
    if (r.riskCategory = "EXTREMELY_LOW" || r.riskCategory = "LOW")
        return "no follow-up recommended"
    return ""
}

; Separate `recommendation` field -- imperative, self-contained.
_GBP_Recommendation(r) {
    if (r.excluded)
        return ""   ; impression already states the action
    if (r.riskCategory = "SURGICAL")
        return "Recommend surgical consultation."
    if (r.followup != "")
        return "Recommend follow-up ultrasound at " r.followup " months."
    return "No follow-up imaging recommended."
}

_GBP_MorphKey(label) {
    if InStr(label, "ball-on-the-wall")
        return "ball_on_wall"
    if InStr(label, "thin stalk")
        return "thin_stalk"
    if InStr(label, "thick / wide")
        return "thick_stalk"
    if InStr(label, "Sessile")
        return "sessile"
    if InStr(label, "wall thickening")
        return "wall_thickening"
    return "unknown"
}

_GBP_Classify(size, morph, wallThick, psc, susp, poor, highEth, priorMm, priorMonths) {
    if psc
        return _GBP_R("EXCLUDED", "", "Primary sclerosing cholangitis -- refer to GI specialty guidelines (AGA/ACG). Gallbladder polyps in PSC carry elevated malignancy risk.", "", "PSC")
    if susp
        return _GBP_R("EXCLUDED", "", "Suspicious features present (wall invasion, liver masses, biliary obstruction, or pathologic lymphadenopathy). Refer to oncologic specialist.", "", "Suspicious features")
    if poor
        return _GBP_R("EXCLUDED", "", "Technically inadequate visualization. Repeat US within 1-2 months with optimized technique, patient preparation, and Doppler; alternatively consider CEUS or MRI.", "", "Poor visualization")

    ; Rapid growth -> surgical (Kamaya 2022, p. 6: ">=4 mm increase in
    ; <=12 months" triggers surgical consultation regardless of current
    ; morphology or size). Decrease-by-4-mm is NOT a source rule; that
    ; branch was removed since the SRU paper says nothing about polyp
    ; regression and surveillance discontinuation.
    if (priorMm > 0 && priorMonths > 0) {
        growth := size - priorMm
        if (growth >= 4 && priorMonths <= 12)
            return _GBP_R("SURGICAL", "Rapid growth", "Rapid growth detected: " growth " mm increase in " Round(priorMonths, 0) " months (>=4 mm in <=12 months). Surgical consultation recommended.", "", "")
    }

    ; Base risk from morphology
    risk := "LOW"
    if (wallThick || morph = "wall_thickening")
        risk := "INDETERMINATE"
    else if (morph = "ball_on_wall" || morph = "thin_stalk")
        risk := "EXTREMELY_LOW"
    else if (morph = "thick_stalk" || morph = "sessile")
        risk := "LOW"
    ; unknown -> LOW (SRU default)

    ; Kamaya 2022: "if known, geographic and genetic patient factors MAY
    ; increase polyp risk stratification UP TO the low risk category."
    ; The wording is permissive, not mandatory -- some radiologists will
    ; apply it, others will not. Rather than silently force-upgrading
    ; extremely-low to low, surface the option as an advisory note and
    ; leave the categorization to the reader.
    ethNote := ""
    if (highEth && risk = "EXTREMELY_LOW")
        ethNote := "High-risk geographic/genetic background ticked. Per Kamaya 2022 you MAY upgrade this polyp's stratification up to Low Risk (optional, not mandatory). At <=6 mm this would still produce a no-follow-up recommendation; the upgrade only changes management for polyps in the 7-9 mm range (low risk -> 12-mo US instead of no follow-up)."
    else if (highEth && (risk = "LOW" || risk = "INDETERMINATE"))
        ethNote := "High-risk geographic/genetic background ticked. No further upgrade is offered: per Kamaya 2022 the ethnicity modifier is capped at Low Risk and this polyp is already at or above that category."

    return _GBP_BySize(risk, size, highEth, ethNote)
}

_GBP_BySize(risk, size, highEth, ethNote := "") {
    upg := ethNote != "" ? " " ethNote : ""
    if (risk = "EXTREMELY_LOW") {
        if (size <= 9)
            return _GBP_R("EXTREMELY_LOW", "Extremely low risk polyp <=9 mm", "No follow-up recommended." upg, "", "")
        if (size <= 14)
            return _GBP_R("EXTREMELY_LOW", "Extremely low risk polyp 10-14 mm", "Follow-up US at 6, 12, and 24 months. If stable at end of schedule, discontinue surveillance (max ~3 years per SRU 2022)." upg, "6, 12, 24", "")
        return _GBP_R("SURGICAL", "Size >=15 mm", "Polyp >=15 mm. Surgical consultation recommended." upg, "", "")
    }
    if (risk = "LOW") {
        if (size <= 6)
            return _GBP_R("LOW", "Low risk polyp <=6 mm", "No follow-up recommended." upg, "", "")
        if (size <= 9)
            return _GBP_R("LOW", "Low risk polyp 7-9 mm", "Follow-up US at 12 months." upg, "12", "")
        if (size <= 14)
            return _GBP_R("LOW", "Low risk polyp 10-14 mm", "Follow-up US at 6, 12, 24, and 36 months. Surgical consultation is an acceptable alternative. If stable at end of schedule, discontinue surveillance (max ~3 years per SRU 2022)." upg, "6, 12, 24, 36", "")
        return _GBP_R("SURGICAL", "Size >=15 mm", "Polyp >=15 mm. Surgical consultation recommended." upg, "", "")
    }
    ; INDETERMINATE
    if (size <= 6)
        return _GBP_R("INDETERMINATE", "Indeterminate risk polyp <=6 mm", "Follow-up US at 6, 12, 24, and 36 months. If stable at end of schedule, discontinue surveillance (max ~3 years per SRU 2022)." upg, "6, 12, 24, 36", "")
    return _GBP_R("SURGICAL", "Indeterminate risk, size >=7 mm", "Indeterminate risk polyp >=7 mm. Surgical consultation recommended." upg, "", "")
}

_GBP_R(code, reason, rec, followup, exclude) {
    cats := _GBP_Cats()
    info := cats.Has(code) ? cats[code] : cats["LOW"]
    return { riskCategory: code
           , riskDesc: info.desc
           , reason: reason
           , recommendation: rec
           , followup: followup
           , excluded: code = "EXCLUDED"
           , excludeReason: exclude }
}
_GBP_Cats() {
    static m := _GBP_BuildCats()
    return m
}
_GBP_BuildCats() {
    m := Map()
    m["EXTREMELY_LOW"] := { desc: "Extremely Low Risk" }
    m["LOW"]           := { desc: "Low Risk" }
    m["INDETERMINATE"] := { desc: "Indeterminate Risk" }
    m["SURGICAL"]      := { desc: "Surgical Consultation Recommended" }
    m["EXCLUDED"]      := { desc: "Algorithm Not Applicable" }
    return m
}
