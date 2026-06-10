; ============================================================
; lib/calc/Bosniak.ahk -- Bosniak 2019 cystic renal mass class.
; ------------------------------------------------------------
; Reference: Silverman SG, Pedrosa I, Ellis JH, et al. Bosniak
; Classification of Cystic Renal Masses, Version 2019: An Update
; Proposal and Needs Assessment. Radiology. 2019;292(2):475-488.
;
; Algorithm priority (highest first):
;  1) enhancing nodule >=4 mm OR any acute-margin protrusion -> IV
;  2) enhancing <=3 mm obtuse protrusion -> III
;  3) thick (>=4 mm) enhancing wall / septa -> III
;  4) minimally thickened (3 mm) enhancing wall / septa -> IIF
;  5) >=4 thin septa -> IIF
;  6) heterogeneously T1 hyper (MRI only) -> IIF
;  7) few thin septa, hyperattenuating CT, T1-hyper MRI,
;     calcification alone -> II
;  8) simple fluid + thin wall + no septa + no enhancement -> I
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

Bosniak_Entry(input) {
    ShowBosniakDialog(input)
    return ""
}

ShowBosniakDialog(text := "") {
    fluid := TextScan.ContainsAny(text, ["\bsimple\s+cyst","\b(?:water|csf)[- ]like\b","\bhypoattenuating\s+cyst"])
    calc  := TextScan.ContainsAny(text, ["\bcalcified\b","\bcalcification"])
    form := RadsForm("Bosniak 2019", 580)

    ; Progressive-disclosure layout. Mandatory context (modality, wall /
    ; septa morphology) sits at the top; every conditional input appears
    ; directly below the field that makes it relevant and is HIDDEN (the
    ; form reflows and resizes), not greyed, until then:
    ;   * IrregMm                          -> only when wall/septa enhance
    ;     (the classifier reads it only inside the `if enh` branch)
    ;   * NodMm / NodMargin / NodEnh       -> only when a protrusion exists
    ;   * NodCalc                          -> protrusion present AND CT
    ;   * HyperCT                          -> CT only
    ;   * T2Hyper / T1Hyper / T1Hetero     -> MRI only
    form.Header("Modality")
    form.Dropdown("Modality", "Imaging:", ["CT", "MRI"], 1)

    form.Header("Cyst wall and septa")
    form.Numeric("WallMm",  "Wall thickness (mm, 0 if thin):", 0)
    form.Numeric("SeptaCnt", "Septa count (0 = none):", 0)
    form.Numeric("SeptaMm", "Max septa thickness (mm):", 0)
    form.Checkbox("EnhWall", "Wall or septa enhance", false)
    form.Numeric("IrregMm", "Irregular protrusion size (<=3 mm obtuse, 0 if none):", 0)

    form.Header("Convex protrusion (nodule)")
    form.Checkbox("HasNod", "Convex protrusion present", false)
    form.Numeric("NodMm",  "Protrusion size (mm):", 0)
    form.Dropdown("NodMargin", "Protrusion margins:"
        , ["Obtuse (wide base)", "Acute (sharp angle)"], 1)
    form.Checkbox("NodEnh",  "Protrusion enhances", false)
    form.Checkbox("NodCalc", "Protrusion is calcified (CT only)", false)

    form.Header("Other features")
    form.Checkbox("HyperCT",  "Homogeneous >=70 HU on noncontrast CT", false)
    form.Checkbox("T2Hyper",  "Homogeneous markedly T2-hyperintense, similar to CSF (noncontrast MRI)", false)
    form.Checkbox("T1Hyper",  "Homogeneous T1 hyperintense, ~2.5x parenchymal signal (noncontrast MRI)", false)
    form.Checkbox("T1Hetero", "Heterogeneously T1 hyperintense on fat-saturated noncontrast MRI", false)
    form.Checkbox("Fluid",    "Simple fluid (-9 to 20 HU or CSF-like T2)", fluid)
    form.Checkbox("CalsOnly", "Calcification only, no other complex features", calc)

    form.OnChange("Modality", _Bos_UpdateForm)
    form.OnChange("HasNod",   _Bos_UpdateForm)
    form.OnChange("EnhWall",  _Bos_UpdateForm)
    _Bos_UpdateForm(form)

    form.SetSubmit(Bosniak_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_Bos_UpdateForm(frm) {
    isCT   := frm.GetValue("Modality") = 1
    hasNod := !!frm.GetValue("HasNod")
    enhW   := !!frm.GetValue("EnhWall")

    ; A hidden control reverts to its default, so the classifier never
    ; reads a stale answer: IrregMm/NodMm -> 0, checkboxes -> unchecked,
    ; NodMargin -> first option -- all of which the classifier provably
    ; treats as "absent / not applicable" on the hidden branch.
    frm.SetVisible("IrregMm", enhW)

    frm.SetVisible("NodMm",     hasNod)
    frm.SetVisible("NodMargin", hasNod)
    frm.SetVisible("NodEnh",    hasNod)
    frm.SetVisible("NodCalc",   hasNod && isCT)

    frm.SetVisible("HyperCT",  isCT)
    frm.SetVisible("T2Hyper",  !isCT)
    frm.SetVisible("T1Hyper",  !isCT)
    frm.SetVisible("T1Hetero", !isCT)
}

Bosniak_OnSubmit(v, form := "") {
    global g_LastSelectedText
    modality := v.Modality = "CT" ? "CT" : "MRI"
    wallMm   := v.WallMm + 0.0
    sCnt     := SafeInt(v.SeptaCnt, 0)
    sMm      := v.SeptaMm + 0.0
    enh      := !!v.EnhWall
    hasNod   := !!v.HasNod
    nodMm    := v.NodMm + 0.0
    nodObtuse := InStr(v.NodMargin, "Obtuse")
    nodMargin := nodObtuse ? "obtuse" : "acute"
    nodEnh   := !!v.NodEnh
    nodCalc  := !!v.NodCalc
    irreg    := v.IrregMm + 0.0
    hyperCT  := !!v.HyperCT
    t2Hyper  := !!v.T2Hyper
    t1Hyper  := !!v.T1Hyper
    t1Hetero := !!v.T1Hetero
    fluid    := !!v.Fluid
    calsOnly := !!v.CalsOnly

    r := _Bos_Classify(modality, wallMm, sCnt, sMm, enh, hasNod, nodMm, nodMargin
                     , nodEnh, nodCalc, irreg, hyperCT, t2Hyper, t1Hyper, t1Hetero, fluid, calsOnly)

    showRisk := Prefs.Get("display", "showMalignancyRisk", true)

    ; --- Impression: report-ready sentence ending with management.
    ; r.reason is sentence-cased ("Minimally thickened enhancing wall..."); lower
    ; the leading character so it reads as prose inside the parenthetical.
    reasonLc := (r.reason != "") ? StrLower(SubStr(r.reason, 1, 1)) SubStr(r.reason, 2) : ""
    impression := "Bosniak " r.category " renal cyst"
    if (reasonLc != "")
        impression .= " (" reasonLc ")"
    if (showRisk && r.risk != "")
        impression .= ", reported malignancy risk " r.risk
    impression .= ". " _Bos_ShortMgmt_Cap(r.category) " per Bosniak v2019."

    ; --- Methodology: full audit detail. ---
    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nCategory: " r.category " (" r.desc ")"
    if (showRisk && r.risk != "")
        method .= "`nMalignancy risk: " r.risk
    method .= "`nReason: " r.reason
    method .= "`nFull management: " r.mgmt

    return MakeResult({
        classification: "Bosniak " r.category,
        impression:     impression,
        recommendation: "",   ; impression already ends with management
        methodology:    method,
        citations:      [{ text: "Silverman SG, Pedrosa I, Ellis JH, et al. "
                                . "Bosniak Classification of Cystic Renal Masses, Version 2019: "
                                . "An Update Proposal and Needs Assessment. "
                                . "Radiology. 2019;292(2):475-488.",
                           url:  "https://pubs.rsna.org/doi/10.1148/radiol.2019182646" }],
        echo:           g_LastSelectedText
    })
}

_Bos_ShortMgmt(cat) {
    if (cat = "I" || cat = "II")
        return "no follow-up recommended"
    if (cat = "IIF")
        return "recommend follow-up imaging at 6 and 12 months, then annually for 5 years"
    if (cat = "III" || cat = "IV")
        return "recommend urology consultation"
    return ""
}

; Sentence-start version (capitalized) for use after a period.
_Bos_ShortMgmt_Cap(cat) {
    if (cat = "I" || cat = "II")
        return "No follow-up recommended"
    if (cat = "IIF")
        return "Recommend follow-up imaging at 6 and 12 months, then annually for 5 years"
    if (cat = "III" || cat = "IV")
        return "Recommend urology consultation"
    return ""
}

_Bos_Classify(modality, wallMm, sCnt, sMm, enh, hasNod, nodMm, nodMargin
             , nodEnh, nodCalc, irreg, hyperCT, t2Hyper, t1Hyper, t1Hetero, fluid, calsOnly) {
    cats := _Bos_Cats()

    ; protrusion evaluation
    if hasNod {
        if nodEnh {
            if (nodMargin = "acute" || nodMm >= 4)
                return _Bos_R("IV", "Enhancing nodule (" nodMm " mm, " nodMargin " margins)", cats)
            if (nodMargin = "obtuse" && nodMm <= 3)
                return _Bos_R("III", "Irregular enhancing wall/septa thickening (" nodMm " mm obtuse protrusion)", cats)
        } else {
            if (modality = "CT" && nodCalc)
                return _Bos_R("II", "Calcified non-enhancing protrusion. Recommend MRI to exclude occult enhancing elements beneath calcification.", cats)
            ; non-enhancing non-calcified protrusion -- fall through
        }
    }

    ; Bosniak III: thick enhancing wall / septa
    if enh {
        if (wallMm >= 4)
            return _Bos_R("III", "Thick enhancing wall (" wallMm " mm >=4 mm)", cats)
        if (sMm >= 4)
            return _Bos_R("III", "Thick enhancing septa (" sMm " mm >=4 mm)", cats)
        if (irreg > 0 && irreg <= 3)
            return _Bos_R("III", "Irregular enhancing wall/septa (<=3 mm obtuse protrusion)", cats)
    }

    ; Bosniak IIF: minimally thickened, many thin septa, or hetero T1
    if enh {
        if (wallMm = 3)
            return _Bos_R("IIF", "Minimally thickened enhancing wall (3 mm)", cats)
        if (sMm = 3)
            return _Bos_R("IIF", "Minimally thickened enhancing septa (3 mm)", cats)
    }
    ; Silverman 2019 IIF "many enhancing septa": Table 3 specifies the septa
    ; must enhance to qualify for IIF (lines 234-236, 731). Non-enhancing thin
    ; septa drop to Bosniak II.
    if (enh && sCnt >= 4 && sMm <= 2)
        return _Bos_R("IIF", "Many (>=4) thin enhancing septa", cats)
    if (modality = "MRI" && t1Hetero)
        return _Bos_R("IIF", "Heterogeneously T1 hyperintense on fat-saturated imaging", cats)

    ; Bosniak II
    if (sCnt >= 1 && sCnt <= 3 && sMm <= 2)
        return _Bos_R("II", "Few thin septa (" sCnt " septa, <=2 mm)", cats)
    if (modality = "CT" && hyperCT)
        return _Bos_R("II", "Homogeneous hyperattenuating mass (>=70 HU on unenhanced CT)", cats)
    ; Silverman 2019 MRI Cat II type 2: homogeneous markedly T2-hyperintense, similar to CSF
    if (modality = "MRI" && t2Hyper)
        return _Bos_R("II", "Homogeneous markedly T2-hyperintense mass (similar to CSF) on noncontrast MRI", cats)
    if (modality = "MRI" && t1Hyper)
        return _Bos_R("II", "Homogeneous T1 hyperintense mass (~2.5x parenchymal signal)", cats)
    if calsOnly
        return _Bos_R("II", "Calcification without other complex features", cats)

    ; Bosniak I
    if (fluid && sCnt = 0 && wallMm <= 2 && !enh)
        return _Bos_R("I", "Simple cyst (thin wall, simple fluid, no septa)", cats)
    if (fluid && sCnt = 0 && wallMm <= 2)
        return _Bos_R("I", "Simple cyst (wall may enhance)", cats)
    if (sCnt = 0 && wallMm <= 2)
        return _Bos_R("I", "Simple-appearing cyst", cats)
    if (sCnt > 0 && sMm <= 2 && !enh)
        return _Bos_R("II", "Thin septa without measurable enhancement", cats)

    return _Bos_R("IIF", "Indeterminate features -- recommend follow-up", cats)
}

_Bos_R(code, reason, cats) {
    info := cats[code]
    return { category: code, desc: info.desc, risk: info.risk, mgmt: info.mgmt, reason: reason }
}
_Bos_Cats() {
    static m := _Bos_BuildCats()
    return m
}
; Malignancy rates and ranges per Silverman SG, Pedrosa I, Ellis JH, et al.
; Bosniak Classification of Cystic Renal Masses, Version 2019. Radiology
; 2019;292(2):475-488 (p. 477-478). Bosniak I has no specific rate published
; -- simple cysts are implicitly benign -- so the field is left empty. IIF
; has no central tendency in the paper; the published 0-38% range across
; resected series is used directly (selection bias acknowledged; Schoots
; cohort analysis infers true rate ~9% but that is not a quote from the
; paper). Display of these values is gated on display.showMalignancyRisk.
_Bos_BuildCats() {
    m := Map()
    m["I"]   := { desc: "Benign simple cyst", risk: ""
                , mgmt: "Benign simple renal cyst requiring no follow-up." }
    m["II"]  := { desc: "Benign cyst", risk: "<1%"
                , mgmt: "Benign Bosniak II renal cyst requiring no follow-up." }
    m["IIF"] := { desc: "Probably benign, follow-up recommended"
                , risk: "0-38% (wide range across resected series; Silverman 2019)"
                , mgmt: "Follow-up imaging at 6 and 12 months, then annually for 5 years to assess for morphologic change." }
    m["III"] := { desc: "Indeterminate"
                , risk: "~50% (range 25-100% across resected series; Silverman 2019)"
                , mgmt: "Urology consultation recommended given intermediate probability of malignancy." }
    m["IV"]  := { desc: "Likely malignant"
                , risk: "~90% (range 56-100% across resected series; Silverman 2019)"
                , mgmt: "Urology consultation recommended; the large majority are malignant." }
    return m
}
