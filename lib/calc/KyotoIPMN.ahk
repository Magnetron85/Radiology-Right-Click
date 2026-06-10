; ============================================================
; lib/calc/KyotoIPMN.ahk -- Kyoto 2024 IPMN management criteria
; ------------------------------------------------------------
; Reference: Ohtsuka T, Fernandez-del Castillo C, Furukawa T,
; et al. International evidence-based Kyoto guidelines for the
; management of intraductal papillary mucinous neoplasm of the
; pancreas. Pancreatology. 2024;24(2):255-270.
;
; HRS (high-risk stigmata): any -> surgical consultation
;   - Enhancing mural nodule >=5 mm or solid component
;   - Main pancreatic duct >=10 mm
;   - Obstructive jaundice (head/uncinate IPMN)
;   - Suspicious / positive cytology on EUS-FNA
;   - Main-duct or mixed type with MPD >=10 mm
;
; WF (worrisome features): EUS-based evaluation; multiple WFs
; markedly raise risk (HR 10.1 for >=3 WFs).
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

KyotoIPMN_Entry(input) {
    ShowKyotoIPMNDialog(input)
    return ""
}

ShowKyotoIPMNDialog(text := "") {
    typeIdx := 1
    if TextScan.ContainsAny(text, ["\bmain[- ]duct\b","\bmain\s+duct\s+IPMN"])
        typeIdx := 2
    if TextScan.ContainsAny(text, ["\bmixed[- ]type\b","\bmixed\s+IPMN"])
        typeIdx := 3

    age := TextScan.Age(text)

    ; cyst size: prefer "cyst[ic lesion] X mm" or "X mm [pancreatic] cyst"
    cystMm := 0
    if RegExMatch(text, "i)(?:cyst|cystic\s+lesion|IPMN)\s+(?:measuring\s+|of\s+|is\s+)?(?:up\s+to\s+)?(\d+(?:\.\d+)?)\s*(mm|cm)", &m)
        cystMm := (m[2] = "cm") ? (m[1] + 0.0) * 10 : m[1] + 0.0
    else if RegExMatch(text, "i)(\d+(?:\.\d+)?)\s*(mm|cm)\s+(?:pancreatic\s+)?cyst", &m)
        cystMm := (m[2] = "cm") ? (m[1] + 0.0) * 10 : m[1] + 0.0

    ; MPD diameter
    mpdMm := 0
    if RegExMatch(text, "i)(?:MPD|main\s+pancreatic\s+duct)\s+(?:measur(?:ing|es?)\s+|of\s+|is\s+|=\s*)?(\d+(?:\.\d+)?)\s*(mm|cm)", &m)
        mpdMm := (m[2] = "cm") ? (m[1] + 0.0) * 10 : m[1] + 0.0

    ; mural nodule
    nodMm := 0
    if RegExMatch(text, "i)(?:mural\s+)?nodule(?:\s+measuring|\s+of|\s+is|\s+size)?\s+(\d+(?:\.\d+)?)\s*(mm|cm)", &m)
        nodMm := (m[2] = "cm") ? (m[1] + 0.0) * 10 : m[1] + 0.0

    jaundice := TextScan.ContainsAny(text, ["\bjaundice\b","\bobstructive\s+jaundice"])
    panc     := TextScan.ContainsAny(text, ["\bacute\s+pancreatitis\b"])
    ca199    := TextScan.ContainsAny(text, ["\bCA\s*19[- ]?9\b.*?(?:elevat|increas)","\belevat.*\bCA\s*19[- ]?9"])

    form := RadsForm("Kyoto IPMN Guidelines (2024)", 580)

    ; Progressive-disclosure layout.
    ;   * Duct type gates the CYST-specific inputs. A pure main-duct (MD)
    ;     IPMN is segmental/diffuse MPD dilation with NO side-branch cyst
    ;     by definition -- so "largest cyst diameter", "thickened/enhancing
    ;     cyst walls", and "cyst growth rate" do not apply (if a cyst is
    ;     present, the lesion is mixed-type, not MD). These three are hidden
    ;     for MD and shown for BD / Mixed. MPD diameter, mural nodule, and
    ;     all clinical/lab features are type-agnostic and stay.
    ;   * NotFit is read by _IPMN_Classify only inside the HRS branch, so it
    ;     appears only once a high-risk stigma is present.
    form.Header("IPMN type and dimensions")
    form.Dropdown("Type", "Type:"
        , ["Branch duct (BD)"
        ,  "Main duct (MD)"
        ,  "Mixed"], typeIdx)
    form.Numeric("CystMm", "Largest cyst diameter (mm):", cystMm > 0 ? Round(cystMm, 1) : 0)
    form.Numeric("MpdMm",  "Main pancreatic duct (mm):", mpdMm > 0 ? Round(mpdMm, 1) : 0)
    form.Numeric("NodMm",  "Enhancing mural nodule (mm, 0 if none):", nodMm > 0 ? Round(nodMm, 1) : 0)
    form.Numeric("Growth", "Cyst growth rate (mm/year, 0 if unknown):", 0)

    form.Header("High-risk stigmata")
    form.Checkbox("Jaundice", "Obstructive jaundice (head/uncinate IPMN)", jaundice)
    form.Checkbox("Cyto", "Suspicious or positive cytology on EUS-FNA", false)
    form.Checkbox("NotFit", "Patient is NOT a surgical candidate (close surveillance instead)", false)

    form.Header("Worrisome features")
    form.Checkbox("WallThick",   "Thickened / enhancing cyst walls", false)
    form.Checkbox("DuctChange",  "Abrupt MPD caliber change with distal pancreatic atrophy", false)
    form.Checkbox("Lymph",       "Lymphadenopathy", false)
    form.Checkbox("CA199",       "Elevated serum CA 19-9", ca199)
    form.Checkbox("Diabetes",    "New-onset or acute exacerbation of DM within past year", false)
    form.Checkbox("Pancreatitis","Acute pancreatitis", panc)

    form.Header("Patient")
    form.Numeric("Age", "Age (years, 0 if unknown):", age)

    ; Type gates the cyst-specific inputs; NotFit is gated by the four
    ; inputs that can raise an HRS (Jaundice, Cyto, NodMm >=5, MpdMm >=10).
    form.OnChange("Type",     _IPMN_UpdateForm)
    form.OnChange("Jaundice", _IPMN_UpdateForm)
    form.OnChange("Cyto",     _IPMN_UpdateForm)
    form.OnChange("NodMm",    _IPMN_UpdateForm)
    form.OnChange("MpdMm",    _IPMN_UpdateForm)
    _IPMN_UpdateForm(form)

    form.SetSubmit(KyotoIPMN_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_IPMN_UpdateForm(frm) {
    ; --- cyst-specific inputs: hidden for pure main-duct (MD) IPMN ---
    ; MD has no side-branch cyst, so cyst diameter / cyst-wall thickening /
    ; cyst growth rate do not apply. Hidden CystMm / Growth reset to blank
    ; (= "not measured"); hidden WallThick resets to unchecked -- so the
    ; cyst-size WF (>=30 mm) and the cyst-wall WF simply don't fire for MD,
    ; which is the correct behavior.
    isMd := InStr(frm.byName["Type"].ctl.Text, "Main")
    frm.SetVisible("CystMm",    !isMd)
    frm.SetVisible("Growth",    !isMd)
    frm.SetVisible("WallThick", !isMd)

    ; --- NotFit: only meaningful when a high-risk stigma is present ---
    ; Surgical candidacy only changes the output inside the HRS branch.
    ; Hidden NotFit resets to unchecked = surgical candidate (the default).
    nod := _IPMN_NumVal(frm, "NodMm")
    mpd := _IPMN_NumVal(frm, "MpdMm")
    hrsPresent := (frm.GetValue("Jaundice") = 1)
               || (frm.GetValue("Cyto") = 1)
               || (nod >= 5)
               || (mpd >= 10)
    frm.SetVisible("NotFit", hrsPresent)
}

_IPMN_NumVal(frm, name) {
    if !frm.byName.Has(name)
        return 0
    txt := frm.byName[name].ctl.Value
    return IsNumber(txt) ? txt + 0 : 0
}

KyotoIPMN_OnSubmit(v, form := "") {
    global g_LastSelectedText
    typ := InStr(v.Type, "Branch") ? "BD"
         : InStr(v.Type, "Main")   ? "MD"
                                   : "MIXED"
    cyst := v.CystMm + 0.0
    mpd  := v.MpdMm  + 0.0
    nod  := v.NodMm  + 0.0
    grow := v.Growth + 0.0
    age  := SafeInt(v.Age, 0)
    notFit := !!v.NotFit

    r := _IPMN_Classify(typ, cyst, mpd, nod, !!v.Jaundice, !!v.Cyto
                      , !!v.WallThick, !!v.DuctChange, !!v.Lymph, grow
                      , !!v.CA199, !!v.Diabetes, !!v.Pancreatitis
                      , age, !notFit)

    typeLabel := (typ = "BD") ? "branch-duct" : (typ = "MD") ? "main-duct" : "mixed"
    sizePhrase := cyst > 0 ? Format("{:.0f}", cyst) " mm " : ""

    ; --- Impression ---
    ; The classifier populates BOTH r.hrsFeatures and r.wfFeatures
    ; independently (e.g. MPD=12 with also-elevated CA 19-9 lights up an
    ; HRS for the duct AND a WF for the CA 19-9). When the category is
    ; HRS, the management is driven by HRS but the WFs are still
    ; clinically reportable -- list them as additional findings rather
    ; than dropping them. The trailing threshold annotation on each
    ; feature (e.g. " (5-9 mm)") is stripped for impression cleanliness.
    nHrs := r.hrsFeatures.Length
    nWf  := r.wfFeatures.Length

    hrsStr := _IPMN_FormatFeatures(r.hrsFeatures)
    wfStr  := _IPMN_FormatFeatures(r.wfFeatures)

    ; Sentence-start capitalization when size is missing.
    leadingDesc := (sizePhrase = "")
        ? (StrUpper(SubStr(typeLabel, 1, 1)) SubStr(typeLabel, 2))
        : sizePhrase . typeLabel

    if (r.category = "HRS") {
        impression := leadingDesc " IPMN with " nHrs " high-risk feature"
                    . (nHrs > 1 ? "s" : "")
        if (hrsStr != "")
            impression .= " (" hrsStr ")"
        if (nWf > 0) {
            impression .= " and " nWf " worrisome feature" (nWf > 1 ? "s" : "")
            if (wfStr != "")
                impression .= " (" wfStr ")"
        }
    } else if (r.category = "WF") {
        impression := leadingDesc " IPMN with " nWf " worrisome feature"
                    . (nWf > 1 ? "s" : "")
        if (wfStr != "")
            impression .= " (" wfStr ")"
    } else {
        ; Per radiologist policy: the form's HRS/WF checkboxes include
        ; clinical and laboratory items (jaundice, CA 19-9, DM, cytology,
        ; acute pancreatitis) that the radiologist often cannot confirm
        ; from imaging alone. When no risk feature fires, qualify as
        ; "no imaging ..." rather than the absolute "without high-risk
        ; features" so the impression doesn't overclaim absence of
        ; clinical features that weren't assessed.
        impression := leadingDesc " IPMN with no imaging high-risk stigmata or worrisome features"
    }

    impression .= " per Kyoto 2024; " _IPMN_ShortMgmt(r)
    if (SubStr(impression, -1) != ".")
        impression .= "."

    ; --- Methodology ---
    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nCategory: " r.category
    method .= "`nReason: " r.reason
    if (r.hrsFeatures.Length > 0)
        method .= "`nHigh-risk stigmata (" r.hrsFeatures.Length "):`n - " JoinArr(r.hrsFeatures, "`n - ")
    if (r.wfFeatures.Length > 0)
        method .= "`nWorrisome features (" r.wfFeatures.Length "):`n - " JoinArr(r.wfFeatures, "`n - ")
    method .= "`nFull management: " r.mgmt
    if (r.surveillance != "")
        method .= "`nSurveillance schedule:`n" r.surveillance
    if (r.risk != "")
        method .= "`nRisk estimate: " r.risk

    advisories := []
    if (r.ageNote != "")
        advisories.Push(r.ageNote)

    return MakeResult({
        classification: r.category,
        impression:     impression,
        recommendation: "",
        advisories:     advisories,
        methodology:    method,
        citations:      [{ text: "Ohtsuka T, Fernandez-del Castillo C, Furukawa T, et al. "
                                . "International evidence-based Kyoto guidelines for the management "
                                . "of intraductal papillary mucinous neoplasm of the pancreas. "
                                . "Pancreatology. 2024;24(2):255-270.",
                           url:  "https://doi.org/10.1016/j.pan.2023.12.009" }],
        echo:           g_LastSelectedText
    })
}

_IPMN_CategoryLabel(cat) {
    if (cat = "HRS")
        return "with high-risk stigmata"
    if (cat = "WF")
        return "with worrisome features"
    return "without high-risk features"
}

; Build a grammatically-clean Oxford-comma list of feature strings for the
; impression. Strips trailing "(threshold-band)" annotations from each
; entry and lower-cases the leading character so they read inline.
_IPMN_FormatFeatures(features) {
    cleaned := []
    for f in features {
        fc := RegExReplace(f, "\s*\([^)]*\)$", "")
        fc := StrLower(SubStr(fc, 1, 1)) SubStr(fc, 2)
        cleaned.Push(fc)
    }
    out := ""
    n := cleaned.Length
    for i, f in cleaned {
        if (i = 1)
            out := f
        else if (i = n)
            out .= (n > 2 ? ", and " : " and ") f
        else
            out .= ", " f
    }
    return out
}

_IPMN_ShortMgmt(r) {
    if (r.category = "HRS")
        return "recommend surgical consultation / multidisciplinary discussion"
    if (r.category = "WF") {
        if (r.wfFeatures.Length >= 3)
            return "recommend surgical consultation and EUS with tissue sampling"
        return "recommend EUS evaluation"
    }
    ; No-risk-features path -- the size-keyed surveillance interval text
    ; ("every 6/12/18 months") is set on r.surveillance by _IPMN_Classify,
    ; NOT on r.mgmt. r.mgmt only holds the generic "Surveillance imaging
    ; per Kyoto 2024 size-based protocol." preamble; look at r.surveillance
    ; for the actual interval.
    surv := r.HasOwnProp("surveillance") ? r.surveillance : ""
    if (surv != "") {
        if InStr(surv, "every 6 months")
            return "recommend MRI/MRCP surveillance every 6 months"
        if InStr(surv, "every 12 months")
            return "recommend MRI/MRCP at 6 months and 12 months, then every 12 months if stable"
        if InStr(surv, "every 18 months")
            return "recommend MRI/MRCP at 6 months, then every 18 months if stable"
    }
    if InStr(r.mgmt, "Cyst size not specified")
        return "recommend MRI/MRCP for baseline measurement"
    return "follow size-based surveillance per Kyoto 2024"
}

_IPMN_Classify(typ, cyst, mpd, nod, jaundice, cyto, wallThick, ductChg, lymph, grow
              , ca199, diabetes, panc, age, surgicalFit) {
    ; Kyoto 2024 (Ohtsuka et al. Pancreatology 2024;24(2):255-270, sections 5.3 + 5.4):
    ;   HRS = (1) obstructive jaundice in head/uncinate cystic lesion,
    ;         (2) enhancing mural nodule >=5 mm or solid component,
    ;         (3) main pancreatic duct >=10 mm,
    ;         (4) suspicious or positive cytology.
    ;   WF  = (1) acute pancreatitis,
    ;         (2) increased serum CA 19-9,
    ;         (3) new-onset or acute exacerbation of DM within past year,
    ;         (4) cyst >=30 mm,
    ;         (5) enhancing mural nodule <5 mm,
    ;         (6) thickened / enhancing cyst walls,
    ;         (7) MPD >=5 mm and <10 mm,
    ;         (8) abrupt change in MPD caliber with distal atrophy,
    ;         (9) lymphadenopathy,
    ;        (10) cyst growth rate >=2.5 mm/year.
    hrs := []
    if (nod >= 5)
        hrs.Push("Enhancing mural nodule / solid component " Round(nod, 1) " mm")
    if (mpd >= 10)
        hrs.Push("Main pancreatic duct " Round(mpd, 1) " mm (>=10 mm)")
    if jaundice
        hrs.Push("Obstructive jaundice (head/uncinate IPMN)")
    if cyto
        hrs.Push("Suspicious or positive cytology on EUS-FNA")

    wf := []
    if (cyst >= 30)
        wf.Push("Cyst size " Round(cyst, 1) " mm (>=30 mm)")
    if (nod > 0 && nod < 5)
        wf.Push("Enhancing mural nodule " Round(nod, 1) " mm (<5 mm)")
    if wallThick
        wf.Push("Thickened / enhancing cyst walls")
    if (mpd >= 5 && mpd < 10)
        wf.Push("MPD " Round(mpd, 1) " mm (5-9 mm)")
    if ductChg
        wf.Push("Abrupt MPD caliber change with distal pancreatic atrophy")
    if lymph
        wf.Push("Lymphadenopathy")
    if (grow >= 2.5)
        wf.Push("Cyst growth rate " Round(grow, 1) " mm/year (>=2.5)")
    if ca199
        wf.Push("Elevated serum CA 19-9")
    if diabetes
        wf.Push("New-onset or acute exacerbation of DM within past year")
    if panc
        wf.Push("Acute pancreatitis")

    cat := ""
    reason := ""
    mgmt := ""
    risk := ""
    surv := ""
    ageNote := ""

    if (hrs.Length > 0) {
        cat := "HRS"
        risk := "~49% prevalence of pancreatic carcinoma at diagnosis"
        if surgicalFit {
            mgmt := "Surgical consultation recommended. Multidisciplinary discussion advised."
            reason := "High-risk stigmata present -- surgical consultation recommended"
        } else {
            mgmt := "High-risk stigmata present but patient may not be a surgical candidate. Close surveillance with EUS every 3-6 months. Multidisciplinary discussion essential."
            reason := "High-risk stigmata present -- close surveillance (not surgical candidate)"
        }
    } else if (wf.Length > 0) {
        cat := "WF"
        ; HGD/IC prevalence from Zelga et al. (Kyoto 2024 ref [91]; source lines 558-562):
        ;   1 WF -> 22%, 2 WF -> 34%, 3 WF -> 59%, >=4 WF -> 100%.
        if (wf.Length >= 4) {
            risk := "4+ worrisome features: ~100% prevalence of HGD/IC (Zelga et al.)"
            mgmt := "Multiple worrisome features. Surgical consultation and EUS with tissue sampling strongly recommended."
            reason := wf.Length " worrisome features -- very high risk"
        } else if (wf.Length = 3) {
            risk := "3 worrisome features: ~59% prevalence of HGD/IC (Zelga et al.)"
            mgmt := "Multiple worrisome features. Surgical consultation and EUS with tissue sampling recommended."
            reason := "3 worrisome features -- significantly elevated risk"
        } else if (wf.Length = 2) {
            risk := "2 worrisome features: ~34% prevalence of HGD/IC (Zelga et al.)"
            mgmt := "EUS evaluation recommended. If EUS confirms worrisome features (mural nodule, suspicious cytology, MPD involvement), surgical consultation is warranted."
            reason := "2 worrisome features -- EUS evaluation recommended"
        } else {
            risk := "1 worrisome feature: ~22% prevalence of HGD/IC (Zelga et al.)"
            mgmt := "EUS evaluation recommended."
            reason := "1 worrisome feature -- EUS evaluation recommended"
        }
    } else {
        cat := "No risk features"
        risk := "~0.05% prevalence of carcinoma at diagnosis; 0.4-1.1% annual transformation risk"
        reason := "No high-risk stigmata or worrisome features"
        ; Surveillance intervals per Kyoto 2024 (source lines 673-678):
        ;   <20 mm   -> 6 months once, then every 18 months
        ;   20-30 mm -> 6 months twice, then every 12 months
        ;   >=30 mm  -> every 6 months (only reachable when cyst >=30 is excluded from WF,
        ;                                which currently is not the case; kept for completeness)
        if (cyst = 0) {
            mgmt := "Cyst size not specified -- recommend MRI/MRCP for baseline measurement."
        } else if (cyst < 20) {
            surv := "Cyst <20 mm:`n  Initial: CT or MRI/MRCP at 6 months`n  Follow-up: every 18 months if stable"
            mgmt := "Surveillance imaging per Kyoto 2024 size-based protocol."
        } else if (cyst < 30) {
            surv := "Cyst 20-30 mm:`n  Initial: 6 months, then 6 months again`n  Follow-up: every 12 months if stable"
            mgmt := "Surveillance imaging per Kyoto 2024 size-based protocol."
        } else {
            surv := "Cyst >=30 mm:`n  Follow-up: every 6 months"
            mgmt := "Close surveillance per Kyoto 2024 size-based protocol; surgical consultation should be considered."
        }
        ; Stop-surveillance per Kyoto 2024 (source lines 682-685). Three
        ; alternative evidence-based rules; ordered MOST-SPECIFIC first so
        ; the recommendation text reflects the strongest subgroup evidence
        ; that applies to this patient:
        ;   - Age >=75 with cyst <30 mm (elderly-specific, allows up to 3 cm)
        ;   - Age >=65 with cyst <15 mm (older + small-cyst-specific)
        ;   - Otherwise: any-age + cyst <20 mm (general)
        if (age >= 75 && cyst > 0 && cyst < 30)
            mgmt .= " In patients >=75 with cyst <3 cm, consider discontinuing surveillance after 5 years of stability, weighing life expectancy and comorbidities."
        else if (age >= 65 && cyst > 0 && cyst < 15)
            mgmt .= " In patients >=65 with cyst <1.5 cm, consider discontinuing surveillance after 5 years of stability."
        else if (cyst > 0 && cyst < 20)
            mgmt .= " May consider discontinuing surveillance after 5 years of stability."
        if (cyst >= 20 && cyst < 30 && age > 0 && age < 65)
            ageNote := "In young/fit patients with 2-3 cm branch-duct IPMN, surgical consultation is recommended to avoid prolonged surveillance."
    }

    return { category: cat, reason: reason, mgmt: mgmt, risk: risk
           , surveillance: surv, ageNote: ageNote
           , hrsFeatures: hrs, wfFeatures: wf }
}
