; ============================================================
; lib/calc/IncidentalThyroid.ahk
; ------------------------------------------------------------
; ACR Incidental Thyroid Findings Committee White Paper (2015).
; Reference: Hoang JK et al. J Am Coll Radiol. 2015;12:143-150.
;
; Size thresholds (no suspicious features, no clinical risk):
;   age < 35  -> >=1.0 cm warrants ultrasound
;   age >= 35 -> >=1.5 cm warrants ultrasound
; Clinical risk factors / suspicious features / PET-avidity
; override size criteria.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

IncidentalThyroid_Entry(input) {
    ShowIncidentalThyroidDialog(input)
    return ""
}

ShowIncidentalThyroidDialog(text := "") {
    sz   := TextScan.Size(text)
    age  := TextScan.Age(text)
    sizeCm := sz.cm > 0 ? Round(sz.cm, 1) : 0

    form := RadsForm("Incidental Thyroid Nodule (ACR 2015)", 520)

    ; Progressive-disclosure layout. Modality is the mandatory context;
    ; PetAvid only applies to PET so it sits directly under its gate and
    ; is HIDDEN (not greyed) for other modalities. The "Size criteria"
    ; section drives the default algorithm path, but every modifier
    ; checkbox short-circuits _Thy_Classify before size/age are read, so
    ; the whole section collapses while any modifier is ticked.
    form.Header("Imaging context")
    form.Dropdown("Modality", "Modality:"
        , ["CT", "MRI", "PET (FDG)", "US (extrathyroidal)"], 1)
    form.Checkbox("PetAvid", "Focal FDG uptake in thyroid (PET only)", false)

    form.Header("Size criteria")
    form.Numeric("SizeCm", "Largest size (cm, 0 if unknown):", sizeCm)
    form.Numeric("Age", "Patient age (years, 0 if unknown):", age)

    form.Header("Modifiers")
    form.Checkbox("Susp"
        , "Suspicious features (CT/MRI: abnormal nodes, invasion; US: microcalcs, marked hypo, irregular margins, taller-than-wide)"
        , false)
    form.Checkbox("ClinRisk"
        , "Clinical risk factors (h/o neck XRT, family hx, MEN2, vocal cord paralysis, rapid growth, pediatric)"
        , false)
    form.Checkbox("LimitedLE"
        , "Limited life expectancy / significant comorbidities"
        , false)

    ; Modality gates PetAvid; the modifier checkboxes (plus PetAvid
    ; itself) gate the size/age section.
    form.OnChange("Modality",  _Thy_UpdateForm)
    form.OnChange("PetAvid",   _Thy_UpdateForm)
    form.OnChange("Susp",      _Thy_UpdateForm)
    form.OnChange("ClinRisk",  _Thy_UpdateForm)
    form.OnChange("LimitedLE", _Thy_UpdateForm)
    _Thy_UpdateForm(form)

    form.SetSubmit(IncidentalThyroid_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_Thy_UpdateForm(frm) {
    ; PetAvid is only meaningful when modality = PET. Hiding it resets
    ; it to unchecked, so the classifier never sees focal FDG uptake
    ; asserted for a CT / MRI / US study.
    isPET := InStr(frm.byName["Modality"].ctl.Text, "PET")
    frm.SetVisible("PetAvid", isPET)

    ; Every modifier short-circuits the size/age path in _Thy_Classify
    ; (ClinRisk -> CLINICAL, PetAvid -> US_FNA / NO_WORKUP_LE,
    ; Susp -> US / NO_WORKUP_LE, LimitedLE -> NO_WORKUP_LE), so the size
    ; criteria are irrelevant while any of them is set. Read PetAvid
    ; AFTER the visibility pass above -- a just-hidden control has
    ; already been reset to its unchecked default.
    bypass := (frm.GetValue("Susp") = 1)
           || (frm.GetValue("PetAvid") = 1)
           || (frm.GetValue("ClinRisk") = 1)
           || (frm.GetValue("LimitedLE") = 1)
    frm.SetSectionVisible("Size criteria", !bypass)
}

IncidentalThyroid_OnSubmit(v, form := "") {
    global g_LastSelectedText
    modality := SubStr(v.Modality, 1, 3)
    if (modality = "PET")
        modality := "PET"
    else if (modality = "US ")
        modality := "US"
    sizeCm := v.SizeCm + 0.0
    age    := SafeInt(v.Age, 0)
    susp   := !!v.Susp
    pet    := !!v.PetAvid
    risk   := !!v.ClinRisk
    le     := !!v.LimitedLE

    r := _Thy_Classify(modality, sizeCm, age, susp, le, risk, pet)

    ; Capitalize the leading "I" when size is missing; "incidental..." starts
    ; the sentence and needs to be a proper sentence-case noun phrase.
    sizeStr := sizeCm > 0 ? Format("{:.1f}", sizeCm) " cm incidental thyroid nodule"
                          : "Incidental thyroid nodule"
    mgmtLc := (r.mgmt != "") ? StrLower(SubStr(r.mgmt, 1, 1)) SubStr(r.mgmt, 2) : ""
    impression := sizeStr "; " mgmtLc " per ACR 2015."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nRecommendation category: " r.desc
    method .= "`nReason: " r.reason
    method .= "`nFull action: " r.mgmt

    return MakeResult({
        classification: r.desc,
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "Hoang JK, Langer JE, Middleton WD, et al. "
                                . "Managing incidental thyroid nodules detected on imaging: "
                                . "white paper of the ACR Incidental Thyroid Findings Committee. "
                                . "J Am Coll Radiol. 2015 Feb;12(2):143-50.",
                           url:  "https://www.acr.org/Clinical-Resources/Incidental-Findings" }],
        echo:           g_LastSelectedText
    })
}

_Thy_Classify(modality, sizeCm, age, susp, le, risk, pet) {
    if risk
        return _Thy_R("CLINICAL", "Clinical risk factors for thyroid cancer present -- standard workup recommended regardless of size")

    ; Per Hoang 2015 white paper (lines 280-284, 253-255), US+FNA is triggered by
    ; FOCAL FDG uptake, not by PET modality alone. Diffuse uptake or PET without
    ; focal uptake should follow standard size/age path. Require the explicit
    ; "focal FDG uptake" flag rather than modality alone.
    if pet {
        if le
            return _Thy_R("NO_WORKUP_LE", "PET-avid thyroid nodule, but limited life expectancy / significant comorbidities")
        return _Thy_R("US_FNA", "Focal FDG uptake in thyroid -- US + FNA recommended regardless of sonographic findings")
    }

    if susp {
        if le
            return _Thy_R("NO_WORKUP_LE", "Suspicious features present, but limited life expectancy")
        reason := (modality = "US")
                ? "Suspicious sonographic features (microcalcifications, marked hypoechogenicity, irregular margins, taller-than-wide)"
                : "Suspicious CT/MRI features (abnormal lymph nodes, local invasion)"
        return _Thy_R("US", reason " -- dedicated thyroid ultrasound recommended")
    }

    if le
        return _Thy_R("NO_WORKUP_LE", "Limited life expectancy or significant comorbidities -- no further evaluation recommended")

    ; size / age criteria
    if (age > 0 && age < 35) {
        threshold := 1.0
        if (sizeCm >= threshold)
            return _Thy_R("US", "Age <35 years with nodule >=" threshold " cm -- dedicated thyroid ultrasound recommended")
        return _Thy_R("NO_WORKUP", "Age <35 years with nodule <" threshold " cm -- no further evaluation recommended")
    }

    threshold := 1.5
    ageStr := age > 0 ? "Age >=35 years" : "Age unknown (using >=35 threshold)"
    if (sizeCm >= threshold)
        return _Thy_R("US", ageStr " with nodule >=" threshold " cm -- dedicated thyroid ultrasound recommended")
    if (sizeCm > 0)
        return _Thy_R("NO_WORKUP", ageStr " with nodule <" threshold " cm -- no further evaluation recommended")
    return _Thy_R("US", "Nodule size unknown -- dedicated thyroid ultrasound recommended to characterize")
}

_Thy_R(code, reason) {
    map := _Thy_Cats()
    info := map.Has(code) ? map[code] : map["NO_WORKUP"]
    return { code: code, desc: info.desc, mgmt: info.mgmt, reason: reason }
}

_Thy_Cats() {
    static m := _Thy_Build()
    return m
}
_Thy_Build() {
    m := Map()
    m["US_FNA"]       := { desc: "Dedicated thyroid US + FNA recommended"
                         , mgmt: "Dedicated thyroid ultrasound with fine-needle aspiration of the PET-avid lesion recommended regardless of sonographic findings." }
    m["US"]           := { desc: "Dedicated thyroid US recommended"
                         , mgmt: "Dedicated thyroid ultrasound recommended. Further management (FNA) based on US findings." }
    m["NO_WORKUP"]    := { desc: "No further evaluation recommended"
                         , mgmt: "No further evaluation recommended per ACR Incidental Thyroid Findings Committee criteria." }
    m["NO_WORKUP_LE"] := { desc: "No further evaluation (limited life expectancy)"
                         , mgmt: "No further evaluation recommended given limited life expectancy / significant comorbidities. Workup may be reconsidered if clinically warranted." }
    m["CLINICAL"]     := { desc: "Clinical risk factors present"
                         , mgmt: "Standard thyroid nodule evaluation recommended regardless of size criteria." }
    return m
}
