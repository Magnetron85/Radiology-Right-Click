; ============================================================
; lib/calc/PIRADS.ahk -- PI-RADS v2.1 prostate MRI scoring
; ------------------------------------------------------------
; Reference: American College of Radiology PI-RADS v2.1.
;
; Peripheral zone: DWI is the dominant sequence.
; Transition zone: T2W is the dominant sequence.
;
; Per PI-RADS v2.1, size is baked INTO the sequence-score definitions:
;   DWI score 4: focal markedly hypointense / markedly hyperintense, <1.5 cm
;   DWI score 5: same as 4 BUT >=1.5 cm OR definite EPE/invasive behavior
;   T2  score 4: lenticular or non-circumscribed moderately hypointense, <1.5 cm
;   T2  score 5: same as 4 BUT >=1.5 cm OR definite EPE/invasive behavior
;
; The radiologist is expected to pick the correct sequence score based on
; what they measure. If we see DWI=4 paired with size >=15 mm, that's a
; contradiction -- the lesion should have been scored DWI=5. We surface a
; warning in that case rather than silently double-count.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

PIRADS_Entry(input) {
    ShowPIRADSDialog(input)
    return ""
}

ShowPIRADSDialog(text := "") {
    sz   := TextScan.Size(text)
    zone := TextScan.ZoneProstate(text)
    zoneIdx := (zone = "TZ") ? 2 : 1
    ; Pre-detect EPE / invasive behavior in the highlighted text. Pattern
    ; covers the most common ways radiologists phrase it.
    epePre := TextScan.ContainsAny(text
        , ["\bextraprostatic\s+extension\b"
         , "\bextra[- ]?prostatic\s+extension\b"
         , "\bEPE\b"
         , "\bgross\s+EPE\b"
         , "\binvasive\s+behavior\b"
         , "\bseminal\s+vesicle\s+invasion\b"
         , "\bSVI\b"])

    ; Form ordered to match the radiologist's mental workflow:
    ;   1. Measure the lesion (size + invasive features) -- these gate
    ;      which sequence scores are even valid per v2.1.
    ;   2. Identify the zone -- determines which sequence is dominant.
    ;   3. Assess sequence morphology.
    form := RadsForm("PI-RADS v2.1", 520)

    form.Header("Lesion size and PI-RADS 5 triggers")
    form.Note("Per v2.1, score-4 morphology becomes PI-RADS 5 if size >=1.5 cm OR definite extraprostatic extension / invasive behavior. Score 4 is removed from the sequence-score lists below when either trigger applies.")
    form.Numeric("Size", "Largest dimension (mm):", sz.mm > 0 ? Round(sz.mm) : 0)
    form.Checkbox("EPE", "Definite extraprostatic extension or invasive behavior present", epePre)

    form.Header("Lesion location")
    form.Dropdown("Zone", "Zone:", ["PZ (peripheral)", "TZ (transition)"], zoneIdx)

    form.Header("Sequence morphology scores (1-5)")
    form.Note("PZ: DWI dominant; DCE upgrades DWI=3 to 4 if positive. TZ: T2 dominant; DWI is a tiebreaker for T2=2/3. T2 is not used in PZ; DCE is not used in TZ -- those fields gray out when the other zone is selected. Score 4 is hidden when size >=1.5 cm or EPE is checked (those triggers force a 5). Score 5 is hidden when size is <1.5 cm AND no EPE -- v2.1 has no defined score-5 morphology without one of those triggers.")
    form.Dropdown("T2",  "T2W score [TZ only -- not used in PZ]:", ["1","2","3","4","5"], 3)
    form.Dropdown("DWI", "DWI / ADC score:", ["1","2","3","4","5"], 3)
    form.Dropdown("DCE", "DCE [PZ only -- not used in TZ]:", ["Negative","Positive"], 1)

    ; Wire dynamic updates:
    ;   Zone   -> enable/disable T2 / DCE per dominant-sequence table.
    ;   Size,
    ;   EPE    -> rebuild T2 and DWI option lists. When size >=15 or EPE
    ;             checked, score 4 is removed from the list -- the v2.1
    ;             definition forces 5 in those cases, so 4 is not a valid
    ;             choice. If the user had 4 selected when the trigger
    ;             fires, the selection auto-bumps to 5 (preserving the
    ;             radiologist's morphology assessment).
    form.OnChange("Zone", _PIRADS_UpdateZone)
    form.OnChange("Size", _PIRADS_UpdateScoreOptions)
    form.OnChange("EPE",  _PIRADS_UpdateScoreOptions)
    _PIRADS_UpdateZone(form)
    _PIRADS_UpdateScoreOptions(form)

    form.SetSubmit(PIRADS_OnSubmit)
    form.AddButtons()
    form.Show()
}

; Filter the T2 and DWI dropdown items based on the v2.1 size + EPE rules.
;   Score 4 is hidden when size >=1.5 cm OR EPE is checked (either trigger
;   plus score-4 morphology IS score 5 by definition -- picking 4 is
;   logically invalid).
;   Score 5 is hidden when size is known AND <1.5 cm AND no EPE (v2.1
;   defines score 5 only relative to one of those triggers; without either,
;   there's no "score-5 morphology" to pick).
; When size is unknown (0 / blank), neither rule fires -- both 4 and 5
; remain available so the radiologist can score before measuring.
;
; Auto-bumping preserves intent:
;   - 4 hidden + currently on 4 -> bump to 5 (the v2.1 promotion target).
;   - 5 hidden + currently on 5 -> bump to 4 (the closest morphology
;     still available; user can revise if the morphology truly is <4).
_PIRADS_UpdateScoreOptions(frm) {
    size := SafeInt(frm.GetValue("Size"), 0)
    epe := !!frm.GetValue("EPE")
    sizeKnown := (size > 0)

    canBe4 := !((sizeKnown && size >= 15) || epe)
    canBe5 := !(sizeKnown && size < 15 && !epe)

    newOpts := []
    for s in ["1", "2", "3", "4", "5"] {
        if (s = "4" && !canBe4)
            continue
        if (s = "5" && !canBe5)
            continue
        newOpts.Push(s)
    }

    for name in ["T2", "DWI"] {
        if !frm.byName.Has(name)
            continue
        ctl := frm.byName[name].ctl
        prevText := ctl.Text
        prevScore := IsNumber(prevText) ? Integer(prevText) : 3

        targetScore := prevScore
        if (!canBe4 && prevScore = 4)
            targetScore := 5
        if (!canBe5 && prevScore = 5)
            targetScore := 4

        ctl.Delete()
        ctl.Add(newOpts)
        for i, opt in newOpts {
            if (opt = String(targetScore)) {
                ctl.Value := i
                break
            }
        }
    }
}

_PIRADS_UpdateZone(frm) {
    isPZ := InStr(frm.byName["Zone"].ctl.Text, "PZ")
    frm.SetEnabled("T2",  !isPZ)
    frm.SetEnabled("DCE", isPZ)
}

PIRADS_OnSubmit(v, form := "") {
    global g_LastSelectedText
    zone := InStr(v.Zone, "PZ") ? "PZ" : "TZ"
    t2   := Integer(v.T2)
    dwi  := Integer(v.DWI)
    dce  := v.DCE = "Positive"
    size := SafeInt(v.Size, 0)
    epe  := !!v.EPE
    r := _PIRADS_Score(zone, t2, dwi, dce, size, epe)

    sizePhrase := size > 0 ? size " mm " : ""
    likeLc := (r.likelihood != "") ? StrLower(SubStr(r.likelihood, 1, 1)) SubStr(r.likelihood, 2) : ""
    mgmtLc := (r.management != "") ? StrLower(SubStr(r.management, 1, 1)) SubStr(r.management, 2) : ""
    impression := sizePhrase "prostate lesion in the " (zone = "PZ" ? "peripheral" : "transition")
               . " zone, PI-RADS " r.score " (" likeLc ")"
    ; If the score was upgraded from 4 to 5 by size or EPE, surface the
    ; reason inline so the radiologist (and downstream reader) can see WHY
    ; this is a 5 rather than a 4.
    if (r.upgradeReason != "")
        impression .= " -- upgraded from sequence score " r.seqScore " due to " r.upgradeReason
    impression .= " per PI-RADS v2.1"
    if (mgmtLc != "")
        impression .= "; " mgmtLc
    if (SubStr(impression, -1) != ".")
        impression .= "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nDominant-sequence morphology score: " r.seqScore
    if (r.upgradeReason != "")
        method .= "`nUpgrade to overall PI-RADS 5: " r.upgradeReason
    method .= "`nFinal PI-RADS " r.score " -- " r.likelihood
    if (r.warning != "")
        method .= "`nWarning: " r.warning
    method .= "`nFull management: " r.management

    advisories := []
    if (r.warning != "")
        advisories.Push(r.warning)

    return MakeResult({
        classification: "PI-RADS " r.score,
        impression:     impression,
        recommendation: "",
        advisories:     advisories,
        methodology:    method,
        citations:      [{ text: "Turkbey B, Rosenkrantz AB, Haider MA, et al. "
                                . "Prostate Imaging Reporting and Data System Version 2.1: 2019 "
                                . "Update of Prostate Imaging Reporting and Data System Version 2. "
                                . "Eur Urol. 2019 Sep;76(3):340-351.",
                           url:  "https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/PI-RADS" }],
        echo:           g_LastSelectedText
    })
}

_PIRADS_Score(zone, t2, dwi, dce, size, epe) {
    ; --- Step 1: derive the morphology-based sequence score per v2.1.
    ; This is what the score would be if the radiologist picked the
    ; sequence number assuming a <1.5 cm lesion with no EPE.
    seqScore := 0
    if (zone = "PZ") {
        if (dwi = 1)
            seqScore := 1
        else if (dwi = 2)
            seqScore := 2
        else if (dwi = 3)
            seqScore := dce ? 4 : 3   ; DCE+ upgrades equivocal DWI to 4
        else if (dwi = 4)
            seqScore := 4
        else if (dwi = 5)
            seqScore := 5
    } else {
        if (t2 = 1)
            seqScore := 1
        else if (t2 = 2)
            seqScore := (dwi >= 4) ? 3 : 2
        else if (t2 = 3)
            seqScore := (dwi = 5) ? 4 : 3
        else if (t2 = 4)
            seqScore := 4
        else if (t2 = 5)
            seqScore := 5
    }

    ; --- Step 2: PI-RADS v2.1 score-5 triggers.
    ; v2.1 score 5: "Same as 4 but >=1.5 cm in greatest dimension OR
    ; definite extraprostatic extension / invasive behavior"
    ; (T2 PZ lines 1164-1166, T2 TZ 1175-1177, DWI 1300-1303 of
    ; references/pirads.md).
    ;
    ; The OR is parsed clinically as two independent triggers:
    ;   (a) Size >=1.5 cm: STRICT to v2.1 -- the literal "Same as 4 but
    ;       >=1.5 cm" anchors the size criterion to score-4 morphology.
    ;       A heterogeneous-signal score-3 lesion measuring 18 mm stays
    ;       at PI-RADS 3 per the table definitions. Size only upgrades
    ;       sequence-4 to overall-5.
    ;   (b) EPE / invasive behavior: LIBERAL -- definite invasive behavior
    ;       is a definitive-cancer finding regardless of how the dominant
    ;       sequence scored the morphology. A "PI-RADS 3 with definite
    ;       EPE" is internally inconsistent (equivocal-likelihood-of-
    ;       cancer + definite-invasion-by-cancer). EPE upgrades any
    ;       morphology score <5 to PI-RADS 5.
    ;
    ; If both triggers fire, list both in the upgrade reason.
    overall := seqScore
    upgradeReason := ""
    triggers := []
    if (epe && seqScore < 5)
        triggers.Push("definite extraprostatic extension / invasive behavior")
    if (seqScore = 4 && size >= 15)
        triggers.Push("size >=1.5 cm (15 mm)")
    if (triggers.Length > 0) {
        overall := 5
        for i, t in triggers {
            upgradeReason .= (i = 1 ? "" : " AND ") t
        }
    }

    ; --- Step 3: warn on user-asserted PI-RADS 5 lacking either trigger.
    ; If the radiologist picked sequence-5 directly but neither size>=15
    ; nor EPE is present, the v2.1 definition is unsatisfied -- nudge
    ; them to confirm one criterion or downgrade.
    warning := ""
    if (seqScore = 5 && size > 0 && size < 15 && !epe) {
        warning := "User-asserted PI-RADS 5 with size <15 mm and no extraprostatic extension / invasive behavior indicated. Per v2.1, score 5 requires >=1.5 cm OR definite EPE/invasive behavior in addition to score-4 morphology. Confirm criteria."
    }

    if (overall < 1)
        overall := 1
    if (overall > 5)
        overall := 5

    info := _PIRADS_Info()[overall]
    return { score: overall
           , seqScore: seqScore
           , upgradeReason: upgradeReason
           , likelihood: info.likelihood
           , cancerRate: info.cancerRate
           , management: info.management
           , warning: warning }
}

_PIRADS_Info() {
    static t := _PIRADS_BuildInfo()
    return t
}
; Qualitative descriptors per PI-RADS v2.1 (Turkbey et al., Eur Urol 2019).
; Numerical csPCa detection rates per category come from external meta-analyses
; (e.g., Park 2018, Mazzone 2021) and are not part of the v2.1 document
; itself, so they are kept off the result -- the radiologist can cite a
; specific cohort if needed. Biopsy modality is intentionally left
; unspecified (v2.1 does not prescribe fusion vs. cognitive).
_PIRADS_BuildInfo() {
    m := Map()
    m[1] := { likelihood: "Clinically significant cancer (csPCa) is highly unlikely"
            , cancerRate: ""
            , management: "Routine clinical follow-up." }
    m[2] := { likelihood: "csPCa is unlikely"
            , cancerRate: ""
            , management: "Routine clinical follow-up." }
    m[3] := { likelihood: "csPCa is equivocal"
            , cancerRate: ""
            , management: "Equivocal -- biopsy decision should consider PSA, clinical factors, and patient preference (PI-RADS v2.1 does not prescribe biopsy)." }
    m[4] := { likelihood: "csPCa is likely"
            , cancerRate: ""
            , management: "Targeted biopsy recommended (modality per local practice)." }
    m[5] := { likelihood: "csPCa is highly likely"
            , cancerRate: ""
            , management: "Targeted biopsy recommended (modality per local practice)." }
    return m
}
