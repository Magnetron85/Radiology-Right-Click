; ============================================================
; lib/calc/AdrenalWashout.ahk -- absolute + relative washout
; ------------------------------------------------------------
; Reference: Mayo-Smith WW et al. J Am Coll Radiol. 2017;14(8):
; 1038-1044.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

AdrenalWashout_Entry(input) {
    ShowAdrenalWashoutDialog(input)
    return ""
}

ShowAdrenalWashoutDialog(text := "") {
    unenh := TextScan.HU(text, "unenhanced")
    enh   := TextScan.HU(text, "enhanced")
    delay := TextScan.HU(text, "delayed")

    form := RadsForm("Adrenal Washout", 460)
    form.Header("CT attenuation (HU)")
    form.Note("Two-phase (enhanced + delayed) gives relative washout only. Three-phase adds absolute washout.")
    form.Numeric("Unenh", "Unenhanced HU (optional):", unenh != "" ? Round(unenh) : "")
    form.Numeric("Enh",   "Portal-venous / enhanced HU:", enh != "" ? Round(enh) : "")
    form.Numeric("Delay", "15-minute delayed HU:", delay != "" ? Round(delay) : "")
    form.SetSubmit(AdrenalWashout_OnSubmit)
    form.AddButtons()
    form.Show()
}

AdrenalWashout_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (v.Enh = "" || v.Delay = "")
        return MakeResult({ impression: "Enhanced and delayed HU are required.",
                            error: "Missing HU phases" })

    unenh := (v.Unenh = "") ? "" : v.Unenh + 0.0
    enh   := v.Enh + 0.0
    delay := v.Delay + 0.0

    absW := ""
    relW := ""

    if (unenh != "") {
        denom := enh - unenh
        if (denom = 0)
            return MakeResult({ impression: "Enhanced HU equals unenhanced HU; cannot compute absolute washout.",
                                error: "Zero washout denominator" })
        absW := ((enh - delay) / denom) * 100
        relW := (enh != 0) ? ((enh - delay) / enh) * 100 : 0
    } else {
        relW := (enh != 0) ? ((enh - delay) / enh) * 100 : 0
    }

    ; ACR 2017 Incidental Adrenal white paper does not prescribe a single
    ; recommendation off washout values alone (other factors -- size, unenh
    ; HU, history of malignancy -- drive actual management). Per user
    ; policy: emit a factual impression without a fabricated recommendation.
    ; The longer descriptive interpretation (with all caveats) lives in
    ; methodology.
    ;
    ; Verdict construction mirrors v1's InterpretAdrenalWashout decision tree:
    ;   1. unenh-HU verdict (<=10 favors adenoma; >43 raises malignancy
    ;      concern). Stacks with the washout verdict.
    ;   2. Enhancement-pattern verdict. Three branches:
    ;        (a) de-enhancement (enh or delay drops below unenh) -> the
    ;            arithmetic washout doesn't physically apply; flag and skip
    ;            the threshold compare.
    ;        (b) non-enhancing (<10 HU change in both phases vs unenh) ->
    ;            mass is likely cyst/hemorrhage, not adenoma; washout values
    ;            are numerically computable but clinically uninformative.
    ;        (c) significant enhancement -> washout threshold compare
    ;            (APW >=60% or RPW >=40%) -> adenoma vs indeterminate.
    ;   3. 2-phase fallback (no unenh) -- only RPW interpretable, flagged
    ;      as limited.
    verdicts := []
    if (unenh != "") {
        if (unenh <= 10)
            verdicts.Push("unenhanced HU <=10 favors adenoma")
        else if (unenh > 43)
            verdicts.Push("unenhanced HU >43 raises concern for malignancy regardless of washout")

        enhChange := enh - unenh
        delChange := delay - unenh
        if (enhChange < 0 || delChange < 0)
            verdicts.Push("unexpected de-enhancement on enhanced or delayed phase; standard washout calculations may not apply")
        else if (Abs(enhChange) < 10 && Abs(delChange) < 10)
            verdicts.Push("<10 HU change in both enhanced and delayed phases vs unenhanced; may represent cyst, hemorrhage, or other non-enhancing lesion")
        else if (absW >= 60)
            verdicts.Push("absolute washout >=60% consistent with adenoma")
        else if (relW >= 40)
            verdicts.Push("relative washout >=40% consistent with adenoma")
        else
            verdicts.Push("washout below adenoma thresholds (APW >=60%, RPW >=40%)")
    } else {
        ; 2-phase only -- relative-washout interpretation, explicitly flagged
        ; as limited by the missing unenhanced phase.
        if (relW >= 40)
            verdicts.Push("relative washout >=40% suggestive of adenoma; interpretation limited without unenhanced HU")
        else
            verdicts.Push("relative washout below 40%; interpretation indeterminate and limited without unenhanced HU")
    }

    verdictStr := ""
    for i, vd in verdicts {
        if (i = 1)
            verdictStr := StrUpper(SubStr(vd, 1, 1)) SubStr(vd, 2)
        else
            verdictStr .= "; " vd
    }
    verdictStr .= " per ACR 2017."

    unenhDisp := unenh != "" ? _Adr_FmtHU(unenh) : ""
    enhDisp := _Adr_FmtHU(enh)
    delayDisp := _Adr_FmtHU(delay)
    absWPct := Round(absW, 0)
    relWPct := Round(relW, 0)

    if (unenh != "") {
        impression := "Adrenal mass: unenhanced " unenhDisp " HU, enhanced " enhDisp
                    . " HU, delayed " delayDisp " HU. Absolute washout " absWPct
                    . "%, relative washout " relWPct "%. " verdictStr
    } else {
        impression := "Adrenal mass: enhanced " enhDisp " HU, delayed " delayDisp
                    . " HU. Relative washout " relWPct "%. " verdictStr
    }

    interp := _InterpretAdrenalWashout(absW = "" ? 0 : absW, relW, unenh, enh, delay)

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    if (unenh != "")
        method .= "`nUnenhanced HU: " unenh
    method .= "`nEnhanced HU: " enh
    method .= "`nDelayed HU: " delay
    if (absW != "")
        method .= "`nAbsolute washout = (E - D) / (E - U) * 100 = " Round(absW, 1) "%"
    method .= "`nRelative washout = (E - D) / E * 100 = " Round(relW, 1) "%"
    if (interp != "")
        method .= "`n`nInterpretation:`n" interp

    return MakeResult({
        classification: (absW != "" && absW >= 60) || relW >= 40 ? "Adenoma-suggestive washout" : "Indeterminate washout",
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "Mayo-Smith WW, Song JH, Boland GL, et al. "
                                . "Management of Incidental Adrenal Masses: A White Paper of the "
                                . "ACR Incidental Findings Committee. "
                                . "J Am Coll Radiol. 2017 Aug;14(8):1038-1044.",
                           url:  "https://www.acr.org/Clinical-Resources/Incidental-Findings" }],
        echo:           g_LastSelectedText
    })
}

; Format an HU value as integer when whole (avoids "10.0 HU" from the
; v.Unenh + 0.0 float coercion); otherwise show 1 decimal.
_Adr_FmtHU(v) {
    if (v = "")
        return ""
    return (v = Floor(v)) ? Integer(v) : Round(v, 1)
}

CalcAdrenalWashout(input) {
    showCit := Prefs.Get("display", "showCitations", true)
    cit := "`nMayo-Smith WW, Song JH, Boland GL, Francis IR, Israel GM, Mazzaglia PJ, Berland LL, Pandharipande PV. Management of Incidental Adrenal Masses: A White Paper of the ACR Incidental Findings Committee. J Am Coll Radiol. 2017 Aug;14(8):1038-1044.`n`n"

    ; Full 3-phase: unenhanced, enhanced, delayed
    if RegExMatch(input
        , "i)(?:(?:unenhanced|non-?enhanced|intrinsic|pre-?contrast|baseline|native)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?).*?(?:(?:enhanced|post-?contrast|arterial|portal\s+venous?|60-?75\s*(?:second|sec)|1-?2\s*min(?:ute)?)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?).*?(?:(?:delayed|late|15\s*min(?:ute)?|10-?15\s*min(?:ute)?|post-?contrast)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?)"
        , &m) {
        u := m[1] + 0, e := m[2] + 0, d := m[3] + 0
        denom := e - u
        if (denom = 0)
            return input "`n`nError: enhanced equals unenhanced, cannot compute absolute washout."
        absW := ((e - d) / denom) * 100
        relW := (e != 0) ? ((e - d) / e) * 100 : 0
        out := input "`n`n"
        out .= "Absolute Washout: " Round(absW, 1) "% ... (Ref adenomas: >=60%)`n"
        out .= "Relative Washout: " Round(relW, 1) "% ... (Ref adenomas: >=40%)`n`n"
        if showCit
            out .= cit
        out .= _InterpretAdrenalWashout(absW, relW, u, e, d)
        return out
    }

    ; Two-phase: enhanced + delayed only (relative washout only)
    if RegExMatch(input
        , "i)(?:(?:enhanced|post-?contrast|arterial|portal\s+venous?|60-?75\s*(?:second|sec)|1-?2\s*min(?:ute)?)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?).*?(?:(?:delayed|late|15\s*min(?:ute)?|10-?15\s*min(?:ute)?|post-?contrast)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?)"
        , &m2) {
        e := m2[1] + 0, d := m2[2] + 0
        relW := (e != 0) ? ((e - d) / e) * 100 : 0
        out := input "`n`n"
        out .= "Relative Washout: " Round(relW, 1) "%`n`n"
        out .= _InterpretAdrenalWashout(0, relW, "", e, d)
        if showCit
            out .= "`n`n" cit
        return out
    }

    return "Invalid input format for adrenal washout calculation.`n"
        . "Sample syntax: Unenhanced: 10 HU, Enhanced: 80 HU, Delayed: 40 HU"
}

_InterpretAdrenalWashout(absW, relW, unenhanced, enhanced, delayed) {
    out := ""
    unenAvail := !(unenhanced = "" || unenhanced = "NULL")
    if unenAvail {
        u := unenhanced + 0
        encChg := enhanced - u
        delChg := delayed  - u

        if (u <= 10)
            out .= "Unenhanced HU less than 10 is typically suggestive of a benign adrenal adenoma. "
        else if (u > 43)
            out .= "Unenhanced HU >43 in a noncalcified, nonhemorrhagic lesion is suspicious for malignancy, regardless of washout characteristics. "

        if (Abs(encChg) >= 10 || Abs(delChg) >= 10) {
            if (encChg < 0 || delChg < 0) {
                out .= "The adrenal mass demonstrates unexpected de-enhancement (decrease in HU in enhanced or delayed phase). This is an atypical finding. "
                out .= "Caution: Standard washout calculations may not be applicable in this case. "
            } else {
                out .= "The adrenal mass demonstrates enhancement. "
                if (absW >= 60 || relW >= 40)
                    out .= "Adrenal washout characteristics are suggestive of a benign adrenal adenoma. "
                else
                    out .= "Washout characteristics are indeterminate. "
            }
        } else {
            out .= "The adrenal mass does not demonstrate significant enhancement (<10 HU change in both enhanced and delayed phases compared to unenhanced). This may represent a cyst, hemorrhage, or other non-enhancing lesion. Further characterization with additional imaging may be necessary. "
        }
    } else {
        out .= "Unenhanced HU value is not available. "
        if (absW >= 60 || relW >= 40)
            out .= "Based on the provided washout values alone, characteristics are suggestive of a benign adrenal adenoma. However, this interpretation is limited without the unenhanced HU value. "
        else
            out .= "Washout characteristics are indeterminate. "
    }
    return Trim(out)
}
