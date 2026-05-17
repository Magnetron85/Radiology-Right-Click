; ============================================================
; lib/calc/AdrenalWashout.ahk -- absolute + relative washout
; ------------------------------------------------------------
; Reference: Mayo-Smith WW et al. J Am Coll Radiol. 2017;14(8):
; 1038-1044.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

AdrenalWashout_Entry(input) {
    return CalcAdrenalWashout(input)
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
