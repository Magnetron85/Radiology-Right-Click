; ============================================================
; lib/calc/HepaticSteatosis.ahk
; ------------------------------------------------------------
; Grading: Guglielmo FF et al. RadioGraphics 2023;43(6):e220181
; (Table 4, rounded from Tang A et al. Radiology 2013;267:422-431).
; Spleen normalization: Sirlin CB. Radiographics 2009;29:1277-80.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk
#Include ..\FormGui.ahk

HepaticSteatosis_Entry(input) {
    ShowHepaticSteatosisDialog(input)
    return ""
}

ShowHepaticSteatosisDialog(text := "") {
    lIp := "", lOp := "", sIp := "", sOp := ""
    if RegExMatch(text, "i)liver[^0-9]*?(?:in[- ]?phase|IP|T1IP)\D*(\d+)\D*(?:out[- ]?of[- ]?phase|OP|OOP|T1OP)\D*(\d+)", &m) {
        lIp := m[1] + 0
        lOp := m[2] + 0
    }
    if RegExMatch(text, "i)spleen[^0-9]*?(?:in[- ]?phase|IP|T1IP)\D*(\d+)\D*(?:out[- ]?of[- ]?phase|OP|OOP|T1OP)\D*(\d+)", &m) {
        sIp := m[1] + 0
        sOp := m[2] + 0
    }

    form := RadsForm("Hepatic Steatosis (Dixon)", 460)
    form.Header("Liver signal")
    form.Numeric("Lip", "Liver IP:", lIp)
    form.Numeric("Lop", "Liver OP:", lOp)
    form.Header("Spleen reference (optional)")
    form.Numeric("Sip", "Spleen IP:", sIp)
    form.Numeric("Sop", "Spleen OP:", sOp)
    form.SetSubmit(HepaticSteatosis_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

HepaticSteatosis_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (v.Lip = "" || v.Lop = "")
        return MakeResult({ impression: "Liver IP and OP signals are required.",
                            error: "Missing liver IP/OP" })

    liverIP := v.Lip + 0.0
    liverOP := v.Lop + 0.0
    if (liverIP = 0)
        return MakeResult({ impression: "Liver IP value cannot be zero.",
                            error: "Zero liver IP" })

    fatFraction := 100 * (liverIP - liverOP) / (2 * liverIP)
    ; Grade the same 1-decimal value the user sees, so a displayed "6.0%"
    ; can never carry a "below the 6% threshold" interpretation.
    grade := _InterpretSteatosis(Round(fatFraction, 1))
    gradeShort := RegExReplace(grade, "^Interpretation:\s*", "")
    gradeShort := RegExReplace(gradeShort, "\.$", "")

    spleenLine := ""
    spleenFF := ""
    if (v.Sip != "" && v.Sop != "") {
        spleenIP := v.Sip + 0.0
        spleenOP := v.Sop + 0.0
        if (spleenIP != 0 && spleenOP != 0) {
            spleenFF := 100 * ((liverIP/spleenIP) - (liverOP/spleenOP)) / (2 * (liverIP/spleenIP))
            spleenLine := "; spleen-normalized FF " Round(spleenFF, 1) "%"
        }
    }

    impression := "Dixon-derived hepatic fat fraction " Round(fatFraction, 1) "%"
               . spleenLine ". " gradeShort "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nLiver IP: " liverIP ", OP: " liverOP
    method .= "`nFat fraction = 100 * (IP - OP) / (2 * IP) = " Round(fatFraction, 1) "%"
    if (spleenFF != "")
        method .= "`nSpleen-normalized FF = " Round(spleenFF, 1) "%"
    method .= "`n" grade

    advisories := []
    borderNote := _SteatosisBorderlineNote(Round(fatFraction, 1))
    if (borderNote != "")
        advisories.Push(borderNote)
    advisories.Push("Dixon IP/OP fat fraction approximates PDFF only at low fat content. "
                  . "Quantify with PDFF (CSE-MRI) when available.")
    if (spleenFF != "")
        advisories.Push("Spleen-normalized FF (Sirlin 2009) is not endorsed by Guglielmo 2023; "
                      . "use as a sanity check, not the primary metric.")

    return MakeResult({
        classification: gradeShort,
        impression:     impression,
        recommendation: "",
        advisories:     advisories,
        methodology:    method,
        citations:      [{ text: "Guglielmo FF, Barr RG, Yokoo T, et al. "
                                . "Liver Fibrosis, Fat, and Iron Evaluation with MRI and Fibrosis "
                                . "and Fat Evaluation with US: A Practical Guide for Radiologists. "
                                . "Radiographics. 2023;43(6):e220181.",
                           url:  "https://pubs.rsna.org/doi/10.1148/rg.220181" },
                         { text: "Tang A, Tan J, Sun M, et al. Nonalcoholic fatty liver disease: "
                                . "MR imaging of liver proton density fat fraction to assess "
                                . "hepatic steatosis. Radiology. 2013;267(2):422-431.",
                           url:  "https://doi.org/10.1148/radiol.12120896" },
                         { text: "Sirlin CB. Invited Commentary on Image-based quantification of "
                                . "hepatic fat. Radiographics. 2009;29:1277-1280.",
                           url:  "https://doi.org/10.1148/027153330290051277" }],
        echo:           g_LastSelectedText,
        paste:          fatFraction != "" ? "fat fraction " Round(fatFraction, 1) "%" : "",
        pasteMode:      ""
    })
}

CalcHepaticSteatosis(input) {
    showCit := Prefs.Get("display","showCitations",true)
    liverNeedle  := "i)liver.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)"
    spleenNeedle := "i)spleen.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)"

    if !RegExMatch(input, liverNeedle, &lm)
        return "Invalid input format for hepatic steatosis calculation.`n"
            . "Sample syntax: Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88"

    liverIP := lm[2] + 0, liverOP := lm[4] + 0
    if (liverIP = 0)
        return "Liver IP value cannot be zero."

    ; Dixon-based fat-fraction approximation (Guglielmo 2023 line 308-315 notes
    ; this is only valid at low fat content and is biased by T1, T2*, and
    ; spectral complexity of fat; PDFF from CSE-MRI is the reference method
    ; when available).
    fatFraction := 100 * (liverIP - liverOP) / (2 * liverIP)
    out := input " (Fat Fraction: " Round(fatFraction, 1) "%)"
    ; Grade the same 1-decimal value that is displayed (see the form path).
    out .= "`n`nFat Fraction " _InterpretSteatosis(Round(fatFraction, 1))
    borderNote := _SteatosisBorderlineNote(Round(fatFraction, 1))
    if (borderNote != "")
        out .= "`nNote: " borderNote
    out .= "`nNote: Dixon IP/OP fat fraction approximates PDFF only at low fat content. Quantify with PDFF (CSE-MRI) when available.`n"

    if RegExMatch(input, spleenNeedle, &sm) {
        spleenIP := sm[2] + 0, spleenOP := sm[4] + 0
        if (spleenIP != 0 && spleenOP != 0) {
            fatPct := 100 * ((liverIP/spleenIP) - (liverOP/spleenOP)) / (2 * (liverIP/spleenIP))
            ; Anchor the replacement to the exact annotation we appended --
            ; replacing the first bare ")" would inject into any parenthesis
            ; the user's own selection happens to contain.
            ffTag := "Fat Fraction: " Round(fatFraction, 1) "%)"
            out := StrReplace(out, ffTag
                 , "Fat Fraction: " Round(fatFraction, 1) "%, Spleen-normalized FF: " Round(fatPct, 1) "%)",, , 1)
            out .= "Spleen-normalized FF " _InterpretSteatosis(Round(fatPct, 1))
            ; The spleen-normalized FF can land in the borderline band when
            ; the primary FF did not; a borderline verdict must never appear
            ; without its explanation (and the note is never printed twice).
            if (borderNote = "") {
                spleenNote := _SteatosisBorderlineNote(Round(fatPct, 1))
                if (spleenNote != "")
                    out .= "`nNote: " spleenNote
            }
            out .= "`nNote: spleen-normalized fat fraction is an older approximation (Sirlin 2009) not endorsed by Guglielmo 2023; use as a sanity check, not as the primary metric.`n"
        }
    }

    if showCit {
        out .= "`n`nCitation 1: Guglielmo FF, Barr RG, Yokoo T, et al. Liver Fibrosis, Fat, and Iron Evaluation with MRI and Fibrosis and Fat Evaluation with US: A Practical Guide for the Radiologist. RadioGraphics 2023;43(6):e220181."
        out .= "`nCitation 2: Tang A, Tan J, Sun M, et al. Nonalcoholic fatty liver disease: MR imaging of liver proton density fat fraction to assess hepatic steatosis. Radiology 2013;267(2):422-431."
        out .= "`nCitation 3: Sirlin CB. Invited Commentary on Image-based quantification of hepatic fat: methods and clinical applications. Radiographics 2009;29:1277-80.`n"
    }
    return out
}

_InterpretSteatosis(ff) {
    ; Grading bands per Guglielmo et al., RadioGraphics 2023 (Table 4) -- the
    ; integer-rounded adaptation of the histology-calibrated Tang 2013
    ; thresholds (6.4% / 17.4% / 22.1%):
    ;   <6%  Normal | 6-17% Mild (G1) | 17-22% Moderate (G2) | >22% Severe (G3).
    ; The 5-6% band is genuinely contested: the common DIAGNOSTIC cutoffs sit
    ; at PDFF >=5% (MASLD trial / FDA convention; 5.56% by MRS in the Dallas
    ; Heart Study, Szczepaniak 2005) -- below the grade 1 threshold. Values
    ; there are reported as borderline rather than "no steatosis": with a
    ; two-point Dixon approximation a sub-1% margin cannot settle the call.
    if (ff < 5)
        return "Interpretation: No significant hepatic steatosis."
    if (ff < 6)
        return "Interpretation: Borderline hepatic steatosis (meets the 5% diagnostic threshold; below the 6% Grade 1 threshold)."
    if (ff < 17)
        return "Interpretation: Mild hepatic steatosis (Grade 1)."
    if (ff <= 22)
        return "Interpretation: Moderate hepatic steatosis (Grade 2)."
    return "Interpretation: Severe hepatic steatosis (Grade 3)."
}

; One-line explanation of the contested 5-6% band, shared by the form
; (advisory) and legacy (Note:) output paths. Empty outside the band.
_SteatosisBorderlineNote(ff) {
    if (ff < 5 || ff >= 6)
        return ""
    return "Fat fractions of 5-6% meet common diagnostic cutoffs for steatosis "
         . "(PDFF >=5% in MASLD trials; 5.56% by MRS, Dallas Heart Study) but fall "
         . "below the histology-calibrated Grade 1 threshold (6% per Guglielmo 2023 "
         . "Table 4, rounded from 6.4% per Tang 2013). Correlate clinically."
}
