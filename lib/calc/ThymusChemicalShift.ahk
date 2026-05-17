; ============================================================
; lib/calc/ThymusChemicalShift.ahk
; ------------------------------------------------------------
; Reference: Priola AM et al. Radiology. 2015 Jan;274(1):238-49.
; doi: 10.1148/radiol.14132665.
; Cutoffs: chemical-shift ratio <= 0.849 and signal-intensity
; index > 8.92 % support hyperplasia.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

ThymusChemicalShift_Entry(input) {
    return CalcThymusChemicalShift(input)
}

CalcThymusChemicalShift(input) {
    needle := "i)thymus.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)(?:.*?paraspinous.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+))?"
    if !RegExMatch(input, needle, &m)
        return "Invalid input format for thymus chemical shift calculation.`n"
            . "Sample syntax: Thymus IP: 100, OP: 80, Paraspinous IP: 90, OP: 85`n"
            . "Or for thymus only: Thymus IP: 100, OP: 80"

    thIP := m[2] + 0, thOP := m[4] + 0
    psIP := (m.Count >= 6 && m[6] != "") ? m[6] + 0 : ""
    psOP := (m.Count >= 8 && m[8] != "") ? m[8] + 0 : ""

    sii := (thIP != 0) ? ((thIP - thOP) / thIP) * 100 : 0
    showAll := Prefs.Get("display","allValues",true)

    out := input
    if (psIP != "" && psOP != "") {
        opRatio := psOP != 0 ? thOP / psOP : 0
        ipRatio := psIP != 0 ? thIP / psIP : 0
        csr := ipRatio != 0 ? opRatio / ipRatio : 0
        if showAll {
            out .= "`n`nChemical Shift Ratio: " Round(csr, 3) " (hyperplasia < 0.849)`n"
            out .= "Thymus Signal Intensity Index (SII): " Round(sii, 2) "% (hyperplasia > 8.92)`n"
        }
        out .= "`n" _InterpretThymus(csr, sii)
    } else {
        if showAll
            out .= "`nThymus Signal Intensity Index (SII): " Round(sii, 2) "%`n"
        out .= "`n" _InterpretThymus("", sii)
    }
    return out
}

_InterpretThymus(csr, sii) {
    showCit := Prefs.Get("display","showCitations",true)
    out := ""
    if (csr != "" && sii != "") {
        if (csr > 0.849 && sii > 8.92)
            out := "Interpretation: Chemical Shift Ratio is greater than 0.849 and Signal Intensity Index is greater than 8.92%. Calculations are in conflict and therefore indeterminate, though probably consistent with thymic hyperplasia with single dual echo technique.`n`n"
        else if (csr <= 0.849 && sii > 8.92)
            out := "Interpretation: Chemical Shift Ratio is less than or equal to 0.849 and Signal Intensity Index is greater than 8.92%. Findings are consistent with thymic hyperplasia with single dual echo technique.`n`n"
        else if (csr > 0.849 && sii <= 8.92)
            out := "Interpretation: Chemical Shift Ratio is greater than 0.849 and Signal Intensity Index is less than or equal to 8.92%. Findings are not consistent with typical thymic hyperplasia with single dual echo technique.`n`n"
        else
            out := "Interpretation: Chemical Shift Ratio is less than or equal to 0.849 and Signal Intensity Index is less than or equal to 8.92%. Calculations are in conflict and therefore indeterminate, possibly thymic hyperplasia with single dual echo technique.`n`n"
    } else if (csr != "" && sii = "") {
        out := "Interpretation: Chemical Shift Ratio is " (csr > 0.849 ? "greater than" : "less than or equal to") " 0.849 with single dual echo technique.`n`n"
    } else if (csr = "" && sii != "") {
        if (sii > 8.92)
            out := "Interpretation: Signal Intensity Index is greater than 8.92%. This suggests thymic hyperplasia with single dual echo technique.`n`n"
        else
            out := "Interpretation: Signal Intensity Index is less than or equal to 8.92%. Findings are not consistent with typical thymic hyperplasia with single dual echo technique.`n`n"
    } else {
        out := "Error: Both Chemical Shift Ratio and Signal Intensity Index are missing.`n`n"
    }
    if showCit
        out .= "Citation: Priola AM, Priola SM, Ciccone G, Evangelista A, Cataldi A, Gned D, Pazè F, Ducco L, Moretti F, Brundu M, Veltri A. Differentiation of rebound and lymphoid thymic hyperplasia from anterior mediastinal tumors with dual-echo chemical-shift MR imaging in adulthood: reliability of the chemical-shift ratio and signal intensity index. Radiology. 2015 Jan;274(1):238-49. doi: 10.1148/radiol.14132665. Epub 2014 Aug 7. PMID: 25105246.`n"
    return out
}
