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
#Include ..\FormGui.ahk

ThymusChemicalShift_Entry(input) {
    ShowThymusDialog(input)
    return ""
}

ShowThymusDialog(text := "") {
    tIp := "", tOp := "", pIp := "", pOp := ""
    if RegExMatch(text, "i)thymus[^0-9]*?(?:in[- ]?phase|IP|T1IP)\D*(\d+)\D*(?:out[- ]?of[- ]?phase|OP|OOP|T1OP)\D*(\d+)", &m) {
        tIp := m[1] + 0
        tOp := m[2] + 0
    }
    if RegExMatch(text, "i)paraspinous[^0-9]*?(?:in[- ]?phase|IP|T1IP)\D*(\d+)\D*(?:out[- ]?of[- ]?phase|OP|OOP|T1OP)\D*(\d+)", &m) {
        pIp := m[1] + 0
        pOp := m[2] + 0
    }

    form := RadsForm("Thymus Chemical Shift", 460)
    form.Header("Thymus signal")
    form.Numeric("Tip", "In-phase (IP):", tIp)
    form.Numeric("Top", "Out-of-phase (OP):", tOp)
    form.Header("Paraspinous reference (optional)")
    form.Numeric("Pip", "Paraspinous IP:", pIp)
    form.Numeric("Pop", "Paraspinous OP:", pOp)
    form.SetSubmit(Thymus_OnSubmit)
    form.AddButtons()
    form.Show()
}

Thymus_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (v.Tip = "" || v.Top = "")
        return MakeResult({ impression: "Thymus IP and OP are required.",
                            error: "Missing thymus IP/OP" })

    thIP := v.Tip + 0.0
    thOP := v.Top + 0.0
    sii := (thIP != 0) ? ((thIP - thOP) / thIP) * 100 : 0

    psIP := (v.Pip = "") ? "" : v.Pip + 0.0
    psOP := (v.Pop = "") ? "" : v.Pop + 0.0
    csr := ""
    if (psIP != "" && psOP != "") {
        opRatio := psOP != 0 ? thOP / psOP : 0
        ipRatio := psIP != 0 ? thIP / psIP : 0
        csr := ipRatio != 0 ? opRatio / ipRatio : 0
    }

    interp := _InterpretThymus(csr, sii)
    ; Strip the trailing citation that _InterpretThymus appends; citation is
    ; now a structured field. Also strip the "Interpretation: " prefix and
    ; trailing newlines to make the impression sentence-ready.
    interpClean := RegExReplace(interp, "Citation:.*$",,, 1)
    interpClean := RegExReplace(interpClean, "^Interpretation:\s*", "")
    interpClean := RTrim(interpClean, " `t`r`n")

    impression := "Thymic chemical-shift MRI: "
    if (csr != "")
        impression .= "CSR " Round(csr, 3) " (hyperplasia <= 0.849), "
    impression .= "SII " Round(sii, 2) "% (hyperplasia > 8.92%). " interpClean
    if (SubStr(impression, -1) != ".")
        impression .= "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nThymus IP: " thIP ", OP: " thOP
    method .= "`nSII = (IP - OP) / IP * 100 = " Round(sii, 2) "%"
    if (csr != "") {
        method .= "`nParaspinous IP: " psIP ", OP: " psOP
        method .= "`nCSR = (Thymus OP / Paraspinous OP) / (Thymus IP / Paraspinous IP) = " Round(csr, 3)
    }

    return MakeResult({
        classification: csr != "" ? Format("CSR {:.3f}, SII {:.2f}%", csr, sii)
                                  : Format("SII {:.2f}%", sii),
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [{ text: "Priola AM, Priola SM, Ciccone G, et al. "
                                . "Differentiation of rebound and lymphoid thymic hyperplasia from "
                                . "anterior mediastinal tumors with dual-echo chemical-shift MR "
                                . "imaging in adulthood. Radiology. 2015 Jan;274(1):238-49.",
                           url:  "https://pubs.rsna.org/doi/10.1148/radiol.14132665" }],
        echo:           g_LastSelectedText
    })
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
