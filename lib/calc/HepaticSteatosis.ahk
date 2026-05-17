; ============================================================
; lib/calc/HepaticSteatosis.ahk
; ------------------------------------------------------------
; Reference: Sirlin CB. Invited Commentary on Image-based
; quantification of hepatic fat. Radiographics 2009;29:1277-80.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

HepaticSteatosis_Entry(input) {
    return CalcHepaticSteatosis(input)
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

    fatFraction := 100 * (liverIP - liverOP) / (2 * liverIP)
    out := input " (Fat Fraction: " Round(fatFraction, 1) "%)"
    out .= "`n`nFat Fraction " _InterpretSteatosis(fatFraction) "`n"

    if RegExMatch(input, spleenNeedle, &sm) {
        spleenIP := sm[2] + 0, spleenOP := sm[4] + 0
        if (spleenIP != 0 && spleenOP != 0) {
            fatPct := 100 * ((liverIP/spleenIP) - (liverOP/spleenOP)) / (2 * (liverIP/spleenIP))
            out := StrReplace(out, ")", ", Fat Percentage: " Round(fatPct, 1) "%)",, , 1)
            out .= "Fat Percentage " _InterpretSteatosis(fatPct) "`n"
        }
    }

    if showCit
        out .= "`n`nSirlin CB. Invited Commentary on Image-based quantification of hepatic fat: methods and clinical applications. Radiographics 2009; 29:1277-80`n"
    return out
}

_InterpretSteatosis(ff) {
    if (ff < 5)
        return "Interpretation: No significant hepatic steatosis."
    if (ff < 15)
        return "Interpretation: Mild hepatic steatosis."
    if (ff < 30)
        return "Interpretation: Moderate hepatic steatosis."
    return "Interpretation: Severe hepatic steatosis."
}
