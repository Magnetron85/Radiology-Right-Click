; ============================================================
; lib/calc/LiverIron.ahk -- MRI R2* -> liver iron content
; ------------------------------------------------------------
; Reference: Guglielmo FF et al. Radiographics. 2023 Jun;43(6):
; e220181. doi: 10.1148/rg.220181.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

LiverIron_Entry(input) {
    return CalcLiverIron(input)
}

CalcLiverIron(input) {
    if !RegExMatch(input, "i)(?:\b|^)(1[.,]5|3[.,]0|1\.5|3\.0|2\.89)(?:\s*-?\s*)?T(?:esla)?", &fs)
        return "Error: Magnetic field strength (1.5T, 2.89T or 3.0T) not found in the input."

    field := StrReplace(fs[1], ",", ".")
    r2pat := "i)R2\*?\s*(?:value|reading|measurement)?(?:[\s:=]+of)?\s*[:=]?\s*(\d+(?:[.,]\d+)?)\s*(?:Hz|hertz|s(?:ec(?:ond)?)?|1/s)"
    if !RegExMatch(input, r2pat, &r)
        return "Error: R2* value not found in the input.`nSample syntax: 1.5T, R2*: 50 Hz"

    r2 := StrReplace(r[1], ",", ".") + 0.0

    if (field = "1.5")
        iron := 0.02603 * r2 - 0.16
    else if (field = "2.89")
        iron := 0.01400 * r2 - 0.03
    else if (field = "3.0")
        iron := 0.01349 * r2 - 0.03
    else
        iron := 0

    out := input "`nEstimated Iron Content: " Round(iron, 1) " mg Fe/g dry liver`n"
    out .= "Interpretation: " _InterpretIronOverload(iron) "`n"
    if Prefs.Get("display","showCitations",true)
        out .= "`nGuglielmo FF, Barr RG, Yokoo T, Ferraioli G, Lee JT, Dillman JR, Horowitz JM, Jhaveri KS, Miller FH, Modi RY, Mojtahed A, Ohliger MA, Pirasteh A, Reeder SB, Shanbhogue K, Silva AC, Smith EN, Surabhi VR, Taouli B, Welle CL, Yeh BM, Venkatesh SK. Liver Fibrosis, Fat, and Iron Evaluation with MRI and Fibrosis and Fat Evaluation with US: A Practical Guide for Radiologists. Radiographics. 2023 Jun;43(6):e220181. doi: 10.1148/rg.220181. PMID: 37227944.`n"
    return out
}

; Severity bands from Guglielmo 2023 Table 6 (mg Fe/g dry liver).
_InterpretIronOverload(lic) {
    if (lic < 1.8)
        return "Normal iron content (<1.8 mg Fe/g)."
    if (lic < 3.2)
        return "Mild iron overload (1.8 - 3.2 mg Fe/g)."
    if (lic < 7.0)
        return "Moderate iron overload (3.2 - 7.0 mg Fe/g)."
    if (lic < 15.0)
        return "Severe iron overload (7.0 - 15.0 mg Fe/g)."
    return "Extreme iron overload (>=15.0 mg Fe/g)."
}
