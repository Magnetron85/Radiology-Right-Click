; ============================================================
; lib/calc/PSADensity.ahk -- PSA density (PSA / prostate volume)
; ------------------------------------------------------------
; Volume preference order:
;   1) Explicit "Volume: NN cc" in text
;   2) "Size: ... (NN cc)" in text
;   3) Computed from 3D dimensions in text -- bullet under 55 cc,
;      else ellipsoid (the v1 convention is preserved).
; ============================================================

#Requires AutoHotkey v2.0
#Include Volumes.ahk

PSADensity_Entry(input) {
    return CalcPSADensity(input)
}

CalcPSADensity(input) {
    volumeMethod := "User Supplied"
    volNotGiven  := true

    if !RegExMatch(input
        , "i)PSA\s*(?:level|value)?:?\s*(\d+(?:\.\d+)?)(?:\s*(?:ng/ml|ng/cc))?"
        , &psa)
        return "Invalid format for PSA density.`nExample:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm"

    psaLevel := psa[1] + 0
    prostateVolume := ""

    if RegExMatch(input
        , "i)(?:(?:volume:?\s*(\d+(?:\.\d+)?)(?:\s*(?:cc|cm3|mL|ml)))|(?:Prostate )?Size:?.*?\((\d+(?:\.\d+)?)\s*cc\))"
        , &vm) {
        if (vm.Count >= 1 && vm[1] != "")
            prostateVolume := vm[1] + 0
        else if (vm.Count >= 2 && vm[2] != "")
            prostateVolume := vm[2] + 0
    }

    if (prostateVolume = "") {
        ; derive from dimensions
        bullet := CalcBulletVolume(input)
        if InStr(bullet, "Invalid input")
            return "Prostate volume not found.`nExample:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm"
        prostateVolume := ParseTrailingCC(bullet)
        volumeMethod := "Bullet Volume"
        if (prostateVolume + 0 >= 55) {
            elli := CalcEllipsoidVolume(input)
            if !InStr(elli, "Invalid input") {
                prostateVolume := ParseTrailingCC(elli)
                volumeMethod := "Ellipsoid Volume"
            }
        }
        volNotGiven := false
    }

    if (prostateVolume = "" || (prostateVolume + 0) = 0)
        return "Could not determine prostate volume.`nExample:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm"

    density := Round(psaLevel / (prostateVolume + 0), 3)
    units := Prefs.Get("display","units",true) ? " ng/mL/cc" : ""

    if !volNotGiven {
        return input
            . "`nProstate volume: " prostateVolume " cc (" volumeMethod ")"
            . "`nPSA Density: " density units
    }
    return input "`nPSA Density: " density units
}

; Pull the "(NN.NN cc)" trailing value out of a CalcVolume result string.
ParseTrailingCC(s) {
    if RegExMatch(s, "\((\d+(?:\.\d+)?)(?:\s*cc)?\)\s*$", &m)
        return m[1]
    return ""
}
