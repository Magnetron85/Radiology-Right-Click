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
#Include ..\FormGui.ahk

PSADensity_Entry(input) {
    ShowPSADensityDialog(input)
    return ""
}

ShowPSADensityDialog(text := "") {
    ; Use the raw regex-captured strings to preserve the user's original
    ; input precision. Converting via `+ 0.0` produces a Float whose
    ; AHK-v2 string representation is the full IEEE 754 round-trip
    ; (e.g. 5.6 -> "5.5999999999999996"), and `Round(value, 2)` always
    ; returns a string formatted to exactly 2 decimal places (so "3.5"
    ; becomes "3.50"). Both behaviors land in the form's Edit text.
    psa := ""
    if RegExMatch(text, "i)PSA\s*(?:level|value)?:?\s*(\d+(?:\.\d+)?)", &m)
        psa := m[1]

    vol := ""
    if RegExMatch(text, "i)volume:?\s*(\d+(?:\.\d+)?)\s*(?:cc|cm3|mL|ml)", &m)
        vol := m[1]

    a := "", b := "", c := "", unit := "cm"
    if RegExMatch(text, "i)(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*(cm|mm)?", &m) {
        a := m[1]
        b := m[2]
        c := m[3]
        if (m[4] != "")
            unit := m[4]
    }

    form := RadsForm("PSA Density", 480)
    form.Header("PSA")
    form.Numeric("PSA", "PSA (ng/mL):", psa)
    form.Header("Volume (provide one)")
    form.Numeric("Vol", "Prostate volume (cc, blank to compute):", vol)
    form.Header("Dimensions (used when volume is blank)")
    form.Note("Bullet volume used below 55 cc, ellipsoid above.")
    form.Numeric("A", "A:", a)
    form.Numeric("B", "B:", b)
    form.Numeric("C", "C:", c)
    form.Dropdown("Unit", "Unit:", ["cm", "mm"], unit = "mm" ? 2 : 1)

    ; The two volume methods are mutually exclusive: either the user supplies
    ; a measured volume, OR the calc derives one from the three dimensions.
    ; Progressive disclosure: when Vol is populated the entire Dimensions
    ; section disappears (the form reflows and shrinks); clearing Vol brings
    ; it back with any previously typed dimensions restored.
    form.OnChange("Vol", _PSAD_UpdateVol)
    _PSAD_UpdateVol(form)

    form.SetSubmit(PSADensity_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_PSAD_UpdateVol(frm) {
    v := frm.GetValue("Vol")
    hasVol := (v != "" && v != 0)
    frm.SetSectionVisible("Dimensions (used when volume is blank)", !hasVol)
}

PSADensity_OnSubmit(v, form := "") {
    global g_LastSelectedText
    if (v.PSA = "")
        return MakeResult({ impression: "Please enter a PSA value.",
                            error: "Missing PSA" })
    synth := "PSA: " v.PSA " ng/mL"
    if (v.Vol != "")
        synth .= "`nVolume: " v.Vol " cc"
    else if (v.A != "" && v.B != "" && v.C != "")
        synth .= "`nSize: " v.A " x " v.B " x " v.C " " v.Unit
    else
        return MakeResult({ impression: "Please provide either a volume or three dimensions.",
                            error: "Missing volume / dimensions" })
    body := CalcPSADensity(synth)

    ; Extract volume and density values out of the raw body for an
    ; impression-ready sentence; preserve the full body in methodology.
    volVal := ""
    volMethod := ""
    if RegExMatch(body, "i)Prostate volume[:]?\s*([\d.]+)\s*cc\s*\(([^)]+)\)", &m) {
        volVal := m[1]
        volMethod := m[2]
    } else if RegExMatch(body, "i)Volume[:]?\s*([\d.]+)\s*cc", &m)
        volVal := m[1]
    else if (v.Vol != "")
        volVal := v.Vol
    densVal := ""
    if RegExMatch(body, "i)PSA Density[:]?\s*([\d.]+)", &m)
        densVal := m[1]

    ; The impression leads with the DERIVED values. Inputs the highlighted
    ; selection already states are NOT restated -- the paste composition
    ; appends the impression right after the user's own selection, so
    ; repeating them reads as duplication ("PSA: 5.6 / PSA 5.6 ng/mL ...").
    ; A computed volume IS new information, so it stays; the PSA is included
    ; only when the selection doesn't carry it (e.g. launched from the menu
    ; with no selection and typed into the form) so the sentence still
    ; stands alone. A user-supplied volume is an input, never restated.
    pasteFrag := ""
    if (densVal != "") {
        ; ng/mL/cc is the clinical-convention unit for PSA density (also
        ; written ng/mL^2 or ng/mL/cm^3); we keep "cc" rather than the
        ; style-guide-default "mL" because PSA density is conventionally
        ; reported this way -- 1 cc == 1 mL but the form is universal.
        densPart := "PSA density " densVal " ng/mL/cc"
        volPart  := (volVal != "" && volMethod != "")
                  ? "prostate volume " volVal " cc" : ""
        selHasPSA := RegExMatch(g_LastSelectedText
                   , "i)PSA\D{0,20}" StrReplace(v.PSA, ".", "\."))
        psaPart := selHasPSA ? "" : "PSA " v.PSA " ng/mL"

        impression := (psaPart != "" ? psaPart ", " : "") . densPart
        if (volPart != "")
            impression .= " (" volPart ", "
                       . RegExReplace(StrLower(volMethod), "\s*volume$", "") " formula)"
        impression .= "."

        ; Short parenthetical fragment: lowercase-compatible, no period.
        pasteFrag := (psaPart != "" ? psaPart ", " : "")
                   . (volPart != "" ? volPart ", " : "") . densPart
    } else {
        impression := body
    }

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`n`nRaw computation:`n" body

    return MakeResult({
        classification: "PSA density",
        impression:     impression,
        recommendation: "",
        methodology:    method,
        citations:      [],
        echo:           g_LastSelectedText,
        paste:          pasteFrag,
        pasteMode:      ""   ; paren after a single-line selection; newline otherwise
    })
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

    ; Two decimals, NOT one: the clinically used PSAD cutoff is 0.15
    ; ng/mL/cc -- rounding to one decimal (0.1 / 0.2) destroys the answer
    ; right at the decision boundary.
    density := Round(psaLevel / (prostateVolume + 0), 2)
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
