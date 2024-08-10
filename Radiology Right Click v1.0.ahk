; ==========================================
; Radiologist's Helper Script
; Radiology Right Click
; Version: 1.0
; Description: This AutoHotkey script provides various calculation tools and utilities
;              for radiologists, including volume calculations, date estimations,
;              and statistical analysis of measurements.
; ==========================================

#NoEnv

#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%

; Global variables
global DisplayUnits := true
global DisplayAllValues := true
global ShowEllipsoidVolume := true
global ShowBulletVolume := true
global ShowPSADensity := true
global ShowPregnancyDates := true
global ShowMenstrualPhase := true
global ShowAdrenalWashout := true
global ShowThymusChemicalShift := true
global ShowHepaticSteatosis := true
global ShowMRILiverIron := true
global ShowStatistics := true
global ShowNumberRange := true
global PauseDuration := 180000
global DarkMode := false
global ShowCalciumScorePercentile := true
global ShowCitations := true
global ShowArterialAge :=true

; Load preferences (keep this function at the top)
LoadPreferencesFromFile() {
    global DisplayUnits, DisplayAllValues, ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity, ShowPregnancyDates, ShowMenstrualPhase, PauseDuration
    global ShowAdrenalWashout, ShowThymusChemicalShift, ShowHepaticSteatosis, ShowMRILiverIron, ShowStatistics, ShowNumberRange, DarkMode
	global ShowCitations, ShowArterialAge

    preferencesFile := A_ScriptDir . "\preferences.ini"
    if (FileExist(preferencesFile)) {
        IniRead, DisplayUnits, %preferencesFile%, Display, DisplayUnits, 1
        IniRead, DisplayAllValues, %preferencesFile%, Display, DisplayAllValues, 1
        IniRead, ShowEllipsoidVolume, %preferencesFile%, Calculations, ShowEllipsoidVolume, 1
        IniRead, ShowBulletVolume, %preferencesFile%, Calculations, ShowBulletVolume, 1
        IniRead, ShowPSADensity, %preferencesFile%, Calculations, ShowPSADensity, 1
        IniRead, ShowPregnancyDates, %preferencesFile%, Calculations, ShowPregnancyDates, 1
        IniRead, ShowMenstrualPhase, %preferencesFile%, Calculations, ShowMenstrualPhase, 1
        IniRead, ShowAdrenalWashout, %preferencesFile%, Calculations, ShowAdrenalWashout, 1
        IniRead, ShowThymusChemicalShift, %preferencesFile%, Calculations, ShowThymusChemicalShift, 1
        IniRead, ShowHepaticSteatosis, %preferencesFile%, Calculations, ShowHepaticSteatosis, 1
        IniRead, ShowMRILiverIron, %preferencesFile%, Calculations, ShowMRILiverIron, 1
        IniRead, ShowStatistics, %preferencesFile%, Calculations, ShowStatistics, 1
        IniRead, ShowNumberRange, %preferencesFile%, Calculations, ShowNumberRange, 1
        IniRead, PauseDuration, %preferencesFile%, Script, PauseDuration, 180000
        IniRead, DarkMode, %preferencesFile%, Display, DarkMode, 0
		IniRead, ShowCalciumScorePercentile, %preferencesFile%, Calculations, ShowCalciumScorePercentile, 1
		IniRead, ShowCitations, %preferencesFile%, Display, ShowCitations, 1
		IniRead, ShowArterialAge, %preferencesFile%, Display, ShowArterialAge, 1
    } else {
        MsgBox, Preferences file not found: %preferencesFile%. File will be created if preferences are edited.
    }

    DisplayUnits := (DisplayUnits = "1")
    DisplayAllValues := (DisplayAllValues = "1")
    ShowEllipsoidVolume := (ShowEllipsoidVolume = "1")
    ShowBulletVolume := (ShowBulletVolume = "1")
    ShowPSADensity := (ShowPSADensity = "1")
    ShowPregnancyDates := (ShowPregnancyDates = "1")
    ShowMenstrualPhase := (ShowMenstrualPhase = "1")
    ShowAdrenalWashout := (ShowAdrenalWashout = "1")
    ShowThymusChemicalShift := (ShowThymusChemicalShift = "1")
    ShowHepaticSteatosis := (ShowHepaticSteatosis = "1")
    ShowMRILiverIron := (ShowMRILiverIron = "1")
    ShowStatistics := (ShowStatistics = "1")
    ShowNumberRange := (ShowNumberRange = "1")
    DarkMode := (DarkMode = "1")
	ShowCalciumScorePercentile := (ShowCalciumScorePercentile = "1")
	ShowCitations := (ShowCitations = "1")
	ShowArterialAge := (ShowArterialAge = "1")
    PauseDuration += 0
}

LoadPreferencesFromFile()

; ==========================================
; Main Script Logic
; ==========================================

; Set up context menu for target applications
#If WinActive("ahk_class Notepad") or WinActive("ahk_exe notepad.exe") or WinActive("ahk_class PowerScribe") or WinActive("ahk_exe PowerScribe.exe") or WinActive("ahk_class PowerScribe360") or WinActive("ahk_exe Nuance.PowerScribe360.exe") or WinActive("ahk_class PowerScribe | Reporting")
*RButton::
{
    SelectedText := GetSelectedText()
    CreateCustomMenu()
    Menu, CustomMenu, Show
    Menu, CustomMenu, DeleteAll
    return
}
#If

CreateCustomMenu() {
    global ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity, ShowPregnancyDates, ShowMenstrualPhase
    global ShowAdrenalWashout, ShowThymusChemicalShift, ShowHepaticSteatosis
    global ShowMRILiverIron, ShowStatistics, ShowNumberRange

    Menu, CustomMenu, Add, Cut, MenuCut
    Menu, CustomMenu, Add, Copy, MenuCopy
    Menu, CustomMenu, Add, Paste, MenuPaste
    Menu, CustomMenu, Add, Delete, MenuDelete
    Menu, CustomMenu, Add
    Menu, CustomMenu, Add, Compare Measurement Sizes, CompareNoduleSizes
    Menu, CustomMenu, Add, Sort Measurement Sizes, SortNoduleSizes
	
	if (ShowCalciumScorePercentile) {
		Menu, CustomMenu, Add, Calculate Calcium Score Percentile, CalculateCalciumScorePercentile
	}
    if (ShowEllipsoidVolume) {
        Menu, CustomMenu, Add, Calculate Ellipsoid Volume, CalculateEllipsoidVolume
    }
    if (ShowBulletVolume) {
        Menu, CustomMenu, Add, Calculate Bullet Volume, CalculateBulletVolume
    }
    if (ShowPSADensity) {
        Menu, CustomMenu, Add, Calculate PSA Density, CalculatePSADensity
    }
    if (ShowPregnancyDates) {
        Menu, CustomMenu, Add, Calculate Pregnancy Dates, CalculatePregnancyDates
    }
    if (ShowMenstrualPhase) {
        Menu, CustomMenu, Add, Calculate Menstrual Phase, CalculateMenstrualPhase
    }
    if (ShowAdrenalWashout) {
        Menu, CustomMenu, Add, Calculate Adrenal Washout, CalculateAdrenalWashout
    }
    if (ShowThymusChemicalShift) {
        Menu, CustomMenu, Add, Calculate Thymus Chemical Shift, CalculateThymusChemicalShift
    }
    if (ShowHepaticSteatosis) {
        Menu, CustomMenu, Add, Calculate Hepatic Steatosis, CalculateHepaticSteatosis
    }
    if (ShowMRILiverIron) {
        Menu, CustomMenu, Add, MRI Liver Iron Content, CalculateIronContent
    }
    if (ShowStatistics) {
        Menu, CustomMenu, Add, Calculate Statistics, Statistics
    }
    if (ShowNumberRange) {
        Menu, CustomMenu, Add, Calculate Number Range, Range
    }
    Menu, CustomMenu, Add
    Menu, CustomMenu, Add, Pause Script, PauseScript
    Menu, CustomMenu, Add, Preferences, ShowPreferences
}

MenuCut:
    Send, ^x
return

MenuCopy:
    Send, ^c
return

MenuPaste:
    Send, ^v
return

MenuDelete:
    Send, {Delete}
return

CalculateCalciumScorePercentile:
    Result := CalculateCalciumScorePercentile(SelectedText)
    ShowResult(Result)
return

GetSelectedText() {
    OldClipboard := ClipboardAll
    Clipboard := ""
    Send, ^c
    ClipWait, 0.1
    if (ErrorLevel)
    {
        Clipboard := OldClipboard
        return ""
    }
    SelectedText := Clipboard
    Clipboard := OldClipboard
    return SelectedText
}

ShowResult(Result) {
    global DarkMode
    options := DarkMode ? 0x1000000 : 0  ; 0x1000000 is for dark mode
    MsgBox, % options, Result, %Result%
    Clipboard := Result
}

CalculateEllipsoidVolume:
    Result := CalculateEllipsoidVolume(SelectedText)
    ShowResult(Result)
return

CalculateBulletVolume:
    Result := CalculateBulletVolume(SelectedText)
    ShowResult(Result)
return

CalculatePSADensity:
    Result := CalculatePSADensity(SelectedText)
    ShowResult(Result)
return

CalculatePregnancyDates:
    Result := CalculatePregnancyDates(SelectedText)
    ShowResult(Result)
return

CalculateMenstrualPhase:
    Result := CalculateMenstrualPhase(SelectedText)
    ShowResult(Result)
return

SortNoduleSizes:
    ProcessedText := ProcessAllNoduleSizes(SelectedText)
    if (ProcessedText != SelectedText)
    {
        ; Preserve leading and trailing spaces only if they existed in the original input
        leadingSpace := (SubStr(SelectedText, 1, 1) == " ") ? " " : ""
        trailingSpace := (SubStr(SelectedText, 0) == " ") ? " " : ""
        Clipboard := leadingSpace . Trim(ProcessedText) . trailingSpace
        Send, ^v
    }
return

CalculateEllipsoidVolume(input) {
    RegExNeedle := "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
    if (RegExMatch(input, RegExNeedle, match)) {
        dimensions := [match1, match2, match3]
        dimensions := SortDimensions(dimensions)
        isMillimeters := (InStr(dimensions[1], ".") = 0) && (InStr(dimensions[2], ".") = 0) && (InStr(dimensions[3], ".") = 0)
        if (isMillimeters) {
            dimensions[1] := dimensions[1] / 10
            dimensions[2] := dimensions[2] / 10
            dimensions[3] := dimensions[3] / 10
        }
        volume := (1/6) * 3.14159265358979323846 * (dimensions[1]) * (dimensions[2]) * (dimensions[3])
        volumeRounded := (volume < 1) ? Round(volume, 3) : Round(volume, 1)
        result := input . " (" . volumeRounded . (DisplayUnits ? " cc" : "") . ")"
        return result
    } else {
        return "Invalid input format for ellipsoid volume calculation.`nSample syntax: 3 x 2 x 1 cm"
    }
}

CalculateBulletVolume(input) {
    RegExNeedle := "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
    
    if (RegExMatch(input, RegExNeedle, match)) {
        dimensions := [match1, match2, match3]
        dimensions := SortDimensions(dimensions)
        
        isMillimeters := (InStr(dimensions[1], ".") = 0) && (InStr(dimensions[2], ".") = 0) && (InStr(dimensions[3], ".") = 0)
        
        if (isMillimeters) {
            dimensions[1] := dimensions[1] / 10
            dimensions[2] := dimensions[2] / 10
            dimensions[3] := dimensions[3] / 10
        }
        
        volume := dimensions[1] * dimensions[2] * dimensions[3] * (5 * 3.14159265358979323846 / 24)
        volumeRounded := (volume < 1) ? Round(volume, 3) : Round(volume, 1)
        
        result := input . " (" . volumeRounded . (DisplayUnits ? " cc" : "") . ")"
        return result
    } else {
        return "Invalid input format for bullet volume calculation.`nSample syntax: 3 x 2 x 1 cm"
    }
}

CalculatePSADensity(input) {
    ; Regular expression for PSA value
    PSARegEx := "i)PSA\s*(?:level|value)?:?\s*(\d+(?:\.\d+)?)\s*(?:ng\/ml|ng\/mL|ng/ml|ng/mL|ng\/cc|ng/cc)?"
    
    ; Regular expression for prostate volume (all formats)
    VolumeRegEx := "i)(?:(?:Calculated\s*)?(?:ellipsoid\s*)?volume:?\s*(\d+(?:\.\d+)?)\s*(?:cc|cm3|mL|ml)|(?:Prostate )?Size:?.*?\((\d+(?:\.\d+)?)\s*cc\)|(\d+(?:\.\d+)?)\s*x\s*\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?\s*cm\s*\((\d+(?:\.\d+)?)\s*cc\))"

    if (RegExMatch(input, PSARegEx, PSAMatch)) {
        PSALevel := PSAMatch1
    } else {
        return "Invalid input format for PSA density calculation.`nSuggested format:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm (24.7 cc)`nor`nPSA level: 4.5 ng/mL, Prostate volume: 30 cc"
    }

    if (RegExMatch(input, VolumeRegEx, VolumeMatch)) {
        ProstateVolume := VolumeMatch1 ? VolumeMatch1 : (VolumeMatch2 ? VolumeMatch2 : (VolumeMatch4 ? VolumeMatch4 : ""))
    } else {
        return "Prostate volume not found in the selected text.`nSuggested format:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm (24.7 cc)`nor`nPSA level: 4.5 ng/mL, Prostate volume: 30 cc"
    }

    if (ProstateVolume == "") {
        return "Prostate volume not found in the selected text.`nSuggested format:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm (24.7 cc)`nor`nPSA level: 4.5 ng/mL, Prostate volume: 30 cc"
    }

    PSADensity := PSALevel / ProstateVolume
    PSADensity := Round(PSADensity, 3)

    result := input . "`nPSA Density: " . PSADensity . (DisplayUnits ? " ng/mL/cc" : "")
    ;if (DisplayAllValues)
    ;    result .= "`nPSA: " . PSALevel . (DisplayUnits ? " ng/mL" : "") . "`nProstate Volume: " . Round(ProstateVolume, 1) . (DisplayUnits ? " cc" : "")
    return result
}

CalculatePregnancyDates(input) {
    LMPRegEx := "i)(?:LMP|Last\s*Menstrual\s*Period).*?(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})"
    GARegEx := "i)(\d+)(?:\s*(?:weeks?|w))?\s*(?:and|&|,|-|;)?\s*(\d+)?(?:\s*(?:days?|d))?(?:\s*[,;-]?\s+(?:as of|on)\s+(today|\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}))?"

    if (RegExMatch(input, LMPRegEx, LMPMatch)) {
        LMPDate := ParseDate(LMPMatch1)
        if (LMPDate = "Invalid Date") {
            return "Invalid LMP date format. Please use MM/DD/YYYY, DD/MM/YYYY, or MM/DD/YY."
        }
        return CalculateDatesFromLMP(LMPDate)
    } else if (RegExMatch(input, GARegEx, GAMatch)) {
        WeeksGA := GAMatch1 + 0
        DaysGA := (GAMatch2 != "") ? GAMatch2 + 0 : 0
        ReferenceDate := (GAMatch3 != "") ? (GAMatch3 = "today" ? A_Now : ParseDate(GAMatch3)) : A_Now
        return CalculateDatesFromGA(WeeksGA, DaysGA, ReferenceDate)
    } else {
        return "Invalid input format for pregnancy date calculation.`nSample syntax:`nLMP: 01/15/2023`nor`nGA: 12 weeks and 3 days as of today"
    }
}

CalculateMenstrualPhase(input) {
    LMPRegEx := "i)(?:LMP|Last\s*Menstrual\s*Period)?\s*:?\s*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})"
    if (RegExMatch(input, LMPRegEx, LMPMatch)) {
        LMPDate := ParseDate(LMPMatch1)
        if (LMPDate = "Invalid Date") {
            return "Invalid LMP date format. Please use MM/DD/YYYY, DD/MM/YYYY, or MM/DD/YY."
        }
        return DetermineMenstrualPhase(LMPDate)
    } else {
        return "Invalid input format for menstrual phase calculation.`nSample syntax: LMP: 05/01/2023"
    }
}

ProcessAllNoduleSizes(input) {
    RegExNeedleComma3 := "\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*"
    RegExNeedleX3 := "\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*"
    RegExNeedleComma2 := "\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*"
    RegExNeedleX2 := "\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*"

    input := ProcessPattern(input, RegExNeedleComma3, 3)
    input := ProcessPattern(input, RegExNeedleX3, 3)
    input := ProcessPattern(input, RegExNeedleComma2, 2)
    input := ProcessPattern(input, RegExNeedleX2, 2)

    return input
}

ProcessPattern(input, RegExNeedle, dimensions) {
    pos := 1
    while (pos := RegExMatch(input, RegExNeedle, match, pos)) {
        if (dimensions == 3) {
            processed := ProcessNoduleSizes(match1, match2, match3)
        } else if (dimensions == 2) {
            processed := ProcessNoduleSizes(match1, match2)
        }
        if (processed != match) {
            input := SubStr(input, 1, pos-1) . processed . SubStr(input, pos+StrLen(match))
        }
        pos += StrLen(processed)
    }
    return input
}

ProcessNoduleSizes(a, b, c := "") {
    ; Convert to numbers for comparison, but keep original strings
    aNum := a + 0
    bNum := b + 0
    cNum := c != "" ? c + 0 : ""

    if (c != "") {
        if (aNum < bNum) {
            temp := a
            a := b
            b := temp
            tempNum := aNum
            aNum := bNum
            bNum := tempNum
        }
        if (bNum < cNum) {
            temp := b
            b := c
            c := temp
            tempNum := bNum
            bNum := cNum
            cNum := tempNum
        }
        if (aNum < bNum) {
            temp := a
            a := b
            b := temp
            tempNum := aNum
            aNum := bNum
            bNum := tempNum
        }
        return " " . Trim(a) . " x " . Trim(b) . " x " . Trim(c) . " "
    } else {
        if (aNum < bNum) {
            temp := a
            a := b
            b := temp
        }
        return " " . Trim(a) . " x " . Trim(b) . " "
    }
}

Statistics:
    Result := CalculateStatistics(SelectedText)
    ShowResult(Result)
return

CalculateStatistics(input) {
    numbers := ExtractNumbers(input)
    count := numbers.Length()
    if (count == 0) {
        return "No numbers found in the selected text."
    }

    result := "Statistics:`n"
    result .= "Count: " . count . "`n"
	result .= "Sum: " . Round(CalculateSum(numbers), 1) . "`n"
	;result .= "Product: " . CalculateProduct(numbers) . "`n"
    result .= "Mean: " . Round(CalculateMean(numbers), 1) . "`n"
    result .= "Median: " . Round(CalculateMedian(numbers), 1) . "`n"
    result .= "Min: " . Round(Min(numbers*), 1) . "`n"
    result .= "Max: " . Round(Max(numbers*), 1) . "`n"

    if (count >= 9) {
        Q1 := Round(CalculateQuartile(numbers, 0.25), 1)
        Q3 := Round(CalculateQuartile(numbers, 0.75), 1)
        IQR := Q3 - Q1
        median := Round(CalculateMedian(numbers), 1)
        result .= "Q1: " . Q1 . "`n"
        result .= "Q3: " . Q3 . "`n"
        result .= "IQR: " . Round(IQR, 1) . "`n"
        result .= "IQR/Median: " . Round(IQR / median, 2) . "`n"
        result .= "Standard Deviation: " . Round(CalculateStandardDeviation(numbers), 1) . "`n"
    }
    return result
}

ExtractNumbers(input) {
    numbers := []
    input := RegExReplace(input, "i)(?:^|\n|\s)(?:slice|observation|sample|number|#|no\.?)\s*\d+:?\s*", "`n")
    RegExNeedle := "(-?\d+(?:\.\d+)?)\s*(?:cm|mm)?"
    pos := 1
    while (pos := RegExMatch(input, RegExNeedle, match, pos)) {
        if (match1 != "") {
            numbers.Push(match1 + 0)
        }
        pos += StrLen(match)
    }
    return numbers
}

CalculateSum(arr) {
    total := 0
    for index, value in arr
        total += value
    return total
}

CalculateProduct(arr) {
    total := 1
    for index, value in arr
        total *= value
    return total
}

CalculateMean(numbers) {
    sum := 0
    for i, num in numbers {
        sum += num
    }
    return Round(sum / numbers.Length(), 2)
}

CalculateMedian(numbers) {
    sortedNumbers := SortArray(numbers)
    count := sortedNumbers.Length()
    if (count == 0) {
        return 0
    } else if (Mod(count, 2) == 0) {
        middle1 := sortedNumbers[count//2]
        middle2 := sortedNumbers[(count//2) + 1]
        return (middle1 + middle2) / 2
    } else {
        return sortedNumbers[Floor(count/2) + 1]
    }
}

SortArray(numbers) {
    sortedNumbers := []
    for index, value in numbers {
        sortedNumbers.Push(value)
    }
    sortedNumbers.Sort()
    return sortedNumbers
}

CalculateQuartile(numbers, percentile) {
    sortedNumbers := SortArray(numbers)
    count := sortedNumbers.Length()
    position := (count - 1) * percentile + 1
    lower := Floor(position)
    upper := Ceil(position)
    if (lower == upper) {
        return Round(sortedNumbers[lower], 2)
    } else {
        return Round(sortedNumbers[lower] + (position - lower) * (sortedNumbers[upper] - sortedNumbers[lower]), 2)
    }
}

CalculateStandardDeviation(numbers) {
    mean := CalculateMean(numbers)
    sumSquaredDiff := 0
    for i, num in numbers {
        diff := num - mean
        sumSquaredDiff += diff * diff
    }
    variance := sumSquaredDiff / (numbers.Length() - 1)
    return Round(Sqrt(variance), 2)
}

CalculateIronContent:
    Result := EstimateIronContent(SelectedText)
    ShowResult(Result)
return

EstimateIronContent(input) {
	global ShowCitations

    RegExMatch(input, "i)(?:\b|^)(1[.,]5|3[.,]0|1\.5|3\.0)(?:\s*-?\s*)?T(?:esla)?", fieldStrength)
    if (!fieldStrength1) {
        return "Error: Magnetic field strength (1.5T or 3.0T) not found in the input."
    }
    fieldStrength1 := StrReplace(fieldStrength1, ",", ".")

    R2StarPattern := "i)R2\*?\s*(?:value|reading|measurement)?(?:[\s:=]+of)?\s*[:=]?\s*(\d+(?:[.,]\d+)?)\s*(?:Hz|hertz|s(?:ec(?:ond)?)?Ã¢ÂÂ»Ã‚Â¹|1/s)"
    RegExMatch(input, R2StarPattern, R2Star)
    if (!R2Star1) {
        return "Error: R2* value not found in the input.`nSample syntax: 1.5T, R2*: 50 Hz"
    }
    R2StarValue := StrReplace(R2Star1, ",", ".")
    R2StarValue += 0

    if (fieldStrength1 == "1.5") {
		ironContent := -0.04 + 2.62 * 10 ** -2 * R2StarValue
    } else {
        ironContent := 1.41 * 10 ** -2 * R2StarValue
    }

    result := input . " (" . Round(ironContent, 2) . " mg Fe/g dry liver)"
	
    if (DisplayAllValues) {
        result .= "`nMagnetic Field Strength: " . fieldStrength1 . "T`n"
        result .= "R2* Value: " . R2StarValue . " Hz`n"
    }
	
	if (ShowCitations=1){
		result .= "`nHernando D, Cook RJ, Qazi N, Longhurst CA, Diamond CA, Reeder SB. Complex confounder-corrected R2* mapping for liver iron quantification with MRI. Eur Radiol. 2021 Jan;31(1):264-275. doi: 10.1007/s00330-020-07123-x. Epub 2020 Aug 12. PMID: 32785766; PMCID: PMC7755713.`n"
	}
    return result
}

Range:
    Result := CalculateRange(SelectedText)
    ShowResult(Result)
return

CalculateRange(input) {
    numbers := []
    unit := ""
    RegExNeedle := "(-?\d+(?:\.\d+)?)\s*((?:cm/s|mm/s|m/s|km/h|mph|cm|mm|Hz|T|mg|m|ml|mL|cc|s|min|hr|days?|weeks?|months?|years?|g|ng|ng/ml|ng/mL|mmol/L|ÃŽÂ¼mol/L|Ã‚Â°?F|Ã‚Â°?C)(?:/(?:day|week|month|year))?)?"
    pos := 1
    while (pos := RegExMatch(input, RegExNeedle, match, pos)) {
        numbers.Push(match1 + 0)
        if (match2 != "" && unit == "") {
            unit := match2
        }
        pos += StrLen(match)
    }

    if (numbers.Length() == 0) {
        return "No numbers found in the selected text."
    }

    minValue := Min(numbers*)
    maxValue := Max(numbers*)
    result := Round(minValue, 1) . " - " . Round(maxValue, 1)
    if (unit != "") {
        result .= " " . unit
    }
    return result
}

CalculateAdrenalWashout:
    Result := CalculateAdrenalWashout(SelectedText)
    ShowResult(Result)
return

CalculateAdrenalWashout(input) {
	global ShowCitations
	
    RegExNeedle := "i)(?:(?:unenhanced|non-?enhanced|intrinsic|pre-?contrast|baseline|native)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?).*?(?:(?:enhanced|post-?contrast|arterial|portal\s+venous?|60-?75\s*(?:second|sec)|1-?2\s*min(?:ute)?)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?).*?(?:(?:delayed|late|15\s*min(?:ute)?|10-?15\s*min(?:ute)?|post-?contrast)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?)"

    if (RegExMatch(input, RegExNeedle, match)) {
        unenhanced := match1
        enhanced := match2
        delayed := match3

        absoluteWashout := ((enhanced - delayed) / (enhanced - unenhanced)) * 100
        relativeWashout := ((enhanced - delayed) / enhanced) * 100

        result := input . "`n`n"
        result .= "Absolute Washout: " . Round(absoluteWashout, 1) . "% ... (Ref adenomas: >=60%)" . "`n"
        result .= "Relative Washout: " . Round(relativeWashout, 1) . "% ... (Ref adenomas: >=40%)" . "`n`n"
		
		if(ShowCitations=1){
				result .= "`nMayo-Smith WW, Song JH, Boland GL, Francis IR, Israel GM, Mazzaglia PJ, Berland LL, Pandharipande PV. Management of Incidental Adrenal Masses: A White Paper of the ACR Incidental Findings Committee. J Am Coll Radiol. 2017 Aug;14(8):1038-1044.`n`n"
		}
		
        result .= InterpretAdrenalWashout(absoluteWashout, relativeWashout, unenhanced, enhanced, delayed)
        return result
    } else {
        RegExNeedle := "i)(?:(?:enhanced|post-?contrast|arterial|portal\s+venous?|60-?75\s*(?:second|sec)|1-?2\s*min(?:ute)?)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\nits?)?).*?(?:(?:delayed|late|15\s*min(?:ute)?|10-?15\s*min(?:ute)?|post-?contrast)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?)"

        if (RegExMatch(input, RegExNeedle, match)) {
            enhanced := match1
            delayed := match2

            relativeWashout := ((enhanced - delayed) / enhanced) * 100

            result := input . "`n`n"
            result .= "Relative Washout: " . Round(relativeWashout, 1) . "%`n`n"
            result .= InterpretAdrenalWashout(0, relativeWashout, unenhanced, enhanced, delayed)
			
			if(ShowCitations=1){
				result .= "`n`n`nMayo-Smith WW, Song JH, Boland GL, Francis IR, Israel GM, Mazzaglia PJ, Berland LL, Pandharipande PV. Management of Incidental Adrenal Masses: A White Paper of the ACR Incidental Findings Committee. J Am Coll Radiol. 2017 Aug;14(8):1038-1044.`n`n"
			}
            return result
        } else {
            return "Invalid input format for adrenal washout calculation.`nSample syntax: Unenhanced: 10 HU, Enhanced: 80 HU, Delayed: 40 HU"
        }
    }
}

InterpretAdrenalWashout(absoluteWashout, relativeWashout, unenhanced, enhanced, delayed) {
    result := ""

    ; Check if unenhanced is null or empty
    if (unenhanced = "" or unenhanced = "NULL") {
        unenhanced := "N/A"
        isUnenhancedAvailable := false
    } else {
        isUnenhancedAvailable := true
        unenhanced += 0  ; Convert to number
    }

    ; Check for enhancing mass
    if (isUnenhancedAvailable) {
        enhancedChange := CheckEnhancement(enhanced, unenhanced)
        delayedChange := CheckEnhancement(delayed, unenhanced)
		
         ; Add interpretation based on unenhanced HU
        if (unenhanced <= 10) {
            result .= "Unenhanced HU less than 10 is typically suggestive of a benign adrenal adenoma. "
        } else if (unenhanced > 43) {
            result .= "Unenhanced HU >43 in a noncalcified, nonhemorrhagic lesion is suspicious for malignancy, regardless of washout characteristics. "
        }
		
        if (Abs(enhancedChange) >= 10 or Abs(delayedChange) >= 10) {
            if (enhancedChange < 0 or delayedChange < 0) {
                result .= "The adrenal mass demonstrates unexpected de-enhancement (decrease in HU in enhanced or delayed phase). This is an atypical finding. "
                result .= "Caution: Standard washout calculations may not be applicable in this case. "
            } else {
                result .= "The adrenal mass demonstrates enhancement. "
                
                ; Interpret washout only if there's positive enhancement
                if (absoluteWashout >= 60 or relativeWashout >= 40) {
                    result .= "Adrenal washout characteristics are suggestive of a benign adrenal adenoma. "
                } else {
                    result .= "Washout characteristics are indeterminate. "
                }
            }
        } else {
            result .= "The adrenal mass does not demonstrate significant enhancement (<10 HU change in both enhanced and delayed phases compared to unenhanced). This may represent a cyst, hemorrhage, or other non-enhancing lesion. Further characterization with additional imaging may be necessary. "
        }

    } else {
        result .= "Unenhanced HU value is not available. "
        
        if (absoluteWashout >= 60 or relativeWashout >= 40) {
            result .= "Based on the provided washout values alone, characteristics are suggestive of a benign adrenal adenoma. However, this interpretation is limited without the unenhanced HU value. "
        } else {
			result .= "Washout characteristics are indeterminate. "
		}
    }

    return Trim(result)
}

; Function to check enhancement
CheckEnhancement(value, baseline) {
    if (value = "" or value = "NULL" or baseline = "N/A")
        return 0
    return (value - baseline)
}

CalculateThymusChemicalShift:
    Result := CalculateThymusChemicalShift(SelectedText)
    ShowResult(Result)
return

CalculateThymusChemicalShift(input) {
    RegExNeedle := "i)thymus.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+).*?paraspinous.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)"

    if (RegExMatch(input, RegExNeedle, match)) {
        thymusInPhase := match2
        thymusOutOfPhase := match4
        paraspinousInPhase := match6
        paraspinousOutOfPhase := match8

        outOfPhaseSignalIntensityIndex := (thymusOutOfPhase) / paraspinousOutOfPhase
        inPhaseSignalIntensityIndex := (thymusInPhase) / paraspinousInPhase
        chemicalShiftRatio := outOfPhaseSignalIntensityIndex / inPhaseSignalIntensityIndex

        result := input
        if (DisplayAllValues) {
            result .= "`nOut of Phase Intensity Ratio: " . Round(outOfPhaseSignalIntensityIndex*100, 1) . "%`n"
            result .= "In Phase Intensity Ratio: " . Round(inPhaseSignalIntensityIndex*100, 1)			. "%`n"
            result .= "Chemical Shift Ratio: " . Round(chemicalShiftRatio, 2) . "`n`n"		
        }
        result .= "`n" . InterpretThymusChemicalShift(chemicalShiftRatio)
        return result
    } else {
        return "Invalid input format for thymus chemical shift calculation.`nSample syntax: Thymus IP: 100, OP: 80, Paraspinous IP: 90, OP: 85"
    }
}

InterpretThymusChemicalShift(chemicalShiftRatio) {
	global ShowCitations

    if (chemicalShiftRatio <= 0.849) {
		result := "Interpretation: Chemical Shift Ratio is less than 0.849. Findings are most consistent with thymic hyperplasia.`n`n`n"
		if (ShowCitations = 1) {
			result .= "Citation: Priola AM, Priola SM, Ciccone G, Evangelista A, Cataldi A, Gned D, Pazè F, Ducco L, Moretti F, Brundu M, Veltri A. Differentiation of rebound and lymphoid thymic hyperplasia from anterior mediastinal tumors with dual-echo chemical-shift MR imaging in adulthood: reliability of the chemical-shift ratio and signal intensity index. Radiology. 2015 Jan;274(1):238-49. doi: 10.1148/radiol.14132665. Epub 2014 Aug 7. PMID: 25105246.`n" 
		}
    } else {
        result := "Interpretation: Chemical Shift Ratio is greater than 0.849. Findings are not consistent with thymic hyperplasia and can be seen with malignancy."
    }
	return result
}


CalculateHepaticSteatosis:
    Result := CalculateHepaticSteatosis(SelectedText)
    ShowResult(Result)
return

CalculateHepaticSteatosis(inputText) {
    global ShowCitations
    RegExNeedleLocal := "i)liver.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)"
    RegExNeedleSpleenLocal := "i)spleen.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)"
    
    if (RegExMatch(inputText, RegExNeedleLocal, matchLocal)) {
        liverInPhaseLocal := matchLocal2
        liverOutOfPhaseLocal := matchLocal4
        
        ; Calculate Fat fraction
        fatFractionLocal := 100 * (liverInPhaseLocal - liverOutOfPhaseLocal) / (2 * liverInPhaseLocal)
        
        resultLocal := inputText . " (Fat Fraction: " . Round(fatFractionLocal, 1) . "%)"
        resultLocal .= "`n`nFat Fraction " . InterpretHepaticSteatosis(fatFractionLocal) . "`n"
        
        ; Check if spleen values are provided
        if (RegExMatch(inputText, RegExNeedleSpleenLocal, matchSpleenLocal)) {
            spleenInPhaseLocal := matchSpleenLocal2
            spleenOutOfPhaseLocal := matchSpleenLocal4
            
            ; Calculate Fat percentage
            fatPercentageLocal := 100 * ((liverInPhaseLocal / spleenInPhaseLocal) - (liverOutOfPhaseLocal / spleenOutOfPhaseLocal)) / (2 * (liverInPhaseLocal / spleenInPhaseLocal))
            
            resultLocal := StrReplace(resultLocal, ")", ", Fat Percentage: " . Round(fatPercentageLocal, 1) . "%)")
            resultLocal .= "Fat Percentage " . InterpretHepaticSteatosis(fatPercentageLocal) . "`n"
        }
        
        if (ShowCitations = 1) {
            resultLocal .= "`n`nSirlin CB. Invited Commentary on Image-based quantification of hepatic fat: methods and clinical applications. Radiographics 2009; 29:1277-80`n"
        }
        
        return resultLocal
    } else {
        return "Invalid input format for hepatic steatosis calculation.`nSample syntax: Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88"
    }
}

InterpretHepaticSteatosis(hepaticFatFraction) {
    if (hepaticFatFraction < 5) {
        return "Interpretation: No significant hepatic steatosis."
    } else if (hepaticFatFraction < 15) {
        return "Interpretation: Mild hepatic steatosis."
    } else if (hepaticFatFraction < 30) {
        return "Interpretation: Moderate hepatic steatosis."
    } else {
        return "Interpretation: Severe hepatic steatosis."
    }
}

CompareNoduleSizes:
    Result := CompareNoduleSizes(SelectedText)
    ShowResult(Result)
return

CompareNoduleSizes(input) {
    RegExNeedle := "i)(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?.*?(?:previous(?:ly)?|prior|before|old|initial).*?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?"
    if (!RegExMatch(input, RegExNeedle, match)) {
        RegExNeedle := "i)(?:previous(?:ly)?|prior|before|old|initial).*?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?.*?(?:now|current(?:ly)?|present|new|recent(?:ly)?|follow[- ]?up).*?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?"
        if (!RegExMatch(input, RegExNeedle, match)) {
            return "Invalid input format. Please provide both current and previous measurements."
        }
        current := match3 . " " . match4
        previous := match1 . " " . match2
    } else {
        current := match1 . " " . match2
        previous := match3 . " " . match4
    }
    return CompareMeasurements(previous, current, input)
}

CompareMeasurements(previous, current, input) {
    prev := ProcessMeasurement(previous)
    curr := ProcessMeasurement(current)

    if (prev.dimensions.Length() != curr.dimensions.Length()) {
        return "Error: Mismatch in number of dimensions between previous and current measurements.`nPrevious: " . previous . "`nCurrent: " . current
    }

    result := input . "`n`n"
    prevLongestDim := 0
    currLongestDim := 0

    Loop % prev.dimensions.Length() {
        prevDim := prev.dimensions[A_Index]
        currDim := curr.dimensions[A_Index]
        
        ; Round to 1 decimal place for display if cm, keep full precision for mm
        prevDisplayDim := (prev.unit == "cm") ? Round(prevDim, 1) : prevDim
        currDisplayDim := (curr.unit == "cm") ? Round(currDim, 1) : currDim
        
        result .= "Dimension " . A_Index . ": "
            . prevDisplayDim . " " . prev.unit . " -> "
            . currDisplayDim . " " . curr.unit . " "

        prevDimCm := (prev.unit == "mm") ? prevDim / 10 : prevDim
        currDimCm := (curr.unit == "mm") ? currDim / 10 : currDim

        prevLongestDim := (prevDimCm > prevLongestDim) ? prevDimCm : prevLongestDim
        currLongestDim := (currDimCm > currLongestDim) ? currDimCm : currLongestDim

        change := (currDimCm / prevDimCm * 100) - 100
        result .= "(" . (change >= 0 ? "+" : "") . Round(change, 1) . "%)`n"
    }

    longestDimChange := (currLongestDim / prevLongestDim * 100) - 100
    result .= "`nLongest dimension change: " . (longestDimChange >= 0 ? "+" : "") . Round(longestDimChange, 1) . "%"

    if (prev.dimensions.Length() == 3) {
        prevVolumeNum := CalculateEllipsoidVolumeNumeric(JoinDimensions(prev.dimensions), prev.unit)
        currVolumeNum := CalculateEllipsoidVolumeNumeric(JoinDimensions(curr.dimensions), curr.unit)
        if (prevVolumeNum != "Invalid input format" && currVolumeNum != "Invalid input format") {
            volumeChange := (currVolumeNum / prevVolumeNum * 100) - 100
            result .= "`nVolume change: " . (volumeChange >= 0 ? "+" : "") . Round(volumeChange, 1) . "%"
            result .= "`nPrevious volume: " . FormatVolume(prevVolumeNum)
            result .= "`nCurrent volume: " . FormatVolume(currVolumeNum)
        }
    }

    return result
}


FormatVolume(volume) {
    return (volume < 1) ? Round(volume, 3) . " cc" : Round(volume, 1) . " cc"
}

CalculateEllipsoidVolumeNumeric(input, unit) {
    RegExNeedle := "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
    if (RegExMatch(input, RegExNeedle, match)) {
        dimensions := [match1 + 0, match2 + 0, match3 + 0]  ; Convert to numbers
        dimensions := SortDimensions(dimensions)
        if (unit == "mm") {
            dimensions[1] := dimensions[1] / 10
            dimensions[2] := dimensions[2] / 10
            dimensions[3] := dimensions[3] / 10
        }
        volume := (1/6) * 3.14159265358979323846 * (dimensions[1]) * (dimensions[2]) * (dimensions[3])
        return volume
    } else {
        return "Invalid input format"
    }
}

ProcessMeasurement(input) {
    dimensions := []
    RegExMatch(input, "i)(\d+(?:\.\d+)?)(?:\s*(?:x|\*)\s*(\d+(?:\.\d+)?))?(?:\s*(?:x|\*)\s*(\d+(?:\.\d+)?))?(?=\s*(cm|mm)?)", match)
    dimensions.Push(match1 + 0)  ; Convert to number to preserve all decimal places
    if (match2 != "")
        dimensions.Push(match2 + 0)
    if (match3 != "")
        dimensions.Push(match3 + 0)
    unit := (match4 != "") ? match4 : ((InStr(match1, ".") > 0 || InStr(match2, ".") > 0 || InStr(match3, ".") > 0) ? "cm" : "mm")
    return {dimensions: dimensions, unit: unit}
}


JoinDimensions(dimensions) {
    return Round(dimensions[1], 1) . (dimensions.Length() > 1 ? " x " . Round(dimensions[2], 1) : "") . (dimensions.Length() > 2 ? " x " . Round(dimensions[3], 1) : "")
}

global PauseDurationChoice

ShowPreferences() {
    global DisplayUnits, DisplayAllValues, ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity, ShowPregnancyDates, ShowMenstrualPhase, PauseDuration
    global ShowAdrenalWashout, ShowThymusChemicalShift, ShowHepaticSteatosis, ShowMRILiverIron, ShowStatistics, ShowNumberRange, DarkMode
	global ShowCitations, ShowArterialAge
	
    if (PauseDuration = 180000)
        currentPauseDuration := "3 minutes"
    else if (PauseDuration = 600000)
        currentPauseDuration := "10 minutes"
    else if (PauseDuration = 1800000)
        currentPauseDuration := "30 minutes"
    else if (PauseDuration = 3600000)
        currentPauseDuration := "1 hour"
    else if (PauseDuration = 36000000)
        currentPauseDuration := "10 hours"
    else
        currentPauseDuration := PauseDuration . " ms"

    Gui, Preferences:New, +LastFound, Preferences
    preferencesHwnd := WinExist()
    
    ;Gui, Add, Text, x10 y10 w200, Display Options:
    ;Gui, Add, Checkbox, x10 y30 w200 vDisplayUnits Checked%DisplayUnits%, Display units with values
    ;Gui, Add, Checkbox, x10 y60 w200 vDisplayAllValues Checked%DisplayAllValues%, Display all calculated values
    ;Gui, Add, Checkbox, x10 y90 w200 vDarkMode Checked%DarkMode% gToggleDarkMode, Dark Mode
    Gui, Add, Text, x10 y10 w200, Select functions to display:
    Gui, Add, Checkbox, x10 y30 w200 vShowEllipsoidVolume Checked%ShowEllipsoidVolume%, Ellipsoid Volume
    Gui, Add, Checkbox, x10 y60 w200 vShowBulletVolume Checked%ShowBulletVolume%, Bullet Volume
    Gui, Add, Checkbox, x10 y90 w200 vShowPSADensity Checked%ShowPSADensity%, PSA Density
    Gui, Add, Checkbox, x10 y120 w200 vShowPregnancyDates Checked%ShowPregnancyDates%, Pregnancy Dates
    Gui, Add, Checkbox, x10 y150 w200 vShowMenstrualPhase Checked%ShowMenstrualPhase%, Menstrual Phase
    Gui, Add, Checkbox, x10 y180 w200 vShowAdrenalWashout Checked%ShowAdrenalWashout%, Adrenal Washout
    Gui, Add, Checkbox, x10 y210 w200 vShowThymusChemicalShift Checked%ShowThymusChemicalShift%, Thymus Chemical Shift
    Gui, Add, Checkbox, x10 y240 w200 vShowHepaticSteatosis Checked%ShowHepaticSteatosis%, Hepatic Steatosis
    Gui, Add, Checkbox, x10 y270 w200 vShowMRILiverIron Checked%ShowMRILiverIron%, MRI Liver Iron Content
    Gui, Add, Checkbox, x10 y300 w200 vShowStatistics Checked%ShowStatistics%, Calculate Statistics
    Gui, Add, Checkbox, x10 y330 w200 vShowNumberRange Checked%ShowNumberRange%, Calculate Number Range
	Gui, Add, Checkbox, x10 y360 w200 vShowCalciumScorePercentile Checked%ShowCalciumScorePercentile%, Calculate Calcium Score Percentile
    Gui, Add, Checkbox, x10 y390 w200 vShowCitations Checked%ShowCitations%, Show Citations in Output
    Gui, Add, Checkbox, x10 y420 w200 vShowArterialAge Checked%ShowArterialAge%, Show Arterial Age
	Gui, Add, Text, x10 y450 w200, Pause Duration (current: %currentPauseDuration%):
    Gui, Add, DropDownList, x10 y480 w200 vPauseDurationChoice, 3 minutes|10 minutes|30 minutes|1 hour|10 hours
    if (PauseDuration = 180000)
        GuiControl, Choose, PauseDurationChoice, 1
    else if (PauseDuration = 600000)
        GuiControl, Choose, PauseDurationChoice, 2
    else if (PauseDuration = 1800000)
        GuiControl, Choose, PauseDurationChoice, 3
    else if (PauseDuration = 3600000)
        GuiControl, Choose, PauseDurationChoice, 4
    else if (PauseDuration = 36000000)
        GuiControl, Choose, PauseDurationChoice, 5
    Gui, Add, Button, x60 y520 w100 gSavePreferences, Save
    
    ApplyDarkMode(preferencesHwnd, DarkMode)
    Gui, Show, w220 h600
}

ToggleDarkMode:
    Gui, Submit, NoHide
    ApplyDarkMode(preferencesHwnd, DarkMode)
return

ApplyDarkMode(hwnd, isDarkMode) {
    static DWMWA_USE_IMMERSIVE_DARK_MODE := 20
    if (isDarkMode) {
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "Int", DWMWA_USE_IMMERSIVE_DARK_MODE, "Int*", true, "Int", 4)
        Gui, Color, 0x202020, 0x202020
        GuiControl, +Background0x202020 +c0xFFFFFF, DisplayUnits
        GuiControl, +Background0x202020 +c0xFFFFFF, DisplayAllValues
        GuiControl, +Background0x202020 +c0xFFFFFF, DarkMode
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowEllipsoidVolume
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowBulletVolume
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowPSADensity
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowPregnancyDates
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowMenstrualPhase
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowAdrenalWashout
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowThymusChemicalShift
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowHepaticSteatosis
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowMRILiverIron
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowStatistics
        GuiControl, +Background0x202020 +c0xFFFFFF, ShowNumberRange
    } else {
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "Int", DWMWA_USE_IMMERSIVE_DARK_MODE, "Int*", false, "Int", 4)
        Gui, Color, Default
        GuiControl, +BackgroundDefault +cDefault, DisplayUnits
        GuiControl, +BackgroundDefault +cDefault, DisplayAllValues
        GuiControl, +BackgroundDefault +cDefault, DarkMode
        GuiControl, +BackgroundDefault +cDefault, ShowEllipsoidVolume
        GuiControl, +BackgroundDefault +cDefault, ShowBulletVolume
        GuiControl, +BackgroundDefault +cDefault, ShowPSADensity
        GuiControl, +BackgroundDefault +cDefault, ShowPregnancyDates
        GuiControl, +BackgroundDefault +cDefault, ShowMenstrualPhase
        GuiControl, +BackgroundDefault +cDefault, ShowAdrenalWashout
        GuiControl, +BackgroundDefault +cDefault, ShowThymusChemicalShift
        GuiControl, +BackgroundDefault +cDefault, ShowHepaticSteatosis
        GuiControl, +BackgroundDefault +cDefault, ShowMRILiverIron
        GuiControl, +BackgroundDefault +cDefault, ShowStatistics
        GuiControl, +BackgroundDefault +cDefault, ShowNumberRange
    }
}

SavePreferences:
    Gui, Submit, NoHide
    global DisplayUnits, DisplayAllValues
    global ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity, ShowPregnancyDates, ShowMenstrualPhase
    global ShowAdrenalWashout, ShowThymusChemicalShift, ShowHepaticSteatosis
    global ShowMRILiverIron, ShowStatistics, ShowNumberRange
    global PauseDuration, DarkMode
	global ShowArterialAge, ShowCitations

    if (PauseDurationChoice = "3 minutes")
        PauseDuration := 180000
    else if (PauseDurationChoice = "10 minutes")
        PauseDuration := 600000
    else if (PauseDurationChoice = "30 minutes")
        PauseDuration := 1800000
    else if (PauseDurationChoice = "1 hour")
        PauseDuration := 3600000
    else if (PauseDurationChoice = "10 hours")
        PauseDuration := 36000000

    SavePreferencesToFile()
    Gui, Destroy
return

SavePreferencesToFile() {
    global DisplayUnits, DisplayAllValues, ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity, ShowPregnancyDates, ShowMenstrualPhase, PauseDuration
    global ShowAdrenalWashout, ShowThymusChemicalShift, ShowHepaticSteatosis, ShowMRILiverIron, ShowStatistics, ShowNumberRange, DarkMode
	global ShowCitations, ShowArterialAge

    IniWrite, %DisplayUnits%, %A_ScriptDir%\preferences.ini, Display, DisplayUnits
    IniWrite, %DisplayAllValues%, %A_ScriptDir%\preferences.ini, Display, DisplayAllValues
    IniWrite, %ShowEllipsoidVolume%, %A_ScriptDir%\preferences.ini, Calculations, ShowEllipsoidVolume
    IniWrite, %ShowBulletVolume%, %A_ScriptDir%\preferences.ini, Calculations, ShowBulletVolume
    IniWrite, %ShowPSADensity%, %A_ScriptDir%\preferences.ini, Calculations, ShowPSADensity
    IniWrite, %ShowPregnancyDates%, %A_ScriptDir%\preferences.ini, Calculations, ShowPregnancyDates
    IniWrite, %ShowMenstrualPhase%, %A_ScriptDir%\preferences.ini, Calculations, ShowMenstrualPhase
    IniWrite, %ShowAdrenalWashout%, %A_ScriptDir%\preferences.ini, Calculations, ShowAdrenalWashout
    IniWrite, %ShowThymusChemicalShift%, %A_ScriptDir%\preferences.ini, Calculations, ShowThymusChemicalShift
    IniWrite, %ShowHepaticSteatosis%, %A_ScriptDir%\preferences.ini, Calculations, ShowHepaticSteatosis
    IniWrite, %ShowMRILiverIron%, %A_ScriptDir%\preferences.ini, Calculations, ShowMRILiverIron
    IniWrite, %ShowStatistics%, %A_ScriptDir%\preferences.ini, Calculations, ShowStatistics
    IniWrite, %ShowNumberRange%, %A_ScriptDir%\preferences.ini, Calculations, ShowNumberRange
    IniWrite, %PauseDuration%, %A_ScriptDir%\preferences.ini, Script, PauseDuration
	IniWrite, %ShowCitations%, %A_ScriptDir%\preferences.ini, Display, ShowCitations
	IniWrite, %ShowArterialAge%, %A_ScriptDir%\preferences.ini, Display, ShowArterialAge
    IniWrite, %DarkMode%, %A_ScriptDir%\preferences.ini, Display, DarkMode
	IniWrite, %ShowCalciumScorePercentile%, %A_ScriptDir%\preferences.ini, Calculations, ShowCalciumScorePercentile
}

PreferencesGuiClose:
PreferencesGuiEscape:
    Gui, Destroy
return

SortDimensions(dimensions) {
    if (dimensions[1] < dimensions[2]) {
        temp := dimensions[1]
        dimensions[1] := dimensions[2]
        dimensions[2] := temp
    }
    if (dimensions[2] < dimensions[3]) {
        temp := dimensions[2]
        dimensions[2] := dimensions[3]
        dimensions[3] := temp
    }
    if (dimensions[1] < dimensions[2]) {
        temp := dimensions[1]
        dimensions[1] := dimensions[2]
        dimensions[2] := temp
    }
    return dimensions
}

ParseDate(dateStr) {
    dateStr := StrReplace(dateStr, ".", "/")
    dateStr := StrReplace(dateStr, "-", "/")
    if (RegExMatch(dateStr, "(\d{1,2})/(\d{1,2})/(\d{2,4})", match)) {
        month := match1
        day := match2
        year := match3
        if (StrLen(year) == 2)
            year := "20" . year
        if (month > 12) {
            temp := month
            month := day
            day := temp
        }
        if (day > 31 || month > 12)
            return "Invalid Date"
        month := SubStr("0" . month, -1)
        day := SubStr("0" . day, -1)
        return year . month . day
    }
    return "Invalid Date"
}

CalculateDatesFromLMP(LMPDate) {
    FormatTime, LMPFormatted, %LMPDate%, MM/dd/yyyy
    EDDDate := DateCalc(LMPDate, 280)
    FormatTime, EDDFormatted, %EDDDate%, MM/dd/yyyy
    GA := DateCalc(A_Now, 0)
    GA -= LMPDate, days
    GAWeeks := Floor(GA / 7)
    GADays := Mod(GA, 7)
    return % "LMP: " . LMPFormatted . "`n"
        . "Estimated Delivery Date: " . EDDFormatted . "`n"
        . "Current Gestational Age: " . GAWeeks . " weeks " . GADays . " days"
}

CalculateDatesFromGA(WeeksGA, DaysGA, ReferenceDate) {
    FetusAgeDays := (WeeksGA * 7) + DaysGA
    LMPDate := DateCalc(ReferenceDate, -FetusAgeDays)
    FormatTime, LMPFormatted, %LMPDate%, MM/dd/yyyy
    EDDDate := DateCalc(LMPDate, 280)
    FormatTime, EDDFormatted, %EDDDate%, MM/dd/yyyy
    CurrentGA := DateCalc(A_Now, 0)
    CurrentGA -= LMPDate, days
    CurrentGAWeeks := Floor(CurrentGA / 7)
    CurrentGADays := Mod(CurrentGA, 7)
    FormatTime, ReferenceDateFormatted, %ReferenceDate%, MM/dd/yyyy
    return % "LMP: " . LMPFormatted . "`n"
        . "Estimated Delivery Date: " . EDDFormatted . "`n"
        . "Gestational Age as of " . ReferenceDateFormatted . ": " . WeeksGA . " weeks " . DaysGA . " days`n"
        . "Current Gestational Age: " . CurrentGAWeeks . " weeks " . CurrentGADays . " days"
}

DetermineMenstrualPhase(LMPDate) {
    DaysSinceLMP := A_Now
    DaysSinceLMP -= LMPDate, days
    CycleDay := Mod(DaysSinceLMP, 28) + 1
    FormatTime, LMPFormatted, %LMPDate%, MM/dd/yyyy
    Result := "LMP: " . LMPFormatted . "`n"
    Result .= "Current Cycle Day: " . CycleDay . "/28`n`n"
    if (CycleDay >= 1 && CycleDay <= 5) {
        Result .= "Menstrual Phase`nExpected endometrial stripe thickness: 1-4 mm"
    } else if (CycleDay >= 6 && CycleDay <= 13) {
        Result .= "Early Proliferative Phase`nExpected endometrial stripe thickness: 5-7 mm"
    } else if (CycleDay == 14) {
        Result .= "Ovulation`nExpected endometrial appearance: Trilaminar, approximately 11 mm"
    } else if (CycleDay >= 15 && CycleDay <= 28) {
        Result .= "Secretory Phase`nExpected endometrial stripe thickness: 7-16 mm"
    } else {
        Result .= "Error: Invalid cycle day calculated"
    }
    return Result
}

CalculateCalciumScorePercentile(input) {
    ; Extract age, sex, race, and calcium score from input
    RegExMatch(input, "i)Age:\s*(\d+)", age)
    RegExMatch(input, "i)Sex:\s*(Male|Female)", sex)
    RegExMatch(input, "i)Race:\s*(White|Black|Hispanic|Chinese|[A-Za-z]+)", race)
    RegExMatch(input, "i)(?:your\s+)?(?:coronary\s+artery\s+)?calcium\s+score(?:\s+is)?:?\s*(\d+(?:\.\d+)?)\s*(?:\(?\s*Agatston\s*\)?)?|(?:total\s+)?calcium\s+score:?\s*(\d+(?:\.\d+)?)", score)
	
	if (!age1) {
        return input . "`n`nError: Age not found or invalid. Please provide age in the format 'Age: 55'."
    }
    if (!sex1) {
        return input . "`n`nError: Sex not found or invalid. Please specify either Male or Female."
    }
    if (!score1 && score1 !=0 ) {
        return input . "`n`nError: Calcium score not found or invalid. Please ensure the score is provided in the format 'YOUR CORONARY ARTERY CALCIUM SCORE: 0.0 (Agatston)'."
    }

    age := age1
    sex := sex1
    race := race1 ? race1 : "Unspecified"
    score := score1

    result := ""
    if (age >= 45 && age <= 84 && (race = "White" || race = "Black" || race = "Hispanic" || race = "Chinese")) {
        result := CalculateMESAScore(age, race, sex, score)
    } else if (age >= 30 && age <= 74) {
        result := CalculateHoffScore(age, sex, score, race)
    } else if (age > 84) {
        result := "Warning: Age is above 84. The calculator may not provide accurate results. Please consult with a healthcare professional for interpretation.`n`n" . CalculateHoffScore(age, sex, score, race)
    } else {
        result := "Error: The calcium score calculators are only valid for ages 30 and above."
    }

    return input . "`n`nCONTEXT:`n" . result
}

CalculateMESAScore(age, race, sex, score) {
	global ShowCitations
	global ShowArterialAge

    url := "https://www.mesa-nhlbi.org/Calcium/input.aspx"
    
    ; First, we need to get the initial page to retrieve some hidden values
    initialResponse := DownloadToString(url)
    if (InStr(initialResponse, "Error:")) {
        return "Initial page load failed: " . initialResponse
    }
    
    ; Extract necessary hidden values
    RegExMatch(initialResponse, "id=""__VIEWSTATE"" value=""([^""]+)""", viewState)
    RegExMatch(initialResponse, "id=""__VIEWSTATEGENERATOR"" value=""([^""]+)""", viewStateGenerator)
    RegExMatch(initialResponse, "id=""__EVENTVALIDATION"" value=""([^""]+)""", eventValidation)
    
    ; Prepare the post data
    postData := "__VIEWSTATE=" . UrlEncode(viewState1)
              . "&__VIEWSTATEGENERATOR=" . viewStateGenerator1
              . "&__EVENTVALIDATION=" . UrlEncode(eventValidation1)
              . "&Age=" . age
              . "&gender=" . GetSexValue(sex)
              . "&Race=" . GetRaceValue(race)
              . "&Score=" . score
              . "&Calculate=Calculate"

    ; Send the calculation request
    response := DownloadToString(url, postData)
    if (InStr(response, "Error:")) {
        return "Calculation request failed: " . response
    }
    
    ; Extract the results
    RegExMatch(response, "id=""Label10""[^>]*>([^<]+)</span>", probability)
    RegExMatch(response, "id=""scoreLabel""[^>]*>([^<]+)</span>", observedScore)
    RegExMatch(response, "id=""percLabel""[^>]*>([^<]+)</span>", percentile)
    
    if (probability1 || observedScore1 || percentile1) {
        result := "Probability of non-zero calcium score: " . probability1 . "`n`n"
        
        result .= "Plaque Burden: " . DeterminePlaqueBurden(score) . "`n`n"
        
        percentileNum := percentile1 + 0
        comparison := DetermineComparison(percentileNum)
        
        result .= "Comparison to people of the same age, race and sex: " . comparison . "`n`n"
        result .= "Observed calcium score of " . score . " Agatston is at percentile " . percentile1 . " for age, race and sex`n`n"
        
        if(ShowArterialAge=1){
			result .= "Arterial Age: " . CalculateCoronaryAge(score) . " years" . "`n`n"
        }
		if(ShowCitations=1) {
			result .= "Citation 1: McClelland RL, et al. Distribution of coronary artery calcium by race, gender, and age: results from the Multi-Ethnic Study of Atherosclerosis (MESA). Circulation. 2006 Jan 3;113(1):30-7.`n`n"
			result .= "Citation 2: McClelland RL, Nasir K, Budoff M, Blumenthal RS, Kronmal RA. Arterial age as a function of coronary artery calcium (from the Multi-Ethnic Study of Atherosclerosis [MESA]). Am J Cardiol. 2009 Jan 1;103(1):59-63."
        }
		
		return result
		
	} else {
		; Save the response for debugging
		FileAppend, %response%, %A_Desktop%\mesa_debug_response.html
		return "Error: Unable to retrieve result from MESA calculator. The response has been saved to your desktop as 'mesa_debug_response.html' for further investigation."
	}
}

CalculateHoffScore(age, sex, score, race) {
	global ShowCitations
	global ShowArterialAge

    ageGroup := GetAgeGroup(age)
    percentile := GetHoffPercentile(ageGroup, sex, score)
    
    result := "Plaque Burden: " . DeterminePlaqueBurden(score) . "`n`n"
    
    comparison := DetermineComparison(percentile)
    result .= "Comparison to people of the same age and sex: " . comparison . "`n`n"
    
    if(ShowArterialAge=1){
		result .= "Arterial Age: " . CalculateCoronaryAge(score) . " years" . "`n`n"
    }
    if(ShowCitations=1){
		result .= "Citation 1: Hoff JA, et al. Age and gender distributions of coronary artery calcium detected by electron beam tomography in 35,246 adults. Am J Cardiol. 2001 Jun 15;87(12):1335-9.`n`n"
		result .= "Citation 2: McClelland RL, Nasir K, Budoff M, Blumenthal RS, Kronmal RA. Arterial age as a function of coronary artery calcium (from the Multi-Ethnic Study of Atherosclerosis [MESA]). Am J Cardiol. 2009 Jan 1;103(1):59-63."
    }
	
    return result
}

GetAgeGroup(age) {
    if (age < 40)
        return "<40"
    else if (age < 45)
        return "40-44"
    else if (age < 50)
        return "45-49"
    else if (age < 55)
        return "50-54"
    else if (age < 60)
        return "55-59"
    else if (age < 65)
        return "60-64"
    else if (age < 70)
        return "65-69"
    else if (age <= 74)
        return "70-74"
    else
        return ">74"
}

GetHoffPercentile(ageGroup, sex, score) {
    percentiles := {"<40": {}, "40-44": {}, "45-49": {}, "50-54": {}, "55-59": {}, "60-64": {}, "65-69": {}, "70-74": {}, ">74": {}}
    
    ; Male percentiles
    percentiles["<40"]["Male"] := [0, 1, 3, 14]
    percentiles["40-44"]["Male"] := [0, 1, 9, 59]
    percentiles["45-49"]["Male"] := [0, 3, 36, 154]
    percentiles["50-54"]["Male"] := [1, 15, 103, 332]
    percentiles["55-59"]["Male"] := [4, 48, 215, 554]
    percentiles["60-64"]["Male"] := [13, 113, 410, 994]
    percentiles["65-69"]["Male"] := [32, 180, 566, 1299]
    percentiles["70-74"]["Male"] := [64, 310, 892, 1774]
    percentiles[">74"]["Male"] := [166, 473, 1071, 1982]
    
    ; Female percentiles
    percentiles["<40"]["Female"] := [0, 0, 1, 3]
    percentiles["40-44"]["Female"] := [0, 0, 1, 4]
    percentiles["45-49"]["Female"] := [0, 0, 2, 22]
    percentiles["50-54"]["Female"] := [0, 0, 5, 55]
    percentiles["55-59"]["Female"] := [0, 1, 23, 121]
    percentiles["60-64"]["Female"] := [0, 3, 57, 193]
    percentiles["65-69"]["Female"] := [1, 24, 145, 410]
    percentiles["70-74"]["Female"] := [3, 52, 210, 631]
    percentiles[">74"]["Female"] := [9, 75, 241, 709]
    
    agePercentiles := percentiles[ageGroup][sex]
    
    if (score <= agePercentiles[1])
        return 25
    else if (score <= agePercentiles[2])
        return 50
    else if (score <= agePercentiles[3])
        return 75
    else if (score <= agePercentiles[4])
        return 90
    else
        return 99
}

CalculateCoronaryAge(score) {
	logScore := Ln(score + 1)
    effectiveAge := round(39.1 + 7.25 * logScore,0)
    return effectiveAge
}

CalculateCoronaryAge95CI(score) {
	logScore := Ln(score + 1)
	arteryAge95CI := round(sqrt((13.7 - 6.5 * logScore) + 0.8 * (logScore ** 2)),0)
    return arteryAge95CI
}

DeterminePlaqueBurden(score) {
    if (score = 0)
        return "None. Risk of coronary artery disease is very low, generally less than 5 percent."
    else if (score > 0 && score <= 10)
        return "Minimal identifiable plaque. Risk of coronary artery disease is very unlikely, less than 10 percent."
    else if (score > 10 && score <= 100)
        return "At least mild atherosclerotic plaque. Mild or minimal coronary narrowings likely."
    else if (score > 100 && score <= 400)
        return "At least moderate atherosclerotic plaque. Mild coronary artery disease highly likely, significant narrowings possible."
    else
        return "Extensive atherosclerotic plaque. High likelihood of at least one significant coronary narrowing."
}

DetermineComparison(percentile) {
    if (percentile <= 25)
        return "Low; between 0 and 25th percentile"
    else if (percentile > 25 && percentile <= 50)
        return "Average; between 25th and 50th percentile"
    else if (percentile > 50 && percentile <= 75)
        return "Average; between 50th and 75th percentile"
    else if (percentile > 75 && percentile <= 90)
        return "High; between 75th and 90th percentile"
    else
        return "Very high; greater than 90th percentile"
}

DownloadToString(url, postData := "") {
    static INTERNET_FLAG_RELOAD := 0x80000000
    static INTERNET_FLAG_SECURE := 0x00800000
    static SECURITY_FLAG_IGNORE_UNKNOWN_CA := 0x00000100
    
    hModule := DllCall("LoadLibrary", "Str", "wininet.dll", "Ptr")
    if (!hModule)
        return "Error: Failed to load wininet.dll. Error code: " . A_LastError
    
    hInternet := DllCall("wininet\InternetOpenA", "Str", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36", "UInt", 1, "Ptr", 0, "Ptr", 0, "UInt", 0, "Ptr")
    if (!hInternet) {
        DllCall("FreeLibrary", "Ptr", hModule)
        return "Error: InternetOpen failed. Error code: " . A_LastError
    }
    
    hConnect := DllCall("wininet\InternetConnectA", "Ptr", hInternet, "Str", "www.mesa-nhlbi.org", "UShort", 443, "Ptr", 0, "Ptr", 0, "UInt", 3, "UInt", 0, "Ptr", 0, "Ptr")
    if (!hConnect) {
        DllCall("wininet\InternetCloseHandle", "Ptr", hInternet)
        DllCall("FreeLibrary", "Ptr", hModule)
        return "Error: InternetConnect failed. Error code: " . A_LastError
    }
    
    flags := INTERNET_FLAG_RELOAD | INTERNET_FLAG_SECURE | SECURITY_FLAG_IGNORE_UNKNOWN_CA
    hRequest := DllCall("wininet\HttpOpenRequestA", "Ptr", hConnect, "Str", (postData ? "POST" : "GET"), "Str", "/Calcium/input.aspx", "Str", "HTTP/1.1", "Ptr", 0, "Ptr", 0, "UInt", flags, "Ptr", 0, "Ptr")
    if (!hRequest) {
        DllCall("wininet\InternetCloseHandle", "Ptr", hConnect)
        DllCall("wininet\InternetCloseHandle", "Ptr", hInternet)
        DllCall("FreeLibrary", "Ptr", hModule)
        return "Error: HttpOpenRequest failed. Error code: " . A_LastError
    }
    
    headers := "Content-Type: application/x-www-form-urlencoded`r`n"
             . "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36`r`n"
             . "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9`r`n"
             . "Accept-Language: en-US,en;q=0.9`r`n"
    DllCall("wininet\HttpAddRequestHeadersA", "Ptr", hRequest, "Str", headers, "UInt", -1, "UInt", 0x10000000, "Int")
    
    VarSetCapacity(buffer, 8192, 0)
    if (postData) {
        VarSetCapacity(postDataBuffer, StrLen(postData), 0)
        StrPut(postData, &postDataBuffer, "UTF-8")
        result := DllCall("wininet\HttpSendRequestA", "Ptr", hRequest, "Ptr", 0, "UInt", 0, "Ptr", &postDataBuffer, "UInt", StrLen(postData), "Int")
    } else {
        result := DllCall("wininet\HttpSendRequestA", "Ptr", hRequest, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt", 0, "Int")
    }
    
    if (!result) {
        errorCode := A_LastError
        DllCall("wininet\InternetCloseHandle", "Ptr", hRequest)
        DllCall("wininet\InternetCloseHandle", "Ptr", hConnect)
        DllCall("wininet\InternetCloseHandle", "Ptr", hInternet)
        DllCall("FreeLibrary", "Ptr", hModule)
        return "Error: HttpSendRequest failed. Error code: " . errorCode
    }
    
    VarSetCapacity(responseText, 1024*1024)  ; Allocate 1MB for the response
    bytesRead := 0
    totalBytesRead := 0
    
    Loop {
        result := DllCall("wininet\InternetReadFile", "Ptr", hRequest, "Ptr", &buffer, "UInt", 8192, "Ptr", &bytesRead, "Int")
        bytesRead := NumGet(bytesRead, 0, "UInt")
        if (bytesRead == 0)
            break
        DllCall("RtlMoveMemory", "Ptr", &responseText + totalBytesRead, "Ptr", &buffer, "Ptr", bytesRead)
        totalBytesRead += bytesRead
    }
    
    responseText := StrGet(&responseText, totalBytesRead, "UTF-8")
    
    DllCall("wininet\InternetCloseHandle", "Ptr", hRequest)
    DllCall("wininet\InternetCloseHandle", "Ptr", hConnect)
    DllCall("wininet\InternetCloseHandle", "Ptr", hInternet)
    DllCall("FreeLibrary", "Ptr", hModule)
    
    return responseText
}


GetRaceValue(race) {
    switch race {
        case "White": return 3
        case "Black": return 0
        case "Hispanic": return 2
        case "Chinese": return 1
    }
}

GetSexValue(sex) {
    switch sex {
        case "Male": return 1
        case "Female": return 0
    }
}

GetLastErrorMessage(errorCode) {
    VarSetCapacity(msg, 1024)
    DllCall("FormatMessage"
        , "UInt", 0x1000      ; FORMAT_MESSAGE_FROM_SYSTEM
        , "Ptr", 0
        , "UInt", errorCode
        , "UInt", 0           ; Default language
        , "Str", msg
        , "UInt", 1024
        , "Ptr", 0)
    return msg
}

SendHttpRequest(url, postData := "") {
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open(postData ? "POST" : "GET", url, true)
        whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        whr.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36")
        whr.Send(postData)
        whr.WaitForResponse()
        return whr.ResponseText
    }
    catch e {
        return "Error: " . e.message
    }
}

UrlEncode(str) {
    oldFormat := A_FormatInteger
    SetFormat, Integer, Hex
    VarSetCapacity(var, StrPut(str, "UTF-8"), 0)
    StrPut(str, &var, "UTF-8")
    StringLower, str, str
    Loop
    {
        code := NumGet(var, A_Index - 1, "UChar")
        If (!code)
            break
        If (code >= 0x30 && code <= 0x39 ; 0-9
            || code >= 0x41 && code <= 0x5A ; A-Z
            || code >= 0x61 && code <= 0x7A) ; a-z
            result .= Chr(code)
        Else
            result .= "%" . SubStr(code + 0x100, -1)
    }
    SetFormat, Integer, %oldFormat%
    return result
}


DateCalc(date, days) {
    date += days, days
    FormatTime, result, %date%, yyyyMMdd
    return result
}

#If
*RButton::
{
    SetTimer, CheckActiveWindowAndShowMenu, -10  ; Set a 10ms delay
    return
}

CheckActiveWindowAndShowMenu:
    if (IsTargetWindowActive()) {
        SelectedText := GetSelectedText()
        CreateCustomMenu()
        if (IsTargetWindowActive()) {  ; Double-check before showing menu
            Menu, CustomMenu, Show
        }
        Menu, CustomMenu, DeleteAll
    } else {
        Send {RButton}  ; Perform default right-click if not in target window
    }
return

IsTargetWindowActive() {
    return WinActive("ahk_class Notepad") 
        or WinActive("ahk_exe notepad.exe") 
        or WinActive("ahk_class PowerScribe") 
        or WinActive("ahk_exe PowerScribe.exe") 
        or WinActive("ahk_class PowerScribe360") 
        or WinActive("ahk_exe Nuance.PowerScribe360.exe") 
        or WinActive("ahk_class PowerScribe | Reporting")
}

^!p::ShowPreferences()

PauseScript() {
    global PauseDuration
    if (PauseDuration < 3600000) {
        pauseMinutes := Floor(PauseDuration / 60000)
        pauseDisplay := pauseMinutes . " minute" . (pauseMinutes != 1 ? "s" : "")
    } else {
        pauseHours := Floor(PauseDuration / 3600000)
        pauseDisplay := pauseHours . " hour" . (pauseHours != 1 ? "s" : "")
    }
    Suspend, On
    SetTimer, ResumeScript, %PauseDuration%
    MsgBox, 0, Script Paused, Script paused for %pauseDisplay%. Click OK to resume immediately.
    Suspend, Off
    SetTimer, ResumeScript, Off
}

ResumeScript:
    Suspend, Off
    SetTimer, ResumeScript, Off
    MsgBox, 0, Script Resumed, The script has been automatically resumed.
return