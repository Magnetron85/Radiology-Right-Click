; ==========================================
; Radiologist's Helper Script
; Radiology Right Click
; Version: 1.04
; Description: This AutoHotkey script provides various calculation tools and utilities
;              for radiologists, including volume calculations, date estimations,
;              and statistical analysis of measurements.
; ==========================================

#NoEnv

#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%

;OCR dependency
#include <Vis2>  ; Equivalent to #include .\lib\Vis2.ahk

; Global variables for standard functions
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
global ShowArterialAge :=
global g_SelectedText := ""
global PauseDurationChoice
global TargetApps := ["ahk_class Notepad", "ahk_exe notepad.exe", "ahk_class PowerScribe", "ahk_exe PowerScribe.exe", "ahk_class PowerScribe360", "ahk_exe Nuance.PowerScribe360.exe", "ahk_class PowerScribe | Reporting"]
global ResultText
global InvisibleControl
global originalMouseX, originalMouseY
global ScanDate, ScanTime, PremedProtocol
global ShowFleischnerCriteria := true
global g_Nodules := []
global g_FleischnerNodules := []
global g_ShowFleischnerCitation := false
global g_ShowFleischnerExclusions := false
global recommendations := {}
global ShowNASCETCalculator := true



;======== Global variables for for OCR
global g_arteryNames := {"lm": "Left Main"
                        ,"li": "Left Main"
                        ,"lit": "Left Main"
                        ,"lad": "Left Anterior Descending"
                        ,"cx": "Left Circumflex"
                        ,"rca": "Right Coronary"
                        ,"ca": "Right Coronary"
                        ,"pda": "Posterior Descending Artery"
                        ,"total": "Total"}
global g_levenshteinThreshold := 1.5
global g_ocrAttempts := 3  ; Number of OCR attempts
global CalciumScoreMonitor := 1
global CalciumScoreX := 0
global CalciumScoreY := 0
global CalciumScoreWidth := 0
global CalciumScoreHeight := 0

; Initialize GDI+
If !pToken := Gdip_Startup()
{
   MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
   ExitApp
}
OnExit, Exit

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
		IniRead, ShowContrastPremedication, %preferencesFile%, Calculations, ShowContrastPremedication, 1
		IniRead, ShowFleischnerCriteria, %preferencesFile%, Calculations, ShowFleischnerCriteria, 1
		IniRead, ShowNASCETCalculator, %preferencesFile%, Calculations, ShowNASCETCalculator, 1
		IniRead, CalciumScoreMonitor, %preferencesFile%, CalciumScore, Monitor, 1
		IniRead, CalciumScoreX, %preferencesFile%, CalciumScore, X, 0
		IniRead, CalciumScoreY, %preferencesFile%, CalciumScore, Y, 0
		IniRead, CalciumScoreWidth, %preferencesFile%, CalciumScore, Width, 0
		IniRead, CalciumScoreHeight, %preferencesFile%, CalciumScore, Height, 0
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
	ShowNASCETCalculator := (ShowNASCETCalculator = "1")
    PauseDuration += 0
}

LoadPreferencesFromFile()
InitializeRecommendations() ; Make sure to call this function before using the Nodule class -- used for fleischner

; ==========================================
; Main Script Logic
; ==========================================

; Set up context menu for target applications
#If IsTargetApp()

; Hook the spacebar to reset the timer if it's pressed after a single period
$Space::
    if (A_PriorKey == "." and A_TimeSincePriorHotkey < 500)
    {
        lastPeriodTime := 0
    }
    SendInput {Space}
return

RButton::
{
        
    ; Activate the window if it's not already active
    if (windowUnderCursor != WinActive("A")) {
        WinActivate, ahk_id %windowUnderCursor%
    }
    
    ; Small delay to ensure window activation and to allow for text selection
    Sleep, 50

    ; Get selected text
    g_SelectedText := GetSelectedText()

    ; Show our custom menu
    CreateCustomMenu()
    Menu, CustomMenu, Show
    Menu, CustomMenu, DeleteAll

    return
}
#If

; Function to check if the current window is a target application
IsTargetApp() {
    MouseGetPos, , , windowUnderCursor
    WinGetClass, windowClass, ahk_id %windowUnderCursor%
    WinGet, windowExe, ProcessName, ahk_id %windowUnderCursor%
    
    for index, app in TargetApps {
        if (windowClass == StrReplace(app, "ahk_class ", "") || windowExe == StrReplace(app, "ahk_exe ", "")) {
            return true
        }
    }
    return false
}

; Function to create the custom right-click menu
CreateCustomMenu() {
    global ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity, ShowPregnancyDates, ShowMenstrualPhase
    global ShowAdrenalWashout, ShowThymusChemicalShift, ShowHepaticSteatosis
    global ShowMRILiverIron, ShowStatistics, ShowNumberRange, DarkMode
    global ShowCalciumScorePercentile, ShowContrastPremedication

    ; Create the menu
    Menu, CustomMenu, Add
    Menu, CustomMenu, DeleteAll

    ; Set menu colors and properties based on Dark Mode
    if (DarkMode) {
        Menu, CustomMenu, Color, 0xA9A9A9
        
    } else {
        Menu, CustomMenu, Color, Default
        
    }

    ; Add standard editing options
    Menu, CustomMenu, Add, Cut, MenuCut
    Menu, CustomMenu, Add, Copy, MenuCopy
    Menu, CustomMenu, Add, Paste, MenuPaste
    Menu, CustomMenu, Add, Delete, MenuDelete
    Menu, CustomMenu, Add

    ; Add custom calculation options
    Menu, CustomMenu, Add, Compare Measurement Sizes, CompareNoduleSizes
    Menu, CustomMenu, Add, Sort Measurement Sizes, SortNoduleSizes
    Menu, CustomMenu, Add, Capture Calcium Score, CaptureCalciumScoreMenu
    
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
    if (ShowContrastPremedication) {
        Menu, CustomMenu, Add, Calculate Contrast Premedication, CalculateContrastPremedication
    }
	if (ShowFleischnerCriteria) {
		Menu, CustomMenu, Add, Calculate Fleischner Criteria, CalculateFleischnerCriteria
	}
	if (ShowNASCETCalculator) {
		Menu, CustomMenu, Add, Calculate NASCET, CalculateNASCET
	}
    Menu, CustomMenu, Add
    Menu, CustomMenu, Add, Pause Script, PauseScript
    Menu, CustomMenu, Add, Preferences, ShowPreferences
}
; Standard editing functions
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



; Custom calculation functions
; Each function now uses g_SelectedText instead of SelectedText

CaptureCalciumScoreMenu:
    result := CaptureCalciumScore()
    if (result)
        ShowCalciumScore(result)
    else
        MsgBox, No calcium scores were captured or the capture was cancelled.
return

ParseUltrasoundMeasurementsMenu:
    result := ParseUltrasoundMeasurements()
    if (result && result.measurements.Count() > 0)
        ShowUltrasoundMeasurements(result)
    else
        MsgBox, No measurements were captured or the capture was cancelled.
return

CalculateCalciumScorePercentile:
    Result := CalculateCalciumScorePercentile(g_SelectedText)
    ShowResult(Result)
return

CalculateEllipsoidVolume:
    Result := CalculateEllipsoidVolume(g_SelectedText)
    ShowResult(Result)
return

CalculateBulletVolume:
    Result := CalculateBulletVolume(g_SelectedText)
    ShowResult(Result)
return

CalculatePSADensity:
    Result := CalculatePSADensity(g_SelectedText)
    ShowResult(Result)
return

CalculatePregnancyDates:
    Result := CalculatePregnancyDates(g_SelectedText)
    ShowResult(Result)
return

CalculateMenstrualPhase:
    Result := CalculateMenstrualPhase(g_SelectedText)
    ShowResult(Result)
return

SortNoduleSizes:
    ProcessedText := ProcessAllNoduleSizes(g_SelectedText)
    if (ProcessedText != g_SelectedText)
    {
        ; Preserve leading and trailing spaces only if they existed in the original input
        leadingSpace := (SubStr(g_SelectedText, 1, 1) == " ") ? " " : ""
        trailingSpace := (SubStr(g_SelectedText, 0) == " ") ? " " : ""
        Clipboard := leadingSpace . Trim(ProcessedText) . trailingSpace
        Send, ^v
    }
return

Statistics:
    Result := CalculateStatistics(g_SelectedText)
    ShowResult(Result)
return

CalculateIronContent:
    Result := EstimateIronContent(g_SelectedText)
    ShowResult(Result)
return

Range:
    Result := CalculateRange(g_SelectedText)
    ShowResult(Result)
return

CalculateAdrenalWashout:
    Result := CalculateAdrenalWashout(g_SelectedText)
    ShowResult(Result)
return

CalculateThymusChemicalShift:
    Result := CalculateThymusChemicalShift(g_SelectedText)
    ShowResult(Result)
return

CalculateHepaticSteatosis:
    Result := CalculateHepaticSteatosis(g_SelectedText)
    ShowResult(Result)
return

CompareNoduleSizes:
    Result := CompareNoduleSizes(g_SelectedText)
    ShowResult(Result)
return

CalculateFleischnerCriteria:
    Result := ProcessNodules(g_SelectedText)
    ShowResult(Result)
return

CalculateNASCET:
    Result := CalculateNASCET(g_SelectedText)
    ShowResult(Result)
return

CloseResultBox:
    Gui, ResultBox:Destroy

return

; Function to get selected text (unchanged)
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
    global DarkMode, originalMouseX, originalMouseY
    
    ; Get current mouse position (relative to the entire desktop)
    CoordMode, Mouse, Screen
    MouseGetPos, originalMouseX, originalMouseY
    
    ; Determine which monitor the mouse is on
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount%
    {
        SysGet, monArea, Monitor, %A_Index%
        if (originalMouseX >= monAreaLeft && originalMouseX <= monAreaRight && originalMouseY >= monAreaTop && originalMouseY <= monAreaBottom)
        {
            activeMonitor := A_Index
            break
        }
    }
    
    ; Get dimensions of the active monitor
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    monitorWidth := workAreaRight - workAreaLeft
    monitorHeight := workAreaBottom - workAreaTop
    
    ; Calculate maximum dimensions for the GUI (50% of monitor width and height)
    maxWidth := monitorWidth * 0.5
    maxHeight := monitorHeight * 0.5
    
    ; Create a temporary GUI to measure text
    Gui, TempMeasure:New, +AlwaysOnTop
    Gui, TempMeasure:Font, s10, Segoe UI
    Gui, TempMeasure:Add, Text, w%maxWidth% wrap, %Result%
    GuiControlGet, TextSize, TempMeasure:Pos, Static1
	GuiControl, Focus, Invisible
    Gui, TempMeasure:Destroy
    
    ; Calculate required width and height
    requiredWidth := TextSizeW + 40  ; Add some padding
    requiredHeight := TextSizeH + 80  ; Add space for button and padding
    
    ; Adjust dimensions if they exceed the maximum
    guiWidth := (requiredWidth > maxWidth) ? maxWidth : requiredWidth
    guiHeight := (requiredHeight > maxHeight) ? maxHeight : requiredHeight
    
    ; Ensure minimum dimensions
    guiWidth := (guiWidth < 300) ? 300 : guiWidth
    guiHeight := (guiHeight < 200) ? 200 : guiHeight
    
    ; Calculate position for the GUI
    xPos := originalMouseX + 10
    yPos := originalMouseY + 10
    
    ; Ensure the GUI doesn't go off-screen
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight
    
    ; Create GUI with AlwaysOnTop option
    Gui, ResultBox:New, +AlwaysOnTop -SysMenu +Owner
    Gui, ResultBox:Margin, 10, 10
    
    ; Apply styling
    if (DarkMode) {
        Gui, ResultBox:Color, 0x2C2C2C, 0x2C2C2C
        textColor := "cE0E0E0"
        buttonOptions := "Background333333 c999999"
    } else {
        Gui, ResultBox:Color, 0xF0F0F0, 0xF0F0F0
        textColor := "c000000"
        buttonOptions := "Background777777 cFFFFFF"
    }
    
    Gui, ResultBox:Font, s10 %textColor%, Segoe UI
    
    ; Add text control with vertical scrollbar and word wrap, and no highlighting
    editHeight := guiHeight - 50  ; Subtract height for button and margins
    Gui, ResultBox:Add, Edit, vResultText ReadOnly -E0x200 +E0x20000 Wrap VScroll w%guiWidth% h%editHeight%, %Result%
    
    ; Add a close button with improved styling
    Gui, ResultBox:Font, s9 bold, Segoe UI
    Gui, ResultBox:Add, Button, gCloseResultBox w90 x10 y+10 %buttonOptions%, Close
    
    ; Show the GUI
    Gui, ResultBox:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Result
    
    ; Add an invisible control for focus
	Gui, ResultBox:Add, Text, Hidden vInvisibleControl

	; Show the GUI
	Gui, ResultBox:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Result

	; Shift focus to the invisible control
	GuiControl, Focus, InvisibleControl

	; Get the position of the close button
	GuiControlGet, ClosePos, ResultBox:Pos, Close
    
    ; Move the mouse over the close button
    MouseMove, % xPos + ClosePosX + ClosePosW/2, % yPos + ClosePosY + ClosePosH/2, 0
    
    ; Copy result to clipboard
    Clipboard := Result
    
    return
}

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
	volNotGiven = 1
	volumeMethod = User Supplied
	
    ; Regular expression for PSA value
    PSARegEx := "i)PSA\s*(?:level|value)?:?\s*(\d+(?:\.\d+)?)\s*(?:ng\/ml|ng\/mL|ng/ml|ng/mL|ng\/cc|ng/cc)?"
    
    ; Regular expression for prostate volume (all formats)
    VolumeRegEx := "i)(?:(?:Calculated\s*)?(?:ellipsoid\s*)?volume:?\s*(\d+(?:\.\d+)?)\s*(?:cc|cm3|mL|ml)|(?:Prostate )?Size:?.*?\((\d+(?:\.\d+)?)\s*cc\)|(\d+(?:\.\d+)?)\s*x\s*\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?\s*cm\s*\((\d+(?:\.\d+)?)\s*cc\))"

    if (RegExMatch(input, PSARegEx, PSAMatch)) {
        PSALevel := PSAMatch1
    } else {
        return "Invalid input format for PSA density calculation.`nSuggested format:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm"
    }

    if (RegExMatch(input, VolumeRegEx, VolumeMatch)) {
        ProstateVolume := VolumeMatch1 ? VolumeMatch1 : (VolumeMatch2 ? VolumeMatch2 : (VolumeMatch4 ? VolumeMatch4 : ""))
    } else {
		volNotGiven=0
        ; Try to calculate volume using CalculateBullettVolume
        bulletResult := CalculateBulletVolume(input)
        if (!InStr(bulletResult, "Invalid input")) {
			ProstateVolume := RegExReplace(bulletResult, "s).*?(\d+(?:\.\d+)?)(?:\s*cc)?\).*", "$1")
            volumeMethod = Bullett Volume
            ; If volume >= 55 cc, use CalculateEllipsoidVolume instead (https://pubs.rsna.org/doi/10.1148/radiol.2501080290#:~:text=Overall%2C%20between%2066%25%20and%2075,bullet%20formula%20is%20highly%20accurate.)
            if (ProstateVolume >= 55) {
				ellipsoidResult := CalculateEllipsoidVolume(input)
				volumeMethod = Ellipsoid Volume
                if (!InStr(ellipsoidResult, "Invalid input")) {
                    ProstateVolume := RegExReplace(ellipsoidResult, "s).*?(\d+(?:\.\d+)?)(?:\s*cc)?\).*", "$1")
                }
            }
			
        } else {
            return "Prostate volume or dimensions not found or invalid in the input.`nSuggested format:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm"
        }
    }

    PSADensity := PSALevel / ProstateVolume
    PSADensity := Round(PSADensity, 3)
	
	if(volNotGiven=0){
		result := input . "`nProstate volume: " . ProstateVolume . " cc " . "- " . volumeMethod . "`nPSA Density: " . PSADensity . (DisplayUnits ? " ng/mL/cc" : "")
	} else {
		result := input . "`nPSA Density: " . PSADensity . (DisplayUnits ? " ng/mL/cc" : "")
	}
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



CalculateThymusChemicalShift(input) {
    RegExNeedle := "i)thymus.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)(?:.*?paraspinous.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+))?"

    if (RegExMatch(input, RegExNeedle, match)) {
        thymusInPhase := match2
        thymusOutOfPhase := match4
        paraspinousInPhase := match6
        paraspinousOutOfPhase := match8

        signalIntensityIndex := ((thymusInPhase-thymusOutOfPhase)/thymusInPhase)*100

        result := input
        if (paraspinousInPhase && paraspinousOutOfPhase) {
            outOfPhaseSignalRatio := (thymusOutOfPhase) / paraspinousOutOfPhase
            inPhaseSignalRatio := (thymusInPhase) / paraspinousInPhase
            chemicalShiftRatio := outOfPhaseSignalRatio / inPhaseSignalRatio

            if (DisplayAllValues) {
                result .= "`n`nChemical Shift Ratio: " . Round(chemicalShiftRatio, 3) . " (hyperplasia < 0.849)`n"    
                result .= "Thymus Signal Intensity Index (SII): " . round(signalIntensityIndex, 2) . "% (hyperplasia > 8.92)`n"
            }
            result .= "`n" . InterpretThymusChemicalShift(chemicalShiftRatio, signalIntensityIndex)
        } else {
            if (DisplayAllValues) {
                result .= "`nThymus Signal Intensity Index (SII): " . round(signalIntensityIndex, 2) . "%`n"
            }
            result .= "`n" . InterpretThymusChemicalShift("", signalIntensityIndex)
        }
        return result
    } else {
        return "Invalid input format for thymus chemical shift calculation.`nSample syntax: Thymus IP: 100, OP: 80, Paraspinous IP: 90, OP: 85`nOr for thymus only: Thymus IP: 100, OP: 80"
    }
}

InterpretThymusChemicalShift(chemicalShiftRatio := "", signalIntensityIndex := "") {
    global ShowCitations
    result := ""

    if (chemicalShiftRatio != "" && signalIntensityIndex != "") {
        ; Both chemical shift ratio and signal intensity index are provided
        if (chemicalShiftRatio > 0.849 && signalIntensityIndex > 8.92) {
            result := "Interpretation: Chemical Shift Ratio is greater than 0.849 and Signal Intensity Index is greater than 8.92%. Calculations are in conflict and therefore indeterminate, though probably consistent with thymic hyperplasia with single dual echo technique.`n`n"
        } else if (chemicalShiftRatio <= 0.849 && signalIntensityIndex > 8.92) {
            result := "Interpretation: Chemical Shift Ratio is less than or equal to 0.849 and Signal Intensity Index is greater than 8.92%. Findings are consistent with thymic hyperplasia with single dual echo technique. `n`n"
        } else if (chemicalShiftRatio > 0.849 && signalIntensityIndex <= 8.92) {
            result := "Interpretation: Chemical Shift Ratio is greater than 0.849 and Signal Intensity Index is less than or equal to 8.92%. Findings are not consistent with with typical thymic hyperplasia with single dual echo technique.`n`n"
        } else {  ; chemicalShiftRatio <= 0.849 && signalIntensityIndex <= 8.92
            result := "Interpretation: Chemical Shift Ratio is less than or equal to 0.849 and Signal Intensity Index is less than or equal to 8.92%. Calculations are in conflict and therefore indeterminate, possibly thymic hyperplasia with single dual echo technique.`n`n"
        }
    } else if (chemicalShiftRatio != "" && signalIntensityIndex == "") {
        ; Only chemical shift ratio is provided
        result := "Interpretation: Chemical Shift Ratio is " . (chemicalShiftRatio > 0.849 ? "greater than" : "less than or equal to") . " 0.849 with single dual echo technique.`n`n"
    } else if (chemicalShiftRatio == "" && signalIntensityIndex != "") {
        ; Only signal intensity index is provided
        if (signalIntensityIndex > 8.92) {
            result := "Interpretation: Signal Intensity Index is greater than 8.92%. This suggests thymic hyperplasia with single dual echo technique.`n`n"
        } else {
            result := "Interpretation: Signal Intensity Index is less than or equal to 8.92%. Findings are not consistent with typical thymic hyperplasia with single dual echo technique.`n`n"
        }
    } else {
        ; Neither chemical shift ratio nor signal intensity index is provided
        result := "Error: Both Chemical Shift Ratio and Signal Intensity Index are missing. At least one value is required for interpretation.`n`n"
    }

    if (ShowCitations = 1) {
        result .= "Citation: Priola AM, Priola SM, Ciccone G, Evangelista A, Cataldi A, Gned D, Pazè F, Ducco L, Moretti F, Brundu M, Veltri A. Differentiation of rebound and lymphoid thymic hyperplasia from anterior mediastinal tumors with dual-echo chemical-shift MR imaging in adulthood: reliability of the chemical-shift ratio and signal intensity index. Radiology. 2015 Jan;274(1):238-49. doi: 10.1148/radiol.14132665. Epub 2014 Aug 7. PMID: 25105246.`n"
    }

    return result
}




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


CompareNoduleSizes(input) {
    static RegExNeedle := "i)(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?.*?(?:previous(?:ly)?|prior|before|old|initial).*?(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?(?:\s*(?:on|dated?)\s*(\d{1,2}/\d{1,2}/\d{2,4}))?"
    
    if (!RegExMatch(input, RegExNeedle, match)) {
        RegExNeedle := "i)(?:previous(?:ly)?|prior|before|old|initial).*?(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?(?:\s*(?:on|dated?)\s*(\d{1,2}/\d{1,2}/\d{2,4}))?.*?(?:now|current(?:ly)?|present|new|recent(?:ly)?|follow[- ]?up).*?(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?"
        if (!RegExMatch(input, RegExNeedle, match)) {
            return "Invalid input format. Please provide both current and previous measurements."
        }
        current := match6 . " " . match7, previous := match2 . " " . match3
        currentDate := match5, previousDate := match1 ? match1 : match4
    } else {
        current := match2 . " " . match3, previous := match5 . " " . match6
        currentDate := match1, previousDate := match4 ? match4 : match7
    }
    
    return CompareMeasurements(previous, current, previousDate, currentDate, input)
}

CompareMeasurements(previous, current, previousDate, currentDate, input) {
    ; Process measurements
    prev := ProcessMeasurement(previous)
    curr := ProcessMeasurement(current)
    
    ; Check for dimension mismatch
    if (prev.dimensions.MaxIndex() != curr.dimensions.MaxIndex()) {
        return "Error: Mismatch in number of dimensions between previous and current measurements.`nPrevious: " . previous . "`nCurrent: " . current
    }

    ; Initialize result string
    result := input . "`n`n"
    result .= "Previous Date: " . (previousDate ? previousDate : "Not provided") . "`n"
    result .= "Current Date: " . (currentDate ? currentDate : "Assumed as today") . "`n`n"

    ; Process dimensions
    prevLongestDim := 0
    currLongestDim := 0
    
    Loop, % prev.dimensions.MaxIndex()
    {
        prevDim := prev.dimensions[A_Index]
        currDim := curr.dimensions[A_Index]
        
        ; Convert to cm if necessary
        prevDimCm := (prev.unit == "mm") ? prevDim / 10 : prevDim
        currDimCm := (curr.unit == "mm") ? currDim / 10 : currDim
        
        ; Track longest dimension
        prevLongestDim := (prevDimCm > prevLongestDim) ? prevDimCm : prevLongestDim
        currLongestDim := (currDimCm > currLongestDim) ? currDimCm : currLongestDim
        
        ; Calculate change
        change := (currDimCm / prevDimCm - 1) * 100
        
        ; Append to result
        result .= "Dimension " . A_Index . ": " 
                . Round(prevDim, 2) . " " . prev.unit 
                . " -> " 
                . Round(currDim, 2) . " " . curr.unit 
                . " (" . (change >= 0 ? "+" : "") . Round(change, 1) . "%)`n"
    }
    
    ; Calculate longest dimension change
    longestDimChange := (currLongestDim / prevLongestDim - 1) * 100
    result .= "`nLongest dimension change: " . (longestDimChange >= 0 ? "+" : "") . Round(longestDimChange, 1) . "%`n"
    
    ; Calculate volumes
    prevVolumeNum := CalculateVolume(prev.dimensions, prev.unit)
    currVolumeNum := CalculateVolume(curr.dimensions, curr.unit)
    
    ; Process volume calculations if valid
    if (prevVolumeNum != "Invalid input" and currVolumeNum != "Invalid input") {
        volumeChange := (currVolumeNum / prevVolumeNum - 1) * 100
        result .= "Volume change: " . (volumeChange >= 0 ? "+" : "") . Round(volumeChange, 1) . "%`n"
        result .= "Previous volume: " . FormatVolume(prevVolumeNum) . "`n"
        result .= "Current volume: " . FormatVolume(currVolumeNum) . "`n"
        
        ; Process dates and calculate time-based metrics
        if (previousDate) {
            parsedPreviousDate := ParseDate(previousDate)
            if (parsedPreviousDate == "Invalid Date") {
                return result . "`nError: Invalid previous date format."
            }
            
            if (!currentDate) {
                FormatTime, currentDate, , MM/dd/yyyy
            }
            parsedCurrentDate := ParseDate(currentDate)
            if (parsedCurrentDate == "Invalid Date") {
                return result . "`nError: Invalid current date format."
            }
            
            ; Calculate time difference
            timeDiff := DateDiff(parsedPreviousDate, parsedCurrentDate) / 365.25
            result .= "Time difference: " . Round(timeDiff, 2) . " years`n"
            
            ; Calculate growth metrics if time difference is positive
            if (timeDiff > 0) {
                ; Calculate doubling time in days
                doublingTime := CalculateDoublingTime(prevVolumeNum, currVolumeNum, timeDiff)
                doublingTimeDays := doublingTime * 365.25
                result .= "Doubling time: " . (doublingTime != "N/A" ? Round(doublingTimeDays, 0) . " days" : doublingTime) . "`n"
                
                ; Calculate exponential growth rate in % per year
                growthRate := CalculateExponentialGrowth(prevVolumeNum, currVolumeNum, timeDiff)
                result .= "Exponential Growth Rate: " . Round(growthRate * 100, 2) . "% per year"
            } else {
                result .= "Note: Doubling time and Growth Rate not calculated due to invalid time difference."
            }
        } else {
            result .= "Note: Doubling time and Growth Rate not calculated due to missing previous date."
        }
    } else {
        result .= "Error: Unable to calculate volume for one or both measurements."
    }
    
    return result
}

CalculateVolume(dimensions, unit) {
    static PI := 3.14159265358979
    
    if (dimensions.MaxIndex() == 1)
        volume := (4/3) * PI * (dimensions[1] / 2) ** 3  ; Sphere
    else if (dimensions.MaxIndex() == 2)
        volume := (4/3) * PI * (dimensions[1] / 2) * (dimensions[2] / 2) * ((dimensions[1] + dimensions[2]) / 4)  ; Ellipsoid with 3rd dim as average
    else if (dimensions.MaxIndex() == 3)
        volume := CalculateEllipsoidVolumeNumeric(dimensions[1] . " x " . dimensions[2] . " x " . dimensions[3], unit)
    else
        return "Invalid input"
    
    return (unit == "mm") ? volume / 1000 : volume  ; Convert to cm³ if necessary
}

CalculateDoublingTime(initialVolume, finalVolume, time) {
	
    growthRate := (finalVolume / initialVolume) ** (1 / time) - 1
    return (growthRate > 0) ? (Ln(2) / Ln(1 + growthRate)) : "N/A"
}

CalculateExponentialGrowth(initialVolume, finalVolume, time) {
    growthRate := Ln(finalVolume/initialVolume) / time
    return growthRate  ; Return as a decimal, will be converted to percentage in the main function
}

DateDiff(date1, date2) {
    EnvSub, date2, %date1%, Days
    return date2
}

FormatVolume(volume) {
    return (volume < 1) ? Round(volume * 1000, 1) . " cu-mm" : Round(volume, 1) . " cc"
}

CalculateEllipsoidVolumeNumeric(input, unit) {
    RegExNeedle := "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
    if (RegExMatch(input, RegExNeedle, match)) {
        dimensions := [match1 + 0, match2 + 0, match3 + 0]  ; Convert to numbers
        dimensions := SortDimensions(dimensions)

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



ShowPreferences() {
    global DisplayUnits, DisplayAllValues, ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity, ShowPregnancyDates, ShowMenstrualPhase, PauseDuration
    global ShowAdrenalWashout, ShowThymusChemicalShift, ShowHepaticSteatosis, ShowMRILiverIron, ShowStatistics, ShowNumberRange, DarkMode
	global ShowCitations, ShowArterialAge
	
	; Determine colors based on dark mode
    if (DarkMode) {
        bgColor := "0x2C2C2C"
        textColor := "cE0E0E0"
        buttonOptions := "Background333333 c999999"
    } else {
        bgColor := "0xF0F0F0"
        textColor := "c000000"
        buttonOptions := "Background777777 cFFFFFF"
    }
	
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

      ; Create preferences GUI
    Gui, Preferences:New, +AlwaysOnTop
    Gui, Preferences:Color, %bgColor%, %bgColor%
    Gui, Preferences:Font, s10 %textColor%, Segoe UI
    preferencesHwnd := WinExist()
    
    ;Gui, Add, Text, x10 y10 w200, Display Options:
    ;Gui, Add, Checkbox, x10 y30 w200 vDisplayUnits Checked%DisplayUnits%, Display units with values
    ;Gui, Add, Checkbox, x10 y60 w200 vDisplayAllValues Checked%DisplayAllValues%, Display all calculated values
    
    Gui, Add, Text, x10 y10 w200, Select functions to display:
	Gui, Add, Checkbox, vDarkMode Checked%DarkMode%, Dark Mode
    Gui, Add, Checkbox, x10 y60 w200 vShowEllipsoidVolume Checked%ShowEllipsoidVolume%, Ellipsoid Volume
    Gui, Add, Checkbox, x10 y90 w200 vShowBulletVolume Checked%ShowBulletVolume%, Bullet Volume
    Gui, Add, Checkbox, x10 y120 w200 vShowPSADensity Checked%ShowPSADensity%, PSA Density
    Gui, Add, Checkbox, x10 y150 w200 vShowPregnancyDates Checked%ShowPregnancyDates%, Pregnancy Dates
    Gui, Add, Checkbox, x10 y180 w200 vShowMenstrualPhase Checked%ShowMenstrualPhase%, Menstrual Phase
    Gui, Add, Checkbox, x10 y210 w200 vShowAdrenalWashout Checked%ShowAdrenalWashout%, Adrenal Washout
    Gui, Add, Checkbox, x10 y240 w200 vShowThymusChemicalShift Checked%ShowThymusChemicalShift%, Thymus Chemical Shift
    Gui, Add, Checkbox, x10 y270 w200 vShowHepaticSteatosis Checked%ShowHepaticSteatosis%, Hepatic Steatosis
    Gui, Add, Checkbox, x10 y300 w200 vShowMRILiverIron Checked%ShowMRILiverIron%, MRI Liver Iron Content
    Gui, Add, Checkbox, x10 y330 w200 vShowStatistics Checked%ShowStatistics%, Calculate Statistics
    Gui, Add, Checkbox, x10 y360 w200 vShowNumberRange Checked%ShowNumberRange%, Calculate Number Range
	Gui, Add, Checkbox, x10 y390 w200 vShowCalciumScorePercentile Checked%ShowCalciumScorePercentile%, Calcium Score Percentile
    Gui, Add, Checkbox, x10 y420 w200 vShowCitations Checked%ShowCitations%, Show Citations in Output
    Gui, Add, Checkbox, x10 y450 w200 vShowArterialAge Checked%ShowArterialAge%, Show Arterial Age
	Gui, Add, Checkbox, x10 y480 w200 vShowContrastPremedication Checked%ShowContrastPremedication%,  Contrast Premedication
	Gui, Add, Checkbox, x10 y510 w200 vShowFleischnerCriteria Checked%ShowFleischnerCriteria%, Fleischner Criteria
	Gui, Add, Checkbox, x10 y540 w200 vShowNASCETCalculator Checked%ShowNASCETCalculator%, NASCET Calculator
	Gui, Add, Text, x10 y570 w200, Calcium Score Table Location:
    Gui, Add, Text, x10 y600 w100, Monitor:
    Gui, Add, Edit, x120 y600 w80 vCalciumScoreMonitor, %CalciumScoreMonitor%
    Gui, Add, Text, x10 y630 w100, X:
    Gui, Add, Edit, x120 y630 w80 vCalciumScoreX, %CalciumScoreX%
    Gui, Add, Text, x10 y660 w100, Y:
    Gui, Add, Edit, x120 y660 w80 vCalciumScoreY, %CalciumScoreY%
    Gui, Add, Text, x10 y690 w100, Width:
    Gui, Add, Edit, x120 y690 w80 vCalciumScoreWidth, %CalciumScoreWidth%
    Gui, Add, Text, x10 y720 w100, Height:
    Gui, Add, Edit, x120 y720 w80 vCalciumScoreHeight, %CalciumScoreHeight%
	Gui, Add, Text, x10 y750 w200, Pause Length (cur: %currentPauseDuration%):
    Gui, Add, DropDownList, x10 y780 w200 vPauseDurationChoice, 3 minutes|10 minutes|30 minutes|1 hour|10 hours
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
    Gui, Add, Button, x60 y820 w100 gSavePreferences, Save
    
    Gui, Show, w220 h920
}



SavePreferences:
    Gui, Submit, NoHide
    global DisplayUnits, DisplayAllValues
    global ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity, ShowPregnancyDates, ShowMenstrualPhase
    global ShowAdrenalWashout, ShowThymusChemicalShift, ShowHepaticSteatosis
    global ShowMRILiverIron, ShowStatistics, ShowNumberRange
    global PauseDuration, DarkMode
	global ShowArterialAge, ShowCitations
	global ShowContrastPremedication
	

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
		
    CalciumScoreMonitor := CalciumScoreMonitor
    CalciumScoreX := CalciumScoreX
    CalciumScoreY := CalciumScoreY
    CalciumScoreWidth := CalciumScoreWidth
    CalciumScoreHeight := CalciumScoreHeight
	
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
	IniWrite, %ShowContrastPremedication%, %A_ScriptDir%\preferences.ini, Calculations, ShowContrastPremedication
	IniWrite, %ShowFleischnerCriteria%, %A_ScriptDir%\preferences.ini, Calculations, ShowFleischnerCriteria
	IniWrite, %ShowNASCETCalculator%, %A_ScriptDir%\preferences.ini, Calculations, ShowNASCETCalculator
	IniWrite, %CalciumScoreMonitor%, %A_ScriptDir%\preferences.ini, CalciumScore, Monitor
    IniWrite, %CalciumScoreX%, %A_ScriptDir%\preferences.ini, CalciumScore, X
    IniWrite, %CalciumScoreY%, %A_ScriptDir%\preferences.ini, CalciumScore, Y
    IniWrite, %CalciumScoreWidth%, %A_ScriptDir%\preferences.ini, CalciumScore, Width
    IniWrite, %CalciumScoreHeight%, %A_ScriptDir%\preferences.ini, CalciumScore, Height
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
    } else if (age >= 30) {
        result := CalculateHoffScore(age, sex, score, race)
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

;=======================================================================
; Contrast Premedication Function

CalculateContrastPremedication() {
    defaultDateTime := GetDefaultDateTime()
    FormatTime, defaultDate, %defaultDateTime%, yyyyMMdd
    FormatTime, defaultTime, %defaultDateTime%, HH:mm

    ; Get current mouse position (relative to the entire desktop)
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY

    ; Determine which monitor the mouse is on
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount%
    {
        SysGet, monArea, Monitor, %A_Index%
        if (mouseX >= monAreaLeft && mouseX <= monAreaRight && mouseY >= monAreaTop && mouseY <= monAreaBottom)
        {
            activeMonitor := A_Index
            break
        }
    }

    ; Get dimensions of the active monitor
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    monitorWidth := workAreaRight - workAreaLeft
    monitorHeight := workAreaBottom - workAreaTop

    ; Calculate GUI dimensions and position
    guiWidth := 250
    guiHeight := 160
    xPos := mouseX + 10
    yPos := mouseY + 10

    ; Ensure the GUI doesn't go off-screen
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight

    ; Create GUI with AlwaysOnTop option and position it near the mouse
    Gui, ContrastPremed:New, +AlwaysOnTop
    Gui, ContrastPremed:Add, Text, x10 y10, Scan Date:
    Gui, ContrastPremed:Add, DateTime, x10 y30 w120 vScanDate Choose%defaultDate%
    Gui, ContrastPremed:Add, Text, x140 y10, Scan Time:
    Gui, ContrastPremed:Add, DropDownList, x140 y30 w100 vScanTime, % CreateTimeList(defaultTime)
    Gui, ContrastPremed:Add, Radio, x10 y60 w200 vPremedProtocol Checked, Prednisone (13-7-1 hour)
    Gui, ContrastPremed:Add, Radio, x10 y80 w200, Methylprednisolone (12-2 hour)
    Gui, ContrastPremed:Add, Checkbox, x10 y100 w200 vIncludeDiphenhydramine Checked, Include Diphenhydramine
    Gui, ContrastPremed:Add, Button, x10 y130 w100 gCalculatePremedTiming, Calculate
    Gui, ContrastPremed:Add, Button, x120 y130 w100 gShowPremedDosages, Show Dosages
    Gui, ContrastPremed:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Contrast Premedication

    return
}

GetDefaultDateTime() {
    defaultDateTime := A_Now
    EnvAdd, defaultDateTime, 13, Hours
    
    ; Extract hours and minutes
    FormatTime, hours, %defaultDateTime%, HH
    FormatTime, minutes, %defaultDateTime%, mm
    
    ; Round up to nearest 5 minutes
    minutes := Ceil(minutes / 5) * 5
    if (minutes = 60) {
        minutes := 0
        hours := hours + 1
    }
    
    ; Handle day change if hours exceed 23
    if (hours >= 24) {
        hours := Mod(hours, 24)
        EnvAdd, defaultDateTime, 1, Days
    }
    
    ; Format the time back into the datetime
    formattedTime := Format("{:02d}{:02d}00", hours, minutes)
    FormatTime, formattedDate, %defaultDateTime%, yyyyMMdd
    return formattedDate . formattedTime
}

CreateTimeList(defaultTime) {
    timeList := ""
    Loop, 24 {
        hour := A_Index - 1
        Loop, 12 {
            minute := (A_Index - 1) * 5
            time := Format("{:02d}:{:02d}", hour, minute)
            timeList .= time . "|"
        }
    }
    ; Remove the trailing pipe
    timeList := RTrim(timeList, "|")
    
    ; Set the default time
    if (defaultTime) {
        timeList := StrReplace(timeList, defaultTime, defaultTime . "||")
    }
    
    return timeList
}

GetScanDateTime(scanDate, scanTime) {
    if (scanDate = "" || scanTime = "") {
        return GetDefaultDateTime()
    }
    return CombineDateTime(scanDate, scanTime)
}

CalculatePremedTiming:
    Gui, ContrastPremed:Submit, NoHide
    scanDateTime := GetScanDateTime(ScanDate, ScanTime)
    if (!scanDateTime) {
        MsgBox, 0, Error, Please enter a valid date and time.
        return
    }
    CalculateAndShowPremedResult(scanDateTime, PremedProtocol, IncludeDiphenhydramine)
return

CombineDateTime(scanDate, scanTime) {
    FormatTime, formattedDate, %scanDate%, yyyyMMdd
    formattedTime := StrReplace(scanTime, ":")
    fullDateTime := formattedDate . formattedTime . "00"
    
    if (!IsValidDateTime(fullDateTime)) {
        return false
    }
    
    return fullDateTime
}

CalculateAndShowPremedResult(scanDateTime, premedProtocol, includeDiphenhydramine) {
    FormatTime, scanDateTimeFormatted, %scanDateTime%, MM/dd/yyyy hh:mm tt
    
    result := "Scan time: " . scanDateTimeFormatted . " Contrast Premedication Schedule:`n`n"
    
    if (premedProtocol = 1) {
        ; Prednisone-based protocol
        result .= FormatPremedTime(scanDateTime, -13, "13", premedProtocol, includeDiphenhydramine)
        result .= FormatPremedTime(scanDateTime, -7, "7", premedProtocol, includeDiphenhydramine)
        result .= FormatPremedTime(scanDateTime, -1, "1", premedProtocol, includeDiphenhydramine)
    } else {
        ; Methylprednisolone-based protocol
        result .= FormatPremedTime(scanDateTime, -12, "12", premedProtocol, includeDiphenhydramine)
        result .= FormatPremedTime(scanDateTime, -2, "2", premedProtocol, includeDiphenhydramine)
    }
    
    result .= "`nNote: Premedication regimens less than 4-5 hours in duration (oral or IV) have not been shown to be effective.`n"
    result .= "If a patient is unable to take oral medication, 200 mg hydrocortisone IV may be substituted for each dose of oral prednisone in the 13-7-1 premedication regimen.`n"
    
    if (ShowCitations) {
        result .= "`nCitation: ACR Manual on Contrast Media. 2023 American College of Radiology. https://www.acr.org/Clinical-Resources/Contrast-Manual`n"
    }
    
    ShowResult(result)
}

FormatPremedTime(scanTime, hoursOffset, label, protocol, includeDiphenhydramine) {
    premedTime := DateAdd(scanTime, hoursOffset, "hours")
    FormatTime, premedTimeFormatted, %premedTime%, MM/dd/yyyy hh:mm tt
    
    medication := GetMedicationInfo(label, protocol, includeDiphenhydramine)
    
    return label . " hours before (" . premedTimeFormatted . "):`n" . medication . "`n`n"
}

GetMedicationInfo(label, protocol, includeDiphenhydramine) {
    if (protocol = 1) {  ; Prednisone-based protocol
        medication := "- Prednisone 50 mg PO"
        if (label = "1" && includeDiphenhydramine) {
            medication .= "`n- Diphenhydramine 50 mg IV, IM, or PO"
        }
    } else {  ; Methylprednisolone-based protocol
        medication := "- Methylprednisolone 32 mg PO"
        if (label = "2" && includeDiphenhydramine) {
            medication .= "`n- Diphenhydramine 50 mg IV, IM, or PO"
        }
    }
    return medication
}

IsValidDateTime(dateTimeStr) {
    try {
        FormatTime, test, %dateTimeStr%, yyyy-MM-dd HH:mm:ss
        return true
    } catch {
        return false
    }
}

ShowPremedDosages:
    Gui, ContrastPremed:Submit, NoHide
    ShowPremedDosages(PremedProtocol, IncludeDiphenhydramine)
return

ShowPremedDosages(premedProtocol, includeDiphenhydramine) {
    dosages := "Contrast Premedication Dosages:`n`n"
    
    if (premedProtocol = 1) {
        dosages .= "Prednisone-based Protocol (13-7-1 hour):`n`n"
        dosages .= "13 hours before:`n- Prednisone 50 mg PO`n`n"
        dosages .= "7 hours before:`n- Prednisone 50 mg PO`n`n"
        dosages .= "1 hour before:`n- Prednisone 50 mg PO`n"
        if (includeDiphenhydramine) {
            dosages .= "- Diphenhydramine 50 mg IV, IM, or PO`n"
        }
    } else {
        dosages .= "Methylprednisolone-based Protocol (12-2 hour):`n`n"
        dosages .= "12 hours before:`n- Methylprednisolone 32 mg PO`n`n"
        dosages .= "2 hours before:`n- Methylprednisolone 32 mg PO`n"
        if (includeDiphenhydramine) {
            dosages .= "- Diphenhydramine 50 mg IV, IM, or PO`n"
        }
    }
    
    dosages .= "`nNotes:`n"
    dosages .= "- Premedication regimens less than 4-5 hours in duration (oral or IV) have not been shown to be effective.`n"
    dosages .= "- If a patient is unable to take oral medication, 200 mg hydrocortisone IV may be substituted for each dose of oral prednisone in the 13-7-1 premedication regimen.`n"
    dosages .= "- Diphenhydramine is considered optional. If a patient is allergic to diphenhydramine, an alternate anti-histamine without cross-reactivity may be considered, or the anti-histamine may be omitted.`n"
    dosages .= "- These dosages are based on the ACR Manual on Contrast Media. Please consult with a healthcare professional for patient-specific recommendations.`n"
    
    if (ShowCitations) {
        dosages .= "`nCitation: ACR Manual on Contrast Media. 2023 American College of Radiology. https://www.acr.org/Clinical-Resources/Contrast-Manual`n"
    }
    
    ShowResult(dosages)
}

DateAdd(datetime, value, unit) {
    EnvAdd, datetime, %value%, %unit%
    return datetime
}

;============= END PREMEDICATION FUNCTION =============================


;=============PARSE ULTRASOUND IMAGE DATA =================

ParseUltrasoundMeasurements() {
    ; Use Vis2's built-in OCR function
    text := OCR()
    
    ; If no text was captured (user cancelled), return
    if (text == "")
        return

    ; Replace @ with 0 in the OCR output
    text := StrReplace(text, "@", "0")

    ; Pre-process the text to separate measurements for each organ
    organMeasurements := PreProcessText(text)
	
	MsgBox %text%

    ; Define the organs to look for with their variants and preferred units
    organs := { liver: {variants: ["liver", "liv"], unit: "cm", standardName: "Liver"}
               , spleen: {variants: ["spleen", "spl"], unit: "cm", standardName: "Spleen"}
               , pancreas: {variants: ["pancreas", "panc"], unit: "cm", standardName: "Pancreas"}
               , right_kidney: {variants: ["right kidney", "r kidney", "r kid", "rk"], unit: "cm", standardName: "Right Kidney"}
               , left_kidney: {variants: ["left kidney", "l kidney", "l kid", "lk"], unit: "cm", standardName: "Left Kidney"}
               , uterus: {variants: ["uterus", "ut"], unit: "cm", standardName: "Uterus"}
               , prostate: {variants: ["prostate", "prost"], unit: "cm", standardName: "Prostate"}
               , right_thyroid: {variants: ["right thyroid", "rt thyroid", "rt"], unit: "cm", standardName: "Right Thyroid"}
               , left_thyroid: {variants: ["left thyroid", "lt thyroid", "lt"], unit: "cm", standardName: "Left Thyroid"}
               , thyroid_isthmus: {variants: ["thyroid isthmus", "isthmus"], unit: "cm", standardName: "Thyroid Isthmus"}
               , right_thyroid_nodule_1: {variants: ["right thyroid nodule 1", "rt nodule 1", "right nodule 1"], unit: "mm", standardName: "Right Thyroid Nodule 1"}
			   , right_thyroid_nodule_2: {variants: ["right thyroid nodule 2", "rt nodule 2", "right nodule 2"], unit: "mm", standardName: "Right Thyroid Nodule 2"}
			   , right_thyroid_nodule_3: {variants: ["right thyroid nodule 3", "rt nodule 3", "right nodule 3"], unit: "mm", standardName: "Right Thyroid Nodule 3"}
			   , right_thyroid_nodule_4: {variants: ["right thyroid nodule 4", "rt nodule 4", "right nodule 4"], unit: "mm", standardName: "Right Thyroid Nodule 4"}
               , left_thyroid_nodule: {variants: ["left thyroid nodule", "lt nodule", "left nodule"], unit: "cm", standardName: "Left Thyroid Nodule"}
               , isthmic_thyroid_nodule: {variants: ["isthmic nodule", "isthmus nodule"], unit: "cm", standardName: "Isthmic Thyroid Nodule"}
               , endometrium: {variants: ["endometrium", "endo"], unit: "mm", standardName: "Endometrium"}
               , common_bile_duct: {variants: ["common bile duct", "cbd", "rcbd"], unit: "mm", standardName: "Common Bile Duct"}
               , gallbladder_wall: {variants: ["gallbladder wall", "gb wall"], unit: "mm", standardName: "Gallbladder Wall"}
               , bladder: {variants: ["bladder", "ubl"], unit: "cm", standardName: "Bladder"}
               , right_ovary: {variants: ["right ovary", "r ovary", "r ov"], unit: "cm", standardName: "Right Ovary"}
               , left_ovary: {variants: ["left ovary", "l ovary", "l ov"], unit: "cm", standardName: "Left Ovary"} }
    
    ; Initialize an object to store the measurements
    measurements := {}
    
    ; Parse the pre-processed text for each organ
    for organName, organText in organMeasurements {
        for organ, info in organs {
            if (HasOrganName(organName, info.variants)) {
                ; Pattern for measurements
                pattern := "i)(\d+(?:\.\d+)?)\s*(mm|cm)?"
                
                measureArr := []
                pos := 1
                while (pos := RegExMatch(organText, pattern, match, pos)) {
                    measure := match1 + 0  ; Convert to number
                    unit := match2 ? match2 : info.unit
                    
                    if (unit == "mm" && info.unit == "cm") {
                        measure := Round(measure / 10, 1)
                    } else if (unit == "cm" && info.unit == "mm") {
                        measure := Round(measure * 10, 0)
                    } else if (info.unit == "cm") {
                        measure := Round(measure, 1)
                    } else {
                        measure := Round(measure, 0)
                    }
                    
                    measureArr.Push(measure)
                    pos += StrLen(match)
                }
                
                if (measureArr.Length() > 0) {
                    measurements[info.standardName] := {measures: measureArr, unit: info.unit}
                }
                
                break  ; Stop checking other organs once a match is found
            }
        }
    }
    
    ; Check for bladder volume
    volumePattern := "i)bladder.*?volume[\s:]*(\d+(?:\.\d+)?)\s*(cm3|cc|ml)"
    if (RegExMatch(text, volumePattern, volMatch)) {
        if (measurements.HasKey("Bladder")) {
            measurements.Bladder.volume := {value: Round(volMatch1 + 0, 1), unit: volMatch2}
        }
    }
    
    ; Determine US type
    usType := DetermineUSType(measurements)
    
    return {measurements: measurements, usType: usType}
}

PreProcessText(text) {
    organMeasurements := {}
    lines := StrSplit(text, "`n", "`r")
    currentOrgan := ""
    currentMeasurement := ""

    for _, line in lines {
        trimmedLine := Trim(line)
        if (trimmedLine != "") {
            if (RegExMatch(trimmedLine, "i)^(.*?(?:thyroid|kidney|ovary|nodule|liver|spleen|pancreas|uterus|prostate|bladder|endo|cbd|gb)\s*\d*).*$", match)) {
                if (currentOrgan != "") {
                    organMeasurements[currentOrgan] := currentMeasurement
                }
                currentOrgan := Trim(match1)
                currentMeasurement := trimmedLine . "`n"
            } else if (currentOrgan != "") {
                currentMeasurement .= trimmedLine . "`n"
            }
        }
    }

    if (currentOrgan != "") {
        organMeasurements[currentOrgan] := currentMeasurement
    }

    return organMeasurements
}

HasOrganName(organName, variants) {
    for _, variant in variants {
        if (InStr(organName, variant)) {
            return true
        }
    }
    return false
}

DetermineUSType(measurements) {
    if (measurements.HasKey("Liver") && measurements.HasKey("Spleen") && measurements.HasKey("Gallbladder Wall") && measurements.HasKey("Common Bile Duct")) {
        if (measurements.HasKey("Left Kidney"))
            return "Complete Abdominal US"
        else
            return "RUQ US"
    } else if (measurements.HasKey("Right Ovary") && measurements.HasKey("Left Ovary") && measurements.HasKey("Uterus") && measurements.HasKey("Endometrium")) {
        return "Pelvic US"
    } else if (measurements.HasKey("Right Kidney") && measurements.HasKey("Left Kidney") && measurements.HasKey("Bladder")) {
        return "Kidney US"
    } else if (measurements.HasKey("Right Thyroid") && measurements.HasKey("Left Thyroid")) {
        return "Thyroid US"
    } else {
        return "Unknown US Type"
    }
}

ShowUltrasoundMeasurements(result) {
    measurements := result.measurements
    usType := result.usType
    
    output := "Ultrasound Measurements:`n"
    output .= "US Type: " . usType . "`n`n"
    
    for organ, data in measurements {
        output .= organ . ": "
        
        for i, measure in data.measures {
            output .= measure . (i < data.measures.Length() ? " x " : "")
        }
        output .= " " . data.unit . "`n"
        
        if (data.HasKey("volume")) {
            output .= "  Volume: " . data.volume.value . " " . data.volume.unit . "`n"
        }
        
        output .= "`n"
    }
    
    ShowResult(output)
}

;=======================================================
LevenshteinDistance(s, t) {
    m := StrLen(s)
    n := StrLen(t)
    d := []

    Loop, % m + 1
    {
        d[A_Index] := []
        d[A_Index, 1] := A_Index - 1
    }

    Loop, % n + 1
        d[1, A_Index] := A_Index - 1

    Loop, % m
    {
        i := A_Index
        Loop, % n
        {
            j := A_Index
            cost := (SubStr(s, i, 1) = SubStr(t, j, 1)) ? 0 : 1
            d[i+1, j+1] := Min(d[i, j+1] + 1, d[i+1, j] + 1, d[i, j] + cost)
        }
    }

    return d[m+1, n+1]
}

CaptureCalciumScore() {
    global CalciumScoreMonitor, CalciumScoreX, CalciumScoreY, CalciumScoreWidth, CalciumScoreHeight

    if (CalciumScoreMonitor > 0 && CalciumScoreWidth > 0 && CalciumScoreHeight > 0) {
        ; Use the specified monitor and coordinates
        SysGet, Mon, Monitor, %CalciumScoreMonitor%
        x := MonLeft + CalciumScoreX
        y := MonTop + CalciumScoreY
        w := CalciumScoreWidth
        h := CalciumScoreHeight

        text := OCR([x, y, w, h])
    } else {
        ; Use the current functionality (user selection)
        text := OCR()
    }

    if (text == "")
        return {report: "No text captured", warning: ""}

    result := ParseCalciumScore(text)
    
    ; Copy only the report text to clipboard
    Clipboard := result.report
    
    return result
}

ParseCalciumScore(text) {
    global g_arteryNames, g_levenshteinThreshold
    lines := StrSplit(RegExReplace(text, "\r\n|\r|\n", "`n"), "`n")
    
    ; Detect Region Agatston format using Levenshtein distance
    if (IsLikelyRegionAgatston(lines[1])) {
        return ParseFlexibleRegionAgatston(lines)
    }
    
    ; Existing code for the original format starts here
    scoreData := {}
    arteries := []
    scores := []
    inArterySection := false
    inScoreSection := false
    ocrTotal := 0

    if (isRegionAgatston) {
        ; Handle Region/Agatston format
        for index, line in lines {
            if (index == 1) ; Skip header
                continue
            
            parts := StrSplit(Trim(line), A_Tab)
            if (parts.Length() >= 2) {
                arteryName := Trim(parts[1])
                score := Round(parts[2] + 0) ; Agatston score, rounded to nearest integer
                
                mappedName := FindBestMatch(arteryName, g_arteryNames)
                if (mappedName && mappedName != "Total") {
                    scoreData[mappedName] := score
                } else if (mappedName == "Total") {
                    ocrTotal := score
                }
            }
        }
    } else {
        ; Handle original format
        arteries := []
        scores := []
        inArterySection := false
        inScoreSection := false
        
        for _, line in lines {
            line := Trim(line)
            if (line == "Artery") {
                inArterySection := true
                continue
            }
            if (line == "Score") {
                inScoreSection := true
                continue
            }
            
            if (inArterySection && !inScoreSection) {
                if (line != "" && !IsExcludedLine(line)) {
                    arteries.Push(line)
                }
            }
            if (inScoreSection) {
                if (RegExMatch(line, "^[\d.]+$")) {
                    scores.Push(line + 0) ; Convert to number
                }
            }
        }
        
        ; Process scores and map to artery names
        Loop, % Min(arteries.Length(), scores.Length())
        {
            arteryName := arteries[A_Index]
            score := Round(scores[A_Index]) ; Round to nearest integer
            mappedName := FindBestMatch(arteryName, g_arteryNames)
            
            if (mappedName && mappedName != "Total") {
                scoreData[mappedName] := score
            } else if (mappedName == "Total") {
                ocrTotal := score
            }
        }
    }

    ; Ensure main arteries are always present
    mainArteries := ["Left Main", "Left Anterior Descending", "Left Circumflex", "Right Coronary"]
    for _, artery in mainArteries {
        if (!scoreData.HasKey(artery)) {
            scoreData[artery] := 0
        }
    }

    ; Calculate total
    calculatedTotal := 0
    for _, score in scoreData {
        calculatedTotal += score
    }

    ; Check if the calculated total is significantly different from OCR total
    warningMessage := ""
    if (ocrTotal > 0 && Abs(calculatedTotal - ocrTotal) > 2) {
        warningMessage := "Warning: Calculated total (" . calculatedTotal . ") differs from OCR total (" . ocrTotal . ") by more than 2 Agatston units. Please review the scores."
    }

    formattedReport := FormatCalciumScoreReport(scoreData, calculatedTotal)
    return {report: formattedReport, warning: warningMessage}
}

IsLikelyRegionAgatston(headerLine) {
    return (LevenshteinDistance(headerLine, "Region Agatston Volume (mm3) Mass (mg)") <= 5)
}

ParseFlexibleRegionAgatston(lines) {
    scoreData := {}
    ocrTotal := 0

    for index, line in lines {
        if (index == 1) ; Skip header
            continue
        
        parts := StrSplit(Trim(line), A_Space)
        if (parts.Length() >= 2) {
            arteryName := ""
            score := 0
            scoreIndex := 0
            
            ; Find the first valid artery name and the corresponding score
            for i, part in parts {
                if (arteryName == "" && FindBestMatch(part, g_arteryNames)) {
                    arteryName := part
                } else if (arteryName != "" && IsNumeric(part)) {
                    score := Ceil(part + 0)
                    scoreIndex := i
                    break
                }
            }
            
            ; If no score found after artery name, check the column right after artery name
            if (score == 0 && scoreIndex == 0 && arteryName != "") {
                arteryIndex := 0
                for i, part in parts {
                    if (part == arteryName) {
                        arteryIndex := i
                        break
                    }
                }
                if (arteryIndex > 0 && parts.Length() > arteryIndex) {
                    potentialScore := parts[arteryIndex + 1]
                    if (IsNumeric(potentialScore)) {
                        score := Ceil(potentialScore + 0)
                    }
                }
            }
            
            mappedName := FindBestMatch(arteryName, g_arteryNames)
            if (mappedName && mappedName != "Total") {
                scoreData[mappedName] := score
            } else if (mappedName == "Total" || LevenshteinDistance(arteryName, "Total") <= 2) {
                ocrTotal := score
            }
        }
    }

    ; Ensure main arteries are always present
    mainArteries := ["Left Main", "Left Anterior Descending", "Left Circumflex", "Right Coronary", "Posterior Descending Artery"]
    for _, artery in mainArteries {
        if (!scoreData.HasKey(artery)) {
            scoreData[artery] := 0
        }
    }

    ; Calculate total
    calculatedTotal := 0
    for _, score in scoreData {
        calculatedTotal += score
    }

    ; Use OCR total if available and different from calculated total
    if (ocrTotal > 0 && ocrTotal != calculatedTotal) {
        calculatedTotal := ocrTotal
    }

    ; Check if the calculated total is significantly different from OCR total
    warningMessage := ""
    if (ocrTotal > 0 && Abs(calculatedTotal - ocrTotal) > 2) {
        warningMessage := "Warning: Calculated total (" . calculatedTotal . ") differs from OCR total (" . ocrTotal . ") by more than 2 Agatston units. Please review the scores."
    }

    formattedReport := FormatCalciumScoreReport(scoreData, calculatedTotal)
    return {report: formattedReport, warning: warningMessage}
}

ParseRegionAgatston(lines) {
    scoreData := {}
    ocrTotal := 0

    for index, line in lines {
        if (index == 1) ; Skip header
            continue
        
        parts := StrSplit(Trim(line), A_Space)
        if (parts.Length() >= 2) {
            arteryName := parts[1]
            score := Ceil(parts[2] + 0) ; Agatston score, rounded up
            
            mappedName := FindBestMatch(arteryName, g_arteryNames)
            if (mappedName && mappedName != "Total") {
                scoreData[mappedName] := score
            } else if (mappedName == "Total") {
                ocrTotal := score
            }
        }
    }

    ; Ensure main arteries are always present
    mainArteries := ["Left Main", "Left Anterior Descending", "Left Circumflex", "Right Coronary"]
    for _, artery in mainArteries {
        if (!scoreData.HasKey(artery)) {
            scoreData[artery] := 0
        }
    }

    ; Calculate total
    calculatedTotal := 0
    for _, score in scoreData {
        calculatedTotal += score
    }

    ; Check if the calculated total is significantly different from OCR total
    warningMessage := ""
    if (ocrTotal > 0 && Abs(calculatedTotal - ocrTotal) > 2) {
        warningMessage := "Warning: Calculated total (" . calculatedTotal . ") differs from OCR total (" . ocrTotal . ") by more than 2 Agatston units. Please review the scores."
    }

    formattedReport := FormatCalciumScoreReport(scoreData, calculatedTotal)
    return {report: formattedReport, warning: warningMessage}
}

FindBestMatch(name, dictionary) {
    global g_levenshteinThreshold
    bestMatch := ""
    minDistance := g_levenshteinThreshold + 1
    StringLower, nameLower, name
    for key, value in dictionary {
        distance := LevenshteinDistance(nameLower, key)
        if (distance < minDistance) {
            minDistance := distance
            bestMatch := value
        }
    }
    return (minDistance <= g_levenshteinThreshold) ? bestMatch : ""
}


FormatCalciumScoreReport(scoreData, calculatedTotal) {
    result := "Calcium Score Report:`n`n"
    
    ; Always display main arteries first
    mainArteries := ["Left Main", "Left Anterior Descending", "Left Circumflex", "Right Coronary"]
    for _, artery in mainArteries {
        result .= artery . ": " . scoreData[artery] . "`n"
    }
    
    ; Display any additional arteries (including PDA)
    for artery, score in scoreData {
        if (!HasValue(mainArteries, artery) && artery != "Total") {
            result .= artery . ": " . score . "`n"
        }
    }
    
    result .= "`nYOUR CORONARY ARTERY CALCIUM SCORE: " . calculatedTotal
    
    return result
}


FindKeyForValue(dictionary, value) {
    for key, val in dictionary {
        if (val == value)
            return key
    }
    return ""
}

HasValue(arr, value) {
    for _, v in arr {
        if (v == value)
            return true
    }
    return false
}

ShowCalciumScore(result) {
    ShowResult(result.report)
    if (result.warning != "") {
        MsgBox, % result.warning
    }
}

IsNumeric(str) {
    return RegExMatch(str, "^\d+(\.\d+)?$")
}

; Example usage:
^!c::  ; Ctrl+Alt+C hotkey
{
    result := CaptureCalciumScore()
    if (result)
        ShowCalciumScore(result)
    else
        MsgBox, No calcium scores were captured or the capture was cancelled.
}

IsExcludedLine(line) {
    excludedLines := ["Lesions", "Son DPWN EF", "Volume/mm", "Equiv.Mass/mg", "Volume (mm3)", "Mass (mg)"]
    for _, excludedLine in excludedLines {
        if (line == excludedLine)
            return true
    }
    return false
}

; ============= END

; =========== FLEISCHNER 2017
class Nodule {
    __New(nString) {
        this.Description := nString
        this.Composition := ""
        this.Calcified := false
        this.mString := ""
        this.Units := ""
        this.Measurements := []
        this.HighRisk := false
        this.Location := ""
        this.Multiplicity := "single"
        this.Perifissural := false
        this.Morphology := ""
        this.GlobalHighRisk := false

        if (!this.ContainsNoduleReference(nString))
            throw Exception("No nodule reference", -1)
        
        this.ParseNoduleProperties(nString)
        this.ExtractMeasurements(nString)

        ; Handle cases where only the largest measurement is given for multiple nodules
        if (InStr(nString, "up to") and InStr(nString, "multiple")) {
            this.Multiplicity := "multiple"
        }
    }

    ContainsNoduleReference(nString) {
        noduleTerms := ["nodule", "nodules", "mass", "masses", "opacity", "opacities", "lesion", "lesions", "micronodule", "micronodules"]
        for _, term in noduleTerms {
            if (InStr(nString, term))
                return true
        }
        return false
    }

    ParseNoduleProperties(nString) {
        words := StrSplit(nString, A_Space)
        for index, word in words {
            if (this.IsMultiplicityWord(word))
                this.Multiplicity := "multiple"
            
            if (this.FuzzyMatch(word, "solid")) {
                if (this.Composition = "")
                    this.Composition := "solid"
                else if (this.Composition = "ground glass")
                    this.Composition := "part solid"
            }
            
            if (this.FuzzyMatch(word, "ground") and index < words.Length() and this.FuzzyMatch(words[index+1], "glass")) {
                if (this.Composition = "")
                    this.Composition := "ground glass"
                else if (this.Composition = "solid")
                    this.Composition := "part solid"
            }
            
            if (this.FuzzyMatch(word, "groundglass")) {
                if (this.Composition = "")
                    this.Composition := "ground glass"
                else if (this.Composition = "solid")
                    this.Composition := "part solid"
            }
            
            if ((this.FuzzyMatch(word, "part") and index < words.Length() and this.FuzzyMatch(words[index+1], "solid")) 
                or this.FuzzyMatch(word, "part-solid") or this.FuzzyMatch(word, "partsolid")) {
                this.Composition := "part solid"
            }
            
            if (this.FuzzyMatch(word, "calcified") or this.FuzzyMatch(word, "calcification") or this.FuzzyMatch(word, "calcifications"))
                this.Calcified := true
            
            if (this.FuzzyMatch(word, "noncalcified") or this.FuzzyMatch(word, "non-calcified"))
                this.Calcified := false

            if (this.FuzzyMatch(word, "emphysema") or this.FuzzyMatch(word, "fibrosis"))
                this.GlobalHighRisk := true

            lobeKeywords := ["upper", "middle", "lower", "lingula", "apical", "basal"]
            for _, keyword in lobeKeywords {
                if (this.FuzzyMatch(word, keyword)) {
                    this.Location .= " " . word
                }
            }
            if (this.FuzzyMatch(word, "right") or this.FuzzyMatch(word, "left")) {
                this.Location .= " " . word
            }

            if (this.FuzzyMatch(word, "perifissural") or this.FuzzyMatch(word, "fissure"))
                this.Perifissural := true

            morphologyKeywords := ["spiculated", "lobulated", "irregular", "smooth"]
            for _, keyword in morphologyKeywords {
                if (this.FuzzyMatch(word, keyword)) {
                    this.Morphology := keyword
                    break
                }
            }
        }
        this.Location := Trim(this.Location)

        if (InStr(nString, "multiple") or InStr(nString, "several") or InStr(nString, "numerous"))
            this.Multiplicity := "multiple"
    }

    IsMultiplicityWord(word) {
        multiplicityWords := ["nodules", "multiple", "several", "few", "numerous", "masses", "opacities", "lesions", "micronodules"]
        for _, w in multiplicityWords {
            if (this.FuzzyMatch(word, w))
                return true
        }
        return false
    }

    FuzzyMatch(word1, word2, threshold := 2) {
        return (LevenshteinDistance(word1, word2) <= threshold)
    }

	ExtractMeasurements(nString) {
	needle := "i)(?:up to|approximately|about|~)?\s*((?:\d*\.)?\d+)\s*(?:x\s*((?:\d*\.)?\d+))?\s*(?:x\s*((?:\d*\.)?\d+))?\s*([cm]m)"
		if (RegExMatch(nString, needle, match)) {
			this.mString := match
			this.Units := match4
			Loop, 3 {
				if (match%A_Index% != "")
					this.Measurements.Push(match%A_Index%)
			}
		}
		else {
			needle := "i)(?:up to|approximately|about|~)?\s*((?:\d*\.)?\d+)\s*([cm]m)"
			if (RegExMatch(nString, needle, match)) {
				this.mString := match
				this.Units := match2
				this.Measurements.Push(match1)
			}
			else if (InStr(nString, "micronodule") or InStr(nString, "micronodules")) {
				; Default size for micronodules when no measurement is given
				this.mString := "5 mm"
					this.Units := "mm"
						this.Measurements.Push(5)
				}
				else {
					; If no specific measurement found, look for any number followed by units
					needle := "i)((?:\d*\.)?\d+)\s*([cm]m)"
					if (RegExMatch(nString, needle, match)) {
						this.mString := match
						this.Units := match2
						this.Measurements.Push(match1)
					}
					else {
					throw Exception("No measurements found", -2)
					}
				}
			}
		; Store original measurements
		this.OriginalMeasurements := this.Measurements.Clone()
		this.OriginalUnits := this.Units
	}

    Size() {
        if (this.Measurements.Length() = 0)
            return 0
        
        totalSize := 0
        for _, measurement in this.Measurements {
            totalSize += measurement
        }
        averageSize := totalSize / this.Measurements.Length()

        ; Convert to mm if necessary
        if (this.Units = "cm")
            return averageSize * 10  ; Convert cm to mm
        return averageSize  ; Already in mm
    }
	
    UpdateMString() {
        if (this.OriginalMeasurements.Length() = 0)
            return

        sizes := []
        for _, measurement in this.OriginalMeasurements {
            sizes.Push(Format("{:.1f}", measurement))
        }
        
        this.mString := Join(sizes, " x ") . " " . this.OriginalUnits
    }

    Category() {
        s := this.Size()  ; s is now always in mm
        
        if (this.Composition == "solid" or this.Composition == "") {
            if (this.Multiplicity == "multiple") {
                if (s < 6)
                    return "multiple_solid_small"
                else if (s >= 6 and s <= 8)
                    return "multiple_solid_medium"
                else
                    return "multiple_solid_large"
            } else {
                if (s < 6)
                    return "single_solid_small"
                else if (s >= 6 and s <= 8)
                    return "single_solid_medium"
                else
                    return "single_solid_large"
            }
        } else if (this.Composition == "ground glass") {
            if (this.Multiplicity == "multiple") {
                if (s < 6)
                    return "multiple_subsolid_small"
                else
                    return "multiple_subsolid_large"
            } else {
                if (s < 6)
                    return "single_gg_small"
                else
                    return "single_gg_large"
            }
        } else if (this.Composition == "part solid") {
            if (this.Multiplicity == "multiple") {
                if (s < 6)
                    return "multiple_subsolid_small"
                else
                    return "multiple_subsolid_large"
            } else {
                if (s < 6)
                    return "single_ps_small"
                else
                    return "single_ps_large"
            }
        }
    }

    Recommendation() {
        category := this.Category()
        risk := this.HighRisk ? "high" : "low"
        
        if (recommendations.HasKey(category)) {
            if (IsObject(recommendations[category]))
                return recommendations[category][risk]
            else
                return recommendations[category]
        }
        return "Unable to determine recommendation."
    }
}

InitializeRecommendations() {
    ObjRawSet(recommendations, "single_solid_small", {low: "No routine follow-up.", high: "Optional CT at 12 months."})
    ObjRawSet(recommendations, "single_solid_medium", {low: "CT at 6-12 months, then consider CT at 18-24 months.", high: "CT at 6-12 months, then CT at 18-24 months."})
    ObjRawSet(recommendations, "single_solid_large", {low: "Consider CT at 3 months, PET/CT, or tissue sampling.", high: "Consider CT at 3 months, PET/CT, or tissue sampling."})
    ObjRawSet(recommendations, "multiple_solid_small", {low: "No routine follow-up.", high: "Optional CT at 12 months."})
    ObjRawSet(recommendations, "multiple_solid_medium", {low: "CT at 3-6 months, then consider CT at 18-24 months.", high: "CT at 3-6 months, then at 18-24 months."})
    ObjRawSet(recommendations, "multiple_solid_large", {low: "CT at 3-6 months, then consider CT at 18-24 months.", high: "CT at 3-6 months, then at 18-24 months."})
    ObjRawSet(recommendations, "single_gg_small", "No routine follow-up.")
    ObjRawSet(recommendations, "single_gg_large", "CT at 6-12 months to confirm persistence, then every 2 years until 5 years.")
    ObjRawSet(recommendations, "single_ps_small", "No routine follow-up.")
    ObjRawSet(recommendations, "single_ps_large", "CT at 3-6 months to confirm persistence. If unchanged and solid component remains <6mm, annual CT should be performed for 5 years.")
    ObjRawSet(recommendations, "multiple_subsolid_small", "CT at 3-6 months. If stable, consider CT at 2 and 4 years.")
	ObjRawSet(recommendations, "multiple_subsolid_large", "CT at 3-6 months. Subsequent management based on the most suspicious nodule(s).")
}

Join(arr, sep) {
    result := ""
    for index, element in arr {
        if (index > 1)
            result .= sep
        result .= element
    }
    return result
}

PreprocessTextNodules(text) {
    sentences := []
    lines := StrSplit(text, "`n", "`r")
    for _, line in lines {
        ; Remove bullet points
        line := RegExReplace(line, "^\s*[\*\-•]\s*", "")
        
        ; Split by periods, but be careful with measurements
        parts := StrSplit(line, ".")
        for index, part in parts {
            part := Trim(part)
            ; Check if this part is a continuation of a measurement
            if (index > 1 && RegExMatch(parts[index-1], "\d+$") && RegExMatch(part, "^\d+"))
                sentences[sentences.Length()] .= "." . part
            else if (part != "")
                sentences.Push(part)
        }
    }
    return sentences
}

ProcessNodules(text) {
    result := ""
    sentences := PreprocessTextNodules(text)
    nodules := []
    globalHighRisk := CheckHighRiskConditions(text)
    isMultiple := false

    for _, sentence in sentences {
        multipleNodulesProbability := CalculateMultipleNodulesProbability(sentence)
        isMultiple := (multipleNodulesProbability >= 0.45) or isMultiple  ; Adjusted threshold and accumulate

        noduleDescriptions := SplitNoduleDescriptions(sentence)
        for _, description in noduleDescriptions {
            try {
                nodule := new Nodule(description)
                nodule.GlobalHighRisk := globalHighRisk
                nodule.Multiplicity := isMultiple ? "multiple" : "single"
                nodules.Push(nodule)
            } catch e {
                ; Skip invalid nodule descriptions
            }
        }
    }

    if (nodules.Length() = 0) {
        ; Check for general mention of micronodules or multiple nodules
        if (InStr(text, "micronodule") or InStr(text, "micronodules") or (InStr(text, "multiple") and InStr(text, "nodule"))) {
            nodule := new Nodule("Multiple nodules measuring up to 5 mm")
            nodule.GlobalHighRisk := globalHighRisk
            nodule.Multiplicity := "multiple"
            nodules.Push(nodule)
        } else {
            result := "Error: No valid nodules found."
            return result
        }
    }

    ; Find the most significant nodule
    mostSignificantNodule := nodules[1]
    for _, nodule in nodules {
        nodule.UpdateMString()  ; Update mString for each nodule
        if (nodule.Size() > mostSignificantNodule.Size() or (nodule.Size() = mostSignificantNodule.Size() and IsMoreSignificant(nodule, mostSignificantNodule))) {
            mostSignificantNodule := nodule
        }
    }

    ; Ensure multiplicity is set correctly
    if (nodules.Length() > 1 or isMultiple or InStr(text, "multiple") or InStr(text, "micronodules")) {
        mostSignificantNodule.Multiplicity := "multiple"
    }

    ; Generate result string
    result := text . "`n`n"  ; Preserve input text at the top
    result .= "Fleischner 2017 assessment:`n"
    if (mostSignificantNodule.Multiplicity == "multiple") {
        result .= "Multiple pulmonary nodules are present. The most significant nodule forming the basis of follow-up:`n"
    } else {
        result .= "A solitary pulmonary nodule is described forming the basis of follow-up:`n"
    }
    esult .= "- Location: " . (mostSignificantNodule.Location ? mostSignificantNodule.Location : "Not specified") . "`n"
    result .= "- Extracted (or inferred) Size: " . mostSignificantNodule.mString . "`n"
    result .= "- Fleischner Size (mm): " . Format("{:.1f} mm", mostSignificantNodule.Size()) . "`n"
    result .= "- Composition (or inferred): " . (mostSignificantNodule.Composition ? mostSignificantNodule.Composition : "solid") . "`n"
    if (mostSignificantNodule.Calcified)
        result .= "- Calcification: Present`n"
    if (mostSignificantNodule.Morphology)
        result .= "- Morphology: " . mostSignificantNodule.Morphology . "`n"
		
    currentDate := A_Now

    result .= "`nFLEISCHNER SOCIETY RECOMMENDATION:`n"
    if ((globalHighRisk || mostSignificantNodule.Morphology == "spiculated") && !mostSignificantNodule.Calcified) {
        mostSignificantNodule.HighRisk := true
        recommendation := mostSignificantNodule.Recommendation()
        result .= AddFollowUpDates(recommendation, currentDate)
    } else {
        if (!mostSignificantNodule.Calcified && mostSignificantNodule.Composition == "solid" || mostSignificantNodule.Composition == "") {
            lowRiskRec := mostSignificantNodule.Recommendation()
            result .= "For low-risk patients: " . AddFollowUpDates(lowRiskRec, currentDate) . "`n"
            mostSignificantNodule.HighRisk := true
            highRiskRec := mostSignificantNodule.Recommendation()
            result .= "`nFor high-risk patients: " . AddFollowUpDates(highRiskRec, currentDate) . "`n"
		} else if (!mostSignificantNodule.Calcified && (mostSignificantNodule.Composition == "part solid" || mostSignificantNodule.Composition == "ground glass")) {
			recommendation := mostSignificantNodule.Recommendation()
			result .= AddFollowUpDates(recommendation, currentDate)
		} else {
            result .= "Incidental calcified nodules do not typically require routine follow up.`n"
        }
    }

    if (globalHighRisk || mostSignificantNodule.Morphology == "spiculated")
        result .= "`n`nNote: Patient has risk factors or morphologic characteristics that may increase the risk of lung cancer."

    if (ShowCitations)
        result .= "`n`nCitation: MacMahon H, Naidich DP, Goo JM, et al. Guidelines for Management of Incidental Pulmonary Nodules Detected on CT Images: From the Fleischner Society 2017. Radiology. 2017;284(1):228-243. doi:10.1148/radiol.2017161659"
   
    return result
}

IsMoreSignificant(nodule1, nodule2) {
    ; Order of significance: solid > part solid > ground glass
    compositionOrder := {"solid": 3, "part solid": 2, "ground glass": 1}
    return compositionOrder[nodule1.Composition] > compositionOrder[nodule2.Composition]
}

SplitNoduleDescriptions(text) {
    descriptions := []
    ; Modified regex to capture nodule descriptions more accurately, including cases where size appears before location and "nodule"
    needle := "i)(?:(?:\d+(?:\.\d+)?\s*(?:x\s*\d+(?:\.\d+)?)*\s*(?:mm|cm))\s*(?:[a-z\s]+\s+)?(?:pulmonary\s+)?(?:nodule|mass|opacity|lesion)s?|(?:(?:multiple\s+)?(?:nodule|mass|opacity|lesion)s?(?:\s+(?:up to|approximately|about|~)?)?\s*(?:\d+(?:\.\d+)?\s*(?:x\s*\d+(?:\.\d+)?)*\s*(?:mm|cm))?|(?:(?:up to|approximately|about|~)?\s*\d+(?:\.\d+)?\s*(?:x\s*\d+(?:\.\d+)?)*\s*(?:mm|cm))?\s*(?:(?:solid|ground glass|part[- ]solid|calcified|noncalcified|spiculated|lobulated|irregular|smooth)?\s*(?:pulmonary\s+)?(?:nodule|mass|opacity|lesion)s?)))(?:\s*\([^)]+\))?"
    
    ; Split the text into lines
    lines := StrSplit(text, "`n", "`r")
    for _, line in lines {
        pos := 1
        while (pos := RegExMatch(line, needle, match, pos)) {
            descriptions.Push(Trim(match))
            pos += StrLen(match)
        }
    }
    
    ; If no specific nodule descriptions found, check for general mentions of nodules
    if (descriptions.Length() = 0) {
        if (RegExMatch(text, "i)(?:multiple\s+)?nodules?.*(?:measure|up to).*(\d+(?:\.\d+)?)\s*(?:x\s*\d+(?:\.\d+)?)*\s*(mm|cm)", match)) {
            descriptions.Push(match)
        } else if (InStr(text, "micronodule") or InStr(text, "micronodules")) {
            descriptions.Push("Multiple micronodules")
        } else if (InStr(text, "multiple") and InStr(text, "nodule")) {
            descriptions.Push("Multiple nodules")
        }
    }
    
    ; If still no descriptions found, return the entire text
    if (descriptions.Length() = 0) {
        descriptions.Push(text)
    }
    
    return descriptions
}

AddFollowUpDates(recommendation, currentDate) {
    followUpPeriods := []
    
    if (InStr(recommendation, "3 months"))
        followUpPeriods.Push({min: 90, max: 90, text: "3 months"})
	if (InStr(recommendation, "3-6 months"))
        followUpPeriods.Push({min: 90, max: 183, text: "3-6 months"})
    if (InStr(recommendation, "6-12 months"))
        followUpPeriods.Push({min: 180, max: 365, text: "6-12 months"})
    if (InStr(recommendation, "18-24 months"))
        followUpPeriods.Push({min: 540, max: 730, text: "18-24 months"})
	if (InStr(recommendation, "2 years until 5 years"))
        followUpPeriods.Push({min: 1095, max: 1825, text: "2 years until 5 years"})
    
    if (followUpPeriods.Length() > 0) {
        recommendation .= "`nFollow-up dates:"
        for _, period in followUpPeriods {
            minDate := DateCalc(currentDate, period.min)
            maxDate := DateCalc(currentDate, period.max)
            FormatTime, formattedMinDate, %minDate%, MMMM yyyy
            FormatTime, formattedMaxDate, %maxDate%, MMMM yyyy
			FormatTime, formattedCurrentDate, %currentDate%, MMMM yyyy
            if (formattedMinDate != formattedMaxDate)
                recommendation .= " " . period.text . " is " . formattedMinDate . " to " . formattedMaxDate . " from " . formattedCurrentDate . ". "
            else
                recommendation .= " " . period.text . " is " . formattedMinDate . " from " . formattedCurrentDate . ". "
        }
    }
    
    return recommendation
}

CheckHighRiskConditions(text) {
    StringLower text, text ; Convert to lowercase
    highRiskTerms := ["emphysema", "fibrosis", "fibrotic", "emphysematous"]
    
    for _, term in highRiskTerms {
        words := StrSplit(text, A_Space)
        for _, word in words {
            if (FuzzyMatchGlobal(word, term))
                return true
        }
    }
    return false
}

FuzzyMatchGlobal(word1, word2, threshold := 2) {
    return (LevenshteinDistance(word1, word2) <= threshold)
}

CalculateMultipleNodulesProbability(text) {
    probability := 0
    
    ; Check for explicit mentions of multiple nodules
    if (InStr(text, "nodules") or InStr(text, "multiple") or InStr(text, "several") or InStr(text, "few"))
        probability += 0.7
    
    ; Check for scattered micronodules
    if (InStr(text, "scattered") and InStr(text, "micronodules"))
        probability += 0.6
    
    ; Check for multiple measurements
    measurementCount := 0
    pos := 1
    while (pos := RegExMatch(text, "i)(\d+(?:\.\d+)?\s*(mm|cm))", match, pos + StrLen(match)))
        measurementCount++
    
    if (measurementCount > 1)
        probability += 0.5  ; Increased weight for multiple measurements
    
    ; Check for multiple locations
    lobeKeywords := ["upper", "middle", "lower", "lingula", "apical", "basal", "right", "left"]
    uniqueLocations := {}
    for _, keyword in lobeKeywords {
        if (InStr(text, keyword))
            uniqueLocations[keyword] := true
    }
    
    if (uniqueLocations.Count() > 1)
        probability += 0.3  ; Increased weight for multiple locations
    
    ; Count number of times "nodule" (singular) is mentioned
    noduleCount := 0
    pos := 1
    while (pos := InStr(text, "nodule", false, pos))
    {
        noduleCount++
        pos += 6  ; length of "nodule"
    }
    
    if (noduleCount > 1)
        probability += 0.4  ; Add probability if "nodule" is mentioned multiple times
    
    return (probability > 1) ? 1 : probability
}
;==============end flesichner

;==============NASCET calculator
CalculateNASCET(input) {
    global ShowCitations

    ; Replace newlines with spaces to handle multi-line input
    rawinput := input 
	input := RegExReplace(input, "\r?\n", " ")

    ; Updated regular expression to match various reporting styles and terminologies
    RegExNeedle := "i)(?:"
                 . "(?:(?:distal|normal|proximal|wider|patent)(?:\s+(?:ica|carotid|lumen|segment|caliber|diameter))?(?:\s*measures\s+|\s*:?\s*)(\d+(?:\.\d+)?)\s*(mm|cm))"
                 . "|"
                 . "(?:(?:ica|internal\s+carotid(?:\s+artery)?|carotid|lumen)(?:\s*measures\s+|\s*:?\s*)(\d+(?:\.\d+)?)\s*(mm|cm))"
                 . ")"
                 . ".*?"  ; This allows for any text between the two measurements
                 . "(?:"
                 . "(?:(?:stenosis|narrowed|stenotic|residual|focal|narrowest|tightest)(?:\s+(?:ica|lumen|segment|diameter|region))?(?:\s*measures\s+|\s*:?\s*)(\d+(?:\.\d+)?)\s*(mm|cm))"
                 . "|"
                 . "(?:(?:narrows|narrowing|stenosis)(?:\s+to)?(?:\s*measures\s+|\s*:?\s*)(\d+(?:\.\d+)?)\s*(mm|cm))"
                 . ")"

    distal := ""
    distalUnit := ""
    stenosis := ""
    stenosisUnit := ""

    if (RegExMatch(input, RegExNeedle, match)) {
        distal := match1 ? match1 : match3
        distalUnit := match2 ? match2 : match4
        stenosis := match5 ? match5 : match7
        stenosisUnit := match6 ? match6 : match8
    } else {
        ; If no match found, try reversing the order
        RegExNeedle := "i)(?:"
                     . "(?:(?:stenosis|narrowed|stenotic|residual|focal|narrowest|tightest)(?:\s+(?:ica|lumen|segment|diameter|region))?(?:\s*measures\s+|\s*:?\s*)(\d+(?:\.\d+)?)\s*(mm|cm))"
                     . "|"
                     . "(?:(?:narrows|narrowing|stenosis)(?:\s+to)?(?:\s*measures\s+|\s*:?\s*)(\d+(?:\.\d+)?)\s*(mm|cm))"
                     . ")"
                     . ".*?"  ; This allows for any text between the two measurements
                     . "(?:"
                     . "(?:(?:distal|normal|proximal|wider|patent)(?:\s+(?:ica|carotid|lumen|segment|caliber|diameter))?(?:\s*measures\s+|\s*:?\s*)(\d+(?:\.\d+)?)\s*(mm|cm))"
                     . "|"
                     . "(?:(?:ica|internal\s+carotid(?:\s+artery)?|carotid|lumen)(?:\s*measures\s+|\s*:?\s*)(\d+(?:\.\d+)?)\s*(mm|cm))"
                     . ")"

        if (RegExMatch(input, RegExNeedle, match)) {
            stenosis := match1 ? match1 : match3
            stenosisUnit := match2 ? match2 : match4
            distal := match5 ? match5 : match7
            distalUnit := match6 ? match6 : match8
        }
    }

    ; If we still don't have both measurements, try to extract any two numbers
    if (!distal || !stenosis) {
        numbers := []
        units := []
        RegExNeedle := "i)(\d+(?:\.\d+)?)\s*(mm|cm)"
        pos := 1
        while (pos := RegExMatch(input, RegExNeedle, match, pos)) {
            numbers.Push(match1)
            units.Push(match2)
            pos += StrLen(match)
        }

        ; If we found exactly two numbers, use them
        if (numbers.Length() == 2) {
            ; Convert both to mm for comparison
            num1 := (units[1] == "cm") ? numbers[1] * 10 : numbers[1]
            num2 := (units[2] == "cm") ? numbers[2] * 10 : numbers[2]

            if (num1 > num2) {
                distal := numbers[1]
                distalUnit := units[1]
                stenosis := numbers[2]
                stenosisUnit := units[2]
            } else {
                distal := numbers[2]
                distalUnit := units[2]
                stenosis := numbers[1]
                stenosisUnit := units[1]
            }
        }
    }

    ; Check if we have both measurements
    if (!distal || !stenosis) {
        return "Unable to extract both measurements. Please check the input format. Sample: The stenosis measures 2 mm. The distal ICA measures 6 mm."
    }

    ; Convert to mm if necessary
    distal := (distalUnit = "cm") ? distal * 10 : distal
    stenosis := (stenosisUnit = "cm") ? stenosis * 10 : stenosis

    ; Calculate NASCET
    nascetValue := (distal - stenosis) / distal * 100

    ; Round to one decimal place
    nascetValue := Round(nascetValue, 1)

    result := rawinput . "`n`n"
	result .= "NASCET Calculation:`n"
    result .= "Distal ICA diameter: " . Round(distal, 1) . " mm`n"
    result .= "Stenosis diameter: " . Round(stenosis, 1) . " mm`n"
    result .= "NASCET value: " . nascetValue . "%`n`n"

    ; Add interpretation
    if (nascetValue < 50)
        result .= "Interpretation: Mild stenosis (<50%)"
    else if (nascetValue >= 50 && nascetValue < 70)
        result .= "Interpretation: Moderate stenosis (50-69%)"
    else
        result .= "Interpretation: Severe stenosis (>=70%)"

    if (ShowCitations) {
        result .= "`n`nCitation: North American Symptomatic Carotid Endarterectomy Trial Collaborators. Beneficial effect of carotid endarterectomy in symptomatic patients with high-grade carotid stenosis. N Engl J Med. 1991;325(7):445-453. doi:10.1056/NEJM199108153250701"
    }

    return result
}
; ==== END NASCET

^!p::ShowPreferences()

^!u::  ; Ctrl+Alt+U hotkey
{
    result := ParseUltrasoundMeasurements()
    if (result && result.measurements.Count() > 0)
        ShowUltrasoundMeasurements(result)
    else
        MsgBox, No measurements were captured or the capture was cancelled.
}

Exit:
    Gdip_Shutdown(pToken)
ExitApp

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
