; ============================================================
; RightClick.ahk -- Radiology Right Click v2 (AHK v2 port)
; ============================================================
; DISCLAIMER -- FOR ENTERTAINMENT, RESEARCH, AND EDUCATIONAL
; USE ONLY. This software is not a medical device, has not
; been validated for clinical use, and must not be relied upon
; for clinical decision-making. All calculations must be
; independently verified by the user before being incorporated
; into any clinical report or used to inform patient care.
; See LICENSE and README.md for full terms.
; ------------------------------------------------------------
; Right-click inside a supported reporting application to open
; a context menu of measurement and analysis helpers. The menu
; works on whatever text is currently selected.
;
; Supported applications are configured in Prefs.activation
; .targetApps (editable via Preferences -> Target Apps...).
; The hardcoded list in TargetApps() below is the fallback when
; that pref is missing.
; ============================================================

#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, OutputDebug
SetWorkingDir A_ScriptDir
CoordMode "Mouse", "Screen"

A_IconTip := "RightClick -- radiology helpers"
if FileExist(A_ScriptDir "\RightClick.ico")
    TraySetIcon(A_ScriptDir "\RightClick.ico")

; ---- foundation
#Include lib\Prefs.ahk
#Include lib\Debug.ahk
#Include lib\Modern.ahk
#Include lib\Util.ahk
#Include lib\CalcResult.ahk
#Include lib\UI.ahk
#Include lib\FormGui.ahk

; ---- text-parse calculators (legacy v1)
#Include lib\calc\Volumes.ahk
#Include lib\calc\PSADensity.ahk
#Include lib\calc\PregnancyDates.ahk
#Include lib\calc\Numerics.ahk
#Include lib\calc\AdrenalWashout.ahk
#Include lib\calc\ThymusChemicalShift.ahk
#Include lib\calc\HepaticSteatosis.ahk
#Include lib\calc\LiverIron.ahk
#Include lib\calc\CalciumScore.ahk
#Include lib\calc\NASCET.ahk
#Include lib\calc\NoduleSizes.ahk
#Include lib\calc\ContrastPremed.ahk
#Include lib\calc\Fleischner.ahk

; ---- form-based RADS calculators (v2.1)
#Include lib\calc\Bosniak.ahk
#Include lib\calc\GBPolyp.ahk
#Include lib\calc\IncidentalAdrenal.ahk
#Include lib\calc\IncidentalThyroid.ahk
#Include lib\calc\KyotoIPMN.ahk
#Include lib\calc\LIRADS.ahk
#Include lib\calc\LungRADS.ahk
#Include lib\calc\ORADSMRI.ahk
#Include lib\calc\ORADSUS.ahk
#Include lib\calc\PIRADS.ahk
#Include lib\calc\TIRADS.ahk
#Include lib\calc\USLIRADS.ahk

; ---- shell
#Include lib\References.ahk
#Include lib\PreferencesWindow.ahk
#Include lib\Menu.ahk

Prefs.Load()
SetAppDarkMode(Prefs.Get("display", "darkMode", false))
OnExit((*) => Prefs.Flush())   ; persist any pending frequency increments

; ------------------------------------------------------------
; Hotkey: right-click in a target window, optionally gated by a
; modifier (Ctrl / Alt / Shift) set in Preferences. Defined via
; the runtime Hotkey() API so the modifier can be changed
; without restarting the script.
; ------------------------------------------------------------
_RegisterActivationHotkeys() {
    HotIf(IsTargetWindowUnderCursor)
    Hotkey("RButton",  HandleRightClick, "Off")
    Hotkey("^RButton", HandleRightClick, "Off")
    Hotkey("!RButton", HandleRightClick, "Off")
    Hotkey("+RButton", HandleRightClick, "Off")
    HotIf()
}

ApplyActivationHotkey() {
    HotIf(IsTargetWindowUnderCursor)
    Hotkey("RButton",  "Off")
    Hotkey("^RButton", "Off")
    Hotkey("!RButton", "Off")
    Hotkey("+RButton", "Off")
    mod := Prefs.Get("activation", "modifier", "none")
    target := (mod = "ctrl")  ? "^RButton"
            : (mod = "alt")   ? "!RButton"
            : (mod = "shift") ? "+RButton"
            : "RButton"
    Hotkey(target, "On")
    HotIf()
}

_RegisterActivationHotkeys()
ApplyActivationHotkey()

HandleRightClick(*) {
    CoordMode "Mouse", "Screen"
    MouseGetPos(&ox, &oy, &hwnd)
    if (hwnd != WinExist("A")) {
        try {
            WinActivate("ahk_id " hwnd)
            ; Wait for activation (bounded) instead of a fixed 30 ms nap --
            ; heavyweight hosts (PowerScribe) can take longer to take focus,
            ; and the Ctrl+C capture that follows lands on the wrong window
            ; if we race ahead.
            WinWaitActive("ahk_id " hwnd, , 0.3)
            MouseMove(ox, oy, 0)
        }
    }
    ShowContextMenu()
}

; ------------------------------------------------------------
; Target windows -- read from Prefs ("activation.targetApps"), with the
; hardcoded default kept here as a fallback in case the pref is empty or
; the file failed to load.
; ------------------------------------------------------------
TargetApps() {
    apps := Prefs.Get("activation", "targetApps", "")
    if (apps is Array && apps.Length > 0)
        return apps
    return [
        Map("type", "class", "value", "Notepad"),
        Map("type", "exe",   "value", "notepad.exe"),
        Map("type", "class", "value", "PowerScribe"),
        Map("type", "exe",   "value", "PowerScribe.exe"),
        Map("type", "class", "value", "PowerScribe360"),
        Map("type", "exe",   "value", "Nuance.PowerScribe360.exe"),
        Map("type", "class", "value", "PowerScribe | Reporting")
    ]
}

; The `*` lets this serve both as a direct call (no args) and as a HotIf
; callback, where AutoHotkey passes the firing hotkey name as the first arg.
IsTargetWindowUnderCursor(*) {
    MouseGetPos(, , &hwnd)
    if (!hwnd)
        return false
    cls := ""
    exe := ""
    try cls := WinGetClass("ahk_id " hwnd)
    try exe := WinGetProcessName("ahk_id " hwnd)
    for app in TargetApps() {
        if !(app is Map) || !app.Has("type") || !app.Has("value")
            continue
        t := app["type"], v := app["value"]
        if (t = "class" && cls = v)
            return true
        if (t = "exe" && exe = v)
            return true
    }
    return false
}
