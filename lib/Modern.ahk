; ============================================================
; lib/Modern.ahk -- Modern Win10/11 visual helpers
; ------------------------------------------------------------
; Provides:
;   ModernFont()                 -> Segoe UI Variable (Win11) or Segoe UI
;   SetDarkTitleBar(hwnd, on)    -> immersive dark mode for the
;                                   title bar via dwmapi
;   SetRoundedCorners(hwnd, m)   -> Win11 rounded corners (no-op
;                                   on earlier builds)
;   ApplyModernChrome(gui, dark) -> convenience: applies both
;   ModernPalette(dark)          -> Map of palette colors
;   IsWin11()                    -> bool
; ============================================================

#Requires AutoHotkey v2.0

ModernFont() {
    return IsWin11() ? "Segoe UI Variable" : "Segoe UI"
}

IsWin11() {
    parts := StrSplit(A_OSVersion, ".")
    return (parts.Length >= 3 && (parts[3] + 0) >= 22000)
}

SetDarkTitleBar(hwnd, enabled := true) {
    static DWMWA_USE_IMMERSIVE_DARK_MODE := 20
    static DWMWA_USE_IMMERSIVE_DARK_MODE_PRE20H1 := 19
    val := enabled ? 1 : 0
    ; Try 20 first (1909+), fall back to 19 (older builds with the
    ; pre-public attribute number)
    try DllCall("dwmapi\DwmSetWindowAttribute"
        , "ptr", hwnd, "uint", DWMWA_USE_IMMERSIVE_DARK_MODE
        , "int*", val, "uint", 4)
    try DllCall("dwmapi\DwmSetWindowAttribute"
        , "ptr", hwnd, "uint", DWMWA_USE_IMMERSIVE_DARK_MODE_PRE20H1
        , "int*", val, "uint", 4)
}

SetRoundedCorners(hwnd, mode := 2) {
    ; 0 = default, 1 = donotround, 2 = round, 3 = roundsmall
    static DWMWA_WINDOW_CORNER_PREFERENCE := 33
    if !IsWin11()
        return
    try DllCall("dwmapi\DwmSetWindowAttribute"
        , "ptr", hwnd, "uint", DWMWA_WINDOW_CORNER_PREFERENCE
        , "int*", mode, "uint", 4)
}

ApplyModernChrome(gui, dark := false) {
    if !gui.HasProp("Hwnd")
        return
    SetDarkTitleBar(gui.Hwnd, dark)
    SetRoundedCorners(gui.Hwnd, 2)
}

; Toggle Windows' dark-mode rendering for *this process's* native popup
; menus, scroll bars, and Common Controls. Uses two undocumented uxtheme.dll
; entry points by ordinal (Win10 1903+):
;   135 = SetPreferredAppMode(AppMode)
;          0 = Default, 1 = AllowDark, 2 = ForceDark, 3 = ForceLight
;   136 = FlushMenuThemes()
; The same mechanism File Explorer / RegEdit / Notepad use for their menus.
; Silently no-ops on Windows builds that don't export the ordinals.
SetAppDarkMode(dark) {
    static initOK := false, setMode := 0, flushMenus := 0
    if !initOK {
        hUx := DllCall("GetModuleHandle", "str", "uxtheme.dll", "ptr")
        if !hUx
            hUx := DllCall("LoadLibrary", "str", "uxtheme.dll", "ptr")
        if hUx {
            setMode    := DllCall("GetProcAddress", "ptr", hUx, "ptr", 135, "ptr")
            flushMenus := DllCall("GetProcAddress", "ptr", hUx, "ptr", 136, "ptr")
        }
        initOK := true
    }
    if !setMode
        return
    mode := dark ? 2 : 3   ; ForceDark / ForceLight
    try DllCall(setMode, "int", mode)
    if flushMenus
        try DllCall(flushMenus)
}

ModernPalette(dark := false) {
    if dark {
        return Map(
            "bg",        "202020",
            "bgAlt",     "2B2B2B",
            "fg",        "F2F2F2",
            "fgMuted",   "C8C8C8",
            "accent",    "60CDFF",
            "border",    "3A3A3A",
            "btnBg",     "2D2D2D",
            "btnFg",     "F2F2F2"
        )
    }
    return Map(
        "bg",        "F3F3F3",
        "bgAlt",     "FFFFFF",
        "fg",        "1A1A1A",
        "fgMuted",   "5C5C5C",
        "accent",    "0067C0",
        "border",    "E5E5E5",
        "btnBg",     "FAFAFA",
        "btnFg",     "1A1A1A"
    )
}
