; ============================================================
; lib/Widget.ahk -- Persistent launcher widget
; ------------------------------------------------------------
; A small, always-on-top, non-intrusive button that opens the
; right-click context menu without needing a (Ctrl+)right-click
; inside a target app. Useful when the user is not over a
; recognized reporting window, or just prefers a visible launcher.
;
;   * Left-CLICK  -> opens the context menu (same as the hotkey).
;   * Left-DRAG   -> moves the widget; the new position is saved.
;   * Right-click -> tiny menu (Open / Hide / Preferences).
;
; The window uses WS_EX_NOACTIVATE so clicking it never steals
; focus from the reporting app -- the text selection survives, so
; the menu's capture (Ctrl+C) still reads the highlighted report
; text exactly as the hotkey path does.
;
; Controlled by Prefs "widget": { enabled, x, y }. Launcher.Apply()
; creates or destroys the widget to match the pref and is called at
; startup and whenever the preference is toggled.
; ============================================================

#Requires AutoHotkey v2.0
#Include Modern.ahk
#Include Util.ahk

class Launcher {
    static gui      := ""
    static ctrlHwnd := 0
    static dragArmed := false

    ; Create or destroy the widget to match the saved preference.
    static Apply() {
        if Prefs.Get("widget", "enabled", false)
            Launcher.Show()
        else
            Launcher.Hide()
    }

    static Show() {
        if (Launcher.gui != "")
            return   ; already shown
        dark := Prefs.Get("display", "darkMode", false)
        pal  := ModernPalette(dark)

        ; WS_EX_TOOLWINDOW (no taskbar button) + WS_EX_NOACTIVATE (clicks
        ; don't activate it, so the reporting app keeps focus + selection).
        g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
        g.MarginX := 0, g.MarginY := 0
        g.BackColor := pal["accent"]

        size := 40
        ; Icon if available, else a bold "RC" label. The control fills the
        ; client area so every click lands on it.
        ico := A_ScriptDir "\RightClick.ico"
        if FileExist(ico) {
            ctrl := g.Add("Picture", "x4 y4 w" (size-8) " h" (size-8), ico)
        } else {
            g.SetFont("s12 Bold cFFFFFF", ModernFont())
            ctrl := g.Add("Text"
                , "x0 y0 w" size " h" size " +Center 0x200", "RC")
        }
        Launcher.ctrlHwnd := ctrl.Hwnd

        pos := Launcher._StartPos(size)
        g.Show("x" pos.x " y" pos.y " w" size " h" size " NoActivate")
        SetRoundedCorners(g.Hwnd, 2)
        ; Slight whole-window translucency so it reads as a non-intrusive
        ; overlay rather than a hard opaque block.
        try WinSetTransparent(232, g)

        g.OnEvent("ContextMenu", (*) => Launcher._RightMenu())
        Launcher.gui := g

        ; One global down-handler; it filters to our window/control.
        OnMessage(0x0201, _Launcher_LButtonDown)   ; WM_LBUTTONDOWN
    }

    static Hide() {
        if (Launcher.gui = "")
            return
        OnMessage(0x0201, _Launcher_LButtonDown, 0)   ; deregister
        try Launcher.gui.Destroy()
        Launcher.gui := ""
        Launcher.ctrlHwnd := 0
    }

    ; Saved position, else a default near the top-right of the work area of
    ; the monitor under the cursor (out of the way of report text).
    static _StartPos(size) {
        sx := Prefs.Get("widget", "x", "")
        sy := Prefs.Get("widget", "y", "")
        MouseGetPos(&mx, &my)
        work := GetWorkAreaAt(mx, my)
        ; size is a logical unit (Gui.Show scales w/h); placement math is
        ; in physical pixels, so use the widget's physical footprint.
        sizePx := Round(size * A_ScreenDPI / 96)
        if (sx = "" || sy = "") {
            return { x: work.right - sizePx - 24, y: work.top + 120 }
        }
        ; Clamp a saved position back on-screen (monitor layout may change).
        clamped := ClampToWorkArea(sx + 0, sy + 0, sizePx, sizePx, work)
        return { x: clamped.x, y: clamped.y }
    }

    static SavePos(x, y) {
        Prefs.Set("widget", "x", x)
        Prefs.Set("widget", "y", y)
        Prefs.Save()
    }

    static _RightMenu() {
        m := Menu()
        m.Add("Open menu", (*) => ShowContextMenu())
        m.Add("Hide widget", (*) => Launcher._HideFromMenu())
        m.Add()
        m.Add("Preferences...", (*) => ShowPreferencesWindow())
        m.Show()
    }

    static _HideFromMenu() {
        Prefs.Set("widget", "enabled", false)
        Prefs.Save()
        Launcher.Hide()
    }
}

; WM_LBUTTONDOWN handler. Initiates the native window move loop (which blocks
; until the button is released); if the window did not actually move, the
; press was a click -> open the menu. Otherwise it was a drag -> save the
; new position. Distinguishing this way means a single left button does both
; "click to open" and "drag to move" with no modifier.
_Launcher_LButtonDown(wParam, lParam, msg, hwnd) {
    if (Launcher.gui = "")
        return
    if (hwnd != Launcher.gui.Hwnd && hwnd != Launcher.ctrlHwnd)
        return

    Launcher.gui.GetPos(&x0, &y0)
    DllCall("ReleaseCapture")
    ; WM_NCLBUTTONDOWN (0xA1) with HTCAPTION (2): enter the system move loop.
    DllCall("SendMessage", "ptr", Launcher.gui.Hwnd, "uint", 0x00A1
          , "ptr", 2, "ptr", 0)
    Launcher.gui.GetPos(&x1, &y1)

    if (Abs(x1 - x0) <= 3 && Abs(y1 - y0) <= 3) {
        ShowContextMenu()   ; treated as a click
    } else {
        Launcher.SavePos(x1, y1)
    }
    return 0
}
