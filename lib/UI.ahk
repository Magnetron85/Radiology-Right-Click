; ============================================================
; lib/UI.ahk -- Result popup window
; ------------------------------------------------------------
; Used by all calculators. Sizes to content (capped at ~50% of
; the active monitor), respects dark mode, applies the modern
; immersive title bar, and copies the result to the clipboard.
; The cursor is parked over the Close button so dismissing the
; popup is a single click -- supports the RSI / fewest-motions
; design rule the script is built around.
; ============================================================

#Requires AutoHotkey v2.0
#Include Modern.ahk
#Include Util.ahk

ShowResult(text) {
    dark    := Prefs.Get("display", "darkMode", false)
    palette := ModernPalette(dark)

    MouseGetPos(&mx, &my)
    work := GetWorkAreaAt(mx, my)
    maxW := Round(work.width  * 0.5)
    maxH := Round(work.height * 0.6)

    ; --- measure text in an offscreen window so we can size to fit
    measure := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")  ; WS_EX_NOACTIVATE
    measure.SetFont("s10", ModernFont())
    tc := measure.Add("Text", "w" (maxW - 60) " wrap", text)
    tc.GetPos(, , &tw, &th)
    measure.Destroy()

    pad  := 16
    btnH := 32
    w := Min(Max(tw + pad*2 + 20, 360), maxW)
    h := Min(Max(th + pad*2 + btnH + 12, 220), maxH)

    pos := ClampToWorkArea(mx + 12, my + 12, w, h, work)

    g := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Result")
    g.MarginX := pad, g.MarginY := pad
    g.BackColor := palette["bg"]
    g.SetFont("s10 c" palette["fg"], ModernFont())

    editH := h - pad*2 - btnH - 12
    editW := w - pad*2
    edit := g.Add("Edit"
        , "ReadOnly Wrap VScroll -E0x200 +Background" palette["bgAlt"]
          . " c" palette["fg"]
          . " w" editW " h" editH
        , text)

    g.SetFont("s10 Bold")
    btn := g.Add("Button"
        , "Default w120 h" btnH " xp y+10 +Background" palette["btnBg"]
          . " c" palette["btnFg"]
        , "Close")
    btn.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close",  (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    ApplyModernChrome(g, dark)
    g.Show("x" pos.x " y" pos.y " w" w " h" h)

    ; park cursor on Close so dismissing is a single click
    btn.GetPos(&bx, &by, &bw, &bh)
    MouseMove(pos.x + bx + bw/2, pos.y + by + bh/2, 0)

    A_Clipboard := text
}
