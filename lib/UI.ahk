; ============================================================
; lib/UI.ahk -- Result popup window
; ------------------------------------------------------------
; Used by all calculators. Accepts either a legacy flat string
; (back-compat path) or a CalcResult struct. The popup renders:
;
;   [echo block, if non-empty]
;   [impression]
;   [advisories]
;   [methodology, if Show methodology is on]
;   [citations + URL links, if Show references is on]
;
; The "Show methodology" and "Show references" checkboxes default
; to the global Prefs values but toggle WINDOW-LOCAL only -- they
; never write back to Prefs (per spec, per-result audit must not
; mutate defaults).
;
; Clipboard policy: the initial clipboard payload is the impression
; (+ separate recommendation, if not already in the impression) and
; nothing else. Radiologists pasting into clinical reports should
; never get methodology or echo by accident. Two explicit buttons:
;
;   [Copy impression]      -> impression + recommendation
;   [Copy with methodology] -> the full rendered text (WYSIWYG)
;
; Other notes:
;   * Cursor parks on Close for one-click dismiss (RSI design rule).
;   * Edit scroll position is preserved when methodology is toggled
;     so the user doesn't lose their place while auditing.
;   * Citation URLs render as native SysLink controls below the Edit
;     so the radiologist can click straight through to the source.
; ============================================================

#Requires AutoHotkey v2.0
#Include Modern.ahk
#Include Util.ahk
#Include CalcResult.ahk

ShowResult(payload) {
    ; Accept legacy string returns transparently.
    result := IsCalcResult(payload) ? payload : LegacyResult(payload)

    dark    := Prefs.Get("display", "darkMode", false)
    palette := ModernPalette(dark)

    ; Local toggle state -- starts at global pref, never writes back.
    state := {
        showMethodology: !!Prefs.Get("display", "showMethodology", true),
        showCitations:   !!Prefs.Get("display", "showCitations",   true),
        result:          result,
        palette:         palette,
        dark:            dark
    }

    MouseGetPos(&mx, &my)
    work := GetWorkAreaAt(mx, my)

    targetW := 640
    maxW := Min(targetW, Round(work.width * 0.55))
    maxH := Round(work.height * 0.80)

    ; Measure the worst-case text (all toggles on) so the window doesn't
    ; clip when the user enables methodology after open.
    measureText := RenderResult(result, { showMethodology: true, showCitations: true })
    measure := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
    measure.SetFont("s10", ModernFont())
    tc := measure.Add("Text", "w" (maxW - 60) " wrap", measureText)
    tc.GetPos(, , &tw, &th)
    measure.Destroy()

    pad     := 16
    chkH    := 22       ; top checkboxes row
    btnH    := 32       ; bottom button row
    ; ALWAYS reserve space for citation links (not gated on the toggle).
    ; Otherwise a user who starts with refs off and then toggles them on
    ; paints SysLink controls into the button row -- the Edit was sized to
    ; fill the would-be link area, leaving nowhere for them to land.
    ; Compute as if showCitations were true so the window has the room.
    linksH  := _MeasureLinksHeight(result, true)
    sectionGap := 8

    w := Min(Max(tw + pad*2 + 24, 460), maxW)
    contentH := th + chkH + sectionGap + linksH + btnH + sectionGap*2
    h := Min(Max(contentH + pad*2, 260), maxH)

    pos := ClampToWorkArea(mx + 12, my + 12, w, h, work)

    g := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Result")
    g.MarginX := pad, g.MarginY := pad
    g.BackColor := palette["bg"]
    g.SetFont("s10 c" palette["fg"], ModernFont())

    ; --- Top row: toggles ---
    cbMethod := g.Add("Checkbox"
        , "x" pad " y" pad " h" chkH " c" palette["fg"]
        , "Show methodology")
    cbMethod.Value := state.showMethodology ? 1 : 0

    cbCit := g.Add("Checkbox"
        , "x+24 yp h" chkH " c" palette["fg"]
        , "Show references")
    cbCit.Value := state.showCitations ? 1 : 0

    ; --- Middle: result text Edit ---
    editY := pad + chkH + sectionGap
    editW := w - pad*2
    editH := h - editY - btnH - sectionGap*2 - linksH - pad

    edit := g.Add("Edit"
        , "ReadOnly Wrap VScroll -E0x200 +Background" palette["bgAlt"]
          . " c" palette["fg"]
          . " x" pad " y" editY " w" editW " h" editH
        , RenderResult(result, state))

    ; --- References block (SysLink controls below the Edit) ---
    state.linkRows := []
    linkY := editY + editH + sectionGap
    if (state.showCitations && result.citations is Array
        && result.citations.Length > 0) {
        for i, c in result.citations {
            row := _AddCitationRow(g, c, pad, linkY, editW, palette)
            if (row != "")
                state.linkRows.Push(row)
            linkY += 22
        }
    }

    ; --- Bottom: copy buttons + close ---
    btnY := h - btnH - pad
    g.SetFont("s10 Bold")

    btnCopyAll := g.Add("Button"
        , "x" pad " y" btnY " w180 h" btnH
          . " +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Copy with methodology")
    btnCopyAll.OnEvent("Click", (*) => _CopyWithMethodology(state))

    btnCopy := g.Add("Button"
        , "x+12 yp w160 h" btnH
          . " +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Copy impression")
    btnCopy.OnEvent("Click", (*) => _CopyImpression(state))

    btnClose := g.Add("Button"
        , "Default x+" (w - pad*2 - 180 - 160 - 12 - 100 - 12) " yp w100 h" btnH
          . " +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Close")
    ; Closing the result window also clears the captured-text cache. This
    ; prevents text from a prior calc from leaking into the next calc's
    ; pre-fill when the user doesn't re-select before right-clicking
    ; again. The cache exists to bridge the case where PowerScribe drops
    ; selection between rapid back-to-back menu invocations, but holding
    ; it past an explicit window close conflates "PS dropped it" with
    ; "user moved on" -- the close is our best heuristic that the user
    ; is done with that input.
    closer := (*) => (_ClearLastSelection(), g.Destroy())
    btnClose.OnEvent("Click", closer)
    g.OnEvent("Close",  closer)
    g.OnEvent("Escape", closer)

    ; --- Wire toggles to re-render ---
    state.g       := g
    state.edit    := edit
    state.cbMethod := cbMethod
    state.cbCit    := cbCit
    state.editY    := editY
    state.editW    := editW
    state.btnY     := btnY
    state.sectionGap := sectionGap

    cbMethod.OnEvent("Click", (*) => _OnToggle(state))
    cbCit.OnEvent("Click",    (*) => _OnToggle(state))

    ApplyModernChrome(g, dark)
    g.Show("x" pos.x " y" pos.y " w" w " h" h)
    WinActivate("ahk_id " g.Hwnd)

    ; Park cursor on Close for one-click dismiss.
    btnClose.GetPos(&bx, &by, &bw, &bh)
    MouseMove(pos.x + bx + bw/2, pos.y + by + bh/2, 0)

    ; Initial clipboard payload is the safe impression-only string.
    A_Clipboard := ClipboardText(result)
}

; Toggle handler -- re-renders the edit with the new state and
; rebuilds the references row. Preserves scroll position via
; EM_GETFIRSTVISIBLELINE / EM_LINESCROLL so the user doesn't lose
; their place while auditing.
_OnToggle(state) {
    state.showMethodology := !!state.cbMethod.Value
    state.showCitations   := !!state.cbCit.Value

    ; Snapshot the top visible line index before mutating the Edit.
    firstLine := SendMessage(0x00CE, 0, 0, state.edit)   ; EM_GETFIRSTVISIBLELINE

    state.edit.Value := RenderResult(state.result, state)

    ; Restore scroll position.
    SendMessage(0x00B6, 0, firstLine, state.edit)        ; EM_LINESCROLL

    ; Destroy and rebuild references row.
    for row in state.linkRows {
        try row.Destroy()
    }
    state.linkRows := []
    if (state.showCitations
        && state.result.citations is Array
        && state.result.citations.Length > 0) {
        state.edit.GetPos(, , , &eh)
        linkY := state.editY + eh + state.sectionGap
        for i, c in state.result.citations {
            row := _AddCitationRow(state.g, c, 16, linkY
                                 , state.editW, state.palette)
            if (row != "")
                state.linkRows.Push(row)
            linkY += 22
        }
    }
}

; Clear the menu's captured-text cache. Called when a result window closes.
; The cache (lib/Menu.ahk:g_LastSelectedText) exists to fall back when an
; immediate re-capture returns empty (PowerScribe dropping selection
; between dialogs); clearing it here means the next right-click starts
; fresh -- no stale text from an earlier calc leaking into the new
; calc's pre-fill / echo.
_ClearLastSelection() {
    global g_LastSelectedText
    g_LastSelectedText := ""
}

_CopyImpression(state) {
    A_Clipboard := ClipboardText(state.result)
    _Flash(state.g, "Impression copied to clipboard")
}

_CopyWithMethodology(state) {
    A_Clipboard := RenderResult(state.result, state)
    _Flash(state.g, "Full result copied to clipboard")
}

; Brief flash in the title bar (no extra control needed) so the user
; gets confirmation without an intrusive MsgBox. The timer can fire AFTER
; the user has closed the result window -- accessing g.Title on a destroyed
; Gui throws "Gui has no window", so wrap the restore in a try.
_Flash(g, msg) {
    orig := ""
    try
        orig := g.Title
    catch
        return
    try
        g.Title := msg
    catch
        return
    SetTimer(_FlashRestore.Bind(g, orig), -1200)
}

_FlashRestore(g, orig) {
    try
        g.Title := orig
    ; If the user already closed the window, do nothing.
}

; Build one citation row -- a single-line clickable SysLink that opens the
; URL in the default browser. The citation text itself renders inside the
; Edit's References block; this control is JUST the click affordance.
; Fixed height (h20) so a long URL doesn't wrap onto a second line and
; overlap the button row below. Returns the SysLink so the toggle handler
; can destroy / recreate it on re-render.
_AddCitationRow(g, citation, x, y, w, palette) {
    if !IsObject(citation)
        return ""
    url := citation.HasOwnProp("url") ? citation.url : ""
    if (url = "")
        return ""
    link := g.Add("Link"
        , "x" x " y" y " w" w " h20 c" palette["fg"]
        , '<a href="' url '">Open reference in browser</a>')
    link.OnEvent("Click", _OnLinkClick.Bind(url))
    return link
}

_OnLinkClick(url, ctrl, ID, href, *) {
    ; Citation URLs are developer-controlled constants, but validate anyway
    ; so every Run() in the app goes through the same http/https allowlist
    ; that user-supplied reference URLs do.
    if !IsValidURL(url)
        return
    try Run(url)
}

; Predict height the citations block needs so the initial window sizes
; correctly. Each row ~22 px; zero if disabled or no URLs.
_MeasureLinksHeight(result, showCitations) {
    if !showCitations || !(result.citations is Array)
        return 0
    n := 0
    for c in result.citations {
        if IsObject(c) && c.HasOwnProp("url") && c.url != ""
            n++
    }
    return n * 22
}
