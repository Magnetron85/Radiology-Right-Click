; ============================================================
; lib/FormGui.ahk -- Shared form builder for RADS calculators
; ------------------------------------------------------------
; Layout strategy:
;
;   * Every helper adds its control(s) at the current y cursor,
;     then queries their actual rendered bottom via GetPos and
;     advances y to one row-gap below the lowest one. This makes
;     the layout self-correcting -- wrapped labels, larger fonts,
;     or DPI scaling no longer cause overlaps or white gaps.
;
;   * Show() displays the gui hidden in AutoSize mode first,
;     measures the natural width/height the controls actually
;     occupied, caps that to the active work area minus a
;     margin, then re-shows positioned near the cursor.
;
;   * In dark mode, every child control gets
;     uxtheme!AllowDarkModeForWindow + SetWindowTheme("DarkMode_
;     Explorer") so dropdowns / edits / scrollbars don't render
;     as white blocks against the dark background.
;
; The helper API itself is unchanged from before:
;
;   form := RadsForm("Title", 480)
;   form.Header("Section")
;   form.Dropdown("V", "Label:", ["a","b"], 1)
;   form.Numeric("N",  "Number:", 0)
;   form.Checkbox("C", "Tickme",  false)
;   form.TextArea("T", "Notes:",  80, "")
;   form.SetSubmit(MyHandler)
;   form.AddButtons()
;   form.Show()
; ============================================================

#Requires AutoHotkey v2.0
#Include Modern.ahk
#Include Util.ahk
#Include UI.ahk
#Include TextScan.ahk

class RadsForm {
    g           := ""
    palette     := ""
    dark        := false
    width       := 480
    leftX       := 18
    labelX      := 18
    fieldX      := 200
    fieldW      := 260
    rowGap      := 6
    sectionGap  := 12
    y           := 14
    titleStr    := ""
    onSubmit    := ""
    saveLabel   := "Calculate"
    controls    := ""
    inputs      := ""   ; ordered list of {vname, label, type} for FormatInputs
    byName      := ""   ; vname -> {ctl, label, default, type} for dependency wiring
    stash       := ""   ; vname -> value cached on disable so re-enable can restore
    currentSection := ""

    __New(title, width := 480) {
        this.titleStr := title
        this.width    := width
        this.fieldW   := width - this.fieldX - 18
        this.dark     := Prefs.Get("display", "darkMode", false)
        this.palette  := ModernPalette(this.dark)
        this.controls := []
        this.inputs   := []
        this.byName   := Map()
        this.stash    := Map()

        this.g := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox +DPIScale", title)
        this.g.MarginX := 16
        this.g.MarginY := 12
        this.g.BackColor := this.palette["bg"]
        this.g.SetFont("s10 c" this.palette["fg"], ModernFont())
    }

    ; Record an input for the auto-generated inputs summary AND the dependency
    ; wiring framework. `ctl` is the input control, `labelCtl` is the optional
    ; Text label control (may be ""), `default` is the value to restore when
    ; the control is disabled or hidden via SetEnabled / SetVisible.
    _RecordInput(vname, label, type, ctl := "", labelCtl := "", default := "") {
        this.inputs.Push({
            vname:   vname,
            label:   _CleanLabel(label),
            type:    type,
            section: this.currentSection
        })
        if (ctl != "")
            this.byName[vname] := { ctl: ctl, label: labelCtl, default: default, type: type }
    }

    ; ---- dependency wiring primitives ---------------------------------------
    ; All operate on the vname used when the control was added. Reset-on-
    ; disable / reset-on-hide is automatic, so disabled/hidden controls return
    ; their default value from Gui.Submit() rather than whatever the user
    ; previously typed -- this neutralizes the well-known AHK v2 gotcha where
    ; Submit() walks all named controls regardless of enabled / visible state.

    ; Wire a Click (checkbox) or Change (dropdown / edit) handler on a named
    ; control. The callback receives this form as its single argument; read
    ; current values via form.GetValue(name).
    OnChange(name, callback) {
        if !this.byName.Has(name)
            return
        ctl := this.byName[name].ctl
        evt := (ctl.Type = "Checkbox" || ctl.Type = "Radio") ? "Click" : "Change"
        ctl.OnEvent(evt, (*) => callback(this))
    }

    SetEnabled(name, enabled) {
        if !this.byName.Has(name)
            return
        rec := this.byName[name]
        if !enabled {
            ; Disabling: stash the user's current value so re-enable can
            ; restore it, then reset the live value (so Gui.Submit returns
            ; the default rather than the stashed value while disabled).
            if !this.stash.Has(name)
                this.stash[name] := rec.ctl.Value
            rec.ctl.Enabled := false
            if (rec.label != "")
                rec.label.Enabled := false
            this._SetMuted(rec, true)
            this._ResetCtl(rec)
        } else {
            rec.ctl.Enabled := true
            if (rec.label != "")
                rec.label.Enabled := true
            this._SetMuted(rec, false)
            ; Restore the user's prior value if one was stashed.
            if this.stash.Has(name) {
                try rec.ctl.Value := this.stash[name]
                this.stash.Delete(name)
            }
        }
    }

    SetVisible(name, visible) {
        if !this.byName.Has(name)
            return
        rec := this.byName[name]
        if !visible {
            if !this.stash.Has(name)
                this.stash[name] := rec.ctl.Value
            rec.ctl.Visible := false
            if (rec.label != "")
                rec.label.Visible := false
            this._ResetCtl(rec)
        } else {
            rec.ctl.Visible := true
            if (rec.label != "")
                rec.label.Visible := true
            if this.stash.Has(name) {
                try rec.ctl.Value := this.stash[name]
                this.stash.Delete(name)
            }
        }
    }

    ; Apply a muted foreground color for visual disabled feedback. Only used
    ; for Checkbox / Radio controls -- those render their label text via the
    ; system theme, which doesn't dim under dark mode when Enabled := false.
    ; Edit / DropDownList / DateTime have reliable native disabled rendering
    ; (the text greys out automatically), so manually recoloring them is
    ; unnecessary and can stick if the OS theme doesn't honor `Opt("+c...")`.
    ; The sibling Text label DOES need recoloring for non-checkbox fields.
    _SetMuted(rec, muted) {
        color := muted ? this.palette["fgMuted"] : this.palette["fg"]
        isCheckLike := (rec.type = "checkbox" || rec.type = "radio")
        if isCheckLike {
            try {
                rec.ctl.Opt("+c" color)
                rec.ctl.Redraw()
            }
        }
        if (rec.label != "") {
            try {
                rec.label.Opt("+c" color)
                rec.label.Redraw()
            }
        }
    }

    GetValue(name) {
        if !this.byName.Has(name)
            return ""
        return this.byName[name].ctl.Value
    }

    SetValue(name, value) {
        if !this.byName.Has(name)
            return
        this.byName[name].ctl.Value := value
    }

    _ResetCtl(rec) {
        try {
            if (rec.default != "")
                rec.ctl.Value := rec.default
            else if (rec.type = "checkbox" || rec.type = "radio")
                rec.ctl.Value := 0
            else if (rec.type = "dropdown")
                rec.ctl.Value := 1
            else if (rec.type = "numeric")
                rec.ctl.Value := 0
            else if (rec.type = "textarea")
                rec.ctl.Value := ""
        }
    }

    ; Declarative shortcut: when ANY of `triggers` (checkbox vnames) is
    ; checked, disable + reset every control named in `affected`. When ALL
    ; triggers are unchecked, re-enable them.
    DisableWhenAnyChecked(triggers, affected) {
        update := (*) => this._ApplyExclusion(triggers, affected)
        for t in triggers {
            if this.byName.Has(t)
                this.byName[t].ctl.OnEvent("Click", update)
        }
    }
    _ApplyExclusion(triggers, affected) {
        active := false
        for t in triggers {
            if this.byName.Has(t) && this.byName[t].ctl.Value {
                active := true
                break
            }
        }
        for a in affected
            this.SetEnabled(a, !active)
    }

    ; Build a "Selected inputs:" block from the saved submit values. Skips
    ; unchecked checkboxes, blank numerics, blank dropdowns, and empty
    ; textareas. Groups by the Header() section the input was declared in.
    FormatInputs(saved) {
        out := ""
        currentSection := ""
        for inp in this.inputs {
            try
                val := saved.%inp.vname%
            catch
                continue

            ; Skip controls that were disabled or hidden at submit time.
            ; SetEnabled(false) / SetVisible(false) reset the control to
            ; its default value, so its carried value isn't a real user
            ; input -- emitting it would surface the first dropdown option
            ; (e.g. "Obtuse (wide base)") even when the gating checkbox
            ; was off.
            if this.byName.Has(inp.vname) {
                ctl := this.byName[inp.vname].ctl
                try {
                    if (!ctl.Enabled || !ctl.Visible)
                        continue
                }
            }

            line := _FormatInputLine(inp, val)
            if (line = "")
                continue

            if (inp.section != currentSection) {
                if (out != "")
                    out .= "`n"
                currentSection := inp.section
                if (currentSection != "")
                    out .= currentSection ":`n"
            }
            out .= line
        }
        return out
    }

    ; ---- internal: advance y to one row-gap below the actual rendered
    ;      bottom of the controls just added.
    _Advance(ctls, extraGap := -1) {
        gap := extraGap >= 0 ? extraGap : this.rowGap
        maxBottom := this.y
        for c in ctls {
            c.GetPos(, &cy, , &ch)
            if (cy + ch > maxBottom)
                maxBottom := cy + ch
            this.controls.Push(c)
        }
        this.y := maxBottom + gap
    }

    Header(label) {
        if (this.y > 14)
            this.y += this.sectionGap
        this.g.SetFont("s10 Bold c" this.palette["fg"])
        t := this.g.Add("Text"
            , "x" this.leftX " y" this.y " w" (this.width - 32) " +Wrap"
            , label)
        this.g.SetFont("s10 Norm c" this.palette["fg"])
        this.currentSection := _CleanLabel(label)
        this._Advance([t], 4)
    }

    Note(text) {
        this.g.SetFont("s9 Italic c" this.palette["fgMuted"])
        t := this.g.Add("Text"
            , "x" this.leftX " y" this.y " w" (this.width - 32) " +Wrap"
            , text)
        this.g.SetFont("s10 Norm c" this.palette["fg"])
        this._Advance([t], 2)
    }

    Spacer(px := 8) {
        this.y += px
    }

    Dropdown(vname, label, choices, selectedIdx := 1) {
        lbl := this.g.Add("Text"
            , "x" this.labelX " y" (this.y + 3)
              . " w" (this.fieldX - this.labelX - 8) " +Wrap"
            , label)
        ctl := this.g.Add("DropDownList"
            , "x" this.fieldX " y" this.y " w" this.fieldW
              . " v" vname " Choose" selectedIdx
              . " Background" this.palette["bgAlt"] " c" this.palette["fg"]
            , choices)
        this._RecordInput(vname, label, "dropdown", ctl, lbl, selectedIdx)
        this._Advance([lbl, ctl])
        return ctl
    }

    DateField(vname, label, defaultYmd := "") {
        lbl := this.g.Add("Text"
            , "x" this.labelX " y" (this.y + 3)
              . " w" (this.fieldX - this.labelX - 8) " +Wrap"
            , label)
        dt := this.g.Add("DateTime"
            , "x" this.fieldX " y" this.y " w160 v" vname
            , "MM/dd/yyyy")
        if (defaultYmd != "")
            dt.Value := defaultYmd
        this._RecordInput(vname, label, "date", dt, lbl, defaultYmd)
        this._Advance([lbl, dt])
        return dt
    }

    Numeric(vname, label, default := 0) {
        lbl := this.g.Add("Text"
            , "x" this.labelX " y" (this.y + 3)
              . " w" (this.fieldX - this.labelX - 8) " +Wrap"
            , label)
        ctl := this.g.Add("Edit"
            , "x" this.fieldX " y" this.y " w" this.fieldW
              . " v" vname
              . " Background" this.palette["bgAlt"] " c" this.palette["fg"]
            , default)
        this._RecordInput(vname, label, "numeric", ctl, lbl, default)
        this._Advance([lbl, ctl])
        return ctl
    }

    ; Two paired numerics on a single row -- useful for related measurements
    ; (e.g. prior size + interval). Each cell is half-width with its own
    ; label + edit; both inputs participate in the dependency framework via
    ; _RecordInput exactly like a single Numeric.
    NumericRow2(v1, label1, v2, label2, default1 := 0, default2 := 0) {
        halfW := (this.width - 40) // 2
        labelW := halfW * 4 // 7         ; ~57% label / ~43% edit per half
        editW := halfW - labelW - 4

        lbl1 := this.g.Add("Text"
            , "x" this.leftX " y" (this.y + 3) " w" labelW " +Wrap"
            , label1)
        ctl1 := this.g.Add("Edit"
            , "x" (this.leftX + labelW + 4) " y" this.y " w" editW
              . " v" v1
              . " Background" this.palette["bgAlt"] " c" this.palette["fg"]
            , default1)

        rightX := this.leftX + halfW + 8
        lbl2 := this.g.Add("Text"
            , "x" rightX " y" (this.y + 3) " w" labelW " +Wrap"
            , label2)
        ctl2 := this.g.Add("Edit"
            , "x" (rightX + labelW + 4) " y" this.y " w" editW
              . " v" v2
              . " Background" this.palette["bgAlt"] " c" this.palette["fg"]
            , default2)

        this._RecordInput(v1, label1, "numeric", ctl1, lbl1, default1)
        this._RecordInput(v2, label2, "numeric", ctl2, lbl2, default2)
        this._Advance([lbl1, ctl1, lbl2, ctl2])
        return [ctl1, ctl2]
    }

    Checkbox(vname, label, default := false) {
        ; Explicit `c<fg>` is required because dark child theming via
        ; SetWindowTheme("DarkMode_Explorer", ...) keeps the box dark but
        ; leaves the *text* rendered in the system default (black). Setting
        ; the option here forces the label color to match the palette.
        ctl := this.g.Add("Checkbox"
            , "x" this.leftX " y" this.y " w" (this.width - 32) " +Wrap"
              . " v" vname " c" this.palette["fg"]
            , label)
        ctl.Value := default ? 1 : 0
        this._RecordInput(vname, label, "checkbox", ctl, "", default ? 1 : 0)
        this._Advance([ctl])
        return ctl
    }

    ; Checkbox in left cell + Numeric (label + edit) in right cell -- one
    ; row total. Useful when a checkbox toggles a feature whose magnitude
    ; is captured in a related numeric (e.g. Lung-RADS "solid component
    ; new/growing" + the component size).
    CheckboxNumericRow(vCheck, labelCheck, vNum, labelNum, defaultCheck := false, defaultNum := 0) {
        halfW := (this.width - 40) // 2
        cb := this.g.Add("Checkbox"
            , "x" this.leftX " y" this.y " w" halfW " +Wrap"
              . " v" vCheck " c" this.palette["fg"]
            , labelCheck)
        cb.Value := defaultCheck ? 1 : 0

        rightX := this.leftX + halfW + 8
        labelW := halfW * 4 // 7
        editW  := halfW - labelW - 4
        lblNum := this.g.Add("Text"
            , "x" rightX " y" (this.y + 3) " w" labelW " +Wrap"
            , labelNum)
        ctlNum := this.g.Add("Edit"
            , "x" (rightX + labelW + 4) " y" this.y " w" editW
              . " v" vNum
              . " Background" this.palette["bgAlt"] " c" this.palette["fg"]
            , defaultNum)

        this._RecordInput(vCheck, labelCheck, "checkbox", cb,     "",     defaultCheck ? 1 : 0)
        this._RecordInput(vNum,   labelNum,   "numeric",  ctlNum, lblNum, defaultNum)
        this._Advance([cb, lblNum, ctlNum])
        return [cb, ctlNum]
    }

    CheckboxRow2(v1, label1, v2, label2, default1 := false, default2 := false) {
        halfW := (this.width - 40) // 2
        c1 := this.g.Add("Checkbox"
            , "x" this.leftX " y" this.y " w" halfW " +Wrap"
              . " v" v1 " c" this.palette["fg"]
            , label1)
        c1.Value := default1 ? 1 : 0
        c2 := this.g.Add("Checkbox"
            , "x" (this.leftX + halfW + 4) " y" this.y " w" halfW " +Wrap"
              . " v" v2 " c" this.palette["fg"]
            , label2)
        c2.Value := default2 ? 1 : 0
        this._RecordInput(v1, label1, "checkbox", c1, "", default1 ? 1 : 0)
        this._RecordInput(v2, label2, "checkbox", c2, "", default2 ? 1 : 0)
        this._Advance([c1, c2])
        return [c1, c2]
    }

    TextArea(vname, label, height := 60, default := "") {
        lbl := this.g.Add("Text"
            , "x" this.leftX " y" this.y " w" (this.width - 32)
            , label)
        this._Advance([lbl], 4)
        ctl := this.g.Add("Edit"
            , "x" this.leftX " y" this.y " w" (this.width - 32) " h" height
              . " -Wrap +VScroll v" vname
              . " Background" this.palette["bgAlt"] " c" this.palette["fg"]
            , default)
        this._RecordInput(vname, label, "textarea", ctl, lbl, default)
        this._Advance([ctl])
        return ctl
    }

    SetSubmit(callback) {
        this.onSubmit := callback
    }

    SetSaveLabel(label) {
        this.saveLabel := label
    }

    AddButtons() {
        btnH := 32
        this.y += 8
        cancelX := this.width - 208
        saveX   := this.width - 112

        btnCancel := this.g.Add("Button"
            , "x" cancelX " y" this.y " w90 h" btnH
              . " +Background" this.palette["btnBg"] " c" this.palette["btnFg"]
            , "Cancel")
        btnCancel.OnEvent("Click", this._OnCancel.Bind(this))

        btnSave := this.g.Add("Button"
            , "Default x" saveX " y" this.y " w94 h" btnH
              . " +Background" this.palette["btnBg"] " c" this.palette["btnFg"]
            , this.saveLabel)
        btnSave.OnEvent("Click", this._OnSave.Bind(this))

        this.g.OnEvent("Close",  this._OnCancel.Bind(this))
        this.g.OnEvent("Escape", this._OnCancel.Bind(this))

        this._Advance([btnCancel, btnSave], 10)
    }

    _OnCancel(*) {
        this.g.Destroy()
    }

    _OnSave(*) {
        try
            saved := this.g.Submit(true)
        catch as e {
            MsgBox("Form submit failed: " e.Message, "RightClick", "Iconx")
            return
        }
        if (this.onSubmit = "")
            return
        try {
            ; Pass the form itself as a 2nd arg so the handler can call
            ; form.FormatInputs(saved) and echo the user's selections into
            ; the result. Handlers that don't want the form can declare
            ; their signature as (v) or ignore the 2nd parameter.
            result := this.onSubmit.Call(saved, this)
            ; Handlers may return a flat string (legacy) or a CalcResult
            ; struct (new). ShowResult accepts both transparently.
            if IsCalcResult(result)
                ShowResult(result)
            else if (result != "" && Type(result) = "String")
                ShowResult(result)
        } catch as e {
            MsgBox("Calculation error:`n`n" e.Message "`n`nAt: "
                   . (e.HasProp("File") ? e.File : "?") ":"
                   . (e.HasProp("Line") ? e.Line : "?")
                 , "RightClick", "Iconx")
        }
    }

    Show() {
        MouseGetPos(&mx, &my)
        work := GetWorkAreaAt(mx, my)

        ; Height comes from this.y, which the _Advance() helpers updated
        ; based on each control's measured rendered bottom. Cap at work area
        ; minus a small margin so the window always fits.
        h := this.y + 14
        w := this.width
        if (w > work.width  - 24)
            w := work.width  - 24
        if (h > work.height - 24)
            h := work.height - 24

        pos := ClampToWorkArea(mx + 12, my + 12, w, h, work)

        ApplyModernChrome(this.g, this.dark)
        this.g.Show("x" pos.x " y" pos.y " w" w " h" h)
        WinActivate("ahk_id " this.g.Hwnd)
    }
}

; ---- helpers used by RadsForm.FormatInputs ---------------------------------

; Strip trailing colon / hint text "(0-2)" / parenthetical examples and
; whitespace from a field label so the summary reads as natural prose.
_CleanLabel(label) {
    s := label
    s := RegExReplace(s, "\s*\(.*?\)\s*$", "")   ; drop trailing parenthetical
    s := RegExReplace(s, ":+\s*$", "")           ; drop trailing colon(s)
    return Trim(s)
}

; Format a single recorded input into a "  - Label: value" line, or "" if
; the value should be skipped (unticked checkbox, blank numeric, etc).
_FormatInputLine(inp, val) {
    label := inp.label
    switch inp.type {
        case "checkbox":
            return val ? "  - " label "`n" : ""
        case "dropdown":
            if (val = "" || val = 0)
                return ""
            ; Suppress noise: dropdowns whose selection is a "no-selection"
            ; sentinel (None / N/A / Not assessed ...) carry no information
            ; for the report summary and just clutter the output.
            if (val = "None")
                return ""
            if (SubStr(val, 1, 3) = "N/A")
                return ""
            if (SubStr(val, 1, 12) = "Not assessed")
                return ""
            return "  - " label ": " val "`n"
        case "numeric":
            if (val = "" || val = 0)
                return ""
            return "  - " label ": " val "`n"
        case "date":
            if (val = "" || val = 0)
                return ""
            try
                ds := FormatTime(val, "MM/dd/yyyy")
            catch
                ds := val
            return "  - " label ": " ds "`n"
        case "textarea":
            if (Trim(val) = "")
                return ""
            ; multi-line textareas get a folded display
            indented := "      " . StrReplace(Trim(val), "`n", "`n      ")
            return "  - " label ":`n" indented "`n"
    }
    return ""
}
