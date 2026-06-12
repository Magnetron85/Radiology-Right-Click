; ============================================================
; lib/Util.ahk -- Shared utilities used across calculators
; ============================================================

#Requires AutoHotkey v2.0
#Include Debug.ahk

; --- text / clipboard --------------------------------------------------------

; timeout is the FIRST Ctrl+C wait. Kept short (0.3 s) so the common case of
; opening the menu with NO selection -- where every copy attempt necessarily
; comes back empty -- returns quickly instead of stalling the menu through a
; long fallback chain. A real selection satisfies ClipWait the instant the
; app updates the clipboard (typically well under the timeout), so shortening
; this does not slow successful captures.
GetSelectedText(timeout := 0.3) {
    callT0 := A_TickCount
    saved := ClipboardAll()
    A_Clipboard := ""

    ; Release physically-held modifier keys so the subsequent Ctrl+C
    ; goes through cleanly. Only release keys that are actually held --
    ; releasing modifiers that aren't held would still send KeyUp events
    ; to the focused app. MS Word in particular treats a standalone
    ; Alt-up as a ribbon-menu activation (the "FEV HLBGIKO RPUS" KeyTip
    ; letters), which intercepts the subsequent Ctrl+C.
    for key in ["LCtrl","RCtrl","LAlt","RAlt","LShift","RShift","LWin","RWin"] {
        if GetKeyState(key, "P")
            SendEvent "{" key " up}"
    }
    Sleep 20
    ; Use SendEvent (not SendInput) -- AHK v1 defaulted to SendEvent and
    ; worked with apps that have low-level keyboard hooks (PowerScribe
    ; and other dictation systems). AHK v2 defaults to SendInput, which
    ; is faster but bypasses the standard Windows message queue --
    ; dictation hooks may not see it, leaving the clipboard empty.
    ;
    ; Give the Ctrl+C a real press duration (20 ms held) -- dictation
    ; hooks sample key state and can miss a zero-duration synthetic
    ; press. SetKeyDelay only affects the current thread.
    SetKeyDelay 10, 20
    SendEvent "^c"

    stage := "ctrl+c"
    waited := ClipWait(timeout, 1)
    if !waited {
        ; Fallback 1: WM_COPY posted straight at the focused control. No
        ; keyboard involved, so keyboard hooks can't eat it; standard
        ; Edit / RichEdit-derived controls honor it.
        try {
            focused := ControlGetFocus("A")
            if (focused != "") {
                SendMessage(0x0301, 0, 0, focused, "A")   ; WM_COPY
                stage := "wm_copy"
                waited := ClipWait(0.15, 1)
            }
        }
    }
    if !waited {
        ; Fallback 2: one more keyboard retry, for apps whose focused
        ; control doesn't honor WM_COPY. Kept short -- the WM_COPY path
        ; above already covers the hook-ate-the-keystroke case, so this is
        ; belt-and-suspenders, not the primary reliability mechanism.
        SetKeyDelay 30, 40
        SendEvent "^c"
        stage := "retry"
        waited := ClipWait(0.2, 1)
    }
    elapsed := A_TickCount - callT0
    if !waited {
        A_Clipboard := saved
        Debug.LogHex("capture"
            , Format("elapsed={}ms clipWait=TIMEOUT len=0", elapsed), "")
        return ""
    }
    text := A_Clipboard
    A_Clipboard := saved
    text := NormalizeReportWhitespace(text)
    Debug.LogHex("capture"
        , Format("elapsed={}ms clipWait=ok stage={} len={}", elapsed, stage, StrLen(text))
        , text)
    return text
}

; PowerScribe and other rich-text editors silently insert non-breaking and
; other exotic Unicode spaces -- U+00A0 (NBSP) is the usual culprit -- into
; measurement and label context, e.g. "12<NBSP>mm". A regex \s matches NONE of
; these, so any size/label parser written with \s silently misses the value;
; that is the recurring "size not detected" bug (LI-RADS etc. only worked
; because they parse via TextScan, which uses \h). Folding these code points
; down to a plain space ONCE, at the single capture point, makes every
; downstream parser -- \s or \h -- see clean ASCII whitespace, so individual
; calculators no longer have to remember to use \h.
NormalizeReportWhitespace(text) {
    if (text = "")
        return text
    ; Non-breaking / exotic SPACE characters -> a regular space.
    text := RegExReplace(text, "[\x{00A0}\x{2000}-\x{200A}\x{202F}\x{205F}\x{3000}]", " ")
    ; Zero-width characters -> removed. These must NOT become a space, or they
    ; would split words ("wash<ZWSP>out" -> "wash out") and break term matches.
    text := RegExReplace(text, "[\x{200B}\x{200C}\x{200D}\x{2060}\x{FEFF}]", "")
    return text
}

; --- math / measurements ----------------------------------------------------

; Sort a 3-element array of numbers in descending order, in place; returns it.
SortDimensions(d) {
    if (d[1] < d[2]) {
        t := d[1], d[1] := d[2], d[2] := t
    }
    if (d.Length >= 3 && d[2] < d[3]) {
        t := d[2], d[2] := d[3], d[3] := t
    }
    if (d[1] < d[2]) {
        t := d[1], d[1] := d[2], d[2] := t
    }
    return d
}

ExtractNumbers(text) {
    nums := []
    ; Strip leading labels like "Slice 4:" or "Sample 2:" so the label
    ; numbers do not get treated as data points.
    text := RegExReplace(text, "i)(?:^|\n|\s)(?:slice|observation|sample|number|#|no\.?)\s*\d+:?\s*", "`n")
    pos := 1
    while (pos := RegExMatch(text, "(-?\d+(?:\.\d+)?)(?:\s*(?:cm|mm)?)", &m, pos)) {
        nums.Push(m[1] + 0)
        pos += m.Len[0] ? m.Len[0] : 1
    }
    return nums
}

JoinArr(arr, sep) {
    out := ""
    for i, v in arr
        out .= (i > 1 ? sep : "") . v
    return out
}

; --- dates ------------------------------------------------------------------

; Accepts MM/DD/YYYY (with - or . as alternate separators). If the first
; number is > 12, treats input as DD/MM/YYYY. 2-digit years are expanded to
; 20YY. Returns "YYYYMMDD" or "" on failure.
ParseDate(dateStr) {
    s := StrReplace(StrReplace(dateStr, ".", "/"), "-", "/")
    if !RegExMatch(s, "^\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\s*$", &m)
        return ""
    month := m[1] + 0
    day   := m[2] + 0
    year  := m[3] + 0
    if (StrLen(m[3]) = 2)
        year += 2000
    if (month > 12) {
        t := month, month := day, day := t
    }
    if (month < 1 || month > 12 || day < 1 || day > 31)
        return ""
    return Format("{:04d}{:02d}{:02d}", year, month, day)
}

DateAddDays(yyyymmdd, days) {
    ; DateAdd needs YYYYMMDDHHMISS; pad if needed
    base := StrLen(yyyymmdd) >= 14 ? yyyymmdd : yyyymmdd . "000000"
    out := DateAdd(base, days, "days")
    return SubStr(out, 1, 8)
}

; days from d1 to d2 (positive if d2 > d1)
DaysBetween(d1, d2) {
    b1 := StrLen(d1) >= 14 ? d1 : d1 . "000000"
    b2 := StrLen(d2) >= 14 ? d2 : d2 . "000000"
    return DateDiff(b2, b1, "days")
}

; --- monitor / window placement ---------------------------------------------

; Returns { left, top, right, bottom, width, height } of the work area of the
; monitor containing the given screen point. Falls back to the primary
; monitor.
GetWorkAreaAt(x, y) {
    count := MonitorGetCount()
    loop count {
        MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
        if (x >= l && x <= r && y >= t && y <= b)
            return { left: l, top: t, right: r, bottom: b, width: r - l, height: b - t }
    }
    MonitorGetWorkArea(MonitorGetPrimary(), &l, &t, &r, &b)
    return { left: l, top: t, right: r, bottom: b, width: r - l, height: b - t }
}

ClampToWorkArea(x, y, w, h, work) {
    if (x + w > work.right)
        x := work.right - w
    if (y + h > work.bottom)
        y := work.bottom - h
    if (x < work.left)
        x := work.left
    if (y < work.top)
        y := work.top
    return { x: x, y: y }
}

; Unit contract (verified against the AHK v2.0 source, GuiType::Show):
; Gui.Show("x.. y.. w.. h..") treats X/Y as PHYSICAL screen pixels but
; multiplies W/H by A_ScreenDPI/96 -- and DPIScale is on by default for
; every Gui. Cursor and monitor coords (MouseGetPos, MonitorGetWorkArea)
; are physical. So clamping must use the window's PHYSICAL footprint
; (w * A_ScreenDPI/96) even though Show is then given the LOGICAL w/h.
; Getting this wrong is invisible at 100% scaling and pushes windows off
; the bottom/right edge at 125-175%.

; Fit a desired LOGICAL (w, h) into the active work area, leaving a
; margin, then place at the PHYSICAL anchor (x, y) and clamp inside.
; Returns { x, y, w, h } with x/y physical and w/h logical -- pass them
; straight to Gui.Show. Use this instead of ClampToWorkArea when the
; window itself may be larger than the monitor (e.g. small laptops).
; Caller decides whether to add scroll bars when the height is capped.
FitWindow(desiredW, desiredH, anchorX, anchorY, work, margin := 24) {
    scale := A_ScreenDPI / 96
    w := Min(desiredW, Floor((work.width  - margin) / scale))
    h := Min(desiredH, Floor((work.height - margin) / scale))
    pos := ClampToWorkArea(anchorX, anchorY, Round(w * scale), Round(h * scale), work)
    return { x: pos.x, y: pos.y, w: w, h: h, capped: (h < desiredH) }
}

; --- string distance --------------------------------------------------------

LevenshteinDistance(s, t) {
    m := StrLen(s), n := StrLen(t)
    if (m = 0)
        return n
    if (n = 0)
        return m
    d := []
    loop m + 1 {
        d.Push([])
        d[A_Index].Length := n + 1
        d[A_Index][1] := A_Index - 1
    }
    loop n + 1
        d[1][A_Index] := A_Index - 1
    loop m {
        i := A_Index
        loop n {
            j := A_Index
            cost := (SubStr(s, i, 1) = SubStr(t, j, 1)) ? 0 : 1
            d[i+1][j+1] := Min(d[i][j+1] + 1, d[i+1][j] + 1, d[i][j] + cost)
        }
    }
    return d[m+1][n+1]
}

FuzzyMatch(w1, w2, threshold := 2) {
    return LevenshteinDistance(w1, w2) <= threshold
}

; --- security: URL / filename validation ------------------------------------

IsValidURL(url) {
    ; Only http / https. No file://, javascript:, ftp:, etc.
    if !RegExMatch(url, "i)^https?://([\w\-]+\.)+[\w\-]+(?::\d+)?(?:[/?#].*)?$")
        return false
    return true
}

NormalizeURL(url) {
    ; Add https:// if no scheme present, then re-validate.
    if !RegExMatch(url, "i)^[a-z][a-z0-9+\-.]*://")
        url := "https://" . url
    return IsValidURL(url) ? url : ""
}

IsValidFileType(ext) {
    static valid := Map("pdf", 1, "doc", 1, "docx", 1, "xls", 1
                      , "xlsx", 1, "ppt", 1, "pptx", 1)
    e := StrLower(LTrim(ext, "."))
    return valid.Has(e)
}

; Convert any value to Integer, returning `default` on failure. v2's
; Integer() throws on empty/malformed strings; this never throws.
SafeInt(v, default := 0) {
    if (Type(v) = "Integer")
        return v
    if (Type(v) = "Float")
        return Integer(v)
    try
        return Integer(v)
    catch
        return default
}

; Returns a sanitized filename (no path separators, no traversal, no
; control chars) suitable for use as a destination filename. Returns ""
; if nothing usable remains.
SafeFilename(name) {
    name := RegExReplace(name, '[\\/:*?"<>|\x00-\x1F]', "")
    name := RegExReplace(name, "\.+", ".")
    name := Trim(name, ". ")
    if (name = "" || name = "." || name = "..")
        return ""
    ; Windows reserved device names (any extension) cannot be created.
    ; Prepend an underscore so the user's intent is preserved.
    SplitPath name, , , , &noExt
    static reserved := Map(
        "CON",1,"PRN",1,"AUX",1,"NUL",1,
        "COM1",1,"COM2",1,"COM3",1,"COM4",1,"COM5",1,
        "COM6",1,"COM7",1,"COM8",1,"COM9",1,
        "LPT1",1,"LPT2",1,"LPT3",1,"LPT4",1,"LPT5",1,
        "LPT6",1,"LPT7",1,"LPT8",1,"LPT9",1)
    if reserved.Has(StrUpper(noExt))
        name := "_" . name
    return SubStr(name, 1, 200)
}

GetUniqueFilePath(dir, fileName) {
    fileName := SafeFilename(fileName)
    if (fileName = "")
        return ""
    SplitPath fileName, , , &ext, &nameNoExt
    newPath := dir . "\" . fileName
    counter := 1
    while FileExist(newPath) {
        newPath := dir . "\" . nameNoExt . " (" . counter . ")"
                 . (ext != "" ? "." . ext : "")
        counter++
    }
    return newPath
}
