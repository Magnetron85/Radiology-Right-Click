; ============================================================
; lib/Debug.ahk -- Conditional logging framework
; ------------------------------------------------------------
; Off by default. To enable, edit preferences.json:
;
;   "debug": {
;       "capture": true,
;       "menu": true,
;       ...
;   }
;
; Each category is independent. When a category is true, its
; messages are appended to debug.log (next to the script). When
; false (or missing), Debug.Log / LogHex are no-ops -- the only
; runtime cost is one Prefs map lookup per call.
;
; Usage:
;   Debug.Log("menu",    "context menu opened")
;   Debug.Log("capture", "got " StrLen(text) " chars")
;   Debug.LogHex("capture", "raw clipboard", text)   ; line + hex dump
;
; Adding a new category: just call Debug.Log with a new name and
; set "debug": {"<name>": true} in preferences.json to enable it.
; ============================================================

#Requires AutoHotkey v2.0

class Debug {

    ; True if logging is enabled for `category`. Cheap -- just a Prefs lookup.
    static Enabled(category) {
        return !!Prefs.Get("debug", category, false)
    }

    ; Append a single line to debug.log. No-op when the category is off.
    static Log(category, msg) {
        if !Debug.Enabled(category)
            return
        Debug._Write(category, msg)
    }

    ; Append a line plus a hex dump of `text` (first 30 code units). Useful
    ; for diagnosing invisible characters in captured input -- non-breaking
    ; spaces, zero-width joiners, BOMs, smart quotes etc. all surface here.
    static LogHex(category, msg, text) {
        if !Debug.Enabled(category)
            return
        line := msg
        if (text != "") {
            line .= "  text=|" SubStr(text, 1, 60) "|  hex="
            loop Min(StrLen(text), 30)
                line .= Format(" {:04X}", Ord(SubStr(text, A_Index, 1)))
        }
        Debug._Write(category, line)
    }

    ; Internal: timestamp + category + line, appended to debug.log. Failures
    ; are swallowed so logging can never break the caller.
    static _Write(category, line) {
        try {
            ts := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
            FileAppend Format("[{}] [{}] {}`r`n", ts, category, line)
                     , A_ScriptDir "\debug.log"
        } catch {
            ; logging must never throw
        }
    }
}
