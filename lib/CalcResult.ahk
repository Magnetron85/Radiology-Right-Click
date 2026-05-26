; ============================================================
; lib/CalcResult.ahk -- Structured calculator output
; ------------------------------------------------------------
; Replaces the legacy flat-string return value with a richer
; struct that lets ShowResult render different regions with
; different visibility/clipboard rules:
;
;   classification    short label ("Lung-RADS 4A", "Low risk")
;   impression        one-sentence prose, the default paste target
;   recommendation    imperative-voice f/u, often duplicated in impression
;   findings          fragment for the Findings section (optional)
;   advisories[]      one-line caveats; not pasted by default
;   methodology       audit-only verbose explanation
;   citations[]       array of { text, url } -- ACR / DOI links
;   echo              raw user selection, shown only when transformed
;   error             error message; if non-empty, only impression renders
;
; Every field is optional. Use MakeResult({}) to construct; the
; helper fills missing fields with safe defaults so RenderResult
; never has to check existence.
;
; RenderResult composes the final display text from a CalcResult
; plus an opts map { showMethodology, showCitations }. The same
; renderer is used for the on-screen popup and the clipboard --
; what the user sees is what they (potentially) copy.
; ============================================================

#Requires AutoHotkey v2.0

; Construct a CalcResult with safe defaults. Pass a Map or
; props-style object; missing fields get filled in.
MakeResult(fields) {
    out := { classification: ""
           , impression:     ""
           , recommendation: ""
           , findings:       ""
           , advisories:     []
           , methodology:    ""
           , citations:      []
           , echo:           ""
           , error:          "" }
    ; Copy any provided fields onto the result. Accept either a
    ; plain object (props) or a Map.
    if (fields is Map) {
        for k, v in fields
            out.%k% := v
    } else if IsObject(fields) {
        for k in out.OwnProps() {
            if fields.HasOwnProp(k)
                out.%k% := fields.%k%
        }
    }
    return out
}

; Wrap a legacy flat-string return into a CalcResult. Used by the
; back-compat path in ShowResult so calcs that haven't been migrated
; yet still display correctly.
LegacyResult(text) {
    return MakeResult({ impression: text })
}

; True if the value is a CalcResult-shaped object. We discriminate
; on a sentinel property rather than Type() because AHK v2 plain
; objects all report as "Object".
IsCalcResult(v) {
    if !IsObject(v)
        return false
    try
        return v.HasOwnProp("impression") && v.HasOwnProp("methodology")
    catch
        return false
}

; ---- Rendering ----------------------------------------------------

; Compose the display text. `opts` controls section visibility:
;   showMethodology  bool
;   showCitations    bool   (also gates the references block / links)
;
; Output regions in order:
;   error (if set)        -> just the error message, nothing else
;   echo                  -> "> <verbatim user text>"
;   impression            -> always shown
;   advisories            -> always shown
;   methodology           -> only if opts.showMethodology
;   citations             -> only if opts.showCitations; URLs included inline
;
; Sanitization (no smart quotes / NBSP / em-dashes etc.) is applied
; to every text-bearing field on the way out -- see _Sanitize. The
; goal is bit-clean ASCII suitable for PowerScribe and downstream
; EHR rendering.
RenderResult(result, opts := "") {
    if !IsCalcResult(result)
        return _Sanitize(IsObject(result) ? "" : result)

    showMethod := IsObject(opts) && opts.HasOwnProp("showMethodology")
                  ? opts.showMethodology : true
    showCit    := IsObject(opts) && opts.HasOwnProp("showCitations")
                  ? opts.showCitations : true

    if (result.error != "")
        return _Sanitize(result.error)

    parts := []

    if (result.echo != "")
        parts.Push("> " _Sanitize(result.echo))

    if (result.impression != "")
        parts.Push(_Sanitize(result.impression))

    if (result.advisories is Array && result.advisories.Length > 0) {
        ad := ""
        for line in result.advisories
            ad .= (ad = "" ? "" : "`n") "Note: " _Sanitize(line)
        parts.Push(ad)
    }

    if (showMethod && result.methodology != "") {
        sep := "---------- Methodology ----------"
        parts.Push(sep "`n" _Sanitize(result.methodology))
    }

    if (showCit && result.citations is Array && result.citations.Length > 0) {
        ref := "---------- References ----------"
        for c in result.citations {
            t := IsObject(c) && c.HasOwnProp("text") ? c.text : (c is String ? c : "")
            u := IsObject(c) && c.HasOwnProp("url")  ? c.url  : ""
            if (t = "" && u = "")
                continue
            ref .= "`n" _Sanitize(t)
            if (u != "")
                ref .= "`n  " u
        }
        parts.Push(ref)
    }

    out := ""
    for p in parts
        out .= (out = "" ? "" : "`n`n") p
    return out
}

; Compose the clipboard payload. Default is impression + recommendation
; (when distinct), nothing else -- never echo, never methodology, never
; citations. Radiologists pasting into clinical reports should not have
; to strip noise. An "Copy with methodology" caller can override by
; passing opts.includeAll := true, which routes through RenderResult.
ClipboardText(result, opts := "") {
    if !IsCalcResult(result)
        return _Sanitize(IsObject(result) ? "" : result)
    if (IsObject(opts) && opts.HasOwnProp("includeAll") && opts.includeAll)
        return RenderResult(result, opts)
    if (result.error != "")
        return _Sanitize(result.error)
    out := _Sanitize(result.impression)
    ; Append recommendation only if it isn't already embedded in the
    ; impression sentence (avoids double-pasting).
    if (result.recommendation != ""
        && !InStr(result.impression, result.recommendation))
        out .= "`n" _Sanitize(result.recommendation)
    return out
}

; ---- Sanitization -------------------------------------------------
; Strip Unicode punctuation that PowerScribe / EHRs render as "?" or
; boxes, normalize whitespace, drop trailing newlines.
_Sanitize(s) {
    if (s = "")
        return ""
    ; Smart quotes -> straight ASCII
    s := StrReplace(s, Chr(0x2018), "'")  ; left single quote
    s := StrReplace(s, Chr(0x2019), "'")  ; right single quote / apostrophe
    s := StrReplace(s, Chr(0x201C), '"')  ; left double quote
    s := StrReplace(s, Chr(0x201D), '"')  ; right double quote
    ; Dashes -> double-hyphen
    s := StrReplace(s, Chr(0x2013), "-")  ; en-dash
    s := StrReplace(s, Chr(0x2014), "--") ; em-dash
    ; Ellipsis -> three dots
    s := StrReplace(s, Chr(0x2026), "...")
    ; Non-breaking and other unicode spaces -> ASCII space
    s := StrReplace(s, Chr(0x00A0), " ")
    s := StrReplace(s, Chr(0x202F), " ")
    s := StrReplace(s, Chr(0x2009), " ")
    s := StrReplace(s, Chr(0x205F), " ")
    ; Bullet -> dash
    s := StrReplace(s, Chr(0x2022), "-")
    ; Trim trailing newlines / whitespace
    return RTrim(s, " `t`r`n")
}
