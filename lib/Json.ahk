; ============================================================
; lib/Json.ahk -- Minimal JSON parser / stringifier (AHK v2)
; ------------------------------------------------------------
; Hand-rolled, audit-sized (no external dependency / supply
; chain). Maps -> objects, Arrays -> arrays. Booleans round
; trip via true/false. Numbers are returned as Number.
; ============================================================

#Requires AutoHotkey v2.0

class Json {

    static Parse(text) {
        ctx := { s: text, i: 1, n: StrLen(text) }
        Json._SkipWS(ctx)
        v := Json._Value(ctx)
        Json._SkipWS(ctx)
        if (ctx.i <= ctx.n)
            throw Error("Trailing content in JSON at pos " ctx.i)
        return v
    }

    static Stringify(value, indent := "") {
        out := ""
        Json._Write(value, &out, indent, "")
        return out
    }

    ; ---------- writer ----------
    static _Write(v, &out, indent, prefix) {
        if (v is Map) {
            if (v.Count = 0) {
                out .= "{}"
                return
            }
            childPrefix := prefix . indent
            out .= "{"
            first := true
            for k, val in v {
                if !first
                    out .= ","
                first := false
                if (indent != "")
                    out .= "`n" childPrefix
                out .= Json._Quote(String(k)) ":"
                if (indent != "")
                    out .= " "
                Json._Write(val, &out, indent, childPrefix)
            }
            if (indent != "")
                out .= "`n" prefix
            out .= "}"
        } else if (v is Array) {
            if (v.Length = 0) {
                out .= "[]"
                return
            }
            childPrefix := prefix . indent
            out .= "["
            first := true
            for val in v {
                if !first
                    out .= ","
                first := false
                if (indent != "")
                    out .= "`n" childPrefix
                Json._Write(val, &out, indent, childPrefix)
            }
            if (indent != "")
                out .= "`n" prefix
            out .= "]"
        } else if (Type(v) = "Integer" || Type(v) = "Float") {
            out .= String(v)
        } else if (Type(v) = "String") {
            out .= Json._Quote(v)
        } else {
            out .= "null"
        }
    }

    static _Quote(s) {
        s := StrReplace(s, "\", "\\")
        s := StrReplace(s, '"', '\"')
        s := StrReplace(s, "`n", "\n")
        s := StrReplace(s, "`r", "\r")
        s := StrReplace(s, "`t", "\t")
        s := StrReplace(s, Chr(8), "\b")
        s := StrReplace(s, Chr(12), "\f")
        ; Escape remaining ASCII control chars (0x00-0x1F minus those above)
        ; per RFC 8259. Most users will never trigger this; defensive only.
        loop 32 {
            c := A_Index - 1
            if (c = 8 || c = 9 || c = 10 || c = 12 || c = 13)
                continue
            s := StrReplace(s, Chr(c), Format("\u{:04x}", c))
        }
        return '"' s '"'
    }

    ; ---------- parser ----------
    static _SkipWS(ctx) {
        while (ctx.i <= ctx.n) {
            c := SubStr(ctx.s, ctx.i, 1)
            if (c = " " || c = "`t" || c = "`n" || c = "`r")
                ctx.i++
            else
                break
        }
    }

    static _Value(ctx) {
        Json._SkipWS(ctx)
        if (ctx.i > ctx.n)
            throw Error("Unexpected end of JSON")
        c := SubStr(ctx.s, ctx.i, 1)
        if (c = "{")
            return Json._Object(ctx)
        if (c = "[")
            return Json._Array(ctx)
        if (c = '"')
            return Json._String(ctx)
        if (c = "t" || c = "f")
            return Json._Bool(ctx)
        if (c = "n")
            return Json._Null(ctx)
        if (c = "-" || InStr("0123456789", c))
            return Json._Number(ctx)
        throw Error("Unexpected char '" c "' at pos " ctx.i)
    }

    static _Object(ctx) {
        m := Map()
        ctx.i++  ; consume '{'
        Json._SkipWS(ctx)
        if (SubStr(ctx.s, ctx.i, 1) = "}") {
            ctx.i++
            return m
        }
        loop {
            Json._SkipWS(ctx)
            if (SubStr(ctx.s, ctx.i, 1) != '"')
                throw Error("Expected string key at pos " ctx.i)
            key := Json._String(ctx)
            Json._SkipWS(ctx)
            if (SubStr(ctx.s, ctx.i, 1) != ":")
                throw Error("Expected ':' at pos " ctx.i)
            ctx.i++
            val := Json._Value(ctx)
            m[key] := val
            Json._SkipWS(ctx)
            c := SubStr(ctx.s, ctx.i, 1)
            if (c = ",") {
                ctx.i++
                continue
            }
            if (c = "}") {
                ctx.i++
                return m
            }
            throw Error("Expected ',' or '}' at pos " ctx.i)
        }
    }

    static _Array(ctx) {
        a := []
        ctx.i++  ; consume '['
        Json._SkipWS(ctx)
        if (SubStr(ctx.s, ctx.i, 1) = "]") {
            ctx.i++
            return a
        }
        loop {
            val := Json._Value(ctx)
            a.Push(val)
            Json._SkipWS(ctx)
            c := SubStr(ctx.s, ctx.i, 1)
            if (c = ",") {
                ctx.i++
                continue
            }
            if (c = "]") {
                ctx.i++
                return a
            }
            throw Error("Expected ',' or ']' at pos " ctx.i)
        }
    }

    static _String(ctx) {
        ctx.i++  ; consume opening quote
        out := ""
        while (ctx.i <= ctx.n) {
            c := SubStr(ctx.s, ctx.i, 1)
            if (c = '"') {
                ctx.i++
                return out
            }
            if (c = "\") {
                ctx.i++
                esc := SubStr(ctx.s, ctx.i, 1)
                ctx.i++
                switch esc {
                    case '"': out .= '"'
                    case "\": out .= "\"
                    case "/": out .= "/"
                    case "b": out .= Chr(8)
                    case "f": out .= Chr(12)
                    case "n": out .= "`n"
                    case "r": out .= "`r"
                    case "t": out .= "`t"
                    case "u":
                        hex := SubStr(ctx.s, ctx.i, 4)
                        ctx.i += 4
                        out .= Chr(Integer("0x" hex))
                    default:
                        throw Error("Bad escape \\" esc " at pos " ctx.i)
                }
            } else {
                out .= c
                ctx.i++
            }
        }
        throw Error("Unterminated string")
    }

    static _Number(ctx) {
        static numChars := "0123456789.eE+-"
        start := ctx.i
        if (SubStr(ctx.s, ctx.i, 1) = "-")
            ctx.i++
        while (ctx.i <= ctx.n) {
            c := SubStr(ctx.s, ctx.i, 1)
            if InStr(numChars, c)
                ctx.i++
            else
                break
        }
        token := SubStr(ctx.s, start, ctx.i - start)
        if (token = "" || token = "-")
            throw Error("Malformed JSON number at position " start)
        try {
            if (InStr(token, ".") || InStr(token, "e") || InStr(token, "E"))
                return token + 0.0
            return Integer(token)
        } catch as e {
            throw Error("Cannot parse JSON number '" token "' at position " start ": " e.Message)
        }
    }

    static _Bool(ctx) {
        if (SubStr(ctx.s, ctx.i, 4) = "true") {
            ctx.i += 4
            return true
        }
        if (SubStr(ctx.s, ctx.i, 5) = "false") {
            ctx.i += 5
            return false
        }
        throw Error("Bad literal at pos " ctx.i)
    }

    static _Null(ctx) {
        if (SubStr(ctx.s, ctx.i, 4) = "null") {
            ctx.i += 4
            return ""
        }
        throw Error("Bad literal at pos " ctx.i)
    }
}
