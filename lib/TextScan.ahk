; ============================================================
; lib/TextScan.ahk -- Shared regex extractors for pre-filling
;                     RADS dialogs from highlighted report text.
; ------------------------------------------------------------
; Each TextScan.X(input) reads a free-text snippet (typically
; the user's highlighted PowerScribe selection) and returns a
; structured value the form helpers can use as a default. Every
; extractor is defensive: returns "" (or 0 for numeric fallback)
; when the field can't be detected, and never throws.
;
; Coverage:
;   Size(text)                -> { mm, cm }    (first dimension found)
;   Age(text)                 -> integer       ("Age: 55", "55-year-old")
;   Sex(text)                 -> "Male" / "Female"
;   Race(text)                -> "White" / "Black" / "Hispanic" / "Chinese"
;   HU(text, phase)           -> integer       (per CT phase)
;   LMP(text)                 -> date string
;   CalciumScore(text)        -> number
;   R2Star(text)              -> number
;   FieldStrength(text)       -> "1.5" / "2.89" / "3.0"
;   Composition(text)         -> "solid" / "ground glass" / "part solid"
;                              / "cystic" / "spongiform" / "mixed"
;   Echogenicity(text)        -> "anechoic" / "hyperechoic" / "isoechoic"
;                              / "hypoechoic" / "very hypoechoic"
;   Shape(text)               -> "taller" / "wider"
;   Margin(text)              -> "smooth" / "lobulated" / "irregular"
;                              / "extrathyroidal" / "ill-defined"
;   ZoneProstate(text)        -> "PZ" / "TZ"
;   ContainsAny(text, list)   -> bool (case-insensitive whole-string)
;
; Whitespace convention: every regex in this file uses \h (horizontal
; whitespace) for token separators, NOT \s. PowerScribe silently inserts
; non-breaking spaces (U+00A0) in measurement and label contexts; \s only
; matches ASCII whitespace and would miss them. \h matches NBSP, NARROW
; NBSP, tab, etc. -- a strict superset of what we want and a strict subset
; of \s (does not match \n / \r), so substitution is safe here.
; ============================================================

#Requires AutoHotkey v2.0

class TextScan {

    ; Returns { mm: N, cm: N } for the first standalone measurement found.
    ; Recognizes "12 mm", "12mm", "1.5 cm", "1.5cm". Multidimensional inputs
    ; (e.g. "3 x 2 x 1 cm") return the LARGEST dimension.
    static Size(text) {
        out := { mm: 0, cm: 0.0 }
        if (text = "")
            return out

        ; Try multi-dimensional first ("a x b x c unit"). Use \h instead
        ; of \s so non-breaking spaces (U+00A0) match -- PowerScribe and
        ; other rich-text editors silently use them in measurement context.
        if RegExMatch(text
            , "i)(\d+(?:\.\d+)?)\h*[x×]\h*(\d+(?:\.\d+)?)(?:\h*[x×]\h*(\d+(?:\.\d+)?))?\h*(mm|cm)"
            , &m) {
            dims := [m[1] + 0.0, m[2] + 0.0]
            if (m[3] != "")
                dims.Push(m[3] + 0.0)
            maxd := dims[1]
            for d in dims
                if (d > maxd)
                    maxd := d
            if (m[4] = "cm") {
                out.cm := maxd
                out.mm := maxd * 10
            } else {
                out.mm := maxd
                out.cm := maxd / 10
            }
            return out
        }

        ; Single dimension. \h handles non-breaking spaces as above.
        if RegExMatch(text, "i)(\d+(?:\.\d+)?)\h*(mm|cm)\b", &m) {
            v := m[1] + 0.0
            if (m[2] = "cm") {
                out.cm := v
                out.mm := v * 10
            } else {
                out.mm := v
                out.cm := v / 10
            }
        }
        return out
    }

    ; Returns an integer age, 0 if not found. Accepts "Age: 55", "55-year-old",
    ; "55 yo", "55 y/o".
    static Age(text) {
        if (text = "")
            return 0
        if RegExMatch(text, "i)\bage\h*[:=]?\h*(\d{1,3})\b", &m)
            return m[1] + 0
        if RegExMatch(text, "i)\b(\d{1,3})[\h-]*(?:year-?old|yo|y\.?o\.?)\b", &m)
            return m[1] + 0
        return 0
    }

    ; Returns "Male" or "Female"; "" if not found.
    static Sex(text) {
        if (text = "")
            return ""
        if RegExMatch(text, "i)\bsex\h*[:=]?\h*(male|female)\b", &m)
            return _Cap(m[1])
        if RegExMatch(text, "i)\b(male|female)\h+patient\b", &m)
            return _Cap(m[1])
        ; pronoun fallback (lower confidence)
        if RegExMatch(text, "i)\b(?:she|her|woman|female)\b")
            return "Female"
        if RegExMatch(text, "i)\b(?:he|his|him|man|male)\b")
            return "Male"
        return ""
    }

    static Race(text) {
        if (text = "")
            return ""
        if RegExMatch(text, "i)\brace\h*[:=]?\h*(white|black|hispanic|chinese|asian)\b", &m) {
            r := _Cap(m[1])
            return (r = "Asian") ? "Chinese" : r
        }
        return ""
    }

    ; Returns HU value for the requested CT phase, or "" if not found.
    ; phase is one of "unenhanced" | "enhanced" | "delayed".
    ; All patterns anchor on a word boundary (\b) to prevent "enhanced" from
    ; matching inside "Unenhanced", and to keep "delayed" from matching
    ; "underlying" or similar substrings.
    static HU(text, phase) {
        if (text = "")
            return ""
        switch phase {
            case "unenhanced":
                pat := "i)\b(?:unenhanced|non-?enhanced|non-?contrast|baseline|pre-?contrast|native)(?:\h+(?:CT|HU|density))?[\h:=]+(-?\d+(?:\.\d+)?)\h*(?:HU)?"
            case "enhanced":
                pat := "i)\b(?:enhanced|post-?contrast|portal\h+venous|arterial|60-?75\h*sec|1-?2\h*min)(?:\h+(?:CT|HU|density))?[\h:=]+(-?\d+(?:\.\d+)?)\h*(?:HU)?"
            case "delayed":
                pat := "i)\b(?:delayed|15\h*min|10-?15\h*min|late)(?:\h+(?:CT|HU|density))?[\h:=]+(-?\d+(?:\.\d+)?)\h*(?:HU)?"
            default:
                return ""
        }
        if RegExMatch(text, pat, &m)
            return m[1] + 0.0
        return ""
    }

    ; Returns LMP date "MM/DD/YYYY" if found, else "".
    static LMP(text) {
        if (text = "")
            return ""
        if RegExMatch(text, "i)(?:LMP|Last\h*Menstrual\h*Period)\h*[:=]?\h*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})", &m)
            return m[1]
        return ""
    }

    ; Returns Agatston score number or "".
    static CalciumScore(text) {
        if (text = "")
            return ""
        pat := "i)(?:coronary\h+artery\h+)?calcium\h+score(?:\h+is)?\h*[:=]?\h*(\d+(?:\.\d+)?)\h*(?:\(?\h*Agatston\h*\)?)?"
        if RegExMatch(text, pat, &m)
            return m[1] + 0.0
        return ""
    }

    static R2Star(text) {
        if (text = "")
            return ""
        ; \h (not \s) so non-breaking spaces don't defeat the parse -- same
        ; reasoning as TextScan.Size, line 44-46.
        if RegExMatch(text, "i)R2\*?\h*(?:value|measurement)?[\h:=]+(\d+(?:[.,]\d+)?)\h*(?:Hz|hertz|/s|1/s)", &m)
            return StrReplace(m[1], ",", ".") + 0.0
        return ""
    }

    ; Returns "1.5" / "2.89" / "3.0" or "".
    static FieldStrength(text) {
        if (text = "")
            return ""
        if RegExMatch(text, "i)\b(1[.,]5|2[.,]89|3[.,]0)\h*T(?:esla)?\b", &m)
            return StrReplace(m[1], ",", ".")
        return ""
    }

    static Composition(text) {
        if (text = "")
            return ""
        ; order matters -- check more specific terms first
        if RegExMatch(text, "i)\bpart[- ]?solid\b")
            return "part solid"
        if RegExMatch(text, "i)\bground[- ]?glass\b|\bGGN\b|\bGGO\b")
            return "ground glass"
        if RegExMatch(text, "i)\bspongiform\b")
            return "spongiform"
        if RegExMatch(text, "i)\bmixed\h+(?:cystic|solid)\b|\bcystic\h+and\h+solid\b")
            return "mixed"
        if RegExMatch(text, "i)\bsolid\b")
            return "solid"
        if RegExMatch(text, "i)\bcystic\b|\bsimple\h+cyst\b")
            return "cystic"
        return ""
    }

    static Echogenicity(text) {
        if (text = "")
            return ""
        if RegExMatch(text, "i)\bvery\h+hypoechoic\b")
            return "very hypoechoic"
        if RegExMatch(text, "i)\banechoic\b")
            return "anechoic"
        if RegExMatch(text, "i)\bhypoechoic\b")
            return "hypoechoic"
        if RegExMatch(text, "i)\bhyperechoic\b")
            return "hyperechoic"
        if RegExMatch(text, "i)\bisoechoic\b")
            return "isoechoic"
        return ""
    }

    static Shape(text) {
        if (text = "")
            return ""
        if RegExMatch(text, "i)\btaller[- ]than[- ]wide\b|\btaller\h+than\h+wide\b")
            return "taller"
        if RegExMatch(text, "i)\bwider[- ]than[- ]tall\b|\bwider\h+than\h+tall\b")
            return "wider"
        return ""
    }

    static Margin(text) {
        if (text = "")
            return ""
        if RegExMatch(text, "i)\bextra[- ]?thyroidal\b|\bextra[- ]?thyroidal\h+extension\b|\bETE\b")
            return "extrathyroidal"
        if RegExMatch(text, "i)\blobulated\b|\bspiculat|\birregular\h+margin")
            return "lobulated"
        if RegExMatch(text, "i)\bill[- ]defined\b|\bindistinct\b")
            return "ill-defined"
        if RegExMatch(text, "i)\bsmooth\h+margin|\bsmooth\h+contour|\bwell[- ]defined\b|\bwell[- ]circumscribed\b")
            return "smooth"
        return ""
    }

    static ZoneProstate(text) {
        if (text = "")
            return ""
        if RegExMatch(text, "i)\bperipheral\h+zone\b|\bPZ\b")
            return "PZ"
        if RegExMatch(text, "i)\btransition\h+zone\b|\bTZ\b")
            return "TZ"
        return ""
    }

    static ContainsAny(text, patterns) {
        if (text = "")
            return false
        for p in patterns {
            if RegExMatch(text, "i)" p)
                return true
        }
        return false
    }
}

_Cap(s) {
    return StrUpper(SubStr(s, 1, 1)) StrLower(SubStr(s, 2))
}
