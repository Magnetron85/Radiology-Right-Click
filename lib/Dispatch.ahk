; ============================================================
; lib/Dispatch.ahk -- Smart-match calculator dispatcher
; ------------------------------------------------------------
; Scores each calculator against the highlighted text and returns
; the best candidates, so the context menu can pin a "Suggested"
; section at the top. The user always confirms with one click --
; nothing is auto-run.
;
; Scoring model (transparent, regex-based -- NOT machine learning,
; NOT edit-distance):
;   * Each calculator declares signals in three weights:
;       entity  (5) -- a distinctive anatomy / diagnosis / system
;                      term that, on its own, points at this tool
;                      ("adrenal", "carotid", "IPMN", "thyroid
;                      nodule"...).
;       feature (2) -- a supporting descriptor ("washout", "APHE",
;                      "hypoechoic", "Agatston"...).
;       data    (1) -- a measurement shape shared by many tools
;                      (three dimensions, HU values, a date...).
;   * Score = sum of matched weights. A calculator is suggested only
;     at score >= 2. Tier: "strong" needs an entity AND score >= 6;
;     otherwise "possible".
;
; Why regex and not Levenshtein: short medical terms have clinically
; DIFFERENT near-neighbours ("adrenal" vs "renal"), so blanket fuzzy
; matching would produce confident WRONG suggestions -- the one thing
; a clinical dispatcher must never do. Speech-recognition variability
; is instead absorbed by curated alternations in the patterns below
; (optional spaces / hyphens / dots, known misrecognitions like
; "speculated" for spiculated).
; ============================================================

#Requires AutoHotkey v2.0
#Include Util.ahk

; Return the ranked suggestions for `text`:
;   [{ cmd, score, tier, why:[terms] }, ...]   (highest score first)
; Up to 3 by score, plus Sort Measurement Sizes appended as a 4th when it
; matched but was outranked. Empty array when nothing scores at threshold.
MatchCalculators(text) {
    if (Trim(text) = "")
        return []

    results := []
    for cmd, signals in _DispatchSignatures() {
        score := 0
        hasEntity := false
        why := []
        for s in signals {
            if RegExMatch(text, "i)" s[1]) {
                score += s[2]
                why.Push(s[3])
                if (s[2] >= 5)
                    hasEntity := true
            }
        }
        if (score >= 2)
            results.Push({ cmd: cmd, score: score
                         , tier: (hasEntity && score >= 6) ? "strong" : "possible"
                         , why: why })
    }

    results := _DispatchSortByScore(results)
    ; Top 3 by score, with one exception: Sort Measurement Sizes is a
    ; transform, not a differential diagnosis -- whenever the selection
    ; contains a shape it can reorder it is applicable, but its low
    ; data-shape score lets entity-rich context ("prostate ...", "... in
    ; the liver") push it below rank 3, and the whole point of the suggested
    ; section is sparing the user the walk through the menu tree. When
    ; outranked it is appended as a fourth suggestion instead of dropped.
    out := []
    for i, r in results {
        if (i <= 3)
            out.Push(r)
        else if (r.cmd = "SortNoduleSizes") {
            out.Push(r)
            break
        }
    }
    return out
}

; Stable insertion sort, score descending. Ties keep insertion order
; (signature-map order), which is deterministic.
_DispatchSortByScore(arr) {
    sorted := []
    for r in arr {
        inserted := false
        for i, e in sorted {
            if (r.score > e.score) {
                sorted.InsertAt(i, r)
                inserted := true
                break
            }
        }
        if !inserted
            sorted.Push(r)
    }
    return sorted
}

; --- shared data-shape fragments (reused across calculators) ----------------
; Three orthogonal dimensions ("3 x 2 x 1", with optional decimals / units /
; non-breaking spaces). Two dimensions. A date. Three+ delimited numbers.
_DISP_3D   := "\d+(?:\.\d+)?\h*[x*\x{00D7}]\h*\d+(?:\.\d+)?\h*[x*\x{00D7}]\h*\d+(?:\.\d+)?"
_DISP_DATE := "\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}"
_DISP_NLIST := "(?:-?\d+(?:\.\d+)?[\s,]+){2,}-?\d+(?:\.\d+)?"

; Signature table: cmd (matches CalculatorRegistry .cmd) -> array of
; [pattern, weight, why-label]. Order here defines tie-break order.
_DispatchSignatures() {
    static m := _DispatchBuildSignatures()
    return m
}

_DispatchBuildSignatures() {
    global _DISP_3D, _DISP_DATE, _DISP_NLIST
    m := Map()

    ; Lung-specific location terms: upper/middle/lower lobes and the lobe
    ; abbreviations only exist in the lung (liver and thyroid lobes are
    ; bare right/left), so "nodule" near one of these is a LUNG nodule --
    ; the most common phrasing radiologists highlight ("6 mm nodule in the
    ; right upper lobe") never says the words "pulmonary nodule".
    lobe := "(?:R[UML]L\b|L[UL]L\b|lingula|(?:right|left)\s+(?:upper|middle|lower)\s+lobe|(?:right|left)\s+lung\b)"
    lungNod := "\bnodules?\b[^.\r\n]{0,80}\b" lobe "|\b" lobe "[^.\r\n]{0,80}\bnodules?"

    ; ---- measurement / volume (low confidence by design; no entity) ----
    m["CalculateEllipsoidVolume"] := [
        [_DISP_3D, 2, "three dimensions"],
        ["\bellipsoid\b", 5, "ellipsoid"],
        ["\bvolume\b", 2, "volume"]]
    m["CalculateBulletVolume"] := [
        [_DISP_3D, 2, "three dimensions"],
        ["\bbullet\b", 5, "bullet"],
        ["\bprostate\b", 2, "prostate"],
        ["\bvolume\b", 2, "volume"]]
    m["CompareNoduleSizes"] := [
        ["\b(?:previous(?:ly)?|prior|interval|compared\s+to)\b", 5, "prior comparison"],
        ["\b(?:now|current(?:ly)?|increased|decreased|enlarg)\b", 2, "interval change"],
        [_DISP_DATE, 1, "date"]]
    ; Patterns mirror what SortSizes_Entry can actually reorder ("A x B x C",
    ; "A, B, C", and 2D "A x B") rather than reusing _DISP_3D, whose ×/*
    ; separators the sort parser does not handle -- a suggestion must never
    ; lead to text the sorter cannot parse. The pair-or-triple alternation is
    ; ONE signal so a triple isn't double-counted. Constraints on the 2D
    ; forms: an x-pair needs a cm/mm unit, so matrix sizes ("256 x 512") and
    ; dosages ("2 x 20 mg") don't trigger it; bare comma PAIRS are excluded
    ; entirely ("May 3, 2026", "images 10, 12" would false-positive on
    ; ordinary prose).
    m["SortNoduleSizes"] := [
        ["\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?(?:\s*x\s*\d+(?:\.\d+)?|\s*(?:cm|mm)\b)", 2, "x-separated dimensions"],
        ["\d+(?:\.\d+)?\s*,\s*\d+(?:\.\d+)?\s*,\s*\d+(?:\.\d+)?", 2, "comma-separated sizes"]]

    ; ---- prostate ----
    m["CalculatePSADensity"] := [
        ["\bPSA\b", 5, "PSA"],
        ["\bprostate\b", 3, "prostate"],
        ["\bng\s*/\s*m[lL]\b", 2, "ng/mL"]]
    m["CalculatePIRADS"] := [
        ["\bPI-?RADS\b", 5, "PI-RADS"],
        ["\bprostate\b", 5, "prostate"],
        ["\b(?:peripheral|transition)\s+zone\b|\bPZ\b|\bTZ\b", 2, "zone"],
        ["\bDWI\b|\bADC\b|\bDCE\b", 2, "MR sequence"]]

    ; ---- adnexal / OB-GYN ----
    m["CalculatePregnancyDates"] := [
        ["\bLMP\b|\blast\s+menstrual\b", 5, "LMP"],
        ["\bgestational\s+age\b|\bGA\b|\bEGA\b|\bEDD\b|\bEDC\b|\bpregnan", 5, "gestational age"],
        ["\bcrown[- ]?rump\b|\bCRL\b", 3, "crown-rump length"],
        ["\bweeks?\b.*\bdays?\b", 1, "weeks/days"]]
    m["CalculateMenstrualPhase"] := [
        ["\bmenstrual\s+phase\b|\bcycle\s+day\b|\bendometri", 5, "menstrual phase"],
        ["\bLMP\b|\blast\s+menstrual\b", 3, "LMP"]]
    m["CalculateORADSMRI"] := [
        ["\badnex\w*\b|\bovar(?:y|ies|ian)\b", 5, "adnexal/ovarian"],
        ["\bMRI?\b|\bDCE\b|\btime[- ]?intensity\b|\bTIC\b", 2, "MRI"],
        ["\bO-?RADS\b", 2, "O-RADS"]]
    m["CalculateORADSUS"] := [
        ["\badnex\w*\b|\bovar(?:y|ies|ian)\b", 5, "adnexal/ovarian"],
        ["\b(?:ultrasound|sonograph|US)\b", 2, "ultrasound"],
        ["\b(?:uni|multi)locular\b|\bpapillary\s+projection\b", 2, "cyst morphology"],
        ["\bO-?RADS\b", 2, "O-RADS"]]

    ; ---- adrenal ----
    m["CalculateAdrenalWashout"] := [
        ["\badrenal\b", 5, "adrenal"],
        ["\bwashout\b|\bAPW\b|\bRPW\b", 3, "washout"],
        ["\bHU\b|\bhounsfield\b", 2, "HU"],
        ["\b(?:un|non[- ]?)enhanced\b|\bdelayed\b|\bportal\s+venous\b", 1, "CT phase"]]
    m["CalculateIncidentalAdrenal"] := [
        ["\badrenal\b", 5, "adrenal"],
        ["\b(?:nodule|mass|adenoma|incidental)\b", 2, "nodule/mass"],
        ["\bHU\b|\bhounsfield\b", 1, "HU"]]

    ; ---- thyroid / neck ----
    m["CalculateTIRADS"] := [
        ["\bTI-?RADS\b", 5, "TI-RADS"],
        ["\bthyroid\s+nodule\b|\bthyroid\b", 5, "thyroid nodule"],
        ["\bhypo[- ]?echoic\b|\bhyper[- ]?echoic\b|\biso[- ]?echoic\b|\becho[- ]?genic\b", 2, "echogenicity"],
        ["\bmicrocalcif|\bspongiform\b|\btaller[- ]than[- ]wide\b", 2, "morphology"]]
    m["CalculateIncidentalThyroid"] := [
        ["\bthyroid\s+nodule\b|\bthyroid\b", 5, "thyroid nodule"],
        ["\bincidental\b|\bPET\b|\bavid\b", 2, "incidental/PET"]]
    m["CalculateThymusChemicalShift"] := [
        ["\bthym(?:us|ic)\b", 5, "thymus"],
        ["\bchemical\s+shift\b|\b(?:in|out)[- ]?of?[- ]?phase\b|\bsignal\s+intensity\s+index\b", 2, "chemical shift"]]

    ; ---- liver / biliary ----
    m["CalculateLIRADS"] := [
        ["\bLI-?RADS\b", 5, "LI-RADS"],
        ["\bhepatic\s+observation\b|\bHCC\b|\bcirrho|\b(?:hepatic|liver)\s+(?:lesion|mass|nodule)\b", 5, "hepatic observation"],
        ["\bAPHE\b|\barterial\s+phase\s+hyper|\bwashout\b|\bcapsule\b", 2, "major feature"]]
    m["CalculateUSLIRADS"] := [
        ["\bUS\s+LI-?RADS\b|\bvisualization\s+score\b", 5, "US LI-RADS"],
        ["\b(?:hepatic|liver)\b.*\b(?:ultrasound|sonograph|surveillance)\b", 5, "liver US surveillance"]]
    m["CalculateHepaticSteatosis"] := [
        ["\bhepatic\s+steatosis\b|\b(?:liver|hepatic)\s+fat\b|\bfat\s+fraction\b|\bsteato|\bfatty\s+(?:liver|infiltrat)", 5, "hepatic steatosis"],
        ["\b(?:in|out)[- ]?of?[- ]?phase\b|\bDixon\b", 2, "in/out-of-phase"]]
    m["CalculateIronContent"] := [
        ["\b(?:liver\s+)?iron\b|\bhemochromatos|\bhemosideros|\bLIC\b", 5, "liver iron"],
        ["\bR2\s*\*|\bT2\s*\*", 3, "R2*/T2*"]]
    m["CalculateGBPolyp"] := [
        ["\bgallbladder\s+polyp\b|\bGB\s+polyp\b|\bgallbladder\b", 5, "gallbladder polyp"]]

    ; ---- pancreas ----
    m["CalculateKyotoIPMN"] := [
        ["\bIPMN\b|\bI\.P\.M\.N\b", 5, "IPMN"],
        ["\bpancreatic\s+cyst|\bmain\s+pancreatic\s+duct\b|\bMPD\b|\bpancrea\w*\b[^.\r\n]{0,40}\bcyst|\bcyst\w*\b[^.\r\n]{0,40}\bpancrea", 5, "pancreatic cyst/MPD"],
        ["\bmural\s+nodule\b|\bbranch[- ]?duct\b|\bmain[- ]?duct\b|\bside[- ]?branch\b", 2, "duct feature"]]

    ; ---- renal ----
    m["CalculateBosniak"] := [
        ["\bbosniak\b", 5, "Bosniak"],
        ; No trailing \b after "cyst" so "renal cystic lesion" matches too;
        ; the proximity branches catch "cystic lesion in the left kidney".
        ["\b(?:renal|kidney)\s+cyst|\bcystic\s+(?:renal|kidney)|\bcyst\w*\b[^.\r\n]{0,40}\bkidney|\bkidney\b[^.\r\n]{0,40}\bcyst", 5, "renal cyst"],
        ["\bsepta(?:tion)?\b|\bwall\s+enhanc", 2, "septa/wall"]]

    ; ---- lung ----
    m["CalculateFleischnerCriteria"] := [
        ["\bfleischner\b", 5, "Fleischner"],
        ["\bpulmonary\s+nodule\b|\blung\s+nodule\b|" lungNod, 5, "pulmonary nodule"],
        ["\bground[- ]?glass\b|\bGGN\b|\bpart[- ]?solid\b|\bsolid\s+nodule\b|\bperifissural\b|\bsubpleural\b", 2, "nodule type"]]
    m["CalculateLungRADS"] := [
        ["\blung-?rads\b", 5, "Lung-RADS"],
        ["\b(?:lung\s+(?:cancer\s+)?screening|LDCT|low[- ]?dose\s+CT)\b", 5, "lung screening"],
        ["\bpulmonary\s+nodule\b|\blung\s+nodule\b|" lungNod, 3, "pulmonary nodule"]]

    ; ---- cardiovascular ----
    m["CalculateCalciumScorePercentile"] := [
        ["\bcalcium\s+score\b|\bagatston\b|\bcoronary\s+(?:artery\s+)?calci|\bCAC\b", 5, "calcium score"],
        ["\bMESA\b|\bpercentile\b", 2, "percentile"]]
    m["CalculateRVLV"] := [
        ["\bRV\s*[:/]\s*LV\b|\bRV/LV\b", 5, "RV/LV"],
        ["\bright\s+ventric.*\bleft\s+ventric|\bleft\s+ventric.*\bright\s+ventric", 5, "RV and LV"],
        ["\bpulmonary\s+embol|\bPE\b|\bstrain\b", 2, "PE/strain"]]

    ; ---- neuro / head ----
    m["CalculateNASCET"] := [
        ["\bNASCET\b", 5, "NASCET"],
        ["\bcarotid\b|\bICA\b", 5, "carotid"],
        ["\bstenos", 2, "stenosis"]]
    m["CalculateICHVolume"] := [
        ["\bintra(?:cerebral|parenchymal)\s+h(?:a?emorrhage)?\b|\bICH\b|\bIPH\b|\bhematoma\b|\bABC/2\b", 5, "intracerebral hemorrhage"],
        [_DISP_3D, 1, "three dimensions"]]

    ; ---- numeric ----
    m["Statistics"] := [
        ["\b(?:mean|median|average|standard\s+deviation|std\s*dev)\b", 5, "statistic term"],
        [_DISP_NLIST, 2, "number list"]]
    m["Range"] := [
        ["\brange\b", 5, "range"],
        [_DISP_NLIST, 1, "number list"]]

    ; ---- scheduling ----
    m["CalculateContrastPremedication"] := [
        ["\bpremedicat|\bcontrast\s+allerg|\bprednisone\b|\bmethylprednisolone\b|\bsteroid\s+prep|\bdiphenhydramine\b|\bbenadryl\b", 5, "contrast premedication"]]
    m["CalculateFollowUpDate"] := [
        ["\bfollow[- ]?up\b.*\b\d+\s*(?:day|week|month|year)s?\b", 5, "follow-up interval"],
        ["\b\d+\s*(?:month|week|year)s?\b.*\bfollow[- ]?up\b", 5, "follow-up interval"],
        ["\brecommend\b.*\bfollow[- ]?up\b", 2, "recommend follow-up"]]

    return m
}
