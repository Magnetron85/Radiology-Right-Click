; ============================================================
; lib/calc/ORADSMRI.ahk -- O-RADS MRI risk stratification
; ------------------------------------------------------------
; Reference: ACR O-RADS MRI Risk Stratification and Management
; System (source-of-truth table: references/orads mri.md).
;
; Lesion types follow the O-RADS MRI lexicon exactly:
;   - No lesion
;   - Unilocular cyst        (incl. the physiologic Score-1 rows:
;                             follicle / hemorrhagic cyst / corpus
;                             luteum +/- hemorrhage, each <=3 cm in
;                             a PREMENOPAUSAL patient)
;   - Multilocular cyst      (no lipid; irregular enhancing septae
;                             or wall = solid tissue by definition)
;   - Lesion with lipid content (dermoid / teratoma)
;   - Lesion with solid tissue  (solid tissue = ENHANCING papillary
;                             projection, mural nodule, irregular
;                             septation/wall, or larger solid portion)
;   - Dilated fallopian tube
;   - Para-ovarian cyst
;
; Scoring rows implemented 1:1 from the table:
;   peritoneal/mesenteric/omental nodularity        -> 5 (overrides all)
;   no lesion                                        -> 1
;   unilocular simple|hemorrhagic <=3cm premenopausal-> 1 (physiologic)
;   unilocular any fluid, NO wall enhancement        -> 2
;   unilocular simple/endometriotic + smooth enh wall-> 2
;   unilocular protein./hemorrhagic/mucinous + wall  -> 3
;   multilocular smooth septae+wall, no lipid        -> 3
;   multilocular irregular septae/wall               -> solid pathway
;   lipid lesion, no enhancing solid                 -> 2
;   lipid lesion + LARGE-VOLUME enhancing solid      -> 4
;   solid tissue dark T2 AND dark DWI                -> 2
;   solid tissue + TIC 1 / 2 / 3                     -> 3 / 4 / 5
;   solid tissue, non-DCE: enh <= myometrium 30-40s  -> 4
;   solid tissue, non-DCE: enh >  myometrium 30-40s  -> 5
;   solid tissue, no TIC assessment                  -> 4 (default)
;   tube simple fluid + thin smooth wall/folds       -> 2
;   tube non-simple fluid (thin wall) OR thick wall  -> 3
;   para-ovarian, thin smooth wall +/- enhancement   -> 2
;
; Ascites WITHOUT peritoneal nodularity is not an independent
; scoring criterion ("with or without ascites" attaches to the
; nodularity row) -- it is captured for the report but does not
; change the score; an advisory flags it.
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\FormGui.ahk

ORADSMRI_Entry(input) {
    ShowORADSMRIDialog(input)
    return ""
}

ShowORADSMRIDialog(text := "") {
    ; ---- lesion-type prefill (specific terms before generic ones) ----
    if TextScan.ContainsAny(text, ["\bno\s+(?:adnexal|ovarian)\s+lesion\b","\bnormal\s+ovari"])
        typeIdx := 1
    else if TextScan.ContainsAny(text, ["\bpara[- ]?ovarian\b","\bpara[- ]?tubal\b"])
        typeIdx := 7
    else if TextScan.ContainsAny(text, ["\bhydrosalpinx\b","\bhematosalpinx\b","\bpyosalpinx\b","\bdilated\s+fallopian\s+tube\b","\bfallopian\s+tube\b"])
        typeIdx := 6
    else if TextScan.ContainsAny(text, ["\bdermoid\b","\bteratoma\b","\bfat[- ]containing\b","\bmacroscopic\s+fat\b","\blipid\b"])
        typeIdx := 4
    else if TextScan.ContainsAny(text, ["\bmultilocular\b","\bmultiloculated\b"])
        typeIdx := 3
    else if TextScan.ContainsAny(text, ["\bunilocular\b","\bsimple\s+cyst\b","\bhemorrhagic\s+cyst\b","\bendometrioma\b"])
        typeIdx := 2
    else
        typeIdx := 5   ; default: "Lesion with solid tissue"

    ; ---- feature prefills ----
    sz := TextScan.Size(text)
    menoIdx := 1
    if TextScan.ContainsAny(text, ["\bpost[- ]?menopausal\b"])
        menoIdx := 3
    else if TextScan.ContainsAny(text, ["\bpre[- ]?menopausal\b"])
        menoIdx := 2
    fluidIdx := 1
    if TextScan.ContainsAny(text, ["\bhemorrhagic\b","\bcorpus\s+luteum\b"])
        fluidIdx := 4
    else if TextScan.ContainsAny(text, ["\bendometrio"])
        fluidIdx := 3
    else if TextScan.ContainsAny(text, ["\bproteinaceous\b","\bmucinous\b"])
        fluidIdx := 5
    else if TextScan.ContainsAny(text, ["\bsimple\s+(?:fluid|cyst)\b"])
        fluidIdx := 2
    wallEnh := TextScan.ContainsAny(text, ["\bwall\s+enhanc","\benhancing\s+wall\b"])
    septIdx := 1
    if TextScan.ContainsAny(text, ["\birregular\s+sept","\bnodular\s+sept"])
        septIdx := 3
    else if TextScan.ContainsAny(text, ["\bsmooth\s+(?:thin\s+)?sept","\bthin\s+sept"])
        septIdx := 2
    perit := TextScan.ContainsAny(text, ["\bperitoneal\s+(?:deposit|nodul|implant)","\bomental\s+caking?","\bcarcinomatos"])
    asc   := TextScan.ContainsAny(text, ["\bascites\b","\bfree\s+pelvic\s+fluid\b"])

    form := RadsForm("O-RADS MRI", 560)

    ; Decision-tree layout with progressive disclosure: pick the lexicon
    ; lesion type, then only that type's questions appear. Menopausal
    ; status + size live in the unilocular section because they are the
    ; gate for the physiologic Score-1 rows -- the ONLY place in the
    ; O-RADS MRI table where patient status changes the score.
    form.Header("Lesion")
    form.Dropdown("LType", "Lesion type:"
        , ["No lesion"
        ,  "Unilocular cyst"
        ,  "Multilocular cyst (no lipid content)"
        ,  "Lesion with lipid content (dermoid / teratoma)"
        ,  "Lesion with solid tissue"
        ,  "Dilated fallopian tube"
        ,  "Para-ovarian cyst"], typeIdx)
    form.Numeric("SizeCm", "Largest diameter (cm, 0 if not measured):"
        , sz.cm > 0 ? Round(sz.cm, 1) : 0)

    form.Header("Unilocular cyst features")
    form.Dropdown("Meno", "Menopausal status:"
        , ["Not specified"
        ,  "Premenopausal"
        ,  "Postmenopausal"], menoIdx)
    form.Dropdown("FluidType", "Fluid content:"
        , ["Not specified"
        ,  "Simple"
        ,  "Endometriotic (endometrioma)"
        ,  "Hemorrhagic (incl. corpus luteum +/- hemorrhage)"
        ,  "Proteinaceous or mucinous"], fluidIdx)
    form.Checkbox("WallEnh", "Smooth enhancing wall", wallEnh)
    form.Note("Simple or hemorrhagic cyst <=3 cm in a PREMENOPAUSAL patient is physiologic -> O-RADS MRI 1 (needs status + size above). Irregular enhancing wall = solid tissue -- use 'Lesion with solid tissue'.")

    form.Header("Multilocular cyst features")
    form.Dropdown("Septae", "Septae / wall morphology:"
        , ["Not specified"
        ,  "Smooth septae and wall (+/- enhancement)"
        ,  "Irregular septae / wall, enhancing (= solid tissue)"], septIdx)
    form.Note("Irregular enhancing septations or wall meet the solid-tissue definition -- the solid-tissue questions below then apply. If lipid content is present, use the lipid lesion type instead.")

    form.Header("Lipid lesion features")
    form.Checkbox("LargeVolSolid", "Large-volume enhancing solid tissue (-> Score 4)", false)
    form.Note("Minimal enhancement of a Rokitansky nodule does NOT upgrade the lesion to Score 4 (source footnote).")

    form.Header("Solid tissue characterization")
    form.Note("Solid tissue = ENHANCING papillary projection, mural nodule, irregular septation / wall, or larger solid portion. A non-enhancing component is not solid tissue -- choose the matching cyst type instead.")
    form.Dropdown("T2",  "Solid tissue T2 signal:"
        , ["Not assessed","Hypointense (dark)","Intermediate","Hyperintense"], 1)
    form.Dropdown("DWI", "Solid tissue DWI signal:"
        , ["Not assessed","Low (dark)","Intermediate","High"], 1)
    form.Dropdown("TIC", "Enhancement assessment (DCE preferred; accuracy is lower without DCE):"
        , ["Not assessed (defaults to Score 4)"
        ,  "TIC type 1 -- low risk: slow progressive rise"
        ,  "TIC type 2 -- intermediate risk: moderate rise + plateau"
        ,  "TIC type 3 -- high risk: brisk rise (> myometrium)"
        ,  "Non-DCE: enhancement <= myometrium at 30-40 s"
        ,  "Non-DCE: enhancement > myometrium at 30-40 s"], 1)

    form.Header("Dilated fallopian tube features")
    form.Dropdown("TubeFluid", "Tube fluid content:"
        , ["Not specified"
        ,  "Simple fluid"
        ,  "Non-simple fluid"], 1)
    form.Dropdown("TubeWall", "Wall / endosalpingeal folds:"
        , ["Not specified"
        ,  "Thin, smooth wall / folds"
        ,  "Thick, smooth wall / folds"], 1)
    form.Note("Enhancing solid tissue in the tube -- use 'Lesion with solid tissue'.")

    form.Header("Para-ovarian cyst features")
    form.Dropdown("ParaWall", "Wall:"
        , ["Not specified"
        ,  "Thin, smooth wall (+/- enhancement)"
        ,  "Other / thick / irregular"], 1)

    form.Header("Extra-ovarian findings")
    form.Checkbox("Perit",   "Peritoneal / mesenteric / omental nodularity or irregular thickening (-> Score 5)", perit)
    form.Checkbox("Ascites", "Ascites (does not independently change the score)", asc)

    form.OnChange("LType",  _OM_UpdateLType)
    form.OnChange("Septae", _OM_UpdateLType)
    form.OnChange("T2",     _OM_UpdateLType)
    form.OnChange("DWI",    _OM_UpdateLType)
    _OM_UpdateLType(form)

    form.SetSubmit(ORADSMRI_OnSubmit)
    form.AddButtons()
    form.Show()
    return form   ; for GUI smoke tests
}

_OM_UpdateLType(frm) {
    label := frm.byName["LType"].ctl.Text
    isNone  := InStr(label, "No lesion")
    isUni   := InStr(label, "Unilocular")
    isMulti := InStr(label, "Multilocular")
    isLipid := InStr(label, "lipid content")
    isSolid := InStr(label, "solid tissue")
    isTube  := InStr(label, "fallopian")
    isPara  := InStr(label, "Para-ovarian")

    frm.SetVisible("SizeCm", !isNone)
    frm.SetSectionVisible("Unilocular cyst features",  isUni)
    frm.SetSectionVisible("Multilocular cyst features", isMulti)
    frm.SetSectionVisible("Lipid lesion features",      isLipid)
    frm.SetSectionVisible("Dilated fallopian tube features", isTube)
    frm.SetSectionVisible("Para-ovarian cyst features", isPara)

    ; The solid-tissue questions apply to the "Lesion with solid tissue"
    ; type AND to a multilocular cyst whose septae/wall are irregular and
    ; enhancing (= solid tissue per the lexicon footnote). Read Septae
    ; AFTER the section toggle above -- if it was just hidden it has been
    ; reset to its default.
    multiIrreg := isMulti && InStr(frm.byName["Septae"].ctl.Text, "Irregular")
    showSolid := isSolid || multiIrreg
    frm.SetSectionVisible("Solid tissue characterization", showSolid)

    ; Homogeneously dark-T2 + dark-DWI solid tissue is Score 2 regardless
    ; of enhancement kinetics -- TIC is moot, hide it.
    if showSolid {
        darkDark := InStr(frm.byName["T2"].ctl.Text, "dark")
                 && InStr(frm.byName["DWI"].ctl.Text, "dark")
        frm.SetVisible("TIC", !darkDark)
    }
}

ORADSMRI_OnSubmit(v, form := "") {
    global g_LastSelectedText
    p := {
        lt:        _OM_LType(v.LType),
        sizeCm:    IsNumber(v.SizeCm) ? v.SizeCm + 0.0 : 0,
        meno:      InStr(v.Meno, "Premeno") ? "pre" : InStr(v.Meno, "Postmeno") ? "post" : "",
        fluid:     _OM_Fluid(v.FluidType),
        wallEnh:   !!v.WallEnh,
        septae:    InStr(v.Septae, "Irregular") ? "irregular"
                 : InStr(v.Septae, "Smooth")    ? "smooth" : "",
        lipidLargeSolid: !!v.LargeVolSolid,
        t2:        _OM_T2(v.T2),
        dwi:       _OM_DWI(v.DWI),
        tic:       _OM_TIC(v.TIC),
        tubeFluid: InStr(v.TubeFluid, "Non-simple") ? "nonsimple"
                 : InStr(v.TubeFluid, "Simple")     ? "simple" : "",
        tubeWall:  InStr(v.TubeWall, "Thin")  ? "thin"
                 : InStr(v.TubeWall, "Thick") ? "thick" : "",
        paraWall:  InStr(v.ParaWall, "Thin")  ? "thin"
                 : InStr(v.ParaWall, "Other") ? "other" : "",
        perit:     !!v.Perit,
        ascites:   !!v.Ascites
    }

    r := _OM_Score(p)
    score := r.score
    info := _OM_Cats()[score]

    showRisk := Prefs.Get("display", "showMalignancyRisk", true)
    descLc := (info.desc != "") ? StrLower(SubStr(info.desc, 1, 1)) SubStr(info.desc, 2) : ""
    mgmtLc := (info.mgmt != "") ? StrLower(SubStr(info.mgmt, 1, 1)) SubStr(info.mgmt, 2) : ""

    sizePhrase := p.sizeCm > 0 ? Format("{:.1f}", p.sizeCm) " cm " : ""
    lesionPhrase := r.HasOwnProp("phrase") && r.phrase != ""
        ? r.phrase : _OM_TypePhrase(p.lt)
    if (p.lt = "none" || score = 1)
        impression := lesionPhrase ", O-RADS MRI " score " (" descLc ")"
    else
        impression := sizePhrase lesionPhrase ", O-RADS MRI " score " (" descLc ")"
    ; Sentence-start capitalization when no size prefix.
    impression := StrUpper(SubStr(impression, 1, 1)) SubStr(impression, 2)
    if (showRisk && info.risk != "" && info.risk != "N/A")
        impression .= ", malignancy risk " info.risk
    impression .= " per O-RADS MRI v2020"
    if (mgmtLc != "")
        impression .= "; " mgmtLc
    if (SubStr(impression, -1) != ".")
        impression .= "."

    method := ""
    if form
        method .= "Selected inputs:`n" form.FormatInputs(v)
    method .= "`nScore: " score " -- " info.desc
    method .= "`nRule applied: " r.reason
    method .= "`nMalignancy risk (PPV): " info.risk
    method .= "`nFull management: " info.mgmt

    advisories := []
    if (p.ascites && !p.perit)
        advisories.Push("Ascites without peritoneal nodularity is not an independent O-RADS MRI scoring criterion; correlate clinically.")
    if (p.lt = "unilocular" && p.meno = "" && p.fluid != "" && p.sizeCm > 0 && p.sizeCm <= 3
        && (p.fluid = "simple" || p.fluid = "hemorrhagic"))
        advisories.Push("Menopausal status not specified: a " p.fluid " cyst <=3 cm would be physiologic (O-RADS MRI 1) in a premenopausal patient. Specify status to apply that rule.")

    return MakeResult({
        classification: "O-RADS MRI " score,
        impression:     impression,
        recommendation: "",
        advisories:     advisories,
        methodology:    method,
        citations:      [{ text: "Thomassin-Naggara I, Poncelet E, Jalaguier-Coudray A, et al. "
                                . "Ovarian-Adnexal Reporting Data System Magnetic Resonance Imaging "
                                . "(O-RADS MRI) Score for Risk Stratification of Sonographically "
                                . "Indeterminate Adnexal Masses. JAMA Netw Open. 2020;3(1):e1919896.",
                           url:  "https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/O-RADS" }],
        echo:           g_LastSelectedText
    })
}

_OM_LType(label) {
    if InStr(label, "No lesion")
        return "none"
    if InStr(label, "Unilocular")
        return "unilocular"
    if InStr(label, "Multilocular")
        return "multilocular"
    if InStr(label, "lipid content")
        return "lipid"
    if InStr(label, "fallopian")
        return "tube"
    if InStr(label, "Para-ovarian")
        return "paraovarian"
    return "solid"
}
_OM_TypePhrase(lt) {
    static m := Map(
        "none",         "no adnexal lesion",
        "unilocular",   "unilocular adnexal cyst",
        "multilocular", "multilocular adnexal cyst",
        "lipid",        "lipid-containing adnexal lesion",
        "solid",        "adnexal lesion with solid tissue",
        "tube",         "dilated fallopian tube",
        "paraovarian",  "para-ovarian cyst")
    return m.Has(lt) ? m[lt] : "adnexal lesion"
}
_OM_T2(label) {
    if InStr(label, "dark")
        return "hypo"
    if (label = "Intermediate")
        return "intermediate"
    if InStr(label, "Hyper")
        return "hyper"
    return ""
}
_OM_DWI(label) {
    if InStr(label, "dark")
        return "low"
    if (label = "Intermediate")
        return "intermediate"
    if (label = "High")
        return "high"
    return ""
}
_OM_TIC(label) {
    if InStr(label, "type 1")
        return 1
    if InStr(label, "type 2")
        return 2
    if InStr(label, "type 3")
        return 3
    if InStr(label, "<= myometrium")
        return "nd-le"
    if InStr(label, "> myometrium")
        return "nd-gt"
    return 0
}
_OM_Fluid(label) {
    if (label = "Simple")
        return "simple"
    if InStr(label, "Endometriotic")
        return "endometriotic"
    if InStr(label, "Hemorrhagic")
        return "hemorrhagic"
    if InStr(label, "Proteinaceous")
        return "proteinaceous"
    return ""
}

; Fill any omitted props with safe defaults so callers (and tests) can pass
; a partial object stating only the fields that matter for a given case.
_OM_Norm(props) {
    base := { lt: "solid", sizeCm: 0, meno: "", fluid: "", wallEnh: false
            , septae: "", lipidLargeSolid: false, t2: "", dwi: "", tic: 0
            , tubeFluid: "", tubeWall: "", paraWall: "", perit: false, ascites: false }
    if IsObject(props) {
        for k, val in base.OwnProps() {
            if props.HasOwnProp(k)
                base.%k% := props.%k%
        }
    }
    return base
}

; ---- classifier ------------------------------------------------------------
; Takes a props object so tests can express each permutation tersely:
;   { lt, sizeCm, meno, fluid, wallEnh, septae, lipidLargeSolid,
;     t2, dwi, tic, tubeFluid, tubeWall, paraWall, perit, ascites }
; Missing fields default via _OM_Norm. Returns { score, reason [, phrase] }.
; Every branch corresponds to one row of the source table
; (references/orads mri.md); reasons quote the row.
_OM_Score(props) {
    p := _OM_Norm(props)
    ; Peritoneal nodularity overrides everything, including "no lesion" --
    ; carcinomatosis without a visible ovarian primary must not score 1.
    if p.perit
        return { score: 5, reason: "Peritoneal, mesenteric or omental nodularity or irregular thickening (with or without ascites)" }

    if (p.lt = "none")
        return { score: 1, reason: "No ovarian lesion", phrase: "no adnexal lesion" }

    if (p.lt = "unilocular") {
        ; Physiologic rows: simple (follicle) or hemorrhagic cyst (incl.
        ; corpus luteum +/- hemorrhage) <=3 cm in a PREMENOPAUSAL patient
        ; -> Score 1. The main grid uses "<= 3 cm" (footnote *** says
        ; "<3cm" -- we follow the grid). Requires an explicit size and
        ; explicit premenopausal status; "not specified" never scores 1.
        ; Wall enhancement does not exclude this rule (a corpus luteum has
        ; an enhancing wall). Endometriotic fluid is NOT physiologic.
        if (p.meno = "pre" && p.sizeCm > 0 && p.sizeCm <= 3) {
            if (p.fluid = "simple")
                return { score: 1, reason: "Follicle: simple cyst <=3 cm in a premenopausal patient"
                       , phrase: "simple cyst (physiologic follicle, premenopausal)" }
            if (p.fluid = "hemorrhagic")
                return { score: 1, reason: "Hemorrhagic cyst (or corpus luteum +/- hemorrhage) <=3 cm in a premenopausal patient"
                       , phrase: "hemorrhagic cyst (physiologic, premenopausal)" }
        }
        if !p.wallEnh
            return { score: 2, reason: "Unilocular cyst, any fluid content, no wall enhancement, no enhancing solid tissue" }
        if (p.fluid = "simple" || p.fluid = "endometriotic")
            return { score: 2, reason: "Unilocular cyst, simple or endometriotic fluid, smooth enhancing wall, no enhancing solid tissue" }
        if (p.fluid = "")
            return { score: 3, reason: "Unilocular cyst, smooth enhancing wall, fluid type not specified -- scored per the proteinaceous/hemorrhagic/mucinous row (conservative)" }
        return { score: 3, reason: "Unilocular cyst, proteinaceous / hemorrhagic / mucinous fluid, smooth enhancing wall, no enhancing solid tissue" }
    }

    if (p.lt = "multilocular") {
        ; Irregular ENHANCING septae/wall meet the solid-tissue definition
        ; (source footnote *) -> score via the solid-tissue rows.
        if (p.septae = "irregular")
            return _OM_SolidScore(p, "Multilocular cyst with irregular enhancing septae / wall (= solid tissue): ")
        return { score: 3, reason: "Multilocular cyst, any fluid, no lipid content, smooth septae and wall, no enhancing solid tissue" }
    }

    if (p.lt = "lipid") {
        if p.lipidLargeSolid
            return { score: 4, reason: "Lesion with lipid content and large-volume enhancing solid tissue" }
        return { score: 2, reason: "Lesion with lipid content, no enhancing solid tissue (minimal Rokitansky-nodule enhancement does not upgrade)" }
    }

    if (p.lt = "tube") {
        if (p.tubeFluid = "simple" && p.tubeWall = "thin")
            return { score: 2, reason: "Dilated fallopian tube, simple fluid, thin smooth wall / endosalpingeal folds, no enhancing solid tissue" }
        if (p.tubeFluid = "nonsimple")
            return { score: 3, reason: "Dilated fallopian tube, non-simple fluid, thin wall / folds, no enhancing solid tissue" }
        if (p.tubeWall = "thick")
            return { score: 3, reason: "Dilated fallopian tube, simple fluid, thick smooth wall / folds, no enhancing solid tissue" }
        return { score: 3, reason: "Dilated fallopian tube, fluid / wall characteristics not fully specified -- Score 3 (conservative within the no-solid tube rows)" }
    }

    if (p.lt = "paraovarian") {
        if (p.paraWall = "thin" || p.paraWall = "")
            return { score: 2, reason: "Para-ovarian cyst, any fluid, thin smooth wall +/- enhancement, no enhancing solid tissue" }
        return { score: 3, reason: "Para-ovarian cyst with atypical (thick / irregular non-enhancing) wall -- not tabulated; Score 3 (conservative). If enhancing solid tissue is present, use 'Lesion with solid tissue'." }
    }

    return _OM_SolidScore(p, "")
}

; Solid-tissue rows (shared by "Lesion with solid tissue" and multilocular-
; with-irregular-septae). Dark-dark check precedes kinetics: homogeneously
; T2-dark AND DWI-dark solid tissue is Score 2 regardless of TIC.
_OM_SolidScore(p, prefix) {
    if (p.t2 = "hypo" && p.dwi = "low")
        return { score: 2, reason: prefix "Solid tissue homogeneously hypointense on T2 AND DWI (dark/dark) -> almost certainly benign" }
    if (p.tic = 1)
        return { score: 3, reason: prefix "Solid tissue (not dark/dark) with low-risk time-intensity curve on DCE MRI" }
    if (p.tic = 2)
        return { score: 4, reason: prefix "Solid tissue (not dark/dark) with intermediate-risk time-intensity curve on DCE MRI" }
    if (p.tic = 3)
        return { score: 5, reason: prefix "Solid tissue (not dark/dark) with high-risk time-intensity curve on DCE MRI" }
    if (p.tic = "nd-le")
        return { score: 4, reason: prefix "Solid tissue enhancing <= myometrium at 30-40 s on non-DCE MRI (note: accuracy is decreased without DCE)" }
    if (p.tic = "nd-gt")
        return { score: 5, reason: prefix "Solid tissue enhancing > myometrium at 30-40 s on non-DCE MRI (note: accuracy is decreased without DCE)" }
    return { score: 4, reason: prefix "Solid tissue (not dark/dark) without enhancement-kinetics assessment -- defaults to intermediate risk; obtain DCE MRI to refine" }
}

_OM_Cats() {
    static m := _OM_BuildCats()
    return m
}
_OM_BuildCats() {
    m := Map()
    m[0] := { desc: "Incomplete",              risk: "N/A",    mgmt: "Repeat MRI with contrast." }
    ; Source: Score 1 PPV is listed as "N/A" -- the system does not assign
    ; a numeric malignancy risk to a normal ovary.
    m[1] := { desc: "Normal ovaries",          risk: "N/A",    mgmt: "No follow-up." }
    m[2] := { desc: "Almost certainly benign", risk: "<0.5%",  mgmt: "No follow-up." }
    m[3] := { desc: "Low risk",                risk: "~5%",    mgmt: "Surveillance MRI in 6-12 months." }
    m[4] := { desc: "Intermediate risk",       risk: "~50%",   mgmt: "Gynecologic oncology consultation." }
    m[5] := { desc: "High risk",               risk: "~90%",   mgmt: "Gynecologic oncology referral." }
    return m
}
