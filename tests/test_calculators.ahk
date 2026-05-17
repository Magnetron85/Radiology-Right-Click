; ============================================================
; tests/test_calculators.ahk -- self-test for every calculator
; ------------------------------------------------------------
; Double-click to run. Loads each calculator, feeds it a known-
; good sample input, captures the output, and reports
; PASS / FAIL / ERROR. Writes tests/test_output.txt and opens
; it in Notepad.
;
; Two calculators are excluded from the batch run:
;   * SortMeasurementSizes  -- requires a real selection and
;                              pastes via Ctrl+V; verify manually.
;   * ContrastPremedication -- opens a modal GUI; verify manually.
; ============================================================

#Requires AutoHotkey v2.0
#SingleInstance Force

#Include ..\lib\Prefs.ahk
#Include ..\lib\Modern.ahk
#Include ..\lib\Util.ahk

#Include ..\lib\calc\Volumes.ahk
#Include ..\lib\calc\PSADensity.ahk
#Include ..\lib\calc\PregnancyDates.ahk
#Include ..\lib\calc\Numerics.ahk
#Include ..\lib\calc\AdrenalWashout.ahk
#Include ..\lib\calc\ThymusChemicalShift.ahk
#Include ..\lib\calc\HepaticSteatosis.ahk
#Include ..\lib\calc\LiverIron.ahk
#Include ..\lib\calc\CalciumScore.ahk
#Include ..\lib\calc\NASCET.ahk
#Include ..\lib\calc\NoduleSizes.ahk
#Include ..\lib\calc\Fleischner.ahk

; Drive Prefs with defaults so showCitations / showArterialAge /
; allValues fire their full output paths.
Prefs.data := Prefs._Defaults()

cases := [
    { name: "EllipsoidVolume",         fn: EllipsoidVolume_Entry,
      input: "3.0 x 2.0 x 1.0 cm",                                                        expect: "cc" },
    { name: "BulletVolume",            fn: BulletVolume_Entry,
      input: "3.0 x 2.0 x 1.0 cm",                                                        expect: "cc" },
    { name: "PSADensity",              fn: PSADensity_Entry,
      input: "PSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm",                                  expect: "PSA Density" },
    { name: "PregnancyDates (LMP)",    fn: PregnancyDates_Entry,
      input: "LMP: 01/15/2025",                                                           expect: "Estimated Delivery" },
    { name: "PregnancyDates (GA)",     fn: PregnancyDates_Entry,
      input: "GA: 12 weeks and 3 days as of today",                                       expect: "Gestational" },
    { name: "MenstrualPhase",          fn: MenstrualPhase_Entry,
      input: "LMP: 05/01/2026",                                                           expect: "Phase" },
    { name: "CompareNoduleSizes",      fn: CompareSizes_Entry,
      input: "Now 2.5 cm, previously 1.5 cm on 01/01/2024",                               expect: "%" },
    { name: "AdrenalWashout (3-phase)", fn: AdrenalWashout_Entry,
      input: "Unenhanced: 10 HU, Enhanced: 80 HU, Delayed: 40 HU",                        expect: "Absolute Washout" },
    { name: "AdrenalWashout (2-phase)", fn: AdrenalWashout_Entry,
      input: "Enhanced: 80 HU, Delayed: 40 HU",                                           expect: "Relative Washout" },
    { name: "ThymusChemicalShift",     fn: ThymusChemicalShift_Entry,
      input: "Thymus IP: 100, OP: 80, Paraspinous IP: 90, OP: 85",                        expect: "Chemical Shift Ratio" },
    { name: "HepaticSteatosis",        fn: HepaticSteatosis_Entry,
      input: "Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88",                              expect: "Fat Fraction" },
    { name: "LiverIron (1.5T)",        fn: LiverIron_Entry,
      input: "1.5T, R2*: 50 Hz",                                                          expect: "Iron Content" },
    { name: "LiverIron (3.0T)",        fn: LiverIron_Entry,
      input: "3.0T, R2*: 70 Hz",                                                          expect: "Iron Content" },
    { name: "Statistics",              fn: Statistics_Entry,
      input: "1.0, 2.5, 3.7, 4.0, 5.2, 6.1, 7.8, 8.4, 9.0, 10.5",                         expect: "Mean" },
    { name: "Range",                   fn: Range_Entry,
      input: "values 5 to 10 cm",                                                         expect: "-" },
    { name: "CalciumScore",            fn: CalciumScore_Entry,
      input: "Age: 55`nSex: Male`nCoronary artery calcium score: 250 (Agatston)",         expect: "Plaque Burden" },
    { name: "NASCET",                  fn: NASCET_Entry,
      input: "Distal ICA = 6 mm, Stenosis = 2 mm",                                        expect: "NASCET:" },
    { name: "Fleischner (solid)",      fn: Fleischner_Entry,
      input: "Solitary 7 mm solid pulmonary nodule in the right upper lobe.",             expect: "right" },
    { name: "Fleischner (multiple)",   fn: Fleischner_Entry,
      input: "Multiple ground glass nodules in both lungs, the largest measuring up to 8 mm.", expect: "8" },
    { name: "PregnancyDates (stale LMP)", fn: PregnancyDates_Entry,
      input: "LMP: 01/15/2024",                                                           expect: "past EDD" }
]

passes := 0
fails  := 0
errors := 0
report := "RightClick v2 calculator self-test`n"
report .= "Run: " FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "`n"
report .= "=================================`n`n"

for c in cases {
    report .= "[" c.name "]`n"
    report .= "  Input:  " StrReplace(c.input, "`n", " | ") "`n"
    try {
        out := c.fn.Call(c.input)
        if (out = "") {
            report .= "  Status: FAIL (no output)`n`n"
            fails++
            continue
        }
        if InStr(out, c.expect) {
            report .= "  Status: PASS`n"
            passes++
        } else {
            report .= "  Status: FAIL (missing expected substring '" c.expect "')`n"
            fails++
        }
        snippet := StrLen(out) > 320 ? SubStr(out, 1, 320) " ..." : out
        report .= "  Output: " StrReplace(snippet, "`n", " | ") "`n`n"
    } catch as e {
        report .= "  Status: ERROR -- " e.Message "`n"
        report .= "  Where:  " (e.HasProp("File") ? e.File : "?") ":" (e.HasProp("Line") ? e.Line : "?") "`n"
        if (e.HasProp("Extra") && e.Extra != "")
            report .= "  Extra:  " e.Extra "`n"
        report .= "`n"
        errors++
    }
}

summary := Format("Summary: {1} pass / {2} fail / {3} error / {4} total"
    , passes, fails, errors, cases.Length)
report := summary "`n`n" report
report .= "----------------------------------`n"
report .= "Excluded (need manual verification):`n"
report .= "  * Sort Measurement Sizes (clipboard paste)`n"
report .= "  * Contrast Premedication (GUI modal)`n"

reportFile := A_ScriptDir "\test_output.txt"
try FileDelete reportFile
FileAppend report, reportFile, "UTF-8"
try Run 'notepad.exe "' reportFile '"'
ExitApp
