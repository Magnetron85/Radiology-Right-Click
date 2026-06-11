# v2.1 -> v2.1.1 changes

## O-RADS MRI: rebuilt to match the source grid

- **Menopausal status and lesion size are now inputs**, and the three physiologic Score-1 rows are modeled correctly: a simple/follicular or hemorrhagic cyst (incl. corpus luteum +/- hemorrhage) **<=3 cm in a premenopausal patient** scores O-RADS MRI 1. The prior version had no status/size input, so it could not reach Score 1 for these and mis-scored them as 2. The same lesion in a postmenopausal patient, or >3 cm, or with unspecified status, correctly does **not** get the physiologic downgrade.
- **Added the lesion types that were missing**: dilated fallopian tube (simple/thin -> 2, non-simple or thick wall -> 3) and para-ovarian cyst (thin smooth wall -> 2).
- **Split the unilocular fluid types** to drive the Score 2 vs 3 wall-enhancement rows exactly (simple/endometriotic -> 2; proteinaceous/hemorrhagic/mucinous -> 3; unspecified -> conservative 3).
- **Multilocular cysts with irregular enhancing septae/wall** now route through the solid-tissue pathway (irregular enhancing septation is "solid tissue" per the source footnote), instead of being scored as a smooth-septate Score 3.
- **Added the non-DCE enhancement rows** (solid tissue enhancing <= myometrium at 30-40 s -> 4; > myometrium -> 5) alongside the DCE time-intensity-curve rows, with an accuracy caveat.
- Ascites without peritoneal nodularity no longer implies anything on its own (it attaches to the nodularity row); an advisory flags it, and another advisory prompts for menopausal status when it would change a borderline cyst's score.
- The classifier was restructured to take a single props object and now returns the specific source row it applied (shown in the methodology block). New permutation suite `tests/test_oradsmri_50.ahk` covers 47 rows including every physiologic-vs-not status/size boundary; all pass.

## Calculator dialog overhaul (progressive disclosure)

- **The form engine (`lib/FormGui.ahk`) now supports true progressive disclosure.** `SetVisible` HIDES an irrelevant field (instead of greying it), collapses its space -- rows below slide up and the window resizes live, flicker-free -- and resets the hidden control to its default so a stale answer can never reach the classifier; re-showing restores the user's prior value. New `SetSectionVisible` toggles an entire header section in one relayout.
- **Every dialog reorganized: mandatory context first, conditional questions appear only when a previous answer makes them relevant.** Examples: Lung-RADS opens with just lesion type / screening round / size -- the solid-component fields appear only for part-solid, airway location only for airway nodules, cyst features only for atypical cysts, and the whole prior-comparison section only for incident rounds. LI-RADS collapses to two checkboxes when LR-NC or TIV is asserted, and its LR-M criteria / ancillary-adjustment blocks expand on demand. PI-RADS hides T2 for PZ and DCE for TZ (where the v2.1 algorithm never reads them). TI-RADS hides echogenicity / shape / margin / foci for cystic and spongiform nodules (auto TR1). O-RADS US/MRI show only the descriptors for the selected lesion category. PSA density hides the three dimension fields the moment a measured volume is typed. Bosniak, gallbladder polyp, incidental adrenal / thyroid, Kyoto IPMN, US LI-RADS, and pregnancy dates gained equivalent gating; simple measurement forms were reordered mandatory-first with optional inputs under clearly-labeled optional headers.
- **Visual refresh**: accent-colored section headers over hairline rules, dark-mode aware, Win11 rounded corners retained.
- **High-DPI fix**: the window-fits-screen cap compared DPI-scaled logical units against physical pixels, so tall dialogs could run off the bottom of the screen on scaled displays (common on laptops); the cap and the position clamp are now unit-correct.
- Hidden blank-default numeric fields now reset to blank, not 0 -- a synthetic 0 read as a real measured value (e.g. 0 HU) in classifiers.
- New layout test suite: `tests/test_gui_smoke.ahk` opens all 28 dialogs, audits control overlap and screen fit, and exercises the Lung-RADS disclosure end-to-end (95 checks); `tests/_reflow_smoke.ahk` covers the reflow engine itself.

## PowerScribe 360 reliability

- **Menu Cut / Copy / Paste / Delete now use SendEvent.** AHK v2's default `Send` (SendInput) bypasses the message queue, so hosts with low-level keyboard hooks (PowerScribe / Dragon) never saw the keystroke -- these menu items silently did nothing there while working fine in Notepad.
- **Selection capture hardened** (`GetSelectedText`): the Ctrl+C now has a real 20 ms press duration, and on a miss the capture falls back to (1) `WM_COPY` posted directly at the focused control (no keyboard involved, immune to hooks) and (2) one slower Ctrl+C retry. Debug log records which stage succeeded.
- **Right-click activation race fixed**: `HandleRightClick` now waits (bounded, 300 ms) for the target window to actually take focus instead of a fixed 30 ms nap, so the capture can't fire against the wrong window in slow hosts.

## Paste-ready clipboard output

- **New paste composition layer** (`PasteText` in `lib/CalcResult.ahk`): when a result window opens, the clipboard now holds the original highlighted text with the result appended per the module's declared style, so select -> run -> paste over the selection yields a clean report line. Modes: `paren` ("...3.1 x 2.2 x 2.8 cm (ellipsoid volume 10.0 mL)."), `inline`, `newline` (default; RADS impressions append on their own line), `replace`. Multi-line selections always fall back to newline composition.
- Per-module parenthetical fragments added: Ellipsoid/Bullet volume, PSA density, adrenal washout, thymus chemical shift, hepatic fat fraction, liver iron (LIC), calcium score percentile, NASCET, statistics, range, and measurement comparison.
- New pref `clipboard.includeSelection` (default `true`); set `false` in `preferences.json` to restore the old result-only clipboard.
- "Copy impression" still copies the result alone; "Copy with methodology" unchanged.

## Bug fixes

- **Adrenal washout: 2-phase submissions crashed** (`Round("")` on the absent absolute-washout value) -- every enhanced+delayed-only calculation died with a "Calculation error" box. Also, the classification can no longer read "Adenoma-suggestive washout" when the verdict says washout does not apply (de-enhancement / non-enhancing lesion).
- **PSA density now reports 2 decimals** -- the clinical cutoff is 0.15 ng/mL/cc; the old 1-decimal rounding (0.1 / 0.2) destroyed the answer at the decision boundary.
- **Calcium score: Hoff percentile band labels were shifted one band low** in methodology and the legacy path ("75th-90th" reported as "50-75%"); bands now match the impression label. Age parsing anchored so "Coverage: 55" can't parse as Age 55.
- **NASCET: the legacy parser's own documented example never matched** (regex required no space between number and unit); units are now captured and cm values normalize to mm everywhere; swapped distal/stenosis entries are rejected instead of reporting a negative stenosis.
- **Measurement comparison ("previously X ... now Y" word order) read the wrong regex groups** -- the current measurement could be built from a date and its date from the unit ("cm").
- **Number Range: unit detection** no longer matches one-letter prefixes ("5 ml" reported unit "m"); alternation is longest-first with a letter-boundary lookahead.
- Hepatic steatosis legacy path: spleen-normalized FF annotation is anchored to the exact appended tag instead of the first `)` in the output (a parenthesis inside the selection corrupted the line).
- Fleischner dual-risk impression now reads as prose ("if low risk, ...; if high risk, ...") instead of label-style "low-risk:" / "high-risk:" headers.
- Result popup: Close button is pinned to the right edge (it overlapped "Copy impression" at minimum window width); minimum width raised to fit the button row.
- Test harness repaired (`_PIRADS_Score` arity, two stale assertions); full suite green: 21/21 calculator + 41/41 algorithm tests.
- ASCII hygiene: accented character removed from the Priola citation (legacy thymus path).

---

# v2.0 -> v2.1 changes

v2.1 adds a second class of calculator -- guided form-based classifiers for the major ACR Reporting and Data Systems and incidental-findings white papers -- alongside the v2.0 text-parse calculators.

## New calculators (form-based RADS / incidental-findings classifiers)

Twelve new classifiers, each pre-filled from the selected text and completed in a dialog:

- **Bosniak** cystic renal mass (v2019) -- Silverman et al. *Radiology* 2019.
- **Gallbladder Polyp** (SRU 2022) -- Kamaya et al. *Radiology* 2022.
- **LI-RADS** CT/MRI (v2018) -- Chernyak et al. *Radiology* 2018.
- **US LI-RADS** (v2024) -- ACR Ultrasound Surveillance.
- **Incidental Adrenal** (ACR 2017) -- Mayo-Smith et al. *JACR* 2017.
- **Incidental Thyroid** (ACR 2015) -- Hoang et al. *JACR* 2015.
- **ACR TI-RADS** -- Tessler et al. *JACR* 2017.
- **Kyoto IPMN** (2024) -- Ohtsuka et al. *Pancreatology* 2024.
- **Lung-RADS** (v2022) -- ACR Committee on Lung-RADS.
- **O-RADS MRI** -- Thomassin-Naggara et al. *JAMA Netw Open* 2020.
- **O-RADS US** (v2022) -- Andreotti et al. *Radiology* 2020.
- **PI-RADS** (v2.1) -- Turkbey et al. *Eur Urol* 2019.

## New structure

- `lib/FormGui.ahk` -- shared guided-input form used by all twelve classifiers.
- `lib/CalcResult.ahk` -- structured result object (impression / classification / recommendation / findings / methodology / citations / advisories), defining the report-ready output contract.
- `lib/TextScan.ahk` -- selected-text parsing and form pre-fill helpers.

## Menu and display

- **New "grouped" (anatomical) menu sorting, now the default** -- calculators are bucketed into anatomical submenus (Renal, Liver / biliary, Adrenal, Thyroid / neck, Pancreas, Lung, Cardiovascular, Adnexal / OB-Gyn, Prostate, Volume + measurement, Numeric, Scheduling). `alphabetical`, `frequency`, and `none` remain available.
- New display toggles: **Show malignancy risk** (per-category validated risk percentages, e.g. TI-RADS) and **Show methodology** (the selected-inputs + reasoning block). Each result popup also has a local methodology checkbox to override per-result without changing the pref.

## Security / quality

- **Citation links now validated.** The "Open reference in browser" links on result popups route through the same `IsValidURL` http/https allowlist that user-supplied reference URLs use, so every `Run()` in the app is validated.
- **`.gitignore` hardening** -- excludes `.claude/` (local agent settings), `debug.log` / `*.log` (may contain captured clipboard text when debug logging is on), the whole dev-only `tests/` tree, `IMPRESSION_STYLE.md` (working draft), Python cruft, and OS junk.
- **`SECURITY.md`** documents the debug-logging caveat (the "stores no PHI" guarantee holds only with debug logging off).
- README documents all twelve new classifiers and the corrected module layout.

---

# v1.22-local -> v2.0 changes

The calculation math, regexes, decision thresholds, and citations all match v1.22 verbatim. The differences below are structural, presentational, or security/quality fixes.

## Structure

- AHK v1 -> AHK v2 syntax throughout (`#Requires AutoHotkey v2.0`).
- Single 3,310-line script split into one main entry + `lib/` modules + one file per calculator under `lib/calc/`. Each calculator exports a `<Name>_Entry(selectedText)` function that the menu calls.
- All globals replaced with a `Prefs` class accessed via `Prefs.Get("section","key", default)`. Frequency data and references live in the same preferences file as everything else.

## Preferences storage

- New `preferences.json` format. On first run, an existing v1 `preferences.ini` is imported into JSON and the old file is renamed to `preferences.ini.imported` so it isn't re-imported.
- Writes are atomic (stage to `.tmp` then move; previous good file is kept as `preferences.backup.json`).
- Missing keys in the JSON fall back to defaults rather than crashing.

## UI

- Modern flat layout with Segoe UI (Segoe UI Variable on Windows 11).
- Native immersive dark title bar when dark mode is on (`DwmSetWindowAttribute` / `DWMWA_USE_IMMERSIVE_DARK_MODE`).
- Rounded corners on Windows 11 (`DWMWA_WINDOW_CORNER_PREFERENCE`).
- Native popup menus follow dark mode too -- toggling the pref calls `uxtheme!SetPreferredAppMode` + `FlushMenuThemes` (the same undocumented ordinals File Explorer and RegEdit use for their dark menus), so the right-click menu, sub-menus, and dropdowns all flip on the next click.
- Result popup measures the text first, sizes to fit (capped at ~50% of the active monitor), parks the cursor over the Close button.

## Security / robustness fixes

- **URL validation**: references are restricted to `http` / `https`. `javascript:`, `file:`, `ftp:`, and similar schemes are rejected.
- **Filename sanitization**: when a reference file is copied into the local `References/` folder, the destination filename is stripped of path separators, control characters, and `..` traversal attempts, then truncated to 200 chars.
- **File-type whitelist**: still enforced (PDF/DOC/DOCX/XLS/XLSX/PPT/PPTX) and the check now runs both on add and on open.
- **Reference open**: re-validates URL / file path / extension before launching, even if the JSON was edited by hand.
- **Clipboard handling**: full `ClipboardAll()` save/restore around `GetSelectedText()` so binary clipboard contents (images, RTF) survive the helper.
- **No shell-string concatenation** for `Run`; everything goes through `Run(value)` with the value already validated.

## Bug fixes (silently applied from v1)

- `ProcessNodules` had a typo `esult .=` instead of `result .=` (v1 line 2679) -- one line of the Fleischner output was being dropped. Fixed.
- v1 read a `ShowSpellCheck` preference and built a "Check Spelling" menu item, but `CheckSpelling` was never defined and the global was never initialised. Removed.
- v1 added a `$Space::` hotkey inside the target-app block that only re-sent `{Space}` and mutated an undeclared `lastPeriodTime` variable. Dead code -- removed.
- Adrenal washout 2-phase fallback regex contained `\nits?` (a stray newline in the middle of "units?") in v1. Replaced with the correct full pattern.
- `CalciumScore` v1 parsed `Race:` but never used it. Removed -- Hoff percentile is computed from age + sex only, matching the v1 actual behavior.
- `Quartile` / `Median` / `StdDev` now guard against division by zero on degenerate inputs.
- `CompareNoduleSizes` guards against zero divisors in growth-rate math.

## Audit-pass fixes (post-port revision 2)

- **Target apps are now editable in Preferences.** `activation.targetApps` lives in `preferences.json`; click "Target Apps..." in Preferences to add/remove entries in `class:WindowClass` or `exe:Process.exe` format. The hardcoded list in `RightClick.ahk` is kept as a fallback when the pref is missing or empty.
- **Frequency writes batched.** `Prefs.IncrementFrequency()` now flips a `dirty` flag instead of writing the JSON file on every menu click. `Prefs.Flush()` is wired to `OnExit` and runs once when the script terminates.
- **`Prefs.RestoreDefaults()` added** as a public API so the Preferences GUI no longer reaches into `Prefs.data` / private `_Defaults()` directly.
- **Sort Measurement Sizes** now saves and restores the user's clipboard around its `Send "^v"`, with a 100 ms settle delay so the paste actually lands before restore.
- **Reference open errors are surfaced**: `Run()` failures in `OpenReference()` now show a MsgBox instead of silently swallowing the exception, and the `uses` counter is only bumped on successful launch.
- **No more "(no references saved)" placeholder** — the References submenu just shows "Add..." / "Manage..." when empty.
- **Contrast Premedication protocol selector** is now a DropDownList instead of a brittle pair of radio buttons (the v1 code keyed off the first radio's checked state).
- **`SafeFilename` rejects Windows reserved device names** (CON, PRN, AUX, NUL, COM1-9, LPT1-9); prefixes an underscore rather than failing the copy silently.
- **JSON output escapes remaining ASCII control chars** as `\uXXXX` per RFC 8259. Practical impact is zero for our data; this is defensive.
- Stale header comment in `lib/Menu.ahk` (was still listing the removed `custom` sort method) and the misplaced doc comment in `lib/Util.ahk` (orphaned above `SafeInt`) both cleaned up.

## Pause script (removed)

- "Pause Script" menu item, the `PauseScript` / `ResumeScript` helpers, and the `script.pauseDuration` preference are gone. v1's pause UX (modal MsgBox blocking the script while it suspended itself) was clunky; suspending via the tray icon serves the same purpose. v1 INI imports no longer read `PauseDuration`.

## Right-click activation modifier (new)

- New preference `activation.modifier`: `none`, `ctrl` (**default**), `alt`, or `shift`. Defaults to Ctrl so plain right-click stays native in PowerScribe / Notepad and the helper menu only appears when you specifically ask for it.
- When set to a modifier, plain right-click goes straight to the app's native context menu; only the modifier + right-click opens the RightClick menu.
- Implemented with the runtime `Hotkey()` API and re-applied immediately when the preference is saved -- no script restart needed.
- "Fn" is intentionally not offered: on most laptops it's handled in the keyboard's firmware and never reaches Windows as a normal modifier, so AutoHotkey can't hook it. Shift is the practical equivalent.

## Menu sorting (post-port revision)

- Default sort changed from `frequency` to `alphabetical` (case-insensitive by display title).
- `frequency` sort still available; uses alphabetical as the tiebreaker so equal-usage items don't reshuffle between sessions.
- `none` still available (registry order).
- **The `custom` sort option and the `customOrder` list are gone.** Old preferences with `sortingMethod = "custom"` are migrated to `alphabetical` on first load, and any stored `customOrder` array is dropped.

## Behavior preserved

- All clinical math (ellipsoid / bullet volume, PSA density, washout, SII / CSR cutoffs, fat fraction thresholds, R2* -> iron coefficients, Hoff percentiles, NASCET formula, Fleischner categorisation and recommendations) is byte-identical to v1.
- All citations preserved verbatim.
- Same default pause durations (3 / 10 / 30 min, 1 / 10 hr), same default sort method ("frequency"), same default custom-order list.
- "Sort Measurement Sizes" still works by pasting over the selection.
- Frequency counts and references carry over from v1 via the one-time import.
