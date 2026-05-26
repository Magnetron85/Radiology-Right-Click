# Radiology Right Click (v2)

> **DISCLAIMER — FOR ENTERTAINMENT, RESEARCH, AND EDUCATIONAL USE ONLY.**
>
> This software is **not** a medical device, has not been validated for clinical use, and **must not be used for clinical decision-making**. All calculations are derived from published formulas and must be independently verified by the user against the source data before they are incorporated into any clinical report or used to inform patient care. Use of this software does not establish a clinician-patient relationship, does not constitute medical advice, and the author makes no warranty, express or implied, that any output is fit for clinical purposes. The user assumes all responsibility for verifying the accuracy and applicability of every value reported. The author disclaims all liability for any clinical, diagnostic, therapeutic, or other decision made on the basis of this software's output.

A right-click context menu of measurement and analysis helpers for radiology dictation workflows. AutoHotkey v2 rewrite of [Radiology Right Click v1.22](https://github.com/Magnetron85/Radiology-Right-Click); see `CHANGES.md` for what's different.

> **License:** [PolyForm Noncommercial 1.0.0](https://polyformproject.org/licenses/noncommercial/1.0.0). Free for individuals and non-commercial / public-benefit organizations (clinics, hospitals, universities, research institutions). Commercial use is not permitted. See `LICENSE`.

---

## Contents

- [Install](#install)
- [Usage at a glance](#usage-at-a-glance)
- [Calculator reference](#calculator-reference)
  - [Volumes and serial measurements](#volumes-and-serial-measurements)
  - [Abdominal MRI](#abdominal-mri)
  - [Cardiovascular](#cardiovascular)
  - [Lung / pulmonary](#lung--pulmonary)
  - [Obstetric / gynecologic](#obstetric--gynecologic)
  - [Numeric utilities](#numeric-utilities)
  - [Scheduling](#scheduling)
- [Preferences](#preferences)
- [Saved references](#saved-references)
- [Configuring target apps](#configuring-target-apps)
- [Activation modifier](#activation-modifier)
- [Project layout](#project-layout)
- [License](#license)

---

## Install

**Option A — run from source (recommended for transparency):**

1. Install [AutoHotkey v2](https://www.autohotkey.com/) (1.5 MB, free).
2. Download or clone this repo to any folder (USB stick, OneDrive, local drive).
3. Double-click `RightClick.ahk`. A tray icon appears in the system tray.

**Option B — run a compiled `.exe`:**

1. Download `RightClick.exe` from the [Releases page](https://github.com/Magnetron85/Radiology-Right-Click/releases) (no AutoHotkey install needed).
2. Place it in a folder you can write to (the script saves `preferences.json` next to itself).
3. Double-click to run.

**To start with Windows**: drop a shortcut to `RightClick.ahk` (or the `.exe`) into `shell:startup` (Win+R, type `shell:startup`, press Enter, paste the shortcut).

---

## Usage at a glance

1. In Notepad or PowerScribe, **highlight** the text containing your input data.
2. **Hold Ctrl and right-click** on the selection (the default activation modifier — change in Preferences if you prefer plain right-click or a different modifier).
3. Pick a calculator from the menu.
4. A small popup shows the result, copies it to the clipboard, and parks the mouse over the **Close** button so dismissing is a single click. Hit Enter, click Close, or press Escape.

The popup text is also on the clipboard, so you can paste it straight into your report.

---

## Calculator reference

Every calculator runs on whatever text was highlighted before you opened the menu. The expected input format is shown in **`code`** blocks below; the calculator's own error message will remind you of the format if it doesn't recognize the input.

### Volumes and serial measurements

#### Ellipsoid Volume
Computes **(1/6) π a b c** from three orthogonal dimensions. Dimensions are sorted automatically (you don't have to enter them largest-first). The script infers cm vs mm: if all three values are whole numbers it assumes mm and divides by 10; any decimal point flags cm.

Selected text:
```
3.0 x 2.0 x 1.0 cm
```
Result popup:
```
3.0 x 2.0 x 1.0 cm (3.1 cc)
```

#### Bullet Volume
Same as above but uses the **5π/24** prostate "bullet" coefficient (Terris–Stamey). Same cm/mm inference.

Selected text:
```
3.5 x 5.4 x 2.5 cm
```
Result popup:
```
3.5 x 5.4 x 2.5 cm (32.1 cc)
```

#### PSA Density
PSA ÷ prostate volume. Volume is taken from any of (in order):
1. an explicit `Volume: N cc` or `Size: ... (N cc)` in the text,
2. otherwise computed from 3D dimensions — bullet volume if <55 cc, ellipsoid if larger (per the v1 convention).

Selected text:
```
PSA: 5.6 ng/mL
Size: 3.5 x 5.4 x 2.5 cm
```
Result popup:
```
PSA: 5.6 ng/mL
Size: 3.5 x 5.4 x 2.5 cm
Prostate volume: 32.1 cc (Bullet Volume)
PSA Density: 0.174 ng/mL/cc
```

#### Compare Measurement Sizes
Compares a current measurement to a prior one in the same selected text. Looks for keywords like `now`, `previously`, `prior`, `previous`, `old`, `current`. Reports per-dimension change, longest-dimension change, computed volume change, time difference (if both dates are present), doubling time, and exponential growth rate.

Selected text:
```
Now 2.5 cm, previously 1.5 cm on 01/01/2024
```
Result popup:
```
Now 2.5 cm, previously 1.5 cm on 01/01/2024

Previous Date: 01/01/2024
Current Date: Assumed as today

Dimension 1: 1.50 cm -> 2.50 cm (+66.7%)

Longest dimension change: +66.7%
Volume change: +363.0%
Previous volume: 1.8 cc
Current volume: 8.2 cc
Time difference: 2.37 years
Doubling time: 392 days
Exponential Growth Rate: 64.99% per year
```

#### Sort Measurement Sizes
**This one pastes its output back over the selection.** Finds every 2D/3D measurement in the highlighted text and rewrites it largest-dimension-first. Useful when a tech enters `1 x 2 x 3 cm` and you want `3 x 2 x 1 cm`. Your prior clipboard is preserved (saved before paste, restored after).

Before:
```
The mass measures 1.5 x 2.8 x 0.9 cm, with adjacent nodule 0.5 x 1.2 cm.
```
After paste:
```
The mass measures 2.8 x 1.5 x 0.9 cm, with adjacent nodule 1.2 x 0.5 cm.
```

### Abdominal MRI

#### Calculate Adrenal Washout
Absolute and relative washout from CT HU values across phases. Handles either three-phase input (unenhanced + enhanced + delayed) or two-phase (enhanced + delayed only — relative washout only). Interprets against the >=60 % absolute / >=40 % relative adenoma thresholds.

Reference: Mayo-Smith WW et al. *J Am Coll Radiol.* 2017;14(8):1038–1044.

Selected text:
```
Unenhanced: 10 HU, Enhanced: 80 HU, Delayed: 40 HU
```
Result popup includes:
```
Absolute Washout: 57.1% ... (Ref adenomas: >=60%)
Relative Washout: 50.0% ... (Ref adenomas: >=40%)
```
followed by the interpretation paragraph and citation.

#### Calculate Thymus Chemical Shift
Signal Intensity Index (SII) and Chemical Shift Ratio (CSR) from dual-echo in-phase/out-of-phase liver-style imaging adapted to the thymus. Reports against the Priola cutoffs (CSR <= 0.849 and SII > 8.92 % suggest hyperplasia).

Reference: Priola AM et al. *Radiology.* 2015;274(1):238–249.

Selected text:
```
Thymus IP: 100, OP: 80, Paraspinous IP: 90, OP: 85
```
Result popup includes:
```
Chemical Shift Ratio: 0.847 (hyperplasia < 0.849)
Thymus Signal Intensity Index (SII): 20.00% (hyperplasia > 8.92)
```
plus the interpretation and citation. Single-tissue input (thymus IP+OP only, no paraspinous) is also accepted and reports SII alone.

#### Calculate Hepatic Steatosis
Two-point Dixon fat fraction. Liver IP + OP only gives the standard fat fraction; if spleen IP + OP are also provided, also reports the spleen-normalized fat percentage. Bands: <5 % none, 5–15 % mild, 15–30 % moderate, ≥30 % severe.

Reference: Sirlin CB. *Radiographics* 2009;29:1277–80.

Selected text:
```
Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88
```
Result popup includes:
```
Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88 (Fat Fraction: 10.0%, Fat Percentage: 9.1%)
Fat Fraction Interpretation: Mild hepatic steatosis.
Fat Percentage Interpretation: Mild hepatic steatosis.
```

#### MRI Liver Iron Content
R2* relaxometry → liver iron concentration using the field-strength-specific calibrations from Guglielmo 2023. Also classifies severity per the same paper's Table 6 (<1.8 normal, 1.8–3.2 mild, 3.2–7.0 moderate, 7.0–15.0 severe, ≥15.0 extreme — all mg Fe/g dry liver).

Reference: Guglielmo FF et al. *Radiographics.* 2023;43(6):e220181.

Selected text:
```
1.5T, R2*: 50 Hz
```
Result popup:
```
1.5T, R2*: 50 Hz
Estimated Iron Content: 1.1 mg Fe/g dry liver
Interpretation: Normal iron content (<1.8 mg Fe/g).
```

### Cardiovascular

#### Calculate Calcium Score Percentile
Age- and sex-stratified percentile lookup against Hoff 2001 (35,246 adults, electron-beam CT). Also reports arterial age from the MESA equation (McClelland 2009: arterial age = 39.1 + 7.25·ln(Agatston+1)). Valid for ages ≥30.

Selected text:
```
Age: 55
Sex: Male
Coronary artery calcium score: 250 (Agatston)
```
Result popup includes:
```
Plaque Burden: At least moderate atherosclerotic plaque. Mild coronary artery disease highly likely...
Comparison to people of the same age and sex: High (75-90%)
Arterial Age: 79 years
```

#### Calculate NASCET
Carotid stenosis percentage from the distal ICA and residual lumen diameters. Reports the percentage plus mild/moderate/severe grade (NASCET cutoffs: <50, 50–69, >=70).

Reference: NASCET Collaborators. *N Engl J Med* 1991;325:445–453.

Selected text:
```
Distal ICA = 6 mm, Stenosis = 2 mm
```
Result popup:
```
NASCET Calculation:
Distal: 6 mm
Stenosis: 2 mm
NASCET: 66.7%
Moderate (50-69%)
```

### Lung / pulmonary

#### Calculate Fleischner Criteria
Parses a sentence describing one or more pulmonary nodules, infers multiplicity, composition (solid / ground-glass / part-solid), calcification, lobar location, morphology, and size, then returns the **2017 Fleischner Society** follow-up recommendation. If risk factors are detected (emphysema, fibrosis, spiculation) the high-risk recommendation is given; otherwise low-risk and high-risk are both shown. Also computes the actual follow-up date range (e.g. "6–12 months is November 2026 to May 2027 from May 2026").

Reference: MacMahon H et al. *Radiology.* 2017;284(1):228–243.

Selected text:
```
Solitary 7 mm solid pulmonary nodule in the right upper lobe.
```
Result popup includes:
```
A solitary pulmonary nodule is described forming the basis of follow-up:
- Location: right upper lobe
- Extracted (or inferred) Size: 7.0 mm
- Fleischner Size (mm): 7.0 mm
- Composition (or inferred): solid

FLEISCHNER SOCIETY RECOMMENDATION:
For low-risk patients: CT at 6-12 months, then consider CT at 18-24 months...
Follow-up dates: 6-12 months is November 2026 to May 2027 from May 2026.
```

### Obstetric / gynecologic

#### Calculate Pregnancy Dates
Naegele's rule. Two input modes:

- **LMP date** → reports estimated delivery date (LMP + 280 d) and current gestational age. If GA exceeds full term + postdates, an annotation flags the LMP as likely stale rather than reporting a clinically meaningless GA.
- **Gestational age (with optional reference date)** → back-calculates LMP, then EDD and current GA.

Selected text (LMP):
```
LMP: 01/15/2025
```
Result popup:
```
LMP: 01/15/2025
Estimated Delivery Date: 10/22/2025
Current Gestational Age: 69 weeks 4 days -- past EDD by 207 days (LMP likely stale)
```

Selected text (GA):
```
GA: 12 weeks and 3 days as of today
```

#### Calculate Menstrual Phase
Estimates cycle day from LMP (28-day cycle) and reports phase plus expected endometrial thickness.

Selected text:
```
LMP: 05/01/2026
```
Result popup:
```
LMP: 05/01/2026
Current Cycle Day: 17/28

Secretory Phase
Expected endometrial stripe thickness: 7-16 mm
```

### Numeric utilities

#### Calculate Statistics
Pulls every number out of the selected text and reports count, sum, mean, median, min, max. If n ≥ 9 it also reports Q1, Q3, IQR, IQR/median, and standard deviation. Recognizes label prefixes like `Slice 4:` or `Sample 2:` and strips them so the slice number isn't counted as data.

Selected text:
```
1.0, 2.5, 3.7, 4.0, 5.2, 6.1, 7.8, 8.4, 9.0, 10.5
```
Result popup:
```
Statistics:
Count: 10
Sum: 58.2
Mean: 5.8
Median: 5.7
Min: 1.0
Max: 10.5
Q1: 3.8
Q3: 8.3
IQR: 4.5
IQR/Median: 0.79
Standard Deviation: 3.1
```

#### Calculate Number Range
Reports `min - max` with the first unit it finds attached.

Selected text:
```
values 5 to 10 cm
```
Result popup:
```
5.0 - 10.0 cm
```

### Scheduling

#### Calculate Contrast Premedication
Opens a dialog (not a popup) where you pick the scan date, scan time, premedication protocol (Prednisone 13-7-1 or Methylprednisolone 12-2 — ACR Manual on Contrast Media), and whether to include diphenhydramine. **Calculate** produces a dated timeline of when each dose is due; **Show Dosages** shows the protocol's dosing card without a timeline.

The text-selection part of the workflow doesn't apply here — this calculator ignores any highlighted text.

---

## Preferences

`Preferences` at the bottom of the right-click menu opens the settings window:

- **Dark mode** — flips title bars, popup menus, and dialogs to a dark palette (uses Windows' native immersive dark mode plus the undocumented `uxtheme!SetPreferredAppMode` for popup menus, the same mechanism File Explorer and RegEdit use).
- **Show citations in output** — toggles the in-line journal citations on each result.
- **Show arterial age** — toggles the MESA arterial age line on the calcium score output.
- **Calculator visibility** — 15 checkboxes, one per toggleable calculator. Compare / Sort Measurement Sizes are always visible.
- **Menu sorting** — `alphabetical` (default, case-insensitive by title), `frequency` (most-used first, alphabetical for ties), or `none` (registry order).
- **Right-click modifier** — `none` (plain right-click), `Ctrl` (default), `Alt`, or `Shift`. Determines what combination opens the helper menu.
- **References...** — opens the reference manager (see below).
- **Target Apps...** — opens the target-app editor (see below).
- **Restore Defaults** — resets toggles, sort method, modifier, and target apps back to defaults. Your saved references and frequency counts are kept.

Preferences are stored in `preferences.json` next to the script. A `preferences.backup.json` is written on every successful save and is restored if the live file is unreadable on the next launch.

A v1 `preferences.ini` next to the script is imported automatically on first run, then renamed to `preferences.ini.imported` so it isn't re-imported.

---

## Saved references

The **References** submenu lets you save:

- a URL (http or https only — `javascript:`, `file:`, `ftp:` schemes are rejected),
- or a document (PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX).

Documents are copied into a local `References/` folder next to the script under a sanitized name (no path separators, no `..` traversal, no Windows reserved device names). Each reference tracks a usage counter; the submenu sorts by usage so the ones you click most appear first. Maximum 15 visible in the submenu — the full list is reachable via **References -> Manage References...**.

---

## Configuring target apps

By default, the helper menu only opens in:

- Notepad (class `Notepad` / exe `notepad.exe`)
- PowerScribe (class `PowerScribe` / exe `PowerScribe.exe`)
- PowerScribe 360 (class `PowerScribe360` / exe `Nuance.PowerScribe360.exe`)
- "PowerScribe | Reporting" (window class)

Open **Preferences -> Target Apps...** to add or remove entries. The format is one entry per line:

```
class:WindowClass
exe:Process.exe
```

For example, to add a custom PACS client called `MyPACS.exe` with window class `Inteleviewer`:

```
class:Notepad
exe:notepad.exe
class:Inteleviewer
exe:MyPACS.exe
class:PowerScribe360
exe:Nuance.PowerScribe360.exe
```

Use Window Spy (ships with AutoHotkey, in the tray menu) to look up an unknown window's class and process name.

---

## Activation modifier

The helper menu is gated behind a modifier so plain right-click stays native in the host application. To change it:

**Preferences -> Right-click modifier** -> pick one of:

- `none` — plain right-click opens the helper menu (and the app's native context menu is suppressed in target windows)
- `Ctrl + right-click` — **default**
- `Alt + right-click`
- `Shift + right-click`

Changes take effect immediately — no script restart needed.

Fn isn't offered. On most laptops it's intercepted by the keyboard firmware before Windows sees it, so AutoHotkey can't hook it. Shift is the practical alternative for a pinky-friendly safety guard.

---

## Project layout

```
RightClick.ahk              Main entry, hotkey wiring, target-app fallback list
lib/
  Json.ahk                  Minimal JSON parser / stringifier
  Prefs.ahk                 JSON-backed preferences + v1 INI import
  Modern.ahk                Dark title bar, rounded corners, dark popup menus
  UI.ahk                    Result popup window
  Menu.ahk                  Right-click menu build + sort
  PreferencesWindow.ahk     Settings GUI + target-app editor
  References.ahk            URL / file reference manager
  Util.ahk                  Shared helpers (regex, dates, clipboard, safety)
  calc/                     One file per calculator (13 total)
```

---

## License

[PolyForm Noncommercial 1.0.0](https://polyformproject.org/licenses/noncommercial/1.0.0). Free for personal, clinical, academic, and research use, including use within hospitals, universities, and other non-commercial / public-benefit organizations. **Commercial use is not permitted** — the software may not be sold, sublicensed for a fee, or incorporated into a commercial product or paid service. For commercial licensing inquiries, contact the author.

See `LICENSE` for the full text.
