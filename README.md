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
- [RADS and incidental-findings classifiers (v2.1)](#rads-and-incidental-findings-classifiers-v21)
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
Two-point Dixon fat fraction. Liver IP + OP only gives the standard fat fraction; if spleen IP + OP are also provided, also reports the spleen-normalized fat fraction. Grading bands follow Guglielmo 2023 (Table 4, rounded from Tang 2013): <5 % none, 5–6 % borderline (meets the ≥5 % diagnostic cutoff — 5.56 % by MRS, Dallas Heart Study — but below the 6 % Grade 1 threshold), 6–17 % mild (Grade 1), 17–22 % moderate (Grade 2), >22 % severe (Grade 3).

References: Guglielmo FF et al. *RadioGraphics* 2023;43(6):e220181; Tang A et al. *Radiology* 2013;267:422–431; Sirlin CB. *Radiographics* 2009;29:1277–80 (spleen normalization).

Selected text:
```
Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88
```
Result popup includes:
```
Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88 (Fat Fraction: 10.0%, Spleen-normalized FF: 9.1%)
Fat Fraction Interpretation: Mild hepatic steatosis (Grade 1).
Spleen-normalized FF Interpretation: Mild hepatic steatosis (Grade 1).
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
Age- and sex-stratified percentile against Hoff 2001 (35,246 adults, electron-beam CT). When race is supplied (ages 45–84) it additionally reports the race-stratified **MESA** percentile (McClelland 2006). Rather than bucketing the paper's bracket-midpoint table, this evaluates the MESA reference model per integer age × sex × race, with log-linear interpolation across score — so a patient is scored at their actual age instead of the bracket midpoint (e.g. white male, age 46, Agatston 45 → 89th, not the old bracket-midpoint 75th). Also reports arterial age (McClelland 2009: 39.1 + 7.25·ln(Agatston+1)). Valid for ages ≥30 (MESA portion 45–84).

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

#### RV/LV Ratio (PE)
Right-to-left ventricular short-axis diameter ratio on CT, a marker of right-heart strain in acute pulmonary embolism. Enter the RV and LV maximum minor-axis diameters (endocardium to interventricular septum); pre-filled from "RV ... mm / LV ... mm" in the selection. Reports the ratio with the published ≥ 1.0 strain threshold noted (≥ 0.9 has also been used on axial images).

Reference: Meinel FG et al. *Am J Med.* 2015;128(7):747–759.

Selected text:
```
RV 42 mm, LV 35 mm
```
Result:
```
RV/LV diameter ratio 1.20 (RV 42 mm, LV 35 mm). Ratio >= 1.0, associated with right ventricular strain/dysfunction in acute pulmonary embolism.
```

### Neuro / head

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

#### ICH Volume (ABC/2)
Intracerebral hemorrhage volume by the **ABC/2** method: A × B × C ÷ 2, where A, B, C are the three orthogonal hematoma diameters in cm (C is the vertical extent, conventionally slice count × thickness), giving volume in mL. Pre-fills the three dimensions from a measurement in the selection; accepts cm or mm. Output is factual (volume only) — no prognostic threshold is asserted.

Reference: Kothari RU et al. *Stroke.* 1996;27(8):1304–1305.

Selected text:
```
4.0 x 3.0 x 2.5 cm
```
Result:
```
Intracerebral hemorrhage volume 15.0 mL by the ABC/2 method (4 x 3 x 2.5 cm).
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

#### Follow-up Date
Computes a recommended follow-up date from a base (study) date plus an interval (days / weeks / months / years). The base date defaults to today, is pre-filled from a MM/DD/YYYY date in the selection when present, and is editable; the interval is pre-filled from phrasing like "6 month follow-up". Month and year arithmetic is calendar-correct — end-of-month dates clamp (Jan 31 + 1 month → Feb 28/29) and leap years are handled.

Selected text:
```
Recommend 6 month follow-up.
```
Result (base date today):
```
Recommended follow-up by 12/10/2026 (6 months from 06/10/2026).
```

---

## RADS and incidental-findings classifiers (v2.1)

v2.1 adds twelve guided classifiers covering the major ACR Reporting and Data Systems and the ACR incidental-findings white papers. Unlike the text-parse calculators above, these **pre-fill a form from the selected text and then open a dialog** where you confirm or complete the inputs before submitting. The result popup returns a report-ready impression, with optional methodology and citation sections (toggle in Preferences). As with every calculator here, the output must be independently verified before use — see the disclaimer at the top.

The example below shows the report-ready impression line for each; the live popup also includes the category breakdown, methodology, and citation.

### Renal

#### Bosniak (cystic renal mass, v2019)
Classifies a cystic renal mass into Bosniak 2019 categories I, II, IIF, III, or IV from wall/septa thickness, enhancement, and convex protrusion (nodule) size/margins, with modality-specific features (CT hyperattenuation; MRI T1/T2 signal). Reports category, malignancy-risk range, and management.

Reference: Silverman SG, Pedrosa I, Ellis JH, et al. *Radiology.* 2019;292(2):475-488.

Selected text:
```
Simple-appearing cyst in the left kidney with thin nonenhancing septa.
```
Result popup includes (after submitting the dialog -- CT, 2 thin septa <=2 mm, no enhancement):
```
Bosniak II renal cyst (few thin septa), reported malignancy risk <1%; no follow-up recommended per Bosniak v2019.
```

### Liver / biliary

#### Gallbladder Polyp (SRU 2022)
Stratifies an incidentally detected gallbladder polyp by size and morphology (pedunculated thin/thick stalk, sessile, ball-on-wall, adjacent wall thickening) into extremely low / low / indeterminate risk or a surgical-consultation indication, with the SRU 2022 follow-up schedule. Exclusion flags (PSC, suspicious features, poor visualization) report that the algorithm does not apply.

Reference: Kamaya A, Fung C, Szpakowski JL, et al. *Radiology.* 2022;305(2):277-289.

Selected text:
```
12 mm pedunculated polyp on the gallbladder wall with a thick stalk.
```
Result popup includes:
```
12 mm pedunculated thick-stalk gallbladder polyp, low risk per SRU 2022; recommend follow-up US at 6, 12, 24, and 36 months.
```

#### LI-RADS (CT / MRI 2018)
Applies the LI-RADS v2018 diagnostic algorithm to a liver observation in an at-risk patient, returning LR-1 through LR-5, LR-M, LR-TIV, or LR-NC from major features (non-rim APHE, size, non-peripheral washout, enhancing capsule, threshold growth, tumor in vein), LR-M criteria, benign overrides, and ancillary-feature counts.

Reference: Chernyak V, Fowler KJ, Kamaya A, et al. *Radiology.* 2018;289(3):816-830.

Selected text:
```
18 mm arterial phase hyperenhancing mass with non-peripheral washout in segment 6.
```
Result popup includes:
```
18 mm hepatic observation, LR-5 (definitely HCC) per LI-RADS v2018; multidisciplinary discussion recommended.
```

#### US LI-RADS (v2024)
Categorizes a surveillance hepatic ultrasound as US-1 (negative), US-2 (subthreshold), or US-3 (positive) from focal-observation presence/size/benignity, new vascular thrombus, and parenchymal distortion, and reports the visualization score (A/B/C) with its management.

Reference: American College of Radiology. *US LI-RADS v2024 -- Ultrasound Surveillance.*

Selected text:
```
15 mm solid focal observation in the right hepatic lobe; surveillance ultrasound.
```
Result popup includes:
```
Surveillance US US-3 (positive), visualization A, per US LI-RADS v2024; diagnostic multiphase CT or MRI recommended.
```

### Adrenal

#### Incidental Adrenal (ACR 2017)
Works up an incidental adrenal mass per the ACR 2017 white paper, branching on known extra-adrenal malignancy, CT attenuation (auto-computes absolute and relative washout when unenhanced/enhanced/delayed HU are entered), MRI chemical shift, prior stability, size, and morphology. Avid enhancement (>=110 HU) adds a pheochromocytoma caution.

Reference: Mayo-Smith WW, Song JH, Boland GL, et al. *J Am Coll Radiol.* 2017;14(8):1038-1044.

Selected text:
```
2.5 cm left adrenal nodule, unenhanced 8 HU. No known malignancy.
```
Result popup includes:
```
25 mm incidental adrenal mass, BENIGN per ACR Incidental Findings 2017; no imaging follow-up required (lipid-rich adenoma, unenhanced 8 HU <=10 HU).
```

### Thyroid / neck

#### Incidental Thyroid (ACR 2015)
Decides whether an incidentally detected thyroid nodule warrants dedicated ultrasound, using the ACR 2015 size/age thresholds (>=1.0 cm if age <35, >=1.5 cm if age >=35). Clinical risk factors, suspicious features, and focal FDG uptake override the size criteria; limited life expectancy yields a no-workup recommendation.

Reference: Hoang JK, Langer JE, Middleton WD, et al. *J Am Coll Radiol.* 2015;12(2):143-150.

Selected text:
```
1.8 cm thyroid nodule incidentally noted on CT in a 58-year-old.
```
Result popup includes:
```
1.8 cm incidental thyroid nodule; dedicated thyroid ultrasound recommended per ACR 2015 (age >=35 with nodule >=1.5 cm).
```

#### ACR TI-RADS
Scores a thyroid nodule with the ACR TI-RADS point system (composition, echogenicity, shape, margin, echogenic foci), maps the total to TR1-TR5, and gives the size-dependent FNA / follow-up recommendation. Cystic and spongiform nodules force TR1.

Reference: Tessler FN, Middleton WD, Grant EG, et al. *J Am Coll Radiol.* 2017;14(5):587-595.

Selected text:
```
1.5 cm solid hypoechoic thyroid nodule with punctate echogenic foci, wider-than-tall, smooth margins.
```
Result popup includes (solid +2, hypoechoic +2, punctate +3 = 7 pts):
```
1.5 cm thyroid nodule, TR5 (highly suspicious) per ACR TI-RADS 2017; FNA recommended (>=1.0 cm for TR5).
```

### Pancreas

#### Kyoto IPMN (2024)
Applies the 2024 Kyoto guidelines for IPMN, separating high-risk stigmata (enhancing mural nodule >=5 mm, MPD >=10 mm, obstructive jaundice, suspicious/positive cytology) from worrisome features (cyst >=30 mm, nodule <5 mm, wall thickening, MPD 5-9 mm, abrupt duct change, lymphadenopathy, growth >=2.5 mm/yr, elevated CA 19-9, new diabetes, acute pancreatitis), and reports the category and management.

Reference: Ohtsuka T, Fernandez-del Castillo C, Furukawa T, et al. *Pancreatology.* 2024;24(2):255-270.

Selected text:
```
22 mm branch-duct pancreatic cyst with a 6 mm enhancing mural nodule.
```
Result popup includes:
```
22 mm branch-duct IPMN with 1 high-risk feature (enhancing mural nodule 6 mm) per Kyoto 2024; recommend surgical consultation / multidisciplinary discussion.
```

### Lung / pulmonary

#### Lung-RADS (v2022)
Assigns a Lung-RADS v2022 category (0, 1, 2, 3, 4A, 4B, 4X, with optional S modifier) to a screening pulmonary nodule from lesion type, mean and solid-component diameter, screening round, prior comparison and growth, benign-feature overrides, and suspicious features (spiculation, lymphadenopathy, etc.).

Reference: ACR Committee on Lung-RADS. *Lung CT Screening Reporting and Data System (Lung-RADS) v2022.* ACR, November 2022.

Selected text:
```
10 mm solid spiculated nodule in the right upper lobe on baseline screening CT.
```
Result popup includes (solid 10 mm baseline = 4A, spiculation upgrades to 4X):
```
10 mm solid pulmonary nodule, Lung-RADS 4X per Lung-RADS v2022; recommend diagnostic CT, PET/CT, or tissue sampling.
```

### Obstetric / gynecologic

#### O-RADS MRI
Risk-stratifies an adnexal lesion on MRI into O-RADS MRI scores 1-5 from lesion type, wall/septal enhancement and fluid type, solid-tissue T2/DWI signal, DCE time-intensity curve, and peritoneal nodularity (which overrides to Score 5). Reports the score, malignancy risk, and management.

Reference: Thomassin-Naggara I, Poncelet E, Jalaguier-Coudray A, et al. *JAMA Netw Open.* 2020;3(1):e1919896.

Selected text:
```
Adnexal lesion with enhancing solid tissue and a type 3 time-intensity curve.
```
Result popup includes:
```
Adnexal lesion, O-RADS MRI 5 (high risk) per O-RADS MRI 2020; gynecologic oncology referral.
```

#### O-RADS US (v2022)
Risk-stratifies an adnexal lesion on ultrasound into O-RADS US scores 1-5 from menstrual status, lesion type and size, inner contour, solid component, classic benign patterns (dermoid, endometrioma, hemorrhagic cyst, hydrosalpinx), papillary-projection count, and color score. Ascites or peritoneal nodularity upgrades any score >=3 to 5.

Reference: Andreotti RF, Timmerman D, Strachowski LM, et al. *Radiology.* 2020;294(1):168-185.

Selected text:
```
Postmenopausal anechoic simple cyst measuring 5 cm with a thin smooth wall.
```
Result popup includes:
```
5.0 cm adnexal lesion, O-RADS US 2 (almost certainly benign) per O-RADS 2022; no follow-up if <10 cm.
```

### Prostate

#### PI-RADS (v2.1)
Scores a prostate MRI lesion per PI-RADS v2.1 using the dominant sequence for the lesion's zone (DWI in the peripheral zone, with DCE upgrading an equivocal DWI=3 to 4; T2 in the transition zone, with DWI as a tiebreaker). Score-4 morphology is promoted to PI-RADS 5 when the lesion is >=1.5 cm or shows definite extraprostatic extension.

Reference: Turkbey B, Rosenkrantz AB, Haider MA, et al. *Eur Urol.* 2019;76(3):340-351.

Selected text:
```
Peripheral zone lesion, 12 mm, markedly hypointense on DWI with low ADC.
```
Result popup includes:
```
12 mm peripheral-zone prostate lesion, PI-RADS 4 (csPCa likely) per PI-RADS v2.1; targeted biopsy recommended.
```

---

## Preferences

`Preferences` at the bottom of the right-click menu opens the settings window:

- **Dark mode** — flips title bars, popup menus, and dialogs to a dark palette (uses Windows' native immersive dark mode plus the undocumented `uxtheme!SetPreferredAppMode` for popup menus, the same mechanism File Explorer and RegEdit use).
- **Show citations in output** — toggles the in-line journal citations on each result.
- **Show arterial age** — toggles the MESA arterial age line on the calcium score output.
- **Calculator visibility** — one checkbox per toggleable calculator (including the v2.1 RADS classifiers). Compare / Sort Measurement Sizes are always visible.
- **Menu sorting** — `alphabetical` (default, case-insensitive by title), `frequency` (most-used first, alphabetical for ties), or `none` (registry order).
- **Right-click modifier** — `none` (plain right-click), `Ctrl` (default), `Alt`, or `Shift`. Determines what combination opens the helper menu.
- **Show floating launcher widget** — a small always-on-top button you can leave anywhere on screen. Left-click opens the menu (no need to be over a reporting window), left-drag moves it (position is remembered), right-click gives Open / Hide / Preferences. It never steals focus, so the highlighted report text is preserved when you click it. Off by default.
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
  CalcResult.ahk            Structured result object (impression / category / citations)
  FormGui.ahk               Guided input form used by the v2.1 RADS classifiers
  TextScan.ahk              Selected-text parsing + form pre-fill helpers
  Menu.ahk                  Right-click menu build + sort (text + grouped/anatomical)
  PreferencesWindow.ahk     Settings GUI + target-app editor
  References.ahk            URL / file reference manager
  Util.ahk                  Shared helpers (regex, dates, clipboard, safety)
  Debug.ahk                 Conditional logging framework (off by default)
  calc/                     One file per calculator (28 total)
```

---

## License

[PolyForm Noncommercial 1.0.0](https://polyformproject.org/licenses/noncommercial/1.0.0). Free for personal, clinical, academic, and research use, including use within hospitals, universities, and other non-commercial / public-benefit organizations. **Commercial use is not permitted** — the software may not be sold, sublicensed for a fee, or incorporated into a commercial product or paid service. For commercial licensing inquiries, contact the author.

See `LICENSE` for the full text.
