# Impression Style Guide

Conventions every calculator's `CalcResult` fields should follow. Goal: a single radiologist (you) can paste the `impression` field straight into a PowerScribe report with zero edits.

**Status: DRAFT — red-pen freely.**

---

## Field semantics

| Field | Audience | Style |
|---|---|---|
| `classification` | radiologist (label) | Short label, exact code/category. "Lung-RADS 4A", "Bosniak IIF", "Low risk", "Cat 0". No prose. |
| `impression` | the dictated report | One sentence, prose. Measurements + verdict + guideline reference + (sometimes) recommendation. **The default clipboard target.** |
| `recommendation` | sometimes a separate report section | Imperative voice, one sentence. "Recommend 3-month LDCT." Often duplicated inside `impression` for stand-alone pasting. |
| `findings` (optional) | the Findings section | Fragment style, no guideline reference, no recommendation. "8 mm part-solid RUL nodule." |
| `advisories` (array) | radiologist on screen | One-liners flagging caveats / source gaps / population-specific notes. Don't paste into a report by default. |
| `methodology` | audit only | Multi-line. Selected inputs + reasoning. Hidden by default in clinical workflow. |
| `citations` (array) | audit/reference | `{ text, url }` per item. Links go to ACR Reporting-and-Data-Systems pages or DOI. |
| `echo` (optional) | UI only | The raw user selection, shown only if the calc *transformed* it (e.g. parsed dimensions). |
| `error` | error path | String. If non-empty, only `impression` (= the error message) is rendered. |

## Impression sentence — required ingredients

In this order:
1. **What the lesion is** with key measurement (`"8 mm part-solid pulmonary nodule"`)
2. **Verdict / category** (`"Lung-RADS 4A"`)
3. **Guideline reference inline** (`"per Lung-RADS v2022"` or `"per SRU 2022"`). Always abbreviated, always inline — preempts "why this f/u?".
4. **Recommendation, if short enough** (`"recommend 3-month LDCT."`)

If the recommendation is long, put it as a second sentence — but always end the `impression` with the recommendation, not the verdict.

### Examples

| Calc type | Good | Bad |
|---|---|---|
| Risk classification | `"20 mm pedunculated gallbladder polyp, low-risk per SRU 2022; recommend follow-up US in 12 months."` | `"Risk: Low. F/U: 12mo."` (telegraphic — looks like a UI dump) |
| Pure measurement | `"Prostate volume 42 mL (ellipsoid); PSA density 0.12 ng/mL/mL."` | `"Volume = 42 mL. Density = 0.12."` (no units, no method) |
| Lookup | `"Coronary calcium score 312, 78th percentile for age/sex (MESA)."` | `"Percentile: 78%."` (no score, no comparator) |
| Recommendation-driven | `"Lung-RADS 4B (15 mm solid nodule) — recommend diagnostic chest CT and/or tissue sampling."` | `"Cat 4B. See ACR."` (vague, unhelpful) |

## Universal formatting rules

1. **ASCII only.** No smart quotes, em-dashes, NBSP, bullet characters, ellipses (`…`). Use `--` for em-dash, `...` for ellipsis, straight quotes. PowerScribe / downstream EHRs choke on Unicode punctuation.
2. **Units lowercase, ACR-standard**: `mm`, `cm`, `mL`, `HU`, `mGy`. Not `ML`, `Mm`, `Hu`.
3. **Space between number and unit**: `8 mm`, not `8mm`. (Already convention but enforce it.)
4. **Single space between sentences.**
5. **No trailing newline.** Single line of impression text terminates without `\n`.
6. **Active voice in recommendations.** `"Recommend X"` not `"X is recommended"` or `"X recommended"`.
7. **Hedging language conventions**:
   - Definitive: `"X meets criteria for Y"` / `"consistent with Y"`
   - Probabilistic: `"likely Y"` / `"compatible with Y"` / `"suspicious for Y"`
   - Uncertain: `"may represent Y"` / `"cannot exclude Y"`
   Avoid `"appears to be"` (vague), `"definitely Y"` (over-claim), `"r/o Y"` (abbreviation belongs in residency, not reports).
8. **Measurements in the impression** must include all dimensions used by the calc — don't bury input dims in methodology only.
9. **Guideline references are abbreviated, not spelled out**: `"per SRU 2022"`, `"per Lung-RADS v2022"`, `"per Bosniak v2019"`, `"per Fleischner 2017"`. The full citation lives in `citations`.

## Per-section style

### Findings (`findings` field)

- Fragment, no verb required: `"8 mm RUL part-solid nodule"`, `"Right ovarian cyst, 4 cm, anechoic"`.
- No guideline reference, no recommendation.
- Only populate this if the calc deals with a discrete imaging finding the radiologist would also describe in the Findings section. Pure-math calcs (PSA density, prostate volume) leave it empty.

### Recommendation (`recommendation` field)

- Self-contained — readable without the impression context.
- Imperative voice: `"Recommend 3-month LDCT."`, `"Recommend MRI/MRCP follow-up at 6 months."`, `"Surgical consultation recommended."`
- Include intervals + modality, never just "follow-up."

### Methodology (`methodology` field)

- Multi-line. Free-form sections OK (`"Selected inputs:`, `"Reason:", `"Intermediate:"`).
- Plain ASCII, no fancy formatting. Use `--` not bullets.
- May include "why this category" reasoning, intermediate computations, formula references.

### Citations (`citations` field, array)

Each citation:
```ahk
{ text: "Kamaya A, Fung C, Szpakowski JL, et al. Management of Incidentally Detected Gallbladder Polyps. Radiology. 2022;305(2):277-289.",
  url:  "https://pubs.rsna.org/doi/10.1148/radiol.213079" }
```

URL preference order:
1. **ACR Reporting and Data Systems page** if applicable (Lung-RADS, LI-RADS, TI-RADS, O-RADS, PI-RADS, etc.).
2. Original article DOI on the journal site (RSNA, NEJM, JACC, etc.).
3. ACR Incidental Findings white-paper page for non-RADS guidelines (adrenal, thyroid incidentaloma).

## Anti-patterns to avoid

| Don't | Do |
|---|---|
| `"This polyp is low risk."` (vague subject) | `"20 mm pedunculated gallbladder polyp, low-risk per SRU 2022..."` |
| `"Risk category: Low"` (UI label exposed) | `"...low-risk..."` (woven into prose) |
| `"f/u in 12 mo"` (chart shorthand) | `"recommend follow-up US in 12 months"` |
| `"As per the ACR Lung-RADS v2022 Assessment System..."` (verbose) | `"per Lung-RADS v2022"` |
| `"Computed using the ellipsoid formula..."` in the impression | Mention `"(ellipsoid)"` parenthetically if needed; full method goes in `methodology` |
| Echoing measurements not used in the calc | Only impression-relevant measurements |

---

**Reviewer notes**: anything below the line is your edit space. Strike, replace, add examples. Once locked, this is the contract every Stage 2 calc rewrite must satisfy.

---
