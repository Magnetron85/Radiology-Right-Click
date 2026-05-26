# Security Policy

## Scope

This script is intended for use on a single clinical workstation by a single radiologist. It does **not**:

- Make network requests of its own
- Authenticate to any service
- Store, transmit, or process PHI
- Run with elevated privileges
- Install services, drivers, or scheduled tasks
- Persist anything outside its own folder (`preferences.json`, `preferences.backup.json`, and the `References/` sub-folder it creates)

The only data it reads is text the user has actively highlighted in the foreground window. That text is copied to the clipboard (via Ctrl+C) and discarded after the calculation. The previous clipboard contents are saved and restored.

> **Debug logging caveat:** diagnostic logging is **off by default**. If you turn it on (`"debug": {"capture": true}` in `preferences.json`), the first 60 characters of the selected text are written to `debug.log` next to the script. On a clinical workstation this can place PHI on disk. Leave debug logging off in clinical use, and do not commit or share `debug.log` (it is git-ignored). The "does not store PHI" guarantee above holds only with debug logging disabled.

## Disclaimer -- not for clinical use

**This software is provided for entertainment, research, and educational purposes only.** It is **not a medical device**, has not been validated for clinical use, and **must not be relied upon for clinical decision-making**. The calculations are derived from published formulas (cited in-line in each calculator) but have not been independently validated. Every reported value must be independently verified by the user against the source data before it is incorporated into a clinical report or used to inform patient care. Use of this software does not establish a clinician-patient relationship and does not constitute medical advice. The author disclaims all liability for any clinical, diagnostic, therapeutic, or other decision made on the basis of this software's output.

## Reporting a vulnerability

If you find a security issue (path traversal, unsafe command execution, clipboard handling that leaks data outside the script, etc.), please report it privately:

- Open a private security advisory at https://github.com/Magnetron85/Radiology-Right-Click/security/advisories/new
- Or email the maintainer at the address listed on the GitHub profile

Please do **not** open a public issue for security reports.

Expected response: acknowledgement within 7 days. Fix or mitigation within 30 days for confirmed issues.

## Supported versions

Only the latest tagged release is supported. v1.x is no longer maintained.

## Threat model assumptions

- The user trusts the contents of their own clipboard.
- The user trusts the references they manually add (URLs and files). The script validates URL schemes (http/https only) and file extensions (PDF/DOC/DOCX/XLS/XLSX/PPT/PPTX), copies referenced files into a local sanitised path, and re-validates before opening, but it cannot guarantee the safety of an arbitrary URL or document the user has chosen to save.
- The user trusts the AutoHotkey v2 runtime they have installed.
