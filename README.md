# 🛡️ Nessus CIS Compliance Toolkit

> **Two-phase approach: generate a clean audit template, then auto-fill it from Nessus scan results.**  
> Works with CIS Benchmarks, DISA STIGs, and any Nessus `.audit` format.

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=flat-square&logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
![Excel](https://img.shields.io/badge/Output-Excel%20.xlsx-217346?style=flat-square&logo=microsoft-excel&logoColor=white)
![CIS](https://img.shields.io/badge/Standard-CIS%20Benchmarks-red?style=flat-square)
![DISA](https://img.shields.io/badge/Standard-DISA%20STIG-blue?style=flat-square)
![Nessus](https://img.shields.io/badge/Tool-Nessus%20%2F%20Tenable-00B388?style=flat-square)

---

## 📋 Table of Contents

- [Why Two Scripts](#-why-two-scripts)
- [What This Toolkit Solves](#-what-this-toolkit-solves)
- [How It Works](#-how-it-works)
- [Installation](#-installation)
- [Phase 1 — audit\_template.py](#-phase-1--audit_templatepy)
- [Phase 2 — report\_generator.py](#-phase-2--report_generatorpy)
- [Column Reference](#-column-reference)
- [Real-Life Use Cases](#-real-life-use-cases)
- [Troubleshooting](#-troubleshooting)
- [Requirements](#-requirements)

---

## 🔀 Why Two Scripts

Most compliance tools try to do everything in one command. This toolkit deliberately separates two concerns that have different timing, ownership, and data sources:

| | `audit_template.py` | `report_generator.py` |
|---|---|---|
| **When** | Before scanning | After scanning |
| **Input** | `.audit` file(s) | `.audit` file(s) + Nessus CSV export(s) |
| **Output** | Blank Excel template | Filled Excel report |
| **Status column** | Empty | PASSED / FAILED / WARNING |
| **Observed Value column** | Empty | Actual system value from scan |
| **Default Value column** | From `.audit` file | From `.audit` file |
| **Use for** | Manual assessments, client templates | Automated scan reporting |

The `.audit` file is the **single source of truth** for benchmark content — descriptions, recommendations, and expected values always come from it, never from the CSV. The CSV contributes only two things: **Status** and **Observed Value**.

---

## 🔥 What This Toolkit Solves

Every security team doing CIS hardening or DISA STIG compliance knows the pain:

| The Old Way 😩 | With This Toolkit ✅ |
|---|---|
| Copy-paste benchmarks from PDF | Extracted automatically from `.audit` file |
| Manually type 226+ check descriptions | Generated in seconds |
| Status and findings in a separate spreadsheet | Auto-merged into one coloured Excel |
| One section at a time | All checks at once, sorted by benchmark ID |
| `N/A` everywhere when data is missing | Clean blank cells — never fake data |
| Rebuild the template for every engagement | Generate once per audit file, reuse always |

**Time saved per engagement: 4–8 hours of manual formatting → under 30 seconds.**

---

## 🔄 How It Works

```
┌──────────────────────────────────────────────────────────────────────┐
│                        TWO-PHASE WORKFLOW                            │
├──────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  PHASE 1 — Template                     PHASE 2 — Report            │
│  ─────────────────                      ─────────────────           │
│                                                                      │
│  .audit file                            .audit file                 │
│       │                                      │                      │
│       ▼                                      ▼                      │
│  audit_template.py          +     report_generator.py               │
│       │                            +         │                      │
│       │                       Nessus CSV     │                      │
│       ▼                            │         ▼                      │
│  template.xlsx                     └──► report.xlsx                 │
│  (blank Status &                        (filled Status &            │
│   Observed Value)                        Observed Value)            │
│                                                                      │
│  ← Share with team for        ← Attach to management report         │
│    manual assessment            or client delivery                  │
│                                                                      │
└──────────────────────────────────────────────────────────────────────┘
```

---

## ⚙️ Installation

### Step 1 — Prerequisites

- **Python 3.8 or higher** → [Download from python.org](https://www.python.org/downloads/)

Verify:

```bash
python --version      # Windows
python3 --version     # macOS / Linux
```

### Step 2 — Get the Scripts

```bash
# Clone
git clone https://github.com/your-org/nessus-cis-excel.git
cd nessus-cis-excel

# Or download individually
curl -O https://raw.githubusercontent.com/your-org/nessus-cis-excel/main/audit_template.py
curl -O https://raw.githubusercontent.com/your-org/nessus-cis-excel/main/report_generator.py
```

### Step 3 — Install Dependencies

```bash
pip install openpyxl
```

**Virtual environment (recommended):**

```bash
python -m venv venv
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows

pip install openpyxl
```

`requirements.txt`:
```
openpyxl>=3.1.0
```

> **Windows tip:** If `pip` is not found, use `python -m pip install openpyxl`

---

## 📄 Phase 1 — `audit_template.py`

Reads one or more Nessus `.audit` files and produces a **blank Excel template**.  
Status and Observed Value are empty — ready for manual assessment or to be filled by Phase 2.

### Supported Audit Formats

- CIS Benchmarks — Windows, Linux, macOS, Microsoft 365, and more
- DISA STIGs — Oracle Database, Windows Server, RHEL, and more
- Any Nessus `.audit` file containing `<custom_item>` or `<report>` blocks

### Usage

```bash
# Single audit file
python audit_template.py --audit CIS_L1.audit

# L1 + L2 combined
python audit_template.py --audit L1.audit L2.audit

# Custom output filename
python audit_template.py --audit CIS_L1.audit -o ClientABC_Template.xlsx

# DISA STIG
python audit_template.py --audit DISA_Oracle_19c.audit -o oracle_stig_template.xlsx
```

### Output

```
audit_template.xlsx
├── Summary              ← Benchmark name + total check count
└── <Benchmark Name>     ← All checks in benchmark order
```

Every data row has:

| Column | Content |
|--------|---------|
| S.NO | Auto-numbered |
| Benchmark | Check ID + title from `.audit` description field |
| Description | Rationale / info from `.audit` info field |
| Status | **Empty** — fill manually or use report_generator.py |
| Default Value | Expected value from `.audit` value_data (blank if not in audit file) |
| Observed Value | **Empty** — fill manually or use report_generator.py |
| Recommendation | Remediation steps from `.audit` solution field |

### What "Default Value" means here

Default Value is extracted from the `.audit` file's `value_data` field and converted to human-readable form:

| Raw audit value | What appears in Excel |
|---|---|
| `@PASSWORD_HISTORY@` | `24 or more passwords` |
| `@MINIMUM_PASSWORD_LENGTH@` | `14 or more characters` |
| `[1..MAX]` | `1 or more` |
| `[15..999]` | `15 to 999` |
| *(not present in audit file)* | *(blank — never N/A)* |

If the audit file has no expected value for a check (common for script-based Linux/Oracle checks), the cell is left **blank**. The toolkit never writes `N/A` or placeholder text.

---

## 📊 Phase 2 — `report_generator.py`

Takes a Nessus `.audit` file plus one or more Nessus compliance **CSV exports** and produces a filled, colour-coded Excel report.

### How to Export CSV from Nessus

```
Nessus  →  My Scans  →  [Select your compliance scan]
        →  Export  →  CSV  →  Download
```

> **Important:** Export as a plain **CSV**, not as HTML or `.nessus`. The report generator handles both Nessus compliance CSV formats:
> - **Compliance scan CSV** — `Name` column contains the check name directly
> - **Vulnerability scan CSV** — `Name` = `"Unix Compliance Checks"` (generic); check name and observed value are embedded in the `Description` field

Both formats are detected and parsed automatically.

### Usage

```bash
# Single host scan
python report_generator.py --audit CIS_L1.audit --csv scan.csv

# Multiple CSV files — one tab per host
python report_generator.py --audit CIS_L1.audit --csv host1.csv host2.csv host3.csv

# Folder of CSVs — one tab per host, auto-detected
python report_generator.py --audit CIS_L1.audit --csv ./scan_results/

# Multiple audit files (L1 + L2)
python report_generator.py --audit L1.audit L2.audit --csv scan.csv

# Custom output filename
python report_generator.py --audit CIS_L1.audit --csv scan.csv -o Report_$(date +%Y%m%d).xlsx
```

### Output — Single CSV

```
report.xlsx
├── Summary              ← Total / Passed / Failed / Warning / Not Evaluated
└── 192.168.1.10         ← Results for the scanned host (tab named by IP)
```

### Output — Multiple CSVs / Folder

```
report.xlsx
├── Summary              ← Per-host breakdown table
├── 192.168.1.10         ← Host 1 results (🔴 red tab = failures exist)
├── 192.168.1.11         ← Host 2 results (🟢 green tab = all passed)
└── 192.168.1.12         ← Host 3 results (🟠 orange tab = warnings only)
```

No Consolidated sheet — each host has its own tab. The Summary tab gives the cross-host overview.

### Row Colour Coding

| Row Colour | Status | Meaning |
|-----------|--------|---------|
| 🟢 Light green | `PASSED` | Control is compliant |
| 🔴 Light red | `FAILED` | Control needs remediation |
| 🟡 Light amber | `WARNING` | Informational / partially compliant |
| ⬜ White / grey | *(blank)* | Check not in scan results |

### Tab Colour Coding

| Tab Colour | Meaning |
|-----------|---------|
| 🔴 Red | One or more FAILED checks in this host's results |
| 🟠 Orange | Warnings present, no failures |
| 🟢 Green | All evaluated controls passed |

### Multi-Host Status Logic

When the same check is reported by multiple hosts (e.g., from a folder of CSVs for the same host scanned twice), the script applies **worst-case** per-tab:

```
Scan A  →  1.1.1  →  PASSED
Scan B  →  1.1.1  →  FAILED
─────────────────────────────
Tab result  →  FAILED   ← any failure wins
```

---

## 📐 Column Reference

Both scripts produce the same 7-column layout:

| # | Column | Template | Report | Source |
|---|--------|----------|--------|--------|
| 1 | S.NO | Auto | Auto | — |
| 2 | Benchmark | ✅ Filled | ✅ Filled | `.audit` description |
| 3 | Description | ✅ Filled | ✅ Filled | `.audit` info field |
| 4 | **Status** | **Empty** | **✅ Filled** | CSV `Risk` column |
| 5 | Default Value | ✅ or blank | ✅ or blank | `.audit` value_data only |
| 6 | **Observed Value** | **Empty** | **✅ Filled** | CSV `Actual Value` section |
| 7 | Recommendation | ✅ Filled | ✅ Filled | `.audit` solution field |

**Key design principles:**

- Default Value **always** comes from the `.audit` file — never from the CSV
- Status and Observed Value **always** come from the CSV — never from the `.audit` file
- When data is not available, the cell is **left blank** — `N/A` is never written

---

## 🌍 Real-Life Use Cases

### 1. Manual Security Assessment — Client Engagement

A consultant needs a professional assessment template before their first client meeting. No scan has been run yet.

```bash
python audit_template.py \
  --audit CIS_AlmaLinux_OS_9_v2_0_0_L1_Server.audit \
  -o "ClientABC_CIS_AlmaLinux_Assessment.xlsx"
```

**Result:** 226-check Excel template with all benchmark text pre-filled. Consultant distributes it to the client team. Each engineer checks their assigned servers and fills in Status and Observed Value. No copy-paste from any PDF.

---

### 2. Automated Weekly Compliance Report — Single Server

A sysadmin runs a Nessus compliance scan every Monday and needs a coloured report ready before the 9am standup.

```bash
python report_generator.py \
  --audit CIS_AlmaLinux_OS_9_v2_0_0_L1_Server.audit \
  --csv   weekly_scan_$(date +%Y%m%d).csv \
  -o "Weekly_Report_$(date +%Y%m%d).xlsx"
```

**Result:** Colour-coded report in under 30 seconds. 177 green rows, 40 red rows. Paste it into the Monday email and done.

---

### 3. Multi-Server Scan — 20 Linux Hosts

An enterprise team scans a fleet of servers and needs individual host results without a consolidated view that hides per-host issues.

```bash
# All CSVs exported into one folder
python report_generator.py \
  --audit CIS_AlmaLinux_OS_9_v2_0_0_L1_Server.audit \
  --csv   ./scan_exports/ \
  -o "Fleet_Compliance_$(date +%Y%m%d).xlsx"
```

**Result:** One Excel workbook with one tab per server (named by IP). Summary tab shows each server's pass/fail count at a glance. Red tabs immediately identify servers that need attention.

---

### 4. DISA STIG Oracle Database Audit

A DBA team needs to assess Oracle 19c compliance against DISA STIG requirements. No scan exists yet — this is a manual walkthrough.

```bash
python audit_template.py \
  --audit DISA_STIG_Oracle_Database_19c_v1r5_Unix.audit \
  -o "Oracle19c_STIG_Assessment.xlsx"
```

**Result:** All 16 STIG controls pre-populated with STIG IDs (e.g., `O19C-00-000200`), full rationale, and remediation steps. DBA team checks each control during the audit walkthrough.

---

### 5. Mixed L1 + L2 Benchmark Coverage

A compliance manager needs a full L1 and L2 CIS benchmark report for an external audit.

```bash
# Template covering both levels
python audit_template.py \
  --audit CIS_L1.audit CIS_L2.audit \
  -o full_benchmark_template.xlsx

# Report after scanning
python report_generator.py \
  --audit CIS_L1.audit CIS_L2.audit \
  --csv   full_scan.csv \
  -o full_benchmark_report.xlsx
```

**Result:** Checks from both L1 and L2 are deduplicated, merged, and sorted by benchmark number into a single sheet.

---

### 6. CI/CD Pipeline — Compliance Gate

A DevSecOps team gates image builds on CIS compliance. If too many controls fail, the build is blocked.

```bash
#!/bin/bash
python report_generator.py \
  --audit CIS_L1.audit \
  --csv   nessus_scan_${BUILD_ID}.csv \
  -o artifacts/compliance_${BUILD_ID}.xlsx

# Count failures (extend the script or use python -c for threshold check)
echo "Compliance report saved as build artifact"
```

---

## ❓ Troubleshooting

### `ModuleNotFoundError: No module named 'openpyxl'`

```bash
pip install openpyxl
```

### `python` is not recognised on Windows

```bash
python3 audit_template.py --audit file.audit
# or
py audit_template.py --audit file.audit
```

### `0 checks found` from `.audit` file

- Confirm it is a **Nessus `.audit`** benchmark definition file — not a `.nessus` scan result
- The file must contain `<custom_item>` or `<report>` blocks
- Check descriptions must start with a benchmark ID:
  - CIS: `1.1.1 (L1) Ensure...`
  - DISA STIG: `O19C-00-000200 - Oracle Database must...`

### CSV matched `0 / N` checks

- The CSV must be exported from a **Compliance scan**, not a Vulnerability scan
- Open the CSV and check: either the `Name` column has check names starting with numbers, **or** the `Description` column contains lines like `"1.1.1 Ensure ..." : [PASSED]`
- Both formats are auto-detected — if neither is present, it's not a compliance export

### Row text still clipped in Excel

- Open in **Microsoft Excel** (not LibreOffice or Google Sheets — both may ignore pre-calculated row heights)
- Manual fix: **Ctrl+A → Home → Format → AutoFit Row Height**

### Default Value shows blank for all checks

This is expected for script-based audit formats such as CIS Linux and DISA STIG Oracle. These checks use multi-line scripts rather than a simple expected value, so no `value_data` is stored in the `.audit` file. Blank is correct — not a bug.

---

## 📁 Repository Structure

```
nessus-cis-excel/
├── audit_template.py      ← Phase 1: blank Excel template from .audit file
├── report_generator.py    ← Phase 2: filled Excel report from .audit + CSV
├── requirements.txt       ← Python dependencies
└── README.md              ← This file
```

---

## 📦 Requirements

| Dependency | Version | Purpose |
|------------|---------|---------|
| Python | 3.8+ | Runtime |
| `openpyxl` | 3.1.0+ | Excel file creation and formatting |

```bash
pip install openpyxl
```

No other dependencies. `beautifulsoup4` is **not required** — use CSV exports, not HTML.

---

## 💡 Quick Reference

**Get the `.audit` file:**
- [Tenable Downloads](https://www.tenable.com/downloads/audit-files) — Tenable account required
- [CIS WorkBench](https://workbench.cisecurity.org) — CIS membership required

**Export CSV from Nessus:**
```
Nessus → My Scans → [Compliance scan] → Export → CSV → Download
```

**L1 + L2 together:**
```bash
# Both scripts accept multiple --audit files
python audit_template.py  --audit L1.audit L2.audit
python report_generator.py --audit L1.audit L2.audit --csv scan.csv
```

**Folder of CSVs:**
```bash
python report_generator.py --audit CIS_L1.audit --csv ./results/
# All *.csv files in the folder are picked up automatically
# One tab per file, tab named by host IP
```

---

## 📄 License

MIT — free to use, modify, and distribute in personal and commercial projects.

---
2. Test `report_generator.py` with a real Nessus compliance CSV — verify matched count
3. Open the Excel in Microsoft Excel and confirm row heights display all content
4. Confirm Default Value cells are blank (not `N/A`) when no value_data exists in the audit file
