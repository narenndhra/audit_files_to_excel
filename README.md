# 🛡️ Nessus CIS Audit → Excel Toolkit

> **Stop copy-pasting CIS benchmarks from PDFs.**  
> Convert Nessus `.audit` files and HTML compliance reports into beautiful, structured Excel workbooks — in seconds.

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=flat-square&logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
![Excel](https://img.shields.io/badge/Output-Excel%20.xlsx-217346?style=flat-square&logo=microsoft-excel&logoColor=white)
![CIS](https://img.shields.io/badge/Standard-CIS%20Benchmarks-red?style=flat-square)
![Nessus](https://img.shields.io/badge/Tool-Nessus%20%2F%20Tenable-00B388?style=flat-square)

---

## 📋 Table of Contents

- [What This Toolkit Solves](#-what-this-toolkit-solves)
- [Tools Overview](#-tools-overview)
- [Real-World Workflow](#-real-world-workflow)
- [Installation](#-installation)
- [Tool 1 — Audit to Excel](#-tool-1--nessus-audit-file--excel)
- [Tool 2 — HTML Report to Excel](#-tool-2--nessus-html-report--excel)
- [Excel Output Structure](#-excel-output-structure)
- [Real-Life Use Cases](#-real-life-use-cases)
- [Troubleshooting](#-troubleshooting)
- [Requirements](#-requirements)

---

## 🔥 What This Toolkit Solves

Every security team doing **CIS hardening** knows the pain:

| The Old Way 😩 | With This Toolkit ✅ |
|---|---|
| Open CIS benchmark PDF | Run one command |
| Copy benchmark title manually | Extracted automatically from `.audit` file |
| Copy 500-word description | Pulled directly from Nessus audit definitions |
| Type Pass/Fail from scan output | Auto-populated from Nessus CSV export |
| Format Excel manually for hours | Professional Excel generated in seconds |
| One section at a time | All 339+ checks at once across 24 category tabs |
| Observations lost in emails | Dedicated `Observation` column ready to fill |

**Time saved per engagement: 4–8 hours of manual work → under 30 seconds.**

---

## 🧰 Tools Overview

| Script | Input | Purpose |
|--------|-------|---------|
| `nessus_audit_to_excel.py` | `.audit` file + optional Nessus CSV | **Manual** baseline template OR **Automatic** merged report with Pass/Fail |
| `nessus_html_to_excel.py` | Nessus HTML compliance report(s) | Colour-coded Pass/Fail Excel from HTML scan exports |

---

## 🔄 Real-World Workflow

### Workflow A — Automated Compliance Reporting

```
┌─────────────────────────────────────────────────────────────────────┐
│                    AUTOMATED WORKFLOW                               │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  1. Download CIS .audit file from Tenable / CIS WorkBench          │
│           │                                                         │
│           ▼                                                         │
│  2. Run Nessus compliance policy scan against target host(s)        │
│           │                                                         │
│           ▼                                                         │
│  3. Export scan results as CSV from Nessus                          │
│           │                                                         │
│           ▼                                                         │
│  4. python nessus_audit_to_excel.py                                 │
│       --audit CIS_WS2025_L1.audit                                   │
│       --csv   nessus_scan_results.csv                               │
│       --output CIS_Report.xlsx                                      │
│           │                                                         │
│           ▼                                                         │
│  5. Open Excel — all 339 checks, colour-coded, observed values      │
│     auto-filled, ready to review and share                          │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

### Workflow B — Manual Assessment Template

```
┌─────────────────────────────────────────────────────────────────────┐
│                     MANUAL WORKFLOW                                 │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  1. Download CIS .audit file                                        │
│           │                                                         │
│           ▼                                                         │
│  2. python nessus_audit_to_excel.py --audit CIS_WS2025_L1.audit    │
│           │                                                         │
│           ▼                                                         │
│  3. Share Excel template with assessment team                       │
│           │                                                         │
│           ▼                                                         │
│  4. Team fills in:                                                  │
│     • Status         (Pass / Fail)                                  │
│     • Observed Value (what they actually found on the system)       │
│     • Observation    (notes, context, exceptions, evidence)         │
│           │                                                         │
│           ▼                                                         │
│  5. Complete hardening report ready for delivery                    │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

### Workflow C — HTML Report Beautification

```
┌─────────────────────────────────────────────────────────────────────┐
│                   HTML REPORT WORKFLOW                              │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  1. Run Nessus compliance scan                                      │
│           │                                                         │
│           ▼                                                         │
│  2. Export as HTML (Compliance) from Nessus                         │
│           │                                                         │
│           ▼                                                         │
│  3. python nessus_html_to_excel.py scan_report.html                 │
│           │                                                         │
│           ▼                                                         │
│  4. Excel with green/red/amber rows, summary tab,                   │
│     remediation steps and affected hosts — client-ready             │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

---

## ⚙️ Installation

### Step 1 — Prerequisites

- **Python 3.8 or higher** → [Download from python.org](https://www.python.org/downloads/)

Verify your install:

```bash
python --version     # Windows
python3 --version    # macOS / Linux
```

---

### Step 2 — Get the Scripts

```bash
# Option A: Clone the repository
git clone https://github.com/your-org/nessus-cis-excel.git
cd nessus-cis-excel

# Option B: Download individual scripts with curl
curl -O https://raw.githubusercontent.com/your-org/nessus-cis-excel/main/nessus_audit_to_excel.py
curl -O https://raw.githubusercontent.com/your-org/nessus-cis-excel/main/nessus_html_to_excel.py
```

---

### Step 3 — Install Dependencies

**Simple global install:**

```bash
pip install openpyxl beautifulsoup4
```

**Recommended — virtual environment (keeps your system Python clean):**

```bash
# Create the environment
python -m venv venv

# Activate it
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows PowerShell

# Install dependencies
pip install openpyxl beautifulsoup4

# When finished, deactivate
deactivate
```

**Install from requirements file:**

```bash
pip install -r requirements.txt
```

`requirements.txt`:

```
openpyxl>=3.1.0
beautifulsoup4>=4.12.0
```

> **Windows tip:** If `pip` is not recognised, use `python -m pip install openpyxl beautifulsoup4`

---

## 🔧 Tool 1 — Nessus Audit File → Excel

**Script:** `nessus_audit_to_excel.py`

Parses Nessus `.audit` files — the CIS benchmark definition files published by Tenable — and produces a structured, multi-tab Excel workbook. Operates in two modes depending on whether you supply a Nessus CSV export.

---

### Mode A — Manual (no CSV)

Generates a clean assessment template with all benchmark text pre-filled. The **Status** and **Observed Value** columns are left blank for your team to complete during the manual check.

```bash
# Single L1 audit file
python nessus_audit_to_excel.py --audit CIS_Windows_Server_2025_L1.audit

# L1 and L2 combined into one workbook
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit CIS_WS2025_L2.audit

# Custom output filename
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit \
  --output "Client_ABC_Assessment_2025.xlsx"
```

---

### Mode B — Automatic (audit + Nessus CSV)

After running a Nessus compliance policy scan, export the results as CSV and provide it alongside the audit file. The script merges both sources — **Status** and **Observed Value** are populated automatically, and rows are colour-coded green/red.

```bash
# Single host
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit \
  --csv   nessus_export.csv \
  --output CIS_Report_AutoFilled.xlsx

# Multiple hosts — worst-case status per check, all observed values shown
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit \
  --csv   dc01.csv dc02.csv web01.csv \
  --output CIS_AllHosts_Report.xlsx

# L1 + L2 with scan results
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit CIS_WS2025_L2.audit \
  --csv   scan_results.csv \
  --output CIS_Full_L1_L2.xlsx
```

---

### How to Export CSV from Nessus

```
Nessus UI  →  My Scans  →  [Select your compliance scan]
           →  Export  →  CSV  →  Download
```

The script automatically detects column name variants across Nessus versions:

| Data Point | Recognised Column Names |
|---|---|
| Check name | `Name`, `Plugin Name`, `Check Name` |
| Pass / Fail | `Risk`, `Status`, `Result` |
| Actual value | `Plugin Output`, `Output`, `Actual Value` |
| Host IP | `Host`, `IP Address` |

---

### Multi-Host Aggregation

When multiple hosts scan the same control, the script applies **worst-case aggregation**:

```
Host A  →  1.1.1  →  PASSED   (Actual: 24)
Host B  →  1.1.1  →  FAILED   (Actual: 0)
Host C  →  1.1.1  →  PASSED   (Actual: 24)
─────────────────────────────────────────
Result  →  FAILED             ← any failure = overall FAILED
Observed Value  →  "24 / 0"  ← all unique values listed
```

This ensures no failing host is hidden by averaging — if one host is non-compliant, the control shows as FAILED so it gets remediated on all hosts.

---

### Column Reference

| Column | Manual Mode | Automatic Mode |
|--------|-------------|----------------|
| S.NO | Auto-numbered | Auto-numbered |
| Benchmark | From `.audit` description field | From `.audit` description field |
| Description | From `.audit` info / rationale | From `.audit` info / rationale |
| **Status** | **Blank** — fill manually | **PASSED / FAILED** — from CSV |
| Default Value | From `.audit` value_data | From `.audit` value_data |
| **Observed Value** | **Blank** — fill manually | **Actual system value** — from CSV |
| Observation | Blank — team notes / evidence | Blank — team notes / evidence |
| Recommendation | From `.audit` solution field | From `.audit` solution field |

---

### Supported CIS Sections → Excel Tabs

Every CIS section is automatically routed to a named tab:

| CIS Section | Tab Name | Typical Check Count |
|-------------|----------|---------------------|
| 1.1.x | `Password Policy` | 7 |
| 1.2.x | `Account Lockout` | 4 |
| 2.2.x | `User Rights` | 37 |
| 2.3.x | `Security Options` | 64 |
| 9.1.x | `Firewall - Domain` | 7 |
| 9.2.x | `Firewall - Private` | 7 |
| 9.3.x | `Firewall - Public` | 9 |
| 17.1–17.9.x | `Audit - *` (8 tabs) | 1–6 each |
| 18.1–18.10.x | `AT - *` (7 tabs) | 3–82 each |
| 19.5 / 19.7.x | `User AT - IE / Win Comp` | 1–7 |

---

## 🔧 Tool 2 — Nessus HTML Report → Excel

**Script:** `nessus_html_to_excel.py`

Parses Nessus HTML compliance scan reports and produces a colour-coded Excel workbook with Pass/Fail rows, remediation steps, and observed values — without any manual copy-paste.

### How to Export HTML from Nessus

```
Nessus UI  →  My Scans  →  [Select scan]
           →  Export  →  HTML (Compliance)  →  Download
```

### Usage

```bash
# Single report
python nessus_html_to_excel.py compliance_report.html

# Multiple reports → one Excel with one tab per report + a summary tab
python nessus_html_to_excel.py scan1.html scan2.html scan3.html

# Wildcard — all HTML files in the current directory
python nessus_html_to_excel.py *.html

# Custom output name
python nessus_html_to_excel.py report.html final_report.xlsx
```

### Column Reference

| Column | Source |
|--------|--------|
| S.NO | Auto-numbered |
| Benchmark | Check title extracted from HTML |
| Description | `Info` section from the HTML body |
| Passed / Failed | Detected from HTML element background colour |
| Remediation | `Solution` section (Impact text stripped automatically) |
| Observed Value | Actual value extracted from `Hosts` section output |
| Affected Host | Host IP from `Hosts` section |

---

## 📊 Excel Output Structure

### `nessus_audit_to_excel.py` workbook layout

```
CIS_Report.xlsx
│
├── 📋 Summary                    ← Total / Passed / Failed / Not Evaluated counts
│                                    Per-category breakdown with pass/fail per section
│
├── 🔐 Password Policy            ←  1.1.x   (7 checks)
├── 🔒 Account Lockout            ←  1.2.x   (4 checks)
├── 👤 User Rights                ←  2.2.x  (37 checks)
├── ⚙️  Security Options           ←  2.3.x  (64 checks)
├── 🔥 Firewall - Domain          ←  9.1.x   (7 checks)
├── 🔥 Firewall - Private         ←  9.2.x   (7 checks)
├── 🔥 Firewall - Public          ←  9.3.x   (9 checks)
├── 📝 Audit - Acct Logon         ← 17.1.x
├── 📝 Audit - Acct Mgmt          ← 17.2.x
├── 📝 Audit - Logon-Logoff       ← 17.5.x
├── 📝 Audit - Object Access      ← 17.6.x
├── 📝 Audit - Policy Change      ← 17.7.x
├── 📝 Audit - System             ← 17.9.x
├── 🖥️  AT - Control Panel         ← 18.1.x
├── 🌐 AT - Network               ← 18.5.x
├── 🖨️  AT - Printers              ← 18.6.x
├── ⚙️  AT - System                ← 18.9.x
└── 🪟 AT - Windows Comp          ← 18.10.x (82 checks)
```

### Row colour coding (Automatic mode only)

| Row Colour | Status | Meaning |
|-----------|--------|---------|
| 🟢 Light green | `PASSED` | Control is compliant on all scanned hosts |
| 🔴 Light red | `FAILED` | Control needs remediation on one or more hosts |
| 🟡 Light amber | `WARNING` | Partially compliant or informational finding |
| ⬜ White / grey | *(blank)* | Not evaluated — manual check required |

### Tab colour coding (Automatic mode only)

| Tab Colour | Meaning |
|-----------|---------|
| 🔴 Red tab | One or more FAILED checks exist in this category |
| 🟠 Orange tab | Warnings present, no failures |
| 🟢 Green tab | All evaluated controls in this category passed |

### Summary tab (Automatic mode)

The Summary tab shows an at-a-glance dashboard:

```
┌──────────────────────────────────────────────┐
│  CIS Windows Server 2025 – Hardening Report  │
│  Benchmark v1.0.0  •  L1 MS  •  Auto Mode   │
├────────────────────┬─────────────────────────┤
│  Total Checks      │  339                    │
│  ✔ Passed          │   52  (green)           │
│  ✘ Failed          │   87  (red)             │
│  ⚠ Warning         │    3  (amber)           │
│  — Not Evaluated   │  197                    │
├────────────────────┼──────────┬──────┬───────┤
│  Category          │  Checks  │  ✔   │  ✘   │
├────────────────────┼──────────┼──────┼───────┤
│  1.1 Password Pol. │    7     │  4   │  3    │
│  1.2 Acct Lockout  │    4     │  1   │  1    │
│  2.3 Security Opt. │   64     │ 18   │ 22    │
│  ...               │  ...     │ ...  │ ...   │
└────────────────────┴──────────┴──────┴───────┘
```

---

## 🌍 Real-Life Use Cases

### 1. Pre-Audit Baseline Assessment

A security consultant is engaged to harden 15 Windows Server 2025 hosts before an ISO 27001 audit.

```bash
python nessus_audit_to_excel.py \
  --audit CIS_Microsoft_Windows_Server_2025_v1_0_0_L1_MS.audit \
  --output "Client_ABC_Baseline_Assessment.xlsx"
```

**Result:** 339-check Excel template distributed to the client team in 10 seconds. Each CIS section is a separate tab. Team fills Status and Observations manually. No PDF copy-paste. No formatting work. The consultant looks professional before the first meeting.

---

### 2. Weekly Automated Compliance Reporting

A system administrator runs a weekly Nessus compliance scan and sends a report to management every Monday.

```bash
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit \
  --csv   weekly_scan_$(date +%Y%m%d).csv \
  --output "Weekly_CIS_Report_$(date +%Y%m%d).xlsx"
```

**Result:** Full colour-coded Excel — green = compliant, red = action needed. Management gets a clear picture. Security team gets exact observed values to remediate. Total time: under 60 seconds including the Nessus CSV export.

---

### 3. Datacenter Hardening — 20 Servers

An enterprise IT team is hardening a datacenter. They scan all servers and need a consolidated report showing which controls fail on any host.

```bash
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit CIS_WS2025_L2.audit \
  --csv   dc01.csv dc02.csv dc03.csv \
          web01.csv web02.csv app01.csv \
  --output "Datacenter_CIS_L1L2_Full_Report.xlsx"
```

**Result:** One Excel with 25 tabs. Every check shows worst-case status across all hosts. The Observed Value column shows all unique values (e.g. `24 / 0 / 0`). Red tab colours instantly highlight problem categories. Summary tab shows per-category breakdown. Entire report ready in under 2 minutes.

---

### 4. Pen Test Report Supplement — Client Deliverable

A pen-test team needs to include a CIS compliance gap analysis alongside their findings. The client expects a formatted Excel attachment, not raw HTML.

```bash
python nessus_html_to_excel.py \
  client_compliance_scan.html \
  "Client_XYZ_CIS_Gap_Analysis_$(date +%Y%m%d).xlsx"
```

**Result:** Professional Excel with colour-coded rows. Summary tab shows pass %, total / passed / failed. Remediation column pulled automatically from Nessus solution text. Ready to attach to the pentest report — zero manual formatting.

---

### 5. CI/CD Pipeline — Golden Image Compliance Gate

A DevSecOps team runs Nessus scans as part of their CI/CD pipeline for Windows golden images. Reports are generated automatically on every build.

```bash
#!/bin/bash
# pipeline_compliance_check.sh

AUDIT_FILE="CIS_WS2025_L1.audit"
SCAN_CSV="nessus_results_${BUILD_NUMBER}.csv"
REPORT="artifacts/CIS_Report_Build_${BUILD_NUMBER}.xlsx"

python nessus_audit_to_excel.py \
  --audit "$AUDIT_FILE" \
  --csv   "$SCAN_CSV" \
  --output "$REPORT"

echo "Compliance report saved as build artifact: $REPORT"
```

**Result:** Every build produces a compliance Excel as an artifact. Developers and security leads can review it directly. Extend the script to fail the build if FAILED count exceeds an acceptable threshold — making CIS compliance part of the release gate.

---

### 6. Compliance Progress Tracking — Month Over Month

A compliance team runs scans monthly and tracks hardening progress over time.

```bash
# January
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit --csv jan_scan.csv \
  --output Reports/CIS_2025_01.xlsx

# February — after first remediation sprint
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit --csv feb_scan.csv \
  --output Reports/CIS_2025_02.xlsx

# March — after second remediation sprint
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit --csv mar_scan.csv \
  --output Reports/CIS_2025_03.xlsx
```

**Result:** Three comparable Excel files showing measurable improvement:

```
January  → 87 FAILED  |  52 PASSED  |  200 Not Evaluated
February → 52 FAILED  |  87 PASSED  |  200 Not Evaluated
March    → 19 FAILED  | 120 PASSED  |  200 Not Evaluated
```

Present this trend to management or auditors as evidence of a maturing security posture.

---

## ❓ Troubleshooting

### `ModuleNotFoundError: No module named 'openpyxl'`

```bash
pip install openpyxl
```

### `ModuleNotFoundError: No module named 'bs4'`

```bash
pip install beautifulsoup4
```

### `python` is not recognised on Windows

```bash
python3 nessus_audit_to_excel.py --audit file.audit
# or
py nessus_audit_to_excel.py --audit file.audit
```

### Script runs but extracts 0 checks from `.audit` file

- Confirm the file is a **Nessus `.audit`** benchmark definition file, not an `.nessus` scan result
- The file must contain `<custom_item>` blocks
- Benchmark descriptions must start with a CIS number like `1.1.1 (L1) Ensure...`

### CSV matched 0 / very few checks

- Ensure the CSV was exported from a **Compliance** scan, not a Vulnerability scan
- Open the CSV in a text editor and confirm the `Name` column contains CIS benchmark names starting with numbers like `1.1.1 (L1) Ensure...`
- Verify accepted column names: `Name`, `Plugin Name`, `Risk`, `Status`, `Plugin Output`

### Row text still clipped in Excel

- Open the file in **Microsoft Excel** (not LibreOffice or Google Sheets, which may ignore pre-calculated row heights)
- If still clipped: **Select All → Home → Format → AutoFit Row Height**

### HTML script finds 0 checks

- Export from Nessus as **HTML (Compliance)** not the default summary HTML
- Open the file in a browser — it must contain colour-coded check title bars, not just a plain table

---

## 📁 Repository Structure

```
nessus-cis-excel/
├── nessus_audit_to_excel.py   ← Audit + CSV converter (Manual & Automatic modes)
├── nessus_html_to_excel.py    ← HTML compliance report converter
├── requirements.txt           ← Python dependencies
└── README.md                  ← This file
```

---

## 📦 Requirements

| Dependency | Minimum Version | Used By | Purpose |
|------------|----------------|---------|---------|
| Python | 3.8+ | Both scripts | Runtime |
| `openpyxl` | 3.1.0+ | Both scripts | Excel file creation and formatting |
| `beautifulsoup4` | 4.12.0+ | HTML script only | HTML parsing |

```bash
# Install everything at once
pip install openpyxl beautifulsoup4
```

---

## 💡 Tips & Reference

**Where to get the `.audit` file:**
- [Tenable Downloads](https://www.tenable.com/downloads/audit-files) — requires a Tenable account
- [CIS WorkBench](https://workbench.cisecurity.org) — requires CIS membership
- Filename example: `CIS_Microsoft_Windows_Server_2025_v1_0_0_L1_MS.audit`

**How to export Nessus CSV:**
```
Nessus → My Scans → [Your compliance scan] → Export → CSV → Download
```

**How to export Nessus HTML:**
```
Nessus → My Scans → [Your compliance scan] → Export → HTML (Compliance) → Download
```

**Running L1 and L2 together:**
```bash
python nessus_audit_to_excel.py \
  --audit CIS_WS2025_L1.audit CIS_WS2025_L2.audit \
  --csv   scan.csv
# Checks from both files are deduplicated and sorted by benchmark number
# L1 and L2 checks appear in the same category tabs in natural order
```

---

## 📄 License

MIT — free to use, modify, and distribute in personal and commercial projects.

---

## 🤝 Contributing

Pull requests are welcome. Before submitting please:

1. Test `nessus_audit_to_excel.py` against a real Nessus `.audit` file and verify check counts
2. Test Automatic mode with a real Nessus compliance CSV export
3. Test `nessus_html_to_excel.py` against a real Nessus HTML compliance export
4. Verify the Excel output opens correctly in Microsoft Excel
5. Confirm all row heights display content without clipping
