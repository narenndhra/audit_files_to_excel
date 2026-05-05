#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         Universal Nessus Audit → Excel  (Manual + Automatic modes)         ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  Works with ANY .audit file type:                                           ║
║    • CIS Benchmarks  (Windows, Linux, macOS, M365, etc.)                   ║
║    • DISA STIG        (Oracle, Windows, RHEL, etc.)                        ║
║    • Any other Nessus .audit format                                         ║
║                                                                              ║
║  MANUAL MODE  — audit only, blank Status + Observed Value columns           ║
║    python audit_to_excel.py --audit file.audit                              ║
║                                                                              ║
║  AUTOMATIC MODE  — audit + Nessus CSV, fully populated                      ║
║    python audit_to_excel.py --audit file.audit --csv scan.csv               ║
║                                                                              ║
║  Multiple audit files:                                                       ║
║    python audit_to_excel.py --audit L1.audit L2.audit --csv scan.csv        ║
╚══════════════════════════════════════════════════════════════════════════════╝

Columns:  S.NO | Benchmark | Description | Status | Default Value |
          Observed Value | Recommendation

Status and Observed Value are blank in Manual mode.
In Automatic mode they are populated from the Nessus CSV export.
"""

import re
import sys
import csv
import math
import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════════════
#  ROW HEIGHT CALCULATOR
#  openpyxl does NOT auto-size rows with wrap_text — we calculate explicitly.
# ══════════════════════════════════════════════════════════════════════════════

FONT_PT       = 9
LINE_H        = FONT_PT * 1.35
MIN_ROW_H     = 18
MAX_ROW_H     = 400
CHARS_PER_COL = 1.05          # Excel column-width unit → characters


def _lines(text: str, col_w: float) -> float:
    if not text:
        return 1.0
    cpl = max(1, col_w * CHARS_PER_COL)
    tot = 0.0
    for para in str(text).split("\n"):
        p = para.strip()
        tot += 0.4 if not p else math.ceil(len(p) / cpl)
    return max(1.0, tot)


def calc_row_height(pairs: list) -> float:
    """pairs = [(text, col_width), ...] for every wrapped cell."""
    mx = max((_lines(t, w) for t, w in pairs), default=1.0)
    return max(MIN_ROW_H, min(mx * LINE_H + 4, MAX_ROW_H))


# ══════════════════════════════════════════════════════════════════════════════
#  AUDIT FILE PARSER  — universal, handles all audit file formats
# ══════════════════════════════════════════════════════════════════════════════

# Matches both <custom_item>...</custom_item> AND <report ...>...</report>
_BLOCK_RE = re.compile(
    r'<(custom_item|report[^>]*)>(.*?)</(custom_item|report)>',
    re.DOTALL | re.IGNORECASE,
)

# Key : value  (quoted or unquoted)
_KV_RE = re.compile(
    r'^\s*([\w]+)\s*:\s*(?:"((?:[^"\\]|\\.)*?)"|([^\n\r]+))',
    re.MULTILINE,
)

# Benchmark ID patterns for different audit types
_BENCHMARK_PATTERNS = [
    re.compile(r'^\d+\.\d+'),                    # CIS:   1.1.1 ...
    re.compile(r'^[A-Z][A-Z0-9]{1,8}-\d{2}-\d{4,}'),  # DISA:  O19C-00-000200, WN19-XX-000001
    re.compile(r'^[A-Z]{2,10}-\d{4,}'),           # Other STIG formats
    re.compile(r'^V-\d{4,}'),                      # Vuln-ID style
]


def _is_real_check(description: str) -> bool:
    """Return True if description looks like a real benchmark/STIG check ID."""
    d = description.strip()
    return any(p.match(d) for p in _BENCHMARK_PATTERNS)


def _parse_block(block: str) -> dict:
    item = {}
    for m in _KV_RE.finditer(block):
        k = m.group(1).lower().strip()
        v = (m.group(2) if m.group(2) is not None else m.group(3) or '').strip()
        v = v.replace('\\"', '"').replace('\\n', '\n').replace('\\t', '\t')
        if k not in item:
            item[k] = v
    return item


def _clean(text: str) -> str:
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def _default_val(value_data: str) -> str:
    if not value_data:
        return ""
    _VARS = {
        "@PASSWORD_HISTORY@":        "24 or more passwords",
        "@MAXIMUM_PASSWORD_AGE@":    "1 to 365 days",
        "@MINIMUM_PASSWORD_AGE@":    "1 or more day(s)",
        "@MINIMUM_PASSWORD_LENGTH@": "14 or more characters",
        "@LOCKOUT_DURATION@":        "15 or more minutes",
        "@LOCKOUT_THRESHOLD@":       "1 to 5 invalid attempts",
        "@LOCKOUT_RESET@":           "15 or more minutes",
    }
    for ph, human in _VARS.items():
        if ph in value_data:
            return human
    val = re.sub(
        r'\[(\d+)\.\.(MAX|\d+)\]',
        lambda m: (f"{m.group(1)} or more" if m.group(2) == "MAX"
                   else f"{m.group(1)} to {m.group(2)}"),
        value_data,
    )
    return val.replace(' || ', ' or ').strip('"').strip()


def _detect_benchmark_name(text: str, filepath: str) -> str:
    """Extract human-readable benchmark name from audit file header."""
    # Try <display_name> tag
    m = re.search(r'<display_name>(.*?)</display_name>', text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # Try # description : line
    m = re.search(r'#\s*description\s*:\s*This .audit is designed against the (.+)', text)
    if m:
        return m.group(1).strip()
    # Fallback to filename
    return Path(filepath).stem.replace('_', ' ')


def _detect_level(filepath: str, text: str) -> str:
    fn = Path(filepath).name.upper()
    if "L2" in fn or "LEVEL_2" in fn or "LEVEL|2" in text:
        return "L2"
    return "L1"


def parse_audit_files(filepaths: list) -> tuple:
    """
    Parse one or more .audit files.
    Returns (checks_list, benchmark_name)
    checks_list items: {benchmark, description, default_value,
                        recommendation, status, observed_value, host}
    """
    all_checks = []
    seen       = set()
    bench_name = ""

    for fp in filepaths:
        text  = Path(fp).read_text(encoding="utf-8", errors="replace")
        level = _detect_level(fp, text)

        # Detect benchmark name from first file
        if not bench_name:
            bench_name = _detect_benchmark_name(text, fp)

        count = 0
        for m in _BLOCK_RE.finditer(text):
            block    = m.group(2)
            item     = _parse_block(block)
            desc_raw = item.get("description", "").strip()

            # Skip: empty, non-benchmark, duplicates
            if not desc_raw or not _is_real_check(desc_raw):
                continue
            if desc_raw in seen:
                continue
            seen.add(desc_raw)

            all_checks.append({
                "benchmark":      desc_raw,
                "description":    _clean(item.get("info", "")),
                "default_value":  _default_val(item.get("value_data", "")),
                "recommendation": _clean(item.get("solution", "")),
                "status":         "",
                "observed_value": "",
                "host":           "",
            })
            count += 1

        print(f"    {Path(fp).name}: {count} checks (level {level})")

    # Natural sort by leading number/ID
    def _sort_key(c):
        d = c["benchmark"]
        # CIS numeric: 1.2.3
        m = re.match(r'^([\d.]+)', d)
        if m:
            return [int(p) if p.isdigit() else p
                    for p in m.group(1).split('.')]
        # DISA STIG: O19C-00-000200
        m2 = re.match(r'^([A-Z0-9-]+)', d)
        return [m2.group(1)] if m2 else [d]

    all_checks.sort(key=_sort_key)
    return all_checks, bench_name


# ══════════════════════════════════════════════════════════════════════════════
#  NESSUS CSV PARSER  &  MERGER
# ══════════════════════════════════════════════════════════════════════════════

_NAME_COLS   = ("Name", "Plugin Name", "Check Name", "name")
_STATUS_COLS = ("Risk", "Status", "Result", "risk", "status")
_OUTPUT_COLS = ("Plugin Output", "Output", "Actual Value",
                "plugin_output", "output")
_HOST_COLS   = ("Host", "IP Address", "host", "ip_address")


def _find_col(headers: list, candidates: tuple) -> str:
    for c in candidates:
        if c in headers:
            return c
    return ""


def _extract_actual(output: str) -> str:
    if not output:
        return ""
    m = re.search(r'Actual\s+Value\s*:\s*(.+?)(?:\n|Policy|$)',
                  output, re.IGNORECASE | re.DOTALL)
    if m:
        return m.group(1).strip()
    for line in output.strip().splitlines():
        line = line.strip()
        if line and not re.match(r'(policy|expected|check)', line, re.I):
            return line[:200]
    return output.strip()[:200]


def _norm_status(raw: str) -> str:
    r = raw.strip().upper()
    if r in ("PASSED", "PASS"):   return "PASSED"
    if r in ("FAILED", "FAIL"):   return "FAILED"
    if r in ("WARNING", "WARN"):  return "WARNING"
    return r or "INFO"


def _norm_id(s: str) -> str:
    """Extract leading benchmark ID from a description for matching."""
    # CIS:  "1.1.1 (L1) Ensure..."  → "1.1.1"
    m = re.match(r'^([\d]+(?:\.[\d]+)*)', s.strip())
    if m:
        return m.group(1)
    # DISA: "O19C-00-000200 - ..."  → "O19C-00-000200"
    m2 = re.match(r'^([A-Z]{1,10}-\d{2}-\d{6})', s.strip())
    if m2:
        return m2.group(1)
    return re.sub(r'[^a-z0-9]', '', s.lower())[:30]


def parse_csv_files(filepaths: list) -> dict:
    """Returns { norm_id → list of {status, observed_value, host} }"""
    results = {}
    for fp in filepaths:
        print(f"    {Path(fp).name}: ", end="", flush=True)
        count = 0
        try:
            with open(fp, newline="", encoding="utf-8-sig", errors="replace") as f:
                sample  = f.read(4096); f.seek(0)
                dialect = csv.Sniffer().sniff(sample, delimiters=",\t")
                reader  = csv.DictReader(f, dialect=dialect)

                hdrs       = reader.fieldnames or []
                name_col   = _find_col(hdrs, _NAME_COLS)
                status_col = _find_col(hdrs, _STATUS_COLS)
                output_col = _find_col(hdrs, _OUTPUT_COLS)
                host_col   = _find_col(hdrs, _HOST_COLS)

                if not name_col or not status_col:
                    print(f"⚠  Cannot find Name/Status columns (got: {hdrs[:6]})")
                    continue

                for row in reader:
                    name   = row.get(name_col, "").strip()
                    status = _norm_status(row.get(status_col, ""))
                    output = _extract_actual(row.get(output_col, ""))
                    host   = row.get(host_col, "").strip()
                    if not name:
                        continue
                    key = _norm_id(name)
                    results.setdefault(key, []).append({
                        "status":         status,
                        "observed_value": output,
                        "host":           host,
                    })
                    count += 1

        except Exception as e:
            print(f"⚠  Error: {e}")
            continue

        print(f"{count} rows")

    return results


def merge_csv(checks: list, csv_results: dict) -> list:
    """Merge CSV results into checks by benchmark ID matching."""
    matched = 0
    for chk in checks:
        key  = _norm_id(chk["benchmark"])
        rows = csv_results.get(key, [])

        if not rows:
            # Fuzzy: try prefix/suffix overlap
            for k, v in csv_results.items():
                if k.startswith(key) or key.startswith(k):
                    rows = v
                    break

        if not rows:
            continue

        statuses = [r["status"] for r in rows]
        obs_vals = list(dict.fromkeys(
            r["observed_value"] for r in rows if r["observed_value"]))
        hosts    = list(dict.fromkeys(r["host"] for r in rows if r["host"]))

        # Worst-case status
        if "FAILED"  in statuses: agg = "FAILED"
        elif "WARNING" in statuses: agg = "WARNING"
        elif all(s == "PASSED" for s in statuses): agg = "PASSED"
        else: agg = statuses[0]

        chk["status"]         = agg
        chk["observed_value"] = "\n".join(obs_vals)
        chk["host"]           = ", ".join(hosts)
        matched += 1

    print(f"    Matched {matched} / {len(checks)} checks to CSV results")
    return checks


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL STYLES
# ══════════════════════════════════════════════════════════════════════════════

C_HDR_BG  = "1F3864"
C_HDR_FG  = "FFFFFF"
C_PASS_BG = "EBF5E1";  C_PASS_FG = "375623"
C_FAIL_BG = "FDE9E8";  C_FAIL_FG = "9C0006"
C_WARN_BG = "FFF3CD";  C_WARN_FG = "7F5200"
C_ALT     = "EBF3FB"
C_WHITE   = "FFFFFF"
C_BORDER  = "8EA9C1"

_S   = Side(style="thin", color=C_BORDER)
_BRD = Border(left=_S, right=_S, top=_S, bottom=_S)

# (header label, col_width, wrap, h-align)
COLUMNS = [
    ("S.NO",           7,   False, "center"),
    ("Benchmark",      50,  True,  "left"),
    ("Description",    70,  True,  "left"),
    ("Status",         14,  False, "center"),
    ("Default Value",  26,  True,  "left"),
    ("Observed Value", 32,  True,  "left"),
    ("Recommendation", 62,  True,  "left"),
]


def _status_style(status: str):
    s = status.upper()
    if s == "PASSED":  return C_PASS_BG, C_PASS_FG
    if s == "FAILED":  return C_FAIL_BG, C_FAIL_FG
    if s == "WARNING": return C_WARN_BG, C_WARN_FG
    return C_WHITE, "595959"


def _hdr(cell, text: str):
    cell.value     = text
    cell.font      = Font(name="Arial", bold=True, color=C_HDR_FG, size=10)
    cell.fill      = PatternFill("solid", fgColor=C_HDR_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                               wrap_text=True)
    cell.border    = _BRD


def _dat(cell, value, bg: str, fg: str, align: str, wrap: bool, bold=False):
    cell.value     = value
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.font      = Font(name="Arial", size=FONT_PT, color=fg, bold=bold)
    cell.alignment = Alignment(horizontal=align, vertical="top", wrap_text=wrap)
    cell.border    = _BRD


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET WRITERS
# ══════════════════════════════════════════════════════════════════════════════

def write_checks_sheet(ws, checks: list, auto_mode: bool):
    # Column widths + header
    ws.row_dimensions[1].height = 30
    for ci, (label, width, wrap, align) in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = width
        _hdr(ws.cell(row=1, column=ci), label)

    ws.freeze_panes  = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    for ri, chk in enumerate(checks, start=1):
        row    = ri + 1
        status = chk["status"]

        if auto_mode and status:
            row_bg, status_fg = _status_style(status)
        else:
            row_bg    = C_ALT if ri % 2 == 0 else C_WHITE
            status_fg = "595959"

        vals = [
            (ri,                       "center", False, False),
            (chk["benchmark"],         "left",   True,  False),
            (chk["description"],       "left",   True,  False),
            (status,                   "center", False, True),
            (chk["default_value"],     "left",   True,  False),
            (chk["observed_value"],    "left",   True,  False),
            (chk["recommendation"],    "left",   True,  False),
        ]

        for ci, (val, align, wrap, bold) in enumerate(vals, start=1):
            label = COLUMNS[ci - 1][0]
            fg    = status_fg if label == "Status" else "000000"
            _dat(ws.cell(row=row, column=ci),
                 val if ci > 1 else ri,
                 row_bg, fg, align, wrap, bold)

        # Calculated row height
        wrapped_pairs = [
            (str(val), COLUMNS[ci - 1][1])
            for ci, (val, _, wrap, _) in enumerate(vals, start=1)
            if wrap
        ]
        ws.row_dimensions[row].height = calc_row_height(wrapped_pairs)


def write_summary_sheet(wb: Workbook, checks: list,
                        bench_name: str, auto_mode: bool):
    ws = wb.create_sheet(title="Summary", index=0)
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 14

    # Title
    ws.merge_cells("A1:B1")
    t           = ws.cell(row=1, column=1, value=bench_name)
    t.font      = Font(name="Arial", bold=True, size=13, color=C_HDR_FG)
    t.fill      = PatternFill("solid", fgColor=C_HDR_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    mode_str = "Automatic Mode  (audit + CSV merged)" if auto_mode else "Manual Mode  (audit only — fill Status & Observed Value)"
    ws.merge_cells("A2:B2")
    sub = ws.cell(row=2, column=1, value=mode_str)
    sub.font      = Font(name="Arial", size=9, italic=True, color="595959")
    sub.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 16

    rows = [("Total Checks", len(checks), C_HDR_BG, C_HDR_FG)]
    if auto_mode:
        p = sum(1 for c in checks if c["status"] == "PASSED")
        f = sum(1 for c in checks if c["status"] == "FAILED")
        w = sum(1 for c in checks if c["status"] == "WARNING")
        n = len(checks) - p - f - w
        rows += [
            ("✔  Passed",        p, "375623",  C_PASS_BG),
            ("✘  Failed",        f, "9C0006",  C_FAIL_BG),
            ("⚠  Warning",       w, "7F5200",  C_WARN_BG),
            ("—  Not Evaluated", n, "595959",  "F2F2F2"),
        ]

    for i, (label, val, fg, bg) in enumerate(rows, start=4):
        for ci, (v, al) in enumerate([(label, "left"), (val, "center")], start=1):
            c = ws.cell(row=i, column=ci, value=v)
            c.font      = Font(name="Arial", bold=True, size=10, color=fg)
            c.fill      = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal=al, vertical="center",
                                    indent=1 if al == "left" else 0)
            c.border    = _BRD
        ws.row_dimensions[i].height = 22


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN BUILDER
# ══════════════════════════════════════════════════════════════════════════════

def build_excel(checks: list, bench_name: str,
                output_path: str, auto_mode: bool):
    wb = Workbook()
    wb.remove(wb.active)

    # Single sheet named after the benchmark (safe for Excel)
    safe_name = re.sub(r'[:/\\?\[\]*]', '-', bench_name)[:31] or "Checks"
    ws = wb.create_sheet(title=safe_name)
    write_checks_sheet(ws, checks, auto_mode)

    if auto_mode:
        p = sum(1 for c in checks if c["status"] == "PASSED")
        f = sum(1 for c in checks if c["status"] == "FAILED")
        ws.sheet_properties.tabColor = (
            "C2212E" if f > 0 else
            "F18C43" if any(c["status"] == "WARNING" for c in checks) else
            "527421"
        )

    write_summary_sheet(wb, checks, bench_name, auto_mode)
    wb.save(output_path)

    print(f"\n[✓] Saved  →  {output_path}")
    print(f"    Benchmark : {bench_name}")
    print(f"    Mode      : {'Automatic (audit + CSV merged)' if auto_mode else 'Manual (audit only)'}")
    print(f"    Checks    : {len(checks)}")
    if auto_mode:
        p = sum(1 for c in checks if c["status"] == "PASSED")
        f = sum(1 for c in checks if c["status"] == "FAILED")
        n = len(checks) - p - f
        print(f"    Results   : {p} PASSED  |  {f} FAILED  |  {n} Not Evaluated")


# ══════════════════════════════════════════════════════════════════════════════
#  CLI
# ══════════════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        description="Convert ANY Nessus .audit file to formatted Excel.\n"
                    "Supports CIS Benchmarks, DISA STIGs, and all other formats.\n"
                    "Add --csv to auto-populate Status from a Nessus scan CSV.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Manual mode (blank Status/Observed Value)
  python audit_to_excel.py --audit CIS_Windows_L1.audit
  python audit_to_excel.py --audit DISA_Oracle_19c.audit

  # Automatic mode (merge with Nessus CSV export)
  python audit_to_excel.py --audit CIS_L1.audit --csv scan.csv
  python audit_to_excel.py --audit L1.audit L2.audit --csv scan.csv

  # Multiple hosts
  python audit_to_excel.py --audit CIS_L1.audit --csv host1.csv host2.csv

  # Custom output filename
  python audit_to_excel.py --audit CIS_L1.audit --csv scan.csv -o report.xlsx
""",
    )
    ap.add_argument("--audit", nargs="+", required=True, metavar="FILE",
                    help=".audit file(s) to convert")
    ap.add_argument("--csv", nargs="*", default=[], metavar="FILE",
                    help="Nessus compliance CSV export(s) — enables Automatic mode")
    ap.add_argument("--output", "-o", default="audit_report.xlsx",
                    help="Output Excel filename  [default: audit_report.xlsx]")
    args = ap.parse_args()

    missing = [f for f in args.audit + (args.csv or []) if not Path(f).exists()]
    if missing:
        for f in missing:
            print(f"[ERROR] File not found: {f}", file=sys.stderr)
        sys.exit(1)

    auto_mode = bool(args.csv)

    print(f"\n{'='*60}")
    print(f"  Mode : {'AUTOMATIC  (audit + CSV)' if auto_mode else 'MANUAL  (audit only)'}")
    print(f"{'='*60}")

    print(f"\n[1] Parsing {len(args.audit)} audit file(s)…")
    checks, bench_name = parse_audit_files(args.audit)
    print(f"    Total unique checks : {len(checks)}")
    print(f"    Benchmark name      : {bench_name}")

    if not checks:
        print("[ERROR] No checks found. Verify the .audit file contains "
              "<custom_item> or <report> blocks with benchmark IDs.")
        sys.exit(1)

    if auto_mode:
        print(f"\n[2] Parsing {len(args.csv)} CSV file(s)…")
        csv_results = parse_csv_files(args.csv)
        print(f"    Total CSV rows : {sum(len(v) for v in csv_results.values())}")
        print(f"\n[3] Merging…")
        checks = merge_csv(checks, csv_results)
    else:
        print("\n    No CSV — Status and Observed Value will be blank.")
        print("    Tip: add --csv scan.csv to auto-populate from Nessus.")

    print(f"\n[{'4' if auto_mode else '2'}] Building Excel…")
    build_excel(checks, bench_name, args.output, auto_mode)


if __name__ == "__main__":
    main()
