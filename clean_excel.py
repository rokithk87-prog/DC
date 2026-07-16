"""
Sales Dataset Cleaner
======================
Usage:
    python clean_excel.py input.xlsx
    python clean_excel.py input.xlsx --output cleaned_output.xlsx

Applies deep, column-specific cleaning rules to a sales dataset:
  - Order ID, Customer Name, Email, Phone, Country
  - Order Date, Product, Quantity, Unit Price, Total, Status, Sales Rep
  - Removes duplicates and blank rows
  - Recalculates Total = Quantity × Unit Price
  - Adds "Data Quality Issues" flag column
  - Outputs "Cleaned Data" + "Cleaning Summary" sheets
"""

import pandas as pd
import numpy as np
import argparse
import os
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


# ── Column-level cleaners ─────────────────────────────────────────────────────

missing_counter = {"n": 0}

def clean_order_id(val):
    issues = []
    if pd.isna(val) or str(val).strip().lower() in ("", "n/a", "na", "none", "null"):
        missing_counter["n"] += 1
        return f"MISSING-{missing_counter['n']:03d}", ["Missing Order ID"]
    v = str(val).strip().upper()
    if v != str(val).strip():
        issues.append("Order ID whitespace stripped")
    return v, issues


def clean_customer_name(val):
    issues = []
    if pd.isna(val) or str(val).strip().lower() in ("", "n/a", "na", "none", "null"):
        return "Unknown Customer", ["Missing customer name"]
    v = str(val).strip().title()
    if v != str(val).strip():
        issues.append("Customer name standardized")
    return v, issues


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def clean_email(val):
    issues = []
    if pd.isna(val) or str(val).strip().lower() in ("", "n/a", "na", "none", "null"):
        return "missing@unknown.com", ["Missing email"]
    v = str(val).strip().lower()
    # Double-@ or other malformed patterns
    if v.count("@") != 1 or not EMAIL_RE.match(v):
        issues.append("Invalid email")
        return "invalid@unknown.com", issues
    return v, issues


def clean_phone(val):
    issues = []
    raw = str(val).strip() if not pd.isna(val) else ""
    if raw.lower() in ("", "n/a", "na", "none", "null", "000-000-0000", "0000000000"):
        return "N/A", []
    # Strip all non-digit characters
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 10:
        formatted = f"+1-{digits[:3]}-{digits[3:6]}-{digits[6:]}"
        return formatted, []
    elif len(digits) == 11 and digits[0] == "1":
        formatted = f"+1-{digits[1:4]}-{digits[4:7]}-{digits[7:]}"
        return formatted, []
    else:
        issues.append("Unrecognized phone format")
        return "N/A", issues


COUNTRY_MAP = {
    "usa": "USA", "us": "USA", "united states": "USA", "united states of america": "USA",
    "uk": "UK", "gb": "UK", "great britain": "UK", "united kingdom": "UK",
    "uae": "UAE", "united arab emirates": "UAE",
    "canada": "Canada", "ca": "Canada",
    "germany": "Germany", "de": "Germany",
    "france": "France", "fr": "France",
    "australia": "Australia", "au": "Australia",
    "india": "India", "in": "India",
}

def clean_country(val):
    issues = []
    if pd.isna(val) or str(val).strip().lower() in ("", "n/a", "na", "none", "null", "xx"):
        return "Unknown", ["Missing/unknown country"]
    v = str(val).strip()
    mapped = COUNTRY_MAP.get(v.lower())
    if mapped:
        return mapped, []
    # Title-case fallback for anything not in map
    result = v.title()
    if result.lower() == v.lower() and result not in COUNTRY_MAP.values():
        issues.append("Unrecognized country")
        return "Unknown", issues
    return result, []


DATE_FORMATS = [
    "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%Y-%m-%d",
    "%d-%m-%Y", "%m-%d-%Y",
    "%d-%b-%Y", "%b %d %Y", "%B %d %Y", "%B %d, %Y",
    "%d %b %Y", "%d %B %Y",
    "%b-%d-%Y", "%B-%d-%Y",
]

def clean_date(val):
    issues = []
    raw = str(val).strip() if not pd.isna(val) else ""
    if raw.lower() in ("", "n/a", "na", "none", "null"):
        return "1900-01-01", ["Missing date"]
    # Try pandas first (handles many formats)
    try:
        parsed = pd.to_datetime(raw, infer_datetime_format=True, dayfirst=False)
        return parsed.strftime("%Y-%m-%d"), []
    except Exception:
        pass
    # Try manual formats
    for fmt in DATE_FORMATS:
        try:
            parsed = datetime.strptime(raw, fmt)
            return parsed.strftime("%Y-%m-%d"), []
        except ValueError:
            continue
    issues.append("Invalid date")
    return "1900-01-01", issues


def clean_product(val):
    issues = []
    if pd.isna(val) or str(val).strip().lower() in ("", "n/a", "na", "none", "null"):
        return "Unknown Product", ["Missing product"]
    v = str(val).strip().title()
    return v, []


WORD_TO_INT = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
}

def clean_quantity(val):
    issues = []
    if pd.isna(val):
        return 1, ["Missing quantity, defaulted to 1"]
    raw = str(val).strip().lower()
    if raw in ("", "n/a", "na", "none", "null", "0"):
        return 1, ["Missing/zero quantity, defaulted to 1"]
    # Word numbers
    if raw in WORD_TO_INT:
        return WORD_TO_INT[raw], []
    # Try numeric
    try:
        n = float(re.sub(r"[^\d.\-]", "", raw))
        n = abs(int(n))
        if n == 0:
            return 1, ["Zero quantity, defaulted to 1"]
        if float(re.sub(r"[^\d.\-]", "", raw)) < 0:
            issues.append("Negative quantity corrected")
        return n, issues
    except (ValueError, TypeError):
        return 1, ["Non-numeric quantity, defaulted to 1"]


CURRENCY_RE = re.compile(r"[^\d.\-]")

def clean_price(val, col_label="Unit Price"):
    issues = []
    if pd.isna(val):
        return 0.0, [f"Missing {col_label}, set to 0.0"]
    raw = str(val).strip()
    if raw.lower() in ("", "n/a", "na", "none", "null", "tbd", "0", "0.0"):
        return 0.0, [f"Missing/zero {col_label}, set to 0.0"]
    cleaned = CURRENCY_RE.sub("", raw)
    try:
        n = float(cleaned)
        if n < 0:
            issues.append(f"Negative {col_label} corrected")
            n = abs(n)
        return round(n, 2), issues
    except (ValueError, TypeError):
        return 0.0, [f"Non-numeric {col_label}, set to 0.0"]


VALID_STATUSES = {"Completed", "Pending", "Shipped", "Refunded", "Cancelled"}

def clean_status(val):
    if pd.isna(val) or str(val).strip().lower() in ("", "n/a", "na", "none", "null", "???"):
        return "Unknown", ["Unknown status"]
    v = str(val).strip().title()
    if v not in VALID_STATUSES:
        return "Unknown", ["Unrecognized status"]
    return v, []


def clean_sales_rep(val):
    if pd.isna(val) or str(val).strip().lower() in ("", "n/a", "na", "none", "null"):
        return "Unassigned", ["Missing sales rep"]
    return str(val).strip().title(), []


# ── Main cleaner ──────────────────────────────────────────────────────────────

# Map from canonical column names to the cleaner functions
COLUMN_CLEANERS = {
    "order id":      clean_order_id,
    "customer name": clean_customer_name,
    "email":         clean_email,
    "phone":         clean_phone,
    "country":       clean_country,
    "order date":    clean_date,
    "product":       clean_product,
    "quantity":      clean_quantity,
    "unit price":    lambda v: clean_price(v, "Unit Price"),
    "status":        clean_status,
    "sales rep":     clean_sales_rep,
}

def normalize_col(name):
    return str(name).strip().lower()


def clean_dataframe(df):
    report = {}
    original_rows = len(df)

    # ── Track fix counts per column ──
    fix_counts = {col: 0 for col in df.columns}

    # ── Remove fully blank rows ──
    df = df.dropna(how="all").reset_index(drop=True)
    blank_rows_removed = original_rows - len(df)

    # ── Remove exact duplicates ──
    before_dupes = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    duplicates_removed = before_dupes - len(df)

    # Reset the missing-order-ID counter
    missing_counter["n"] = 0

    # ── Apply column-level cleaning ──
    col_map = {normalize_col(c): c for c in df.columns}
    issues_col = [""] * len(df)

    for canonical, cleaner in COLUMN_CLEANERS.items():
        actual = col_map.get(canonical)
        if actual is None:
            continue
        new_vals = []
        for i, val in enumerate(df[actual]):
            cleaned, issues = cleaner(val)
            new_vals.append(cleaned)
            if issues:
                fix_counts[actual] = fix_counts.get(actual, 0) + 1
                existing = issues_col[i]
                issues_col[i] = (existing + ", " + ", ".join(issues)).strip(", ")
        df[actual] = new_vals

    # ── Recalculate Total ──
    qty_col   = col_map.get("quantity")
    price_col = col_map.get("unit price")
    total_col = col_map.get("total")

    if qty_col and price_col:
        recalc_totals = [
            round(float(q) * float(p), 2)
            for q, p in zip(df[qty_col], df[price_col])
        ]
        if total_col:
            df[total_col] = recalc_totals
            fix_counts[total_col] = len(recalc_totals)
        else:
            df["Total"] = recalc_totals

    # ── Add Data Quality Issues column ──
    df["Data Quality Issues"] = issues_col

    report["original_rows"]      = original_rows
    report["final_rows"]         = len(df)
    report["duplicates_removed"] = duplicates_removed
    report["blank_rows_removed"] = blank_rows_removed
    report["fix_counts"]         = fix_counts

    return df, report


# ── Excel Writer ──────────────────────────────────────────────────────────────

def write_output(df, report, output_path):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)

        # Cleaning Summary sheet
        summary_rows = [
            ["Metric", "Value"],
            ["Total rows before cleaning", report["original_rows"]],
            ["Total rows after cleaning",  report["final_rows"]],
            ["Duplicates removed",         report["duplicates_removed"]],
            ["Blank rows removed",         report["blank_rows_removed"]],
            ["", ""],
            ["Cells fixed per column", ""],
        ]
        for col, cnt in report["fix_counts"].items():
            summary_rows.append([col, cnt])

        summary_df = pd.DataFrame(summary_rows[1:], columns=summary_rows[0])
        summary_df.to_excel(writer, sheet_name="Cleaning Summary", index=False)

    # ── Styling ──
    wb = load_workbook(output_path)

    # Style Cleaned Data header
    ws = wb["Cleaned Data"]
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(color="FFFFFF", bold=True, name="Arial", size=11)
    for cell in ws[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center")
    # Alternate row shading
    light = PatternFill("solid", fgColor="EEF2F7")
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        if row_idx % 2 == 0:
            for cell in row:
                cell.fill = light
    # Auto-width
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 45)

    # Style Summary sheet
    ws2 = wb["Cleaning Summary"]
    hdr2_fill = PatternFill("solid", fgColor="048A81")
    hdr2_font = Font(color="FFFFFF", bold=True, name="Arial", size=11)
    for cell in ws2[1]:
        cell.fill = hdr2_fill
        cell.font = hdr2_font
    section_fill = PatternFill("solid", fgColor="D9E8E7")
    for row in ws2.iter_rows(min_row=2):
        if row[0].value == "Cells fixed per column":
            for cell in row:
                cell.fill = section_fill
                cell.font = Font(bold=True, name="Arial")
        else:
            for cell in row:
                cell.font = Font(name="Arial")
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 20

    wb.save(output_path)
    print(f"💾 Saved → {output_path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Sales dataset deep cleaner")
    parser.add_argument("input", help="Path to input .xlsx or .csv file")
    parser.add_argument("--output", help="Output path (optional)")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"❌ File not found: {args.input}")
        return

    output_path = args.output or (os.path.splitext(args.input)[0] + "_cleaned.xlsx")

    print(f"📂 Loading: {args.input}")
    if args.input.endswith(".csv"):
        df = pd.read_csv(args.input, dtype=str)
    else:
        df = pd.read_excel(args.input, dtype=str)
    print(f"✅ Loaded {df.shape[0]} rows × {df.shape[1]} columns")

    print("🧹 Cleaning...")
    df_clean, report = clean_dataframe(df)

    write_output(df_clean, report, output_path)

    print("\n" + "=" * 50)
    print("         CLEANING COMPLETE")
    print("=" * 50)
    print(f"  Original rows:      {report['original_rows']}")
    print(f"  Cleaned rows:       {report['final_rows']}")
    print(f"  Duplicates removed: {report['duplicates_removed']}")
    print(f"  Blank rows removed: {report['blank_rows_removed']}")
    print(f"\n  Fixes per column:")
    for col, cnt in report["fix_counts"].items():
        if cnt:
            print(f"    {col}: {cnt}")
    print("=" * 50)


if __name__ == "__main__":
    main()
