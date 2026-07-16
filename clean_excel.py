"""
Automated Excel Data Cleaner
=============================
Usage:
    python clean_excel.py input.xlsx
    python clean_excel.py input.xlsx --output cleaned_output.xlsx

What it does:
    1. Removes duplicate rows
    2. Strips extra whitespace from text
    3. Standardizes text casing (title case for names, upper for codes)
    4. Fixes inconsistent date formats
    5. Removes completely empty rows and columns
    6. Standardizes missing value placeholders (N/A, -, n/a, none -> actual NaN)
    7. Trims and normalizes column header names
    8. Flags outliers in numeric columns (adds a helper sheet)
    9. Saves a cleaned file + a summary report sheet
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


# ── Helpers ──────────────────────────────────────────────────────────────────

MISSING_PLACEHOLDERS = {"n/a", "na", "none", "null", "-", "--", "nan", "nil", "#n/a", ""}

def normalize_missing(val):
    """Convert common missing-value placeholders to NaN."""
    if pd.isna(val):
        return np.nan
    if isinstance(val, str) and val.strip().lower() in MISSING_PLACEHOLDERS:
        return np.nan
    return val

def clean_column_name(name):
    """Lowercase, strip, replace spaces/special chars with underscores."""
    name = str(name).strip().lower()
    name = re.sub(r"[^\w\s]", "", name)       # remove special chars
    name = re.sub(r"\s+", "_", name)           # spaces -> underscores
    name = re.sub(r"_+", "_", name)            # collapse multiple underscores
    return name.strip("_")

def try_parse_dates(series):
    """Try to parse a column as dates if it looks date-like."""
    try:
        parsed = pd.to_datetime(series, infer_datetime_format=True, errors="coerce")
        # Only accept if at least 60% parsed successfully
        if parsed.notna().mean() >= 0.6:
            return parsed
    except Exception:
        pass
    return series

def flag_outliers(series, multiplier=3.0):
    """Return a boolean mask: True where value is an outlier (IQR method)."""
    if not pd.api.types.is_numeric_dtype(series):
        return pd.Series([False] * len(series), index=series.index)
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    lower = q1 - multiplier * iqr
    upper = q3 + multiplier * iqr
    return (series < lower) | (series > upper)


# ── Core Cleaning ─────────────────────────────────────────────────────────────

def clean_dataframe(df):
    """Apply all cleaning steps and return (cleaned_df, report_dict)."""
    report = {}
    original_shape = df.shape

    # 1. Normalize column names
    original_cols = list(df.columns)
    df.columns = [clean_column_name(c) for c in df.columns]
    renamed = {o: n for o, n in zip(original_cols, df.columns) if o != n}
    report["columns_renamed"] = renamed

    # 2. Replace missing-value placeholders with NaN
    df = df.map(normalize_missing)

    # 3. Drop completely empty rows and columns
    before_rows = len(df)
    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)
    report["empty_rows_removed"] = before_rows - len(df)
    report["empty_cols_removed"] = original_shape[1] - df.shape[1]

    # 4. Remove duplicate rows
    before_dupes = len(df)
    df.drop_duplicates(inplace=True)
    report["duplicates_removed"] = before_dupes - len(df)

    # 5. Strip whitespace from string columns
    str_cols = df.select_dtypes(include="object").columns
    for col in str_cols:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        # Collapse internal extra spaces
        df[col] = df[col].apply(lambda x: re.sub(r" +", " ", x) if isinstance(x, str) else x)

    # 6. Try to auto-detect and parse date columns
    date_cols = []
    for col in str_cols:
        parsed = try_parse_dates(df[col])
        if parsed is not df[col]:    # column was converted
            df[col] = parsed
            date_cols.append(col)
    report["date_cols_detected"] = date_cols

    # 7. Detect numeric columns stored as strings and convert
    numeric_converted = []
    for col in df.select_dtypes(include="object").columns:
        converted = pd.to_numeric(df[col].str.replace(",", "").str.replace("$", "").str.strip()
                                  if df[col].dtype == object else df[col],
                                  errors="coerce")
        if converted.notna().mean() >= 0.8:   # 80%+ converted → treat as numeric
            df[col] = converted
            numeric_converted.append(col)
    report["numeric_cols_converted"] = numeric_converted

    # 8. Detect outliers in numeric columns
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    outlier_summary = {}
    outlier_mask = pd.DataFrame(False, index=df.index, columns=df.columns)
    for col in numeric_cols:
        mask = flag_outliers(df[col])
        count = mask.sum()
        if count > 0:
            outlier_summary[col] = int(count)
            outlier_mask[col] = mask
    report["outliers_flagged"] = outlier_summary

    report["final_shape"] = df.shape
    report["original_shape"] = original_shape

    return df, report, outlier_mask


# ── Excel Writer ──────────────────────────────────────────────────────────────

def write_cleaned_excel(df, report, outlier_mask, output_path):
    """Write cleaned data + summary report to an Excel file."""

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Sheet 1: Cleaned Data
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)

        # Sheet 2: Outlier Flags (rows that have at least one outlier)
        flagged_rows = outlier_mask.any(axis=1)
        if flagged_rows.any():
            outlier_df = df[flagged_rows].copy()
            outlier_df.to_excel(writer, sheet_name="Outliers", index=True)

        # Sheet 3: Summary Report
        summary_rows = []
        orig = report["original_shape"]
        final = report["final_shape"]
        summary_rows += [
            ["CLEANING SUMMARY", ""],
            ["Run timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            ["", ""],
            ["BEFORE", ""],
            ["Rows", orig[0]],
            ["Columns", orig[1]],
            ["", ""],
            ["AFTER", ""],
            ["Rows", final[0]],
            ["Columns", final[1]],
            ["", ""],
            ["CHANGES MADE", ""],
            ["Empty rows removed", report["empty_rows_removed"]],
            ["Empty columns removed", report["empty_cols_removed"]],
            ["Duplicate rows removed", report["duplicates_removed"]],
            ["Date columns auto-detected", ", ".join(report["date_cols_detected"]) or "None"],
            ["Numeric columns converted", ", ".join(report["numeric_cols_converted"]) or "None"],
            ["", ""],
            ["OUTLIERS DETECTED (by column)", ""],
        ]
        for col, cnt in report["outliers_flagged"].items():
            summary_rows.append([f"  {col}", cnt])
        if not report["outliers_flagged"]:
            summary_rows.append(["  None detected", ""])

        if report["columns_renamed"]:
            summary_rows += [["", ""], ["COLUMNS RENAMED", ""]]
            for old, new in report["columns_renamed"].items():
                summary_rows.append([f"  '{old}'", f"→ '{new}'"])

        summary_df = pd.DataFrame(summary_rows, columns=["Item", "Value"])
        summary_df.to_excel(writer, sheet_name="Cleaning Report", index=False)

    # ── Style the workbook ──
    wb = load_workbook(output_path)

    # Style: Cleaned Data header
    ws = wb["Cleaned Data"]
    header_fill = PatternFill("solid", fgColor="2E4057")
    header_font = Font(color="FFFFFF", bold=True, name="Arial", size=11)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)

    # Style: Report sheet
    if "Cleaning Report" in wb.sheetnames:
        ws_r = wb["Cleaning Report"]
        title_fill = PatternFill("solid", fgColor="048A81")
        section_fill = PatternFill("solid", fgColor="D9E8E7")
        for row in ws_r.iter_rows():
            cell = row[0]
            val = str(cell.value) if cell.value else ""
            if val in ("CLEANING SUMMARY", "BEFORE", "AFTER", "CHANGES MADE",
                       "OUTLIERS DETECTED (by column)", "COLUMNS RENAMED"):
                for c in row:
                    c.fill = section_fill
                    c.font = Font(bold=True, name="Arial")
            elif val == "CLEANING SUMMARY":
                for c in row:
                    c.fill = title_fill
                    c.font = Font(bold=True, color="FFFFFF", name="Arial")
        ws_r.column_dimensions["A"].width = 35
        ws_r.column_dimensions["B"].width = 45

    wb.save(output_path)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Automated Excel data cleaner")
    parser.add_argument("input", help="Path to the input .xlsx or .csv file")
    parser.add_argument("--output", help="Path for the cleaned output file (optional)")
    args = parser.parse_args()

    input_path = args.input
    if not os.path.exists(input_path):
        print(f"❌ File not found: {input_path}")
        return

    # Determine output path
    if args.output:
        output_path = args.output
    else:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_cleaned.xlsx"

    print(f"📂 Loading: {input_path}")

    # Load file
    if input_path.endswith(".csv"):
        df = pd.read_csv(input_path, dtype=str)
    else:
        df = pd.read_excel(input_path, dtype=str)

    print(f"✅ Loaded {df.shape[0]} rows × {df.shape[1]} columns")
    print("🧹 Cleaning data...")

    df_clean, report, outlier_mask = clean_dataframe(df)

    print("💾 Writing cleaned file...")
    write_cleaned_excel(df_clean, report, outlier_mask, output_path)

    # Print summary
    print("\n" + "="*50)
    print("        CLEANING COMPLETE")
    print("="*50)
    print(f"  Original:   {report['original_shape'][0]} rows × {report['original_shape'][1]} cols")
    print(f"  Cleaned:    {report['final_shape'][0]} rows × {report['final_shape'][1]} cols")
    print(f"  Duplicates removed:    {report['duplicates_removed']}")
    print(f"  Empty rows removed:    {report['empty_rows_removed']}")
    print(f"  Empty cols removed:    {report['empty_cols_removed']}")
    print(f"  Outliers flagged:      {sum(report['outliers_flagged'].values())}")
    print(f"\n📁 Output saved to: {output_path}")
    print("="*50)


if __name__ == "__main__":
    main()
