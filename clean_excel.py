"""
clean_excel.py — Core cleaning logic & analytics engine
"""

import re
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Constants & Mappings ──────────────────────────────────────────────────────

WORD_TO_NUM = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10
}

COUNTRY_MAP = {
    "usa": "USA", "us": "USA", "uk": "UK", "gb": "UK", "uae": "UAE",
    "canada": "Canada", "ca": "Canada", "germany": "Germany", "de": "Germany",
    "france": "France", "fr": "France", "australia": "Australia", "au": "Australia",
    "india": "India", "in": "India"
}

MISSING_PLACEHOLDERS = {"", "n/a", "null", "none", "-", "nan", "nil", "#n/a", "tbd"}

def is_missing(val):
    if pd.isna(val):
        return True
    if isinstance(val, str) and val.strip().lower() in MISSING_PLACEHOLDERS:
        return True
    return False

def get_id_gen():
    i = 0
    def gen():
        nonlocal i
        i += 1
        return f"MISSING-{i:03d}"
    return gen

# ── Analytics Engine ──────────────────────────────────────────────────────────

def generate_analytics(df: pd.DataFrame):
    """Compute data quality metrics for the dashboard."""
    analytics = {}
    
    # Missing Values
    missing_matrix = df.map(is_missing)
    analytics['total_missing'] = int(missing_matrix.sum().sum())
    analytics['missing_by_col'] = missing_matrix.sum().to_dict()
    
    # Duplicates
    analytics['total_duplicates'] = int(df.duplicated().sum())
    
    # Columns Analyzed
    analytics['total_columns'] = len(df.columns)
    analytics['total_rows'] = len(df)
    
    # Invalid Emails
    email_col = 'Email' if 'Email' in df.columns else None
    invalid_emails = 0
    if email_col:
        for val in df[email_col]:
            if not is_missing(val):
                parts = str(val).split('@')
                if len(parts) != 2 or '.' not in parts[-1] or '@@' in str(val):
                    invalid_emails += 1
    analytics['invalid_emails'] = invalid_emails
    
    # Invalid Phones
    phone_col = 'Phone' if 'Phone' in df.columns else None
    invalid_phones = 0
    if phone_col:
        for val in df[phone_col]:
            if not is_missing(val):
                digits = re.sub(r'\D', '', str(val))
                if len(digits) not in [10, 11]:
                    invalid_phones += 1
    analytics['invalid_phones'] = invalid_phones
    
    # Outliers (IQR on numeric cols)
    outliers = {}
    total_outliers = 0
    for col in df.select_dtypes(include=[np.number]).columns:
        q1 = df[col].quantile(0.25)
        q3 = df[col].quantile(0.75)
        iqr = q3 - q1
        if iqr > 0:
            mask = (df[col] < (q1 - 3 * iqr)) | (df[col] > (q3 + 3 * iqr))
            count = int(mask.sum())
            if count > 0:
                outliers[col] = count
                total_outliers += count
    analytics['outliers'] = outliers
    analytics['total_outliers'] = total_outliers
    
    # Health Score (100 - penalties) - FIXED CALCULATION
    penalty = 0
    if len(df) > 0:
        total_cells = len(df) * len(df.columns)
        penalty += (analytics['total_missing'] / total_cells) * 50  # 50% weight
        penalty += (analytics['total_duplicates'] / len(df)) * 20  # 20% weight
        penalty += (analytics['total_outliers'] / len(df)) * 10    # 10% weight
        penalty += (analytics['invalid_emails'] / len(df)) * 10    # 10% weight
        penalty += (analytics['invalid_phones'] / len(df)) * 10    # 10% weight
        
    analytics['health_score'] = int(max(0, 100 - penalty))
    
    return analytics

# ── Main cleaning function ────────────────────────────────────────────────────

def clean_dataframe(df: pd.DataFrame):
    report = {
        "original_shape": df.shape,
        "final_shape": df.shape,
        "duplicates_removed": 0,
        "empty_rows_removed": 0,
        "empty_cols_removed": 0,
        "fixed_counts": {}
    }
    
    df.columns = [str(c).strip() for c in df.columns]
    before_rows = len(df)
    df.dropna(how="all", inplace=True)
    report["blank_rows_removed"] = before_rows - len(df)

    before_dupes = len(df)
    df.drop_duplicates(inplace=True)
    report["duplicates_removed"] = before_dupes - len(df)

    fixed_counts = {col: 0 for col in df.columns}
    missing_id_gen = get_id_gen()
    cleaned_rows = []

    for _, row in df.iterrows():
        issues = []
        fixed = set()

        # ── Order ID ──
        oid = row.get('Order ID', np.nan)
        if is_missing(oid):
            issues.append("Generated missing Order ID")
            fixed.add('Order ID')
            oid = missing_id_gen()
        else:
            cleaned_oid = re.sub(r'\s+', '', str(oid).upper())
            if cleaned_oid != str(oid).strip():
                fixed.add('Order ID')
            oid = cleaned_oid

        # ── Email ──
        email = row.get('Email', np.nan)
        if is_missing(email):
            email = None
        else:
            cleaned_email = str(email).strip().lower()
            valid = True
            if "@@" in cleaned_email: valid = False
            parts = cleaned_email.split('@')
            if len(parts) != 2 or not parts[0] or not parts[1] or '.' not in parts[-1]: valid = False

            if not valid:
                email = None
            elif cleaned_email != str(email).strip().lower():
                fixed.add('Email')
                email = cleaned_email

        # ── Customer Name ──
        cname = row.get('Customer Name', np.nan)
        if is_missing(cname):
            if email:  # Derive from email
                prefix = str(email).split('@')[0]
                clean_prefix = re.sub(r'[._\-+]', ' ', prefix).strip()
                if clean_prefix:
                    cname = clean_prefix.title()
                    issues.append("Derived customer name from email")
                    fixed.add('Customer Name')
            if is_missing(cname):
                issues.append("Missing customer name")
                fixed.add('Customer Name')
                cname = "Unknown Customer"
        else:
            cleaned_cname = str(cname).strip().title()
            if cleaned_cname != str(cname).strip():
                fixed.add('Customer Name')
            cname = cleaned_cname

        # ── Phone ──
        phone = row.get('Phone', np.nan)
        if is_missing(phone):
            phone = None
        else:
            digits = re.sub(r'\D', '', str(phone))
            if digits == "0000000000" or str(phone).strip() == "000-000-0000":
                phone = None
            elif len(digits) == 10:
                cleaned_phone = f"+1-{digits[:3]}-{digits[3:6]}-{digits[6:]}"
                if cleaned_phone != str(phone).strip(): fixed.add('Phone')
                phone = cleaned_phone
            elif len(digits) == 11 and digits.startswith('1'):
                cleaned_phone = f"+1-{digits[1:4]}-{digits[4:7]}-{digits[7:]}"
                if cleaned_phone != str(phone).strip(): fixed.add('Phone')
                phone = cleaned_phone
            else:
                phone = None

        # ── Country ──
        country = row.get('Country', np.nan)
        if is_missing(country) or str(country).strip().lower() == 'xx':
            issues.append("Missing country")
            fixed.add('Country')
            country = "Unknown"
        else:
            c = str(country).strip().lower()
            cleaned_country = COUNTRY_MAP.get(c, str(country).strip().title())
            if cleaned_country != str(country).strip(): fixed.add('Country')
            country = cleaned_country

        # ── Order Date ──
        odate = row.get('Order Date', np.nan)
        if is_missing(odate):
            issues.append("Missing date")
            fixed.add('Order Date')
            odate = pd.Timestamp("1900-01-01")
        else:
            odate_str = str(odate).strip()
            parsed_date = pd.NaT
            match = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', odate_str)
            if match and int(match.group(1)) > 12:
                parsed_date = pd.to_datetime(odate_str, dayfirst=True, errors='coerce')
            else:
                parsed_date = pd.to_datetime(odate_str, errors='coerce')
                if pd.isna(parsed_date): parsed_date = pd.to_datetime(odate_str, dayfirst=True, errors='coerce')

            if pd.isna(parsed_date):
                issues.append("Invalid date")
                fixed.add('Order Date')
                odate = pd.Timestamp("1900-01-01")
            else:
                if parsed_date.strftime("%Y-%m-%d") != odate_str: fixed.add('Order Date')
                odate = parsed_date

        # ── Product ──
        prod = row.get('Product', np.nan)
        if is_missing(prod):
            issues.append("Missing product")
            fixed.add('Product')
            prod = "Unknown Product"
        else:
            cleaned_prod = str(prod).strip().title()
            if cleaned_prod != str(prod).strip(): fixed.add('Product')
            prod = cleaned_prod

        # ── Quantity ──
        qty = row.get('Quantity', np.nan)
        if is_missing(qty):
            issues.append("Assumed quantity=1")
            fixed.add('Quantity')
            qty = 1
        else:
            qty_str = str(qty).strip().lower()
            if qty_str in WORD_TO_NUM:
                fixed.add('Quantity')
                qty = WORD_TO_NUM[qty_str]
            else:
                try:
                    val = float(qty_str.replace(',', ''))
                    if val < 0: val = abs(val); fixed.add('Quantity')
                    if val == 0: val = 1; issues.append("Assumed quantity=1"); fixed.add('Quantity')
                    cleaned_qty = int(val)
                    if qty_str != str(cleaned_qty): fixed.add('Quantity')
                    qty = cleaned_qty
                except ValueError:
                    issues.append("Assumed quantity=1")
                    fixed.add('Quantity')
                    qty = 1

        # ── Unit Price ──
        uprice = row.get('Unit Price', np.nan)
        if is_missing(uprice):
            issues.append("Assumed unit price=0.0")
            fixed.add('Unit Price')
            uprice = 0.0
        else:
            up_str = str(uprice).strip().lower()
            cleaned_up_str = re.sub(r'[\$£€]', '', up_str)
            cleaned_up_str = re.sub(r'\busd\b|\baed\b|\bgbp\b|\beur\b', '', cleaned_up_str)
            cleaned_up_str = cleaned_up_str.replace(',', '').strip()
            try:
                val = float(cleaned_up_str)
                if val < 0: val = abs(val); fixed.add('Unit Price')
                cleaned_up = round(val, 2)
                if cleaned_up == 0.0: issues.append("Assumed unit price=0.0"); fixed.add('Unit Price')
                if str(uprice).strip().lower() != f"{cleaned_up:.2f}": fixed.add('Unit Price')
                uprice = cleaned_up
            except ValueError:
                issues.append("Assumed unit price=0.0")
                fixed.add('Unit Price')
                uprice = 0.0

        # ── Total ── (Always Recalculate)
        total_orig = row.get('Total', np.nan)
        total = round(qty * uprice, 2)
        if is_missing(total_orig):
            fixed.add('Total')
        else:
            tot_str_clean = re.sub(r'[\$£€]', '', str(total_orig).strip().lower())
            tot_str_clean = re.sub(r'\busd\b|\baed\b|\bgbp\b|\beur\b', '', tot_str_clean).replace(',', '').strip()
            try:
                val_rounded = round(abs(float(tot_str_clean)), 2)
                if val_rounded != total: fixed.add('Total')
            except ValueError:
                fixed.add('Total')

        # ── Status ──
        status = row.get('Status', np.nan)
        valid_statuses = {"Completed", "Pending", "Shipped", "Refunded", "Cancelled"}
        st_raw = str(status).strip() if not is_missing(status) else ""
        if is_missing(status) or st_raw in {"???", "NULL"}:
            issues.append("Unknown status")
            fixed.add('Status')
            status = "Unknown"
        else:
            st = st_raw.title()
            if st in valid_statuses:
                if st != st_raw: fixed.add('Status')
                status = st
            else:
                issues.append("Unknown status")
                fixed.add('Status')
                status = "Unknown"

        # ── Sales Rep ──
        srep = row.get('Sales Rep', np.nan)
        if is_missing(srep):
            issues.append("Unassigned sales rep")
            fixed.add('Sales Rep')
            srep = "Unassigned"
        else:
            cleaned_srep = str(srep).strip().title()
            if cleaned_srep != str(srep).strip(): fixed.add('Sales Rep')
            srep = cleaned_srep

        for col in fixed:
            if col in fixed_counts: fixed_counts[col] += 1

        cleaned_rows.append([
            oid, cname, email, phone, country, odate, prod, qty, uprice, total, status, srep, ", ".join(issues)
        ])

    cleaned_df = pd.DataFrame(cleaned_rows, columns=[
        'Order ID', 'Customer Name', 'Email', 'Phone', 'Country', 'Order Date',
        'Product', 'Quantity', 'Unit Price', 'Total', 'Status', 'Sales Rep', 'Data Quality Issues'
    ])
    
    cleaned_df['Order Date'] = pd.to_datetime(cleaned_df['Order Date'], errors='coerce')
    cleaned_df['Quantity'] = pd.to_numeric(cleaned_df['Quantity'], errors='coerce').astype(int)
    cleaned_df['Unit Price'] = pd.to_numeric(cleaned_df['Unit Price'], errors='coerce').astype(float)
    cleaned_df['Total'] = pd.to_numeric(cleaned_df['Total'], errors='coerce').astype(float)

    report["fixed_counts"] = fixed_counts
    report["final_shape"] = cleaned_df.shape
    
    return cleaned_df, report

# ── Excel writer ──────────────────────────────────────────────────────────────

def _auto_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try: max_len = max(max_len, len(str(cell.value or "")))
            except: pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)

def write_cleaned_excel(df_clean: pd.DataFrame, report: dict, path: str):
    wb = Workbook()

    # Sheet 1: Cleaned Data
    ws1 = wb.active
    ws1.title = "Cleaned Data"
    ws1.sheet_view.showGridLines = True
    for j, col in enumerate(df_clean.columns, 1):
        cell = ws1.cell(row=1, column=j, value=str(col))
        cell.fill = PatternFill("solid", fgColor="F1F5F9")
        cell.font = Font(bold=True, color="1E293B", name="Calibri", size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, (_, row) in enumerate(df_clean.iterrows(), 2):
        for j, col in enumerate(df_clean.columns, 1):
            val = row[col]
            if pd.isna(val): val = None
            elif isinstance(val, pd.Timestamp): val = val.to_pydatetime()
            cell = ws1.cell(row=i, column=j, value=val)
            if col == 'Order Date' and val: cell.number_format = 'YYYY-MM-DD'
            elif col == 'Quantity' and val: cell.number_format = '0'
            elif col in ['Unit Price', 'Total'] and val: cell.number_format = '0.00'
    ws1.freeze_panes = "A2"
    _auto_width(ws1)

    # Sheet 2: Cleaning Report
    ws3 = wb.create_sheet("Cleaning Report")
    ws3.sheet_view.showGridLines = False
    fill_hdr = PatternFill("solid", fgColor="F1F5F9")
    font_hdr = Font(bold=True, color="2563EB", name="Calibri", size=11)
    font_key = Font(bold=True, color="1E293B", name="Calibri", size=11)
    font_val = Font(color="475569", name="Calibri", size=11)

    rows = [
        ("Metric", "Value"),
        ("Original rows", report["original_shape"][0]),
        ("Final rows", report["final_shape"][0]),
        ("Duplicates removed", report["duplicates_removed"]),
        ("", ""),
        ("Cells Fixed Per Column", ""),
    ]
    for col, count in report.get("fixed_counts", {}).items():
        rows.append((col, count))

    for r_idx, (key, val) in enumerate(rows, 1):
        c1 = ws3.cell(row=r_idx, column=1, value=key)
        c2 = ws3.cell(row=r_idx, column=2, value=val)
        if r_idx == 1:
            c1.fill = fill_hdr; c1.font = font_hdr
            c2.fill = fill_hdr; c2.font = font_hdr
        else:
            c1.font = font_key; c2.font = font_val

    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 60
    wb.save(path)
