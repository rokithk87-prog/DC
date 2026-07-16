"""
clean_excel.py — Core cleaning logic for Excel Cleaner Streamlit app.
"""

import re
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Column name normalisation ─────────────────────────────────────────────────

def _normalise_col(name: str) -> str:
    """snake_case, strip specials, collapse spaces."""
    s = str(name).strip()
    s = re.sub(r"[^\w\s]", "", s)          # remove punctuation
    s = re.sub(r"\s+", "_", s)             # spaces → underscore
    s = re.sub(r"_+", "_", s).strip("_")  # collapse underscores
    return s.lower()


# ── Date detection ────────────────────────────────────────────────────────────

_DATE_PATTERNS = [
    r"\d{4}-\d{2}-\d{2}",          # ISO
    r"\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4}",  # DMY / MDY
    r"\d{1,2}\s+\w+\s+\d{4}",      # "12 Jan 2023"
]
_DATE_RE = re.compile("|".join(_DATE_PATTERNS))


def _looks_like_date_col(series: pd.Series) -> bool:
    sample = series.dropna().astype(str).head(20)
    if len(sample) == 0:
        return False
    hits = sample.apply(lambda x: bool(_DATE_RE.search(x))).sum()
    return hits / len(sample) >= 0.6


def _try_parse_dates(series: pd.Series) -> pd.Series:
    try:
        parsed = pd.to_datetime(series, infer_datetime_format=True, errors="coerce")
        if parsed.notna().sum() / max(series.notna().sum(), 1) >= 0.5:
            return parsed
    except Exception:
        pass
    return series


# ── Outlier detection (IQR) ───────────────────────────────────────────────────

def _flag_outliers(series: pd.Series) -> pd.Series:
    """Return boolean mask of IQR outliers."""
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return pd.Series(False, index=series.index)
    lo, hi = q1 - 3 * iqr, q3 + 3 * iqr
    return (series < lo) | (series > hi)


# ── Main cleaning function ────────────────────────────────────────────────────

def clean_dataframe(df: pd.DataFrame):
    """
    Clean *df* in-place (caller should pass a copy).

    Returns
    -------
    df_clean    : pd.DataFrame
    report      : dict
    outlier_mask: pd.DataFrame (bool, same shape as df_clean)
    """
    report = {
        "original_shape": df.shape,
        "duplicates_removed": 0,
        "empty_rows_removed": 0,
        "empty_cols_removed": 0,
        "columns_renamed": {},
        "date_cols_detected": [],
        "numeric_cols_converted": [],
        "outliers_flagged": {},
        "final_shape": df.shape,
    }

    # 1. Drop fully-empty columns
    before_cols = set(df.columns)
    df.dropna(axis=1, how="all", inplace=True)
    # also drop cols that are entirely empty strings
    empty_str_cols = [c for c in df.columns if df[c].astype(str).str.strip().eq("").all()]
    df.drop(columns=empty_str_cols, inplace=True)
    report["empty_cols_removed"] = len(before_cols) - len(df.columns)

    # 2. Drop fully-empty rows
    before_rows = len(df)
    df.replace("", np.nan, inplace=True)
    df.dropna(how="all", inplace=True)
    report["empty_rows_removed"] = before_rows - len(df)

    # 3. Remove duplicate rows
    before_dedup = len(df)
    df.drop_duplicates(inplace=True)
    report["duplicates_removed"] = before_dedup - len(df)

    # 4. Normalise column names
    new_cols = {}
    seen = {}
    for col in df.columns:
        norm = _normalise_col(col)
        if not norm:
            norm = "col"
        # deduplicate
        if norm in seen:
            seen[norm] += 1
            norm = f"{norm}_{seen[norm]}"
        else:
            seen[norm] = 0
        if norm != str(col):
            new_cols[str(col)] = norm
        new_cols.setdefault(str(col), norm)

    # rebuild rename map (only those that actually changed)
    rename_map = {old: new for old, new in new_cols.items() if old != new}
    if rename_map:
        df.rename(columns=rename_map, inplace=True)
    report["columns_renamed"] = rename_map

    # 5. Strip whitespace from all string cells
    str_cols = df.select_dtypes(include="object").columns
    for col in str_cols:
        df[col] = df[col].astype(str).str.strip().replace("nan", np.nan)

    # 6. Detect & parse date columns
    for col in list(df.select_dtypes(include="object").columns):
        if _looks_like_date_col(df[col]):
            parsed = _try_parse_dates(df[col])
            if hasattr(parsed, "dt"):
                df[col] = parsed
                report["date_cols_detected"].append(col)

    # 7. Convert remaining object cols that look numeric
    for col in list(df.select_dtypes(include="object").columns):
        cleaned = df[col].astype(str).str.replace(r"[,\$£€%]", "", regex=True).str.strip()
        numeric = pd.to_numeric(cleaned, errors="coerce")
        ratio = numeric.notna().sum() / max(df[col].notna().sum(), 1)
        if ratio >= 0.7:
            df[col] = numeric
            report["numeric_cols_converted"].append(col)

    # 8. Flag outliers in numeric columns (IQR × 3)
    outlier_mask = pd.DataFrame(False, index=df.index, columns=df.columns)
    for col in df.select_dtypes(include=[np.number]).columns:
        mask = _flag_outliers(df[col].dropna())
        full_mask = pd.Series(False, index=df.index)
        full_mask.update(mask)
        outlier_mask[col] = full_mask
        n = int(full_mask.sum())
        if n:
            report["outliers_flagged"][col] = n

    df.reset_index(drop=True, inplace=True)
    outlier_mask.reset_index(drop=True, inplace=True)
    report["final_shape"] = df.shape

    return df, report, outlier_mask


# ── Excel writer ──────────────────────────────────────────────────────────────

def _header_style(ws, n_cols: int):
    """Bold dark header row."""
    fill = PatternFill("solid", fgColor="1E293B")
    font = Font(bold=True, color="F1F5F9", name="Calibri", size=10)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="334155")
    border = Border(bottom=thin)
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = fill
        cell.font = font
        cell.alignment = align
        cell.border = border


def _auto_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)


def _df_to_sheet(ws, df: pd.DataFrame, *, highlight_mask: pd.DataFrame | None = None):
    """Write df to ws with optional cell highlighting."""
    # Write header
    for j, col in enumerate(df.columns, 1):
        ws.cell(row=1, column=j, value=str(col))
    _header_style(ws, len(df.columns))

    orange_fill = PatternFill("solid", fgColor="7C2D12")
    alt_fill    = PatternFill("solid", fgColor="0F172A")
    alt_fill2   = PatternFill("solid", fgColor="0F1D2E")
    default_font = Font(name="Calibri", size=10, color="CBD5E1")

    col_index = {col: j for j, col in enumerate(df.columns, 1)}

    for i, (_, row) in enumerate(df.iterrows(), 2):
        row_fill = alt_fill if i % 2 == 0 else alt_fill2
        for col in df.columns:
            j = col_index[col]
            val = row[col]
            # convert timestamps for Excel
            if isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = default_font
            cell.fill = row_fill
            if highlight_mask is not None and highlight_mask.at[i - 2, col]:
                cell.fill = orange_fill

    ws.freeze_panes = "A2"
    _auto_width(ws)


def write_output(
    df_clean: pd.DataFrame,
    report: dict,
    outlier_mask: pd.DataFrame,
    path: str,
):
    """Write a 3-sheet Excel workbook: Cleaned Data, Outliers, Cleaning Report."""
    wb = Workbook()

    # ── Sheet 1: Cleaned Data ─────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Cleaned Data"
    ws1.sheet_view.showGridLines = False
    _df_to_sheet(ws1, df_clean, highlight_mask=outlier_mask)

    # ── Sheet 2: Outliers ─────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Outliers")
    ws2.sheet_view.showGridLines = False
    outlier_rows = df_clean[outlier_mask.any(axis=1)].copy()
    if outlier_rows.empty:
        ws2.cell(row=1, column=1, value="No outliers detected.")
    else:
        _df_to_sheet(ws2, outlier_rows)

    # ── Sheet 3: Cleaning Report ───────────────────────────────────────────────
    ws3 = wb.create_sheet("Cleaning Report")
    ws3.sheet_view.showGridLines = False

    fill_hdr = PatternFill("solid", fgColor="1E293B")
    fill_row = PatternFill("solid", fgColor="0F172A")
    font_hdr = Font(bold=True, color="7DD3FC", name="Calibri", size=10)
    font_val = Font(color="CBD5E1", name="Calibri", size=10)
    font_key = Font(bold=True, color="94A3B8", name="Calibri", size=10)

    def _write(r, c, val, bold=False, header=False):
        cell = ws3.cell(row=r, column=c, value=val)
        cell.fill = fill_hdr if header else fill_row
        cell.font = font_hdr if header else (font_key if bold else font_val)
        cell.alignment = Alignment(vertical="center")
        return cell

    rows = [
        ("Metric", "Value"),
        ("Original rows", report["original_shape"][0]),
        ("Original columns", report["original_shape"][1]),
        ("Final rows", report["final_shape"][0]),
        ("Final columns", report["final_shape"][1]),
        ("Duplicates removed", report["duplicates_removed"]),
        ("Empty rows removed", report["empty_rows_removed"]),
        ("Empty columns removed", report["empty_cols_removed"]),
        ("Date columns detected", ", ".join(report["date_cols_detected"]) or "None"),
        ("Numeric cols converted", ", ".join(report["numeric_cols_converted"]) or "None"),
        ("Columns renamed", str(report["columns_renamed"]) if report["columns_renamed"] else "None"),
        ("Outliers flagged", str(report["outliers_flagged"]) if report["outliers_flagged"] else "None"),
    ]

    for r_idx, (key, val) in enumerate(rows, 1):
        header = r_idx == 1
        _write(r_idx, 1, key, bold=not header, header=header)
        _write(r_idx, 2, val, header=header)

    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 60
    ws3.freeze_panes = "A2"

    wb.save(path)


