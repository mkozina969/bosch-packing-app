# streamlit_app.py
"""
Bosch Merger – dual mode (Packing Lists + Invoices)
Upload 1–50 .xlsx files and export a single XLSX per run.

Modes:
- Packing lists → export only: ProductNumber, DeliveredQuantity, PkgIdentNumber_2
- Invoices      → export only: ProductNumber, UnitPrice, Quantity, DesAdvRef_Date
  (If a requested column appears twice in the spec, we auto-deduplicate headers.)
"""

import io
import re
import unicodedata
from datetime import datetime
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st

# --------------------------- Normalization helpers ---------------------------

def _strip_accents(s: str) -> str:
    if not isinstance(s, str):
        return s
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def _norm_key(s: str) -> str:
    """Lenient header normalization for matching across subtle naming differences."""
    if s is None:
        return ""
    s = str(s)
    s = _strip_accents(s).lower()
    s = s.replace("\n", " ")
    s = re.sub(r"[\s\-]+", "_", s)      # spaces & dashes → underscore
    s = re.sub(r"[^a-z0-9_]+", "", s)     # keep only a-z0-9_
    s = re.sub(r"_+", "_", s).strip("_") # collapse underscores
    return s


# --------------------------- Table detection ---------------------------

def _find_header_row_and_map(raw: pd.DataFrame, target_headers: List[str]) -> Tuple[Optional[int], Dict[str, str]]:
    """Return (header_row_index, mapping target→original_column_name) if all targets found on some row.
    Matching is done on normalized keys.
    """
    # for detection we deduplicate targets by normalized key
    targets_norm = [_norm_key(t) for t in target_headers]
    uniq_targets_norm = list(dict.fromkeys(targets_norm))  # preserve order

    for i, row in raw.iterrows():
        values = [v if v is not None else "" for v in row.values]
        norm_map: Dict[str, str] = {}
        for v in values:
            k = _norm_key(v)
            if k and k not in norm_map:
                norm_map[k] = str(v)
        if all(tk in norm_map for tk in uniq_targets_norm):
            # Build mapping for original (non-deduped) requested names → original column header
            mapping: Dict[str, str] = {}
            for t, tk in zip(target_headers, targets_norm):
                mapping[t] = norm_map[tk]
            return i, mapping
    return None, {}


def _extract_table_any_sheet(xl: pd.ExcelFile, target_headers: List[str]) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Try each sheet until we find one whose row contains all target headers.
    Returns (table_df, mapping target→original)
    """
    for name in xl.sheet_names:
        try:
            raw = xl.parse(name, header=None, dtype=object)
        except Exception:
            continue
        hdr_idx, mapping = _find_header_row_and_map(raw, target_headers)
        if hdr_idx is not None:
            tbl = raw.iloc[hdr_idx + 1 :].copy()
            # set real headers from that row
            tbl.columns = [str(c) for c in raw.iloc[hdr_idx].values]
            # drop fully-empty columns
            tbl = tbl.dropna(axis=1, how="all")
            return tbl, mapping
    return pd.DataFrame(), {}


def _extract_subset_from_file(file, targets: List[str]) -> pd.DataFrame:
    xl = pd.ExcelFile(file)
    tbl, mapping = _extract_table_any_sheet(xl, targets)

    # Prepare output columns (dedupe requested targets in header, but keep order)
    out_cols: List[str] = []
    seen = set()
    for t in targets:
        if t not in seen:
            out_cols.append(t)
            seen.add(t)
    out = pd.DataFrame(columns=out_cols)

    # Fill columns from mapping
    for t in out_cols:
        src_col = mapping.get(t, None)
        if src_col and src_col in tbl.columns:
            out[t] = tbl[src_col]
        else:
            out[t] = pd.NA

    out["Source_File"] = getattr(file, "name", "uploaded.xlsx")
    return out


# --------------------------- UI ---------------------------

st.set_page_config(page_title="Bosch Packing & Invoices Merger", layout="wide")
st.title("Bosch Merger – Packing Lists & Invoices")

mode = st.radio("Choose mode", ["Packing lists", "Invoices"], horizontal=True)

if mode == "Packing lists":
    TARGETS = ["ProductNumber", "DeliveredQuantity", "PkgIdentNumber_2"]
else:
    TARGETS = ["ProductNumber", "UnitPrice", "Quantity", "DesAdvRef_Date"]

with st.sidebar:
    st.header("Settings")
    keep_source_file = st.checkbox("Add Source_File column", value=True)
    if mode == "Packing lists":
        drop_blank = st.checkbox("Drop rows with blank DeliveredQuantity", value=False)
    else:
        drop_blank = st.checkbox("Drop rows with blank Quantity", value=False)
        parse_dates = st.checkbox("Parse DesAdvRef_Date as date", value=True)

uploaded = st.file_uploader(
    f"Upload {mode} .xlsx files (1–50)", type=["xlsx"], accept_multiple_files=True
)

if not uploaded:
    st.info("Upload 1–50 .xlsx files to begin.")
    st.stop()
if len(uploaded) > 50:
    st.error("Please upload at most 50 files.")
    st.stop()

# Extract from each file
dfs: List[pd.DataFrame] = []
with st.spinner("Reading and extracting tables…"):
    for up in uploaded:
        part = _extract_subset_from_file(up, TARGETS)
        if not keep_source_file and "Source_File" in part.columns:
            part = part.drop(columns=["Source_File"]) 
        dfs.append(part)

merged = pd.concat(dfs, ignore_index=True)

# Typing / cleaning per mode
if mode == "Packing lists":
    if "DeliveredQuantity" in merged.columns:
        merged["DeliveredQuantity"] = pd.to_numeric(merged["DeliveredQuantity"], errors="coerce")
else:
    if "UnitPrice" in merged.columns:
        merged["UnitPrice"] = pd.to_numeric(merged["UnitPrice"], errors="coerce")
    if "Quantity" in merged.columns:
        merged["Quantity"] = pd.to_numeric(merged["Quantity"], errors="coerce")
    if "DesAdvRef_Date" in merged.columns and 'parse_dates' in locals() and parse_dates:
        # Try to parse common date formats; leave as-is if parsing fails
        merged["DesAdvRef_Date"] = pd.to_datetime(merged["DesAdvRef_Date"], errors="coerce")

# Drop blank rows option
if drop_blank:
    key_col = "DeliveredQuantity" if mode == "Packing lists" else "Quantity"
    if key_col in merged.columns:
        merged = merged[merged[key_col].notna()].reset_index(drop=True)

st.success(f"Merged rows: {len(merged):,}")
st.dataframe(merged.head(50))

# Download
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as xw:
    merged.to_excel(xw, index=False, sheet_name="Merged")
buf.seek(0)
fname = f"bosch_{'packing' if mode=='Packing lists' else 'invoices'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
st.download_button("Download XLSX", data=buf, file_name=fname,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
