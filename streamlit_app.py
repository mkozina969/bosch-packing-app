# streamlit_app.py
"""
Bosch Packing – minimal merge (auto-extract 3 columns)
Upload 1–50 .xlsx pakirnih lista → app nađe tablicu, izvuče:
  ProductNumber, DeliveredQuantity, PkgIdentNumber_2
i spoji sve u jedan XLSX.
"""

import io
import os
import re
import unicodedata
from datetime import datetime
from typing import List, Tuple

import pandas as pd
import streamlit as st

# ---------- helpers ----------
def _strip_accents(s: str) -> str:
    if not isinstance(s, str):
        return s
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = _strip_accents(s).replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

TARGET_COLS = ["ProductNumber", "DeliveredQuantity", "PkgIdentNumber_2"]

def find_header_row(df: pd.DataFrame) -> int | None:
    """Nađi red koji sadrži SVE target header-e (točno ime, case sensitive)."""
    for i, row in df.iterrows():
        vals = [str(v) if v is not None else "" for v in row.values]
        if all(t in vals for t in TARGET_COLS):
            return i
    return None

def extract_table_any_sheet(xl: pd.ExcelFile) -> pd.DataFrame | None:
    """Probaj svaku tab-icu: pronađi header redak i vrati tablicu od idućeg retka naniže."""
    for name in xl.sheet_names:
        try:
            raw = xl.parse(name, header=None, dtype=object)
        except Exception:
            continue
        hdr = find_header_row(raw)
        if hdr is not None:
            tbl = raw.iloc[hdr + 1 :].copy()
            tbl.columns = [str(c) for c in raw.iloc[hdr].values]
            # očisti prazne kolone
            tbl = tbl.dropna(axis=1, how="all")
            return tbl
    return None

def extract_three_cols_from_file(file) -> pd.DataFrame:
    """Vrati DataFrame sa samo 3 target kolone + Source_File (ako postoje)."""
    xl = pd.ExcelFile(file)
    tbl = extract_table_any_sheet(xl)
    if tbl is None:
        # fallback: prazno s target kolonama
        return pd.DataFrame(columns=TARGET_COLS + ["Source_File"])
    out = pd.DataFrame()
    for col in TARGET_COLS:
        out[col] = tbl[col] if col in tbl.columns else pd.NA
    # tipiziraj DeliveredQuantity
    if "DeliveredQuantity" in out.columns:
        out["DeliveredQuantity"] = pd.to_numeric(out["DeliveredQuantity"], errors="coerce")
    out["Source_File"] = getattr(file, "name", "uploaded.xlsx")
    # izbaci redove bez ProductNumber-a
    out = out[out["ProductNumber"].notna()].reset_index(drop=True)
    return out

# ---------- UI ----------
st.set_page_config(page_title="Bosch Packing – 3-col merge", layout="wide")
st.title("Bosch Packing – ProductNumber, DeliveredQuantity, PkgIdentNumber_2")

st.caption("Upload 1–50 .xlsx pakirnih lista → automatski ekstrakt 3 kolone → jedan XLSX.")

with st.sidebar:
    keep_source_file = st.checkbox("Add Source_File column", value=True)
    drop_blank_qty = st.checkbox("Drop rows where DeliveredQuantity is blank", value=False)

uploaded = st.file_uploader(
    "Upload Bosch packing-list .xlsx files (1–50)", type=["xlsx"], accept_multiple_files=True
)

if not uploaded:
    st.info("Uploadaj 1–50 .xlsx datoteka.")
    st.stop()

if len(uploaded) > 50:
    st.error("Maksimalno 50 datoteka odjednom.")
    st.stop()

dfs: List[pd.DataFrame] = []
with st.spinner("Reading and extracting…"):
    for up in uploaded:
        df = extract_three_cols_from_file(up)
        if not keep_source_file and "Source_File" in df.columns:
            df = df.drop(columns=["Source_File"])
        dfs.append(df)

merged = pd.concat(dfs, ignore_index=True)

if drop_blank_qty and "DeliveredQuantity" in merged.columns:
    merged = merged[merged["DeliveredQuantity"].notna()].reset_index(drop=True)

st.success(f"Extracted rows: {len(merged):,}")
st.dataframe(merged.head(50))

# download
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as xw:
    merged.to_excel(xw, index=False, sheet_name="Merged")
buf.seek(0)
fname = f"bosch_packing_3cols_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
st.download_button("Download XLSX", data=buf, file_name=fname,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
