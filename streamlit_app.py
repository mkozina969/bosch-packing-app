# streamlit_app.py
"""
Bosch Packing List Merger – local upload version
Upload 1–50 .xlsx datoteka (packing list), mapiraj kolone i preuzmi jedan merged XLSX.
"""

import io
import re
import unicodedata
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from difflib import get_close_matches


# --------------------------- Helpers ---------------------------

def _strip_accents(s: str) -> str:
    if not isinstance(s, str):
        return s
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def normalize_header(s: str) -> str:
    """Lowercase, trim, remove accents, collapse spaces, keep basic symbols."""
    if s is None:
        return ""
    s = str(s)
    s = _strip_accents(s)
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    s = re.sub(r"[^a-z0-9 _\-/]", "", s)
    s = s.replace("  ", " ")
    return s


def pick_best_sheet(xl: pd.ExcelFile) -> str:
    """Pick sheet with most non-empty cells."""
    best_sheet = xl.sheet_names[0]
    best_score = -1
    for name in xl.sheet_names:
        try:
            df = xl.parse(name, dtype=object)
        except Exception:
            continue
        non_empty = int(df.notna().sum().sum())
        if non_empty > best_score:
            best_score = non_empty
            best_sheet = name
    return best_sheet


def read_tabular(file) -> Tuple[pd.DataFrame, str]:
    """Read best-looking sheet and return DataFrame + sheet name."""
    xl = pd.ExcelFile(file)
    sheet = pick_best_sheet(xl)
    df = xl.parse(sheet, dtype=object)
    df.columns = [str(c) for c in df.columns]
    return df, sheet


def auto_map_targets_to_sources(targets: List[str], all_source_cols: List[str]) -> Dict[str, str]:
    """Exact normalized match first, then close match via difflib."""
    norm_to_src = {normalize_header(c): c for c in all_source_cols}
    mapping: Dict[str, str] = {}
    for t in targets:
        nt = normalize_header(t)
        if nt in norm_to_src:
            mapping[t] = norm_to_src[nt]
            continue
        close = get_close_matches(nt, list(norm_to_src.keys()), n=1, cutoff=0.82)
        mapping[t] = norm_to_src[close[0]] if close else ""
    return mapping


def enforce_order(df: pd.DataFrame, ordered_cols: List[str]) -> pd.DataFrame:
    for col in ordered_cols:
        if col not in df.columns:
            df[col] = pd.NA
    return df[ordered_cols]


# --------------------------- UI ---------------------------

st.set_page_config(page_title="Bosch Packing List Merger", layout="wide")
st.title("Bosch Packing List Merger – XLSX → one export")

st.caption(
    "Upload 1–50 .xlsx packing list datoteka, definiraj točne export kolone, mapiraj izvore → ciljeve i preuzmi čisti merged XLSX."
)

with st.sidebar:
    st.header("Settings")
    keep_source_file = st.checkbox("Add Source_File column", value=True, help="Dodaj ime originalne datoteke svakoj liniji.")
    drop_all_blank_rows = st.checkbox(
        "Drop rows where all mapped targets are blank", value=True,
        help="Ako su sve ciljane kolone prazne u retku, odbaci taj red."
    )

    st.subheader("Your EXACT target headers (comma-separated)")
    targets_text = st.text_area(
        "Primjer: Parcel No, Bosch Material, Quantity, Weight (kg)",
        height=90,
        placeholder="Upiši točne nazive kolona, odvojene zarezom (redoslijed = redoslijed u exportu)",
    )

    def parse_targets(txt: str) -> List[str]:
        return [p.strip() for p in (txt or "").split(",") if p.strip()]


uploaded = st.file_uploader(
    "Upload Bosch packing-list .xlsx files (1–50)", type=["xlsx"], accept_multiple_files=True
)

if not uploaded:
    st.info("Uploadaj 1–50 .xlsx datoteka za početak.")
    st.stop()

if len(uploaded) > 50:
    st.error("Maksimalno 50 datoteka odjednom.")
    st.stop()

# Read all files
dataframes: List[pd.DataFrame] = []
file_infos: List[Tuple[str, str]] = []

with st.spinner("Reading Excel files…"):
    for up in uploaded:
        df, sheet = read_tabular(up)
        df = df.dropna(axis=1, how="all")  # drop fully empty columns
        if keep_source_file:
            df["Source_File"] = up.name
        dataframes.append(df)
        file_infos.append((up.name, sheet))

st.success(f"Loaded {len(uploaded)} files.")
with st.expander("Detected sheets per file", expanded=False):
    for fname, sheet in file_infos:
        st.write(f"**{fname}** → sheet: `{sheet}`")

# Union of source columns (in first-seen order)
source_cols_union: List[str] = []
seen = set()
for df in dataframes:
    for c in df.columns:
        if c not in seen:
            source_cols_union.append(c)
            seen.add(c)

# Step 1: target headers
targets = parse_targets(targets_text)
if not targets:
    st.info("Unesi točne ciljne kolone u sidebaru (comma-separated) za nastavak.")
    st.stop()

# Step 2: mapping UI
st.subheader("Map targets → source columns")
st.caption("Auto-suggest je ponuđen, možeš ručno promijeniti izbor.")

automap = auto_map_targets_to_sources(targets, source_cols_union)
col_map: Dict[str, str] = {}

for t in targets:
    default_choice = automap.get(t, "")
    choices = [""] + source_cols_union
    col_map[t] = st.selectbox(
        f"Source for → **{t}**",
        options=choices,
        index=(choices.index(default_choice) if default_choice in choices else 0),
    )

# Step 3: build & download
st.subheader("Build merged export")
if st.button("Build & Download XLSX", type="primary"):
    with st.spinner("Merging and formatting…"):
        merged = pd.concat(dataframes, ignore_index=True)

        out_df = pd.DataFrame()
        for t in targets:
            src = col_map.get(t, "")
            out_df[t] = merged[src] if src and src in merged.columns else pd.NA

        if drop_all_blank_rows:
            mask_all_blank = out_df.apply(lambda r: all((pd.isna(v) or str(v).strip() == "") for v in r), axis=1)
            out_df = out_df.loc[~mask_all_blank].reset_index(drop=True)

        if keep_source_file and "Source_File" in merged.columns:
            out_df["Source_File"] = merged["Source_File"]

        ordered = targets + (["Source_File"] if keep_source_file else [])
        out_df = enforce_order(out_df, ordered)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="Merged")
        buffer.seek(0)

        fname = f"bosch_packing_merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    st.success(f"Done. {len(out_df):,} rows in export.")
    st.download_button(
        "Download merged XLSX",
        data=buffer,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
