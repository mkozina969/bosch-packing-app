"""
Streamlit app: Merge Bosch packing-list XLSX files directly from a GitHub repo
with EXACT user-chosen columns (order preserved).

How to run locally:
1) pip install -r requirements.txt
2) Set GitHub access in environment or .streamlit/secrets.toml (see below)
3) streamlit run streamlit_app.py

Secrets / env vars (either .env, environment, or .streamlit/secrets.toml):
  GITHUB_OWNER   = your GitHub username or org (e.g., "mkozina969")
  GITHUB_REPO    = repo name holding XLSX files (e.g., "bosch-packing-data")
  GITHUB_BRANCH  = branch (default: "main")
  GITHUB_PATH    = subfolder where XLSX live (e.g., "data/bosch_packing")
  GITHUB_TOKEN   = (optional) PAT for private repo; omit for public

Notes:
- Lists up to 50 .xlsx files from the given path, lets you pick any subset
  and produces a single merged export with your exact target headers.
- Column matching is case-insensitive and accent/whitespace-normalized.
- If a target column is left unmapped, the export fills it with blanks.
- "Refresh Git cache" clears in-app caches (in case you pushed new files).
"""

import io
import os
import re
import base64
import unicodedata
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd
import requests
import streamlit as st
from difflib import get_close_matches

# --------------------------- Config ---------------------------

def get_cfg():
    owner = os.getenv("GITHUB_OWNER") or st.secrets.get("GITHUB_OWNER", "")
    repo = os.getenv("GITHUB_REPO") or st.secrets.get("GITHUB_REPO", "")
    branch = os.getenv("GITHUB_BRANCH") or st.secrets.get("GITHUB_BRANCH", "main")
    path = os.getenv("GITHUB_PATH") or st.secrets.get("GITHUB_PATH", "data/bosch_packing")
    token = os.getenv("GITHUB_TOKEN") or st.secrets.get("GITHUB_TOKEN", "")
    return owner, repo, branch, path, token

# --------------------------- Helpers ---------------------------

def _strip_accents(s: str) -> str:
    if not isinstance(s, str):
        return s
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def normalize_header(s: str) -> str:
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


def read_tabular_from_bytes(content: bytes) -> Tuple[pd.DataFrame, str]:
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
        sheet = pick_best_sheet(xl)
        df = xl.parse(sheet, dtype=object)
        df.columns = [str(c) for c in df.columns]
        return df, sheet
    except Exception as e:
        raise RuntimeError(f"Failed reading Excel: {e}")


def auto_map_targets_to_sources(targets: List[str], all_source_cols: List[str]) -> Dict[str, str]:
    norm_to_src = {normalize_header(c): c for c in all_source_cols}
    mapping: Dict[str, str] = {}
    for t in targets:
        nt = normalize_header(t)
        if nt in norm_to_src:
            mapping[t] = norm_to_src[nt]
            continue
        candidates = list(norm_to_src.keys())
        close = get_close_matches(nt, candidates, n=1, cutoff=0.82)
        mapping[t] = norm_to_src[close[0]] if close else ""
    return mapping


def enforce_order(df: pd.DataFrame, ordered_cols: List[str]) -> pd.DataFrame:
    for col in ordered_cols:
        if col not in df.columns:
            df[col] = pd.NA
    return df[ordered_cols]

# --------------------------- GitHub API ---------------------------

API_BASE = "https://api.github.com"
RAW_BASE = "https://raw.githubusercontent.com"


def _headers(token: str) -> Dict[str, str]:
    h = {"Accept": "application/vnd.github+json"}
    if token:
        h["Authorization"] = f"Bearer {token}"
    return h

@st.cache_data(ttl=600)
def gh_list_xlsx(owner: str, repo: str, path: str, branch: str, token: str) -> List[Dict]:
    """List .xlsx files in a given path on GitHub (single directory, not recursive)."""
    if not owner or not repo or not path:
        return []
    url = f"{API_BASE}/repos/{owner}/{repo}/contents/{path}"
    r = requests.get(url, params={"ref": branch}, headers=_headers(token), timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"GitHub list error {r.status_code}: {r.text}")
    items = r.json()
    files = []
    for it in items:
        if it.get("type") == "file" and it.get("name", "").lower().endswith(".xlsx"):
            files.append({
                "name": it.get("name"),
                "path": it.get("path"),
                "download_url": it.get("download_url"),
                "sha": it.get("sha"),
                "size": it.get("size"),
            })
    files.sort(key=lambda x: x["name"].lower())
    return files

@st.cache_data(ttl=600)
def gh_fetch_bytes(owner: str, repo: str, branch: str, file_path: str, token: str) -> bytes:
    """Download file bytes. Tries download_url first; falls back to raw URL.
    Note: For Git LFS private files, raw URL with Authorization header usually works.
    """
    url = f"{API_BASE}/repos/{owner}/{repo}/contents/{file_path}"
    r = requests.get(url, params={"ref": branch}, headers=_headers(token), timeout=60)
    if r.status_code == 200 and isinstance(r.json(), dict):
        j = r.json()
        dl = j.get("download_url")
        if dl:
            r2 = requests.get(dl, headers=_headers(token), timeout=120)
            if r2.status_code == 200:
                return r2.content
        if j.get("encoding") == "base64" and j.get("content"):
            try:
                return base64.b64decode(j["content"])  # type: ignore
            except Exception:
                pass
    raw_url = f"{RAW_BASE}/{owner}/{repo}/{branch}/{file_path}"
    r3 = requests.get(raw_url, headers=_headers(token), timeout=120)
    if r3.status_code != 200:
        raise RuntimeError(f"GitHub raw download failed {r3.status_code}: {r3.text}")
    return r3.content

# --------------------------- UI ---------------------------

st.set_page_config(page_title="Bosch Packing – Git Merge", layout="wide")
st.title("Bosch Packing List Merger – from GitHub")

st.caption("Read XLSX files from a GitHub repo, map to your exact columns, and export one clean XLSX.")

with st.sidebar:
    st.header("GitHub Source")
    owner, repo, branch, path, token = get_cfg()
    owner = st.text_input("Owner / Org", owner or "")
    repo = st.text_input("Repo", repo or "")
    col1, col2 = st.columns(2)
    with col1:
        branch = st.text_input("Branch", branch or "main")
    with col2:
        path = st.text_input("Folder path", path or "data/bosch_packing")
    token_masked = st.text_input("Token (optional, for private)", value=("***" if token else ""), type="password")
    if token_masked and token_masked != "***":
        token = token_masked

    st.button("Refresh Git cache", on_click=lambda: (st.cache_data.clear()))

    st.divider()
    st.header("Export Columns")
    keep_source_file = st.checkbox("Add Source_File column", value=True)
    drop_all_blank_rows = st.checkbox("Drop rows with all targets blank", value=True)
    targets_text = st.text_area("Your EXACT headers (comma-separated)", height=90,
                                placeholder="Parcel No, Bosch Material, Quantity, Weight (kg)")

    def parse_targets(txt: str) -> List[str]:
        return [p.strip() for p in (txt or "").split(",") if p.strip()]

if not owner or not repo or not path:
    st.info("Fill in GitHub Owner, Repo, and Folder path in the sidebar.")
    st.stop()

try:
    files = gh_list_xlsx(owner, repo, path, branch, token)
except Exception as e:
    st.error(f"GitHub listing failed: {e}")
    st.stop()

if not files:
    st.warning("No .xlsx files found in the provided path.")
    st.stop()

st.success(f"Found {len(files)} XLSX files in GitHub → `{path}` on `{branch}`.")

names = [f["name"] for f in files]
selection = st.multiselect("Pick files to merge (max 50)", options=names, default=names[: min(10, len(names))])
if not selection:
    st.info("Select at least one file.")
    st.stop()
if len(selection) > 50:
    st.error("Please select at most 50 files.")
    st.stop()

# Download & read chosen files
dataframes: List[pd.DataFrame] = []
file_infos: List[Tuple[str, str]] = []

with st.spinner("Downloading and reading selected files from GitHub…"):
    for sel in selection:
        fmeta = next(x for x in files if x["name"] == sel)
        content = gh_fetch_bytes(owner, repo, branch, fmeta["path"], token)
        df, sheet = read_tabular_from_bytes(content)
        df = df.dropna(axis=1, how="all")
        if keep_source_file:
            df["Source_File"] = fmeta["name"]
        dataframes.append(df)
        file_infos.append((fmeta["name"], sheet))

with st.expander("Detected sheets per file", expanded=False):
    for fname, sheet in file_infos:
        st.write(f"**{fname}** → sheet: `{sheet}`")

# Build union of source columns
source_cols_union: List[str] = []
seen = set()
for df in dataframes:
    for c in df.columns:
        if c not in seen:
            source_cols_union.append(c)
            seen.add(c)

# Step 1 — targets
targets = parse_targets(targets_text)
if not targets:
    st.info("Enter your exact export headers in the sidebar to continue.")
    st.stop()

# Step 2 — mapping UI
st.subheader("Map targets → source columns")
st.caption("We auto-suggest; you can override.")

automap = auto_map_targets_to_sources(targets, source_cols_union)
col_map: Dict[str, str] = {}
for t in targets:
    default_choice = automap.get(t, "")
    choices = [""] + source_cols_union
    col_map[t] = st.selectbox(
        f"Source for → **{t}**", choices=choices,
        index=(choices.index(default_choice) if default_choice in choices else 0)
    )

# Step 3 — build
st.subheader("Build merged export")
build = st.button("Build & Download XLSX", type="primary")

if build:
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
