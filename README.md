import os
import re
import pandas as pd
from collections import Counter

# -------------------------------------------------------------------
# CONFIG
# -------------------------------------------------------------------
DATA_DIR = "data"

cobra_files = [
    "Cobra-Abrams STS 2022.xlsx",
    "Cobra-Abrams STS.xlsx",
    "Cobra-ARV.xlsx",
    "Cobra-ARV30.xlsx",
    "Cobra-DE-MSHORAD I2.xlsx",
    "Cobra-M-LIDS 21.xlsx",
    "Cobra-M-LIDS.xlsx",
    "Cobra-M-SHORAD ILS YR3.xlsx",
    "Cobra-Stryker Bulgaria 150.xlsx",
    "Cobra-Stryker C4ISR - F0162.xlsx",
    "Cobra-Stryker C5ISR - F0010.xlsx",
    "Cobra-Stryker LES DO-012 F008 H325 Yr2.xlsx",
    "Cobra-Stryker LES DO-025.xlsx",
    "Cobra-Stryker SES - F0010.xlsx",
    "Cobra-Stryker SES - F0162.xlsx",
    "Cobra-XM30.xlsx",
    "John G Weekly CAP OLY 12.07.2025.xlsx",
]

PENSKE_PATH = os.path.join(DATA_DIR, "OpenPlan_Activity-Penske.xlsx")

# -------------------------------------------------------------------
# HELPERS
# -------------------------------------------------------------------

def guess_program_from_filename(fname: str) -> str:
    """
    Very simple heuristic: strip 'Cobra-' prefix and extension, keep middle.
    e.g. 'Cobra-Abrams STS.xlsx' -> 'ABRAMS STS'
    """
    base = os.path.splitext(os.path.basename(fname))[0]
    if base.lower().startswith("cobra-"):
        prog = base[6:]
    else:
        prog = base
    return prog.upper().strip()

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with simplified column names (no spaces / case)."""
    mapping = {
        c: c.strip().upper().replace(" ", "_").replace("-", "_")
        for c in df.columns
    }
    return df.rename(columns=mapping)

def get_core_cols_info(df: pd.DataFrame):
    cols = set(df.columns)
    candidates = {
        "SUB_TEAM"  : [c for c in cols if c.startswith("SUB") and "TEAM" in c],
        "COST_SET"  : [c for c in cols if "COST" in c and "SET" in c],
        "DATE"      : [c for c in cols if "DATE" in c],
        "HOURS"     : [c for c in cols if "HOUR" in c],
    }
    # pick first candidate for each, or None
    chosen = {k: (v[0] if v else None) for k, v in candidates.items()}
    return chosen

def summarize_cobra_file(path: str):
    full = os.path.join(DATA_DIR, path)
    try:
        raw = pd.read_excel(full)
    except Exception as e:
        return {
            "file": path,
            "program_guess": guess_program_from_filename(path),
            "rows": None,
            "cols": None,
            "format_key": None,
            "has_core_cols": False,
            "subteam_col": None,
            "costset_col": None,
            "date_col": None,
            "hours_col": None,
            "min_date": None,
            "max_date": None,
            "error": str(e),
        }

    df = normalize_cols(raw)
    core = get_core_cols_info(df)

    has_core = all(core.values())
    min_date = max_date = None
    if core["DATE"] is not None:
        tmp = pd.to_datetime(df[core["DATE"]], errors="coerce")
        if tmp.notna().any():
            min_date = tmp.min()
            max_date = tmp.max()

    # format key = frozenset of column names so we can cluster formats
    fmt_key = "|".join(sorted(df.columns))

    return {
        "file": path,
        "program_guess": guess_program_from_filename(path),
        "rows": len(df),
        "cols": len(df.columns),
        "format_key": fmt_key,
        "has_core_cols": has_core,
        "subteam_col": core["SUB_TEAM"],
        "costset_col": core["COST_SET"],
        "date_col": core["DATE"],
        "hours_col": core["HOURS"],
        "min_date": min_date,
        "max_date": max_date,
        "error": None,
    }

def sample_subteams(df: pd.DataFrame, subteam_col: str, n=10):
    vals = (
        df[subteam_col]
        .dropna()
        .astype(str)
        .value_counts()
        .head(n)
        .index.tolist()
    )
    return vals

def safe_read_cobra(path: str):
    """Read & normalize Cobra, returning df + core col mapping."""
    full = os.path.join(DATA_DIR, path)
    df = normalize_cols(pd.read_excel(full))
    core = get_core_cols_info(df)
    return df, core

# -------------------------------------------------------------------
# A. COBRA FILE SUMMARIES
# -------------------------------------------------------------------
cobra_summaries = [summarize_cobra_file(f) for f in cobra_files]
cobra_summary_df = pd.DataFrame(cobra_summaries)

print("\n===== Cobra file summary (one row per file) =====")
display(
    cobra_summary_df[
        [
            "file",
            "program_guess",
            "rows",
            "cols",
            "has_core_cols",
            "subteam_col",
            "costset_col",
            "date_col",
            "hours_col",
            "min_date",
            "max_date",
            "error",
        ]
    ]
)

print("\nDistinct Cobra format signatures (by columns):")
fmt_counts = Counter(cobra_summary_df["format_key"].dropna())
for i, (fmt, cnt) in enumerate(fmt_counts.items(), start=1):
    print(f"\n--- Format {i} (used by {cnt} file(s)) ---")
    cols = fmt.split("|")
    print(cols)

# -------------------------------------------------------------------
# B. PENSKE PROGRAM / SUBTEAM OVERVIEW
# -------------------------------------------------------------------
DATA = pd.read_excel(PENSKE_PATH)
DATA = normalize_cols(DATA)

print("\n===== Penske unique PROGRAM values =====")
print(DATA["PROGRAM"].dropna().unique())

penske_prog_summary = (
    DATA.groupby("PROGRAM")
    .agg(
        rows=("PROGRAM", "size"),
        unique_subteams=("SUBTEAM", lambda s: s.dropna().nunique()),
        example_subteams=("SUBTEAM", lambda s: list(s.dropna().unique())[:5]),
    )
    .reset_index()
)

print("\n===== Penske program summary =====")
display(penske_prog_summary)

# -------------------------------------------------------------------
# C. CROSS-MAPPING: Cobra SUB_TEAM vs Penske SubTeam / Cntrl_Acct
# -------------------------------------------------------------------
def map_cobra_to_penske_subteams(cobra_file: str, top_n=10):
    """
    For a given Cobra file:
      - infer program guess from filename
      - compare Cobra SUB_TEAM codes to Penske SUBTEAM / CNTRL_ACCT for that program
    """
    row = cobra_summary_df.loc[cobra_summary_df["file"] == cobra_file].iloc[0]
    if not row["has_core_cols"]:
        print(f"\n[{cobra_file}] skipped: no core cols detected.")
        return None

    df, core = safe_read_cobra(cobra_file)
    st_col = core["SUB_TEAM"]
    prog_guess = row["program_guess"]

    subteams = (
        df[st_col]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )

    penske_prog = DATA[DATA["PROGRAM"].astype(str).str.upper() == prog_guess]

    if penske_prog.empty:
        print(f"\n[{cobra_file}] No Penske rows found where PROGRAM == '{prog_guess}'.")
        return None

    penske_sub = (
        penske_prog["SUBTEAM"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
        if "SUBTEAM" in penske_prog.columns
        else []
    )

    penske_ca_cols = [c for c in penske_prog.columns if "CNTRL_ACCT" in c]
    penske_ca_vals = []
    for c in penske_ca_cols:
        penske_ca_vals.extend(
            penske_prog[c].dropna().astype(str).str.strip().unique().tolist()
        )

    # simple exact-match coverage stats
    st_set = set(subteams)
    sub_match = st_set.intersection(set(penske_sub))
    ca_match = st_set.intersection(set(penske_ca_vals))

    print(f"\n===== Mapping for Cobra file: {cobra_file} =====")
    print(f"  Program guess: {prog_guess}")
    print(f"  # Cobra SUB_TEAMs: {len(st_set)}")
    print(f"  # Penske rows for program: {len(penske_prog)}")
    print(f"  # Exact matches vs Penske SUBTEAM: {len(sub_match)}")
    print(f"  # Exact matches vs Penske control-account fields: {len(ca_match)}")

    # small sample table of a few subteams and where they match
    sample = list(st_set)[:top_n]
    rows = []
    for s in sample:
        rows.append(
            {
                "cobra_sub_team": s,
                "match_in_penske_SUBTEAM": s in penske_sub,
                "match_in_penske_CntrlAcct": s in penske_ca_vals,
            }
        )
    sample_df = pd.DataFrame(rows)
    display(sample_df)

    return {
        "cobra_file": cobra_file,
        "program_guess": prog_guess,
        "n_cobra_subteams": len(st_set),
        "n_penske_rows": len(penske_prog),
        "n_match_sub": len(sub_match),
        "n_match_ca": len(ca_match),
    }

# Run mapping for a few key programs you care about first:
mapping_summaries = []
for f in [
    "Cobra-Abrams STS 2022.xlsx",
    "Cobra-Abrams STS.xlsx",
    "Cobra-XM30.xlsx",
]:
    res = map_cobra_to_penske_subteams(f)
    if res is not None:
        mapping_summaries.append(res)

if mapping_summaries:
    print("\n===== Cross-mapping summary (Cobra -> Penske) =====")
    display(pd.DataFrame(mapping_summaries))
