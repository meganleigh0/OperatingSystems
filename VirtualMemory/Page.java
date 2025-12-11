# =========================================================
# EVMS Deck Generator – Multi-Program, Clean Layout
# =========================================================
import os
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ---------------------------------------------------------
# CONFIG
# ---------------------------------------------------------
DATA_DIR   = "data"          # where Cobra exports live
OUTPUT_DIR = "EVMS_Output"   # where xlsx and pptx will go

# If you have a template deck, point to it here; otherwise will use a blank Presentation()
THEME_PATH = os.path.join(DATA_DIR, "Theme.pptx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Map “program name in outputs” -> “Cobra Excel file name”
PROGRAM_CONFIG = {
    "Abrams_STS_2022": "Cobra-Abrams STS 2022.xlsx",
    "Abrams_STS":      "Cobra-Abrams STS.xlsx",
    "ARV":             "Cobra-ARV.xlsx",
    "ARV30":           "Cobra-ARV30.xlsx",
    "DE_MSHORAD_I2":   "Cobra-DE-MSHORAD I2.xlsx",
    "M-LIDS_21":       "Cobra-M-LIDS 21.xlsx",
    "M-LIDS":          "Cobra-M-LIDS.xlsx",
    "M-SHORAD_ILS_YR3":"Cobra-M-SHORAD ILS YR3.xlsx",
    "Stryker_Bulgaria_150": "Cobra-Stryker Bulgaria 150.xlsx",
    "Stryker_C4ISR_F0162":  "Cobra-Stryker C4ISR -F0162.xlsx",
    "Stryker_CSISR_F0010":  "Cobra-Stryker CSISR - F0010.xlsx",
    "Stryker_LES_DO012_F008_Yr2": "Cobra-Stryker LES DO-012 F008 H325 Yr2.xlsx",
    "Stryker_LES_DO025":          "Cobra-Stryker LES DO-025.xlsx",
    "Stryker_SES_F0010":          "Cobra-Stryker SES - F0010.xlsx",
    "Stryker_SES_F0162":          "Cobra-Stryker SES - F0162.xlsx",
    "XM30":                       "Cobra-XM30.xlsx",
    "JohnG_CAP_OLY":              "John G Weekly CAP OLY 12.07.2025.xlsx",
}

# EVMS cost sets – adjust if your Cobra exports use different names
COSTSET_BCWS = "BCWS"
COSTSET_BCWP = "BCWP"
COSTSET_ACWP = "ACWP"

# Y-axis range for EV plot
YMIN, YMAX = 0.75, 1.25

# ---------------------------------------------------------
# Helpers – Cobra normalization + EVMS calculations
# ---------------------------------------------------------
def load_cobra(path):
    """Load first sheet of a Cobra export."""
    return pd.read_excel(path)

def normalize_cobra(df_raw):
    """
    Try to normalize a Cobra DataFrame to standard columns:
    SUBTEAM, COSTSET, DATE, HOURS.
    Returns None if not possible (we'll log and skip).
    """
    df = df_raw.copy()
    # Standardize column names
    df.columns = [str(c).strip() for c in df.columns]

    colmap = {}
    for c in df.columns:
        c_upper = c.upper().replace(" ", "").replace("_", "")
        if c_upper in ["SUBTEAM", "SUB_TEAM"]:
            colmap[c] = "SUBTEAM"
        elif c_upper in ["COSTSET", "COST-SET", "COST_SET"]:
            colmap[c] = "COSTSET"
        elif c_upper.startswith("DATE"):
            colmap[c] = "DATE"
        elif c_upper in ["HOURS", "HRS"]:
            colmap[c] = "HOURS"

    df = df.rename(columns=colmap)
    required = ["SUBTEAM", "COSTSET", "DATE", "HOURS"]

    if not all(c in df.columns for c in required):
        missing = [c for c in required if c not in df.columns]
        raise ValueError(f"Could not normalize Cobra file – missing logical columns: {missing}")

    # Basic cleaning
    df = df[required].copy()
    df["DATE"] = pd.to_datetime(df["DATE"])
    df["HOURS"] = pd.to_numeric(df["HOURS"], errors="coerce").fillna(0.0)
    df["SUBTEAM"] = df["SUBTEAM"].astype(str).str.strip()

    return df

def compute_ev_timeseries(df_norm):
    """
    Aggregates BCWS, BCWP, ACWP by week/month and computes
    Monthly & Cumulative CPI/SPI.
    """
    # Filter to cost sets we care about
    mask = df_norm["COSTSET"].isin([COSTSET_BCWS, COSTSET_BCWP, COSTSET_ACWP])
    df = df_norm[mask].copy()

    if df.empty:
        raise ValueError("No BCWS/BCWP/ACWP rows found after filtering COSTSET")

    # Map costset into separate numeric columns
    pivot = (
        df.pivot_table(
            index="DATE",
            columns="COSTSET",
            values="HOURS",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
        .sort_values("DATE")
    )

    for col in [COSTSET_BCWS, COSTSET_BCWP, COSTSET_ACWP]:
        if col not in pivot.columns:
            pivot[col] = 0.0

    pivot = pivot.set_index("DATE").resample("W-MON").sum()  # weekly on Mondays

    # Monthly indices (per-period)
    bcws = pivot[COSTSET_BCWS]
    bcwp = pivot[COSTSET_BCWP]
    acwp = pivot[COSTSET_ACWP]

    # Avoid divide-by-zero
    cpi_m = np.where(acwp > 0, bcwp / acwp, np.nan)
    spi_m = np.where(bcws > 0, bcwp / bcws, np.nan)

    # Cumulative series – THIS is what you care about being smooth
    bcws_cum = bcws.cumsum()
    bcwp_cum = bcwp.cumsum()
    acwp_cum = acwp.cumsum()

    cpi_c = np.where(acwp_cum > 0, bcwp_cum / acwp_cum, np.nan)
    spi_c = np.where(bcws_cum > 0, bcwp_cum / bcws_cum, np.nan)

    evdf = pd.DataFrame(
        {
            "DATE": pivot.index,
            "BCWS": bcws.values,
            "BCWP": bcwp.values,
            "ACWP": acwp.values,
            "CPI_M": cpi_m,
            "SPI_M": spi_m,
            "CPI_C": cpi_c,
            "SPI_C": spi_c,
        }
    )

    # Small outlier guard for plotting (does NOT change metrics)
    evdf["CPI_M_clip"] = evdf["CPI_M"].clip(YMIN, YMAX)
    evdf["SPI_M_clip"] = evdf["SPI_M"].clip(YMIN, YMAX)
    evdf["CPI_C_clip"] = evdf["CPI_C"].clip(YMIN, YMAX)
    evdf["SPI_C_clip"] = evdf["SPI_C"].clip(YMIN, YMAX)

    return evdf

def get_status_dates(evdf):
    """Current (CTD) = last date; LSD = previous."""
    dates = sorted(evdf["DATE"].unique())
    if not dates:
        raise ValueError("No EV dates available")
    curr = dates[-1]
    prev = dates[-2] if len(dates) > 1 else dates[-1]
    return curr, prev

def extract_program_metrics(evdf, curr_date, prev_date):
    """Build small SPI/CPI/BEI table for CTD and LSD (BEI = NaN placeholder)."""
    row_ctd = evdf.loc[evdf["DATE"] == curr_date].iloc[-1]
    row_lsd = evdf.loc[evdf["DATE"] == prev_date].iloc[-1]

    metrics = pd.DataFrame(
        {
            "Metric": ["SPI", "CPI", "BEI"],
            "CTD": [row_ctd["SPI_C"], row_ctd["CPI_C"], np.nan],
            "LSD": [row_lsd["SPI_C"], row_lsd["CPI_C"], np.nan],
            "Comments / Root Cause & Corrective Actions": ["" for _ in range(3)],
        }
    )
    return metrics

def build_subteam_labor_table(df_norm):
    """BAC/EAC/VAC per SUBTEAM + Program Manpower summary."""
    # Treat BCWS as BAC proxy, and ACWP + ETC as EAC proxy if you have it.
    # For now: BAC = sum BCWS, EAC = sum ACWP (you can refine later).
    mask = df_norm["COSTSET"].isin([COSTSET_BCWS, COSTSET_ACWP])
    tmp = df_norm[mask].copy()

    pivot = (
        tmp.pivot_table(
            index="SUBTEAM",
            columns="COSTSET",
            values="HOURS",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )

    if COSTSET_BCWS not in pivot.columns:
        pivot[COSTSET_BCWS] = 0.0
    if COSTSET_ACWP not in pivot.columns:
        pivot[COSTSET_ACWP] = 0.0

    pivot["BAC"] = pivot[COSTSET_BCWS]
    pivot["EAC"] = pivot[COSTSET_ACWP]
    pivot["VAC"] = pivot["BAC"] - pivot["EAC"]

    labor_df = pivot[["SUBTEAM", "BAC", "EAC", "VAC"]].copy()
    labor_df = labor_df.sort_values("SUBTEAM").reset_index(drop=True)
    labor_df["Comments / Root Cause & Corrective Actions"] = ""

    # Program manpower summary
    demand_hours = labor_df["BAC"].sum()
    actual_hours = labor_df["EAC"].sum()
    pct_var = (actual_hours / demand_hours) if demand_hours > 0 else np.nan

    manpower_df = pd.DataFrame(
        {
            "Demand Hours": [demand_hours],
            "Actual Hours": [actual_hours],
            "% Var": [pct_var],
            "Next Mo BCWS Hours": [0.0],
            "Next Mo ETC Hours": [0.0],
            "Comments / Root Cause & Corrective Actions": [""],
        }
    )

    return labor_df, manpower_df

# ---------------------------------------------------------
# Helpers – PowerPoint construction
# ---------------------------------------------------------
def load_template():
    if os.path.exists(THEME_PATH):
        return Presentation(THEME_PATH)
    return Presentation()

def get_layout(prs, preferred_names):
    """
    Try to find a layout by name substring; fall back to the first non-bumper layout.
    This avoids accidentally picking up the 'Edit Bumper Sticker' layout.
    """
    for layout in prs.slide_layouts:
        name = layout.name.lower()
        if any(p.lower() in name for p in preferred_names):
            return layout
    # Fallback: simple title + content, avoiding anything with 'bumper' in the name
    for layout in prs.slide_layouts:
        if "bumper" not in layout.name.lower():
            return layout
    return prs.slide_layouts[0]

def add_ev_plot_slide(prs, program, evdf, metrics_tbl):
    slide = prs.slides.add_slide(get_layout(prs, ["Title", "Content"]))
    title = slide.shapes.title
    title.text = f"{program} EVMS Trend Overview"

    # Add chart as picture from Matplotlib
    fig, ax = plt.subplots(figsize=(7, 4))

    # Colored bands
    ax.axhspan(YMIN, 0.9, facecolor="#ffcccc", alpha=0.5)   # red
    ax.axhspan(0.9, 0.95, facecolor="#fff2cc", alpha=0.5)   # yellow
    ax.axhspan(0.95, 1.05, facecolor="#c6efce", alpha=0.5)  # green
    ax.axhspan(1.05, YMAX, facecolor="#cfe2ff", alpha=0.5)  # blue

    ax.scatter(evdf["DATE"], evdf["CPI_M"], s=10, label="Monthly CPI", color="gold")
    ax.scatter(evdf["DATE"], evdf["SPI_M"], s=10, label="Monthly SPI", color="black")
    ax.plot(evdf["DATE"], evdf["CPI_C"], label="Cumulative CPI", linewidth=2, color="blue")
    ax.plot(evdf["DATE"], evdf["SPI_C"], label="Cumulative SPI", linewidth=2, color="gray")

    ax.set_ylim(YMIN, YMAX)
    ax.set_xlabel("Month")
    ax.set_ylabel("EV Indices")
    ax.legend(fontsize=8)
    ax.grid(True, axis="y", alpha=0.3)

    fig.tight_layout()
    img_path = os.path.join(OUTPUT_DIR, f"{program}_ev_plot.png")
    fig.savefig(img_path, dpi=200)
    plt.close(fig)

    # Place image on slide
    left = Inches(0.5)
    top = Inches(1.5)
    slide.shapes.add_picture(img_path, left, top, height=Inches(3.5))

    # Metric table on right
    rows, cols = metrics_tbl.shape
    tbl_left = Inches(6.0)
    tbl_top = Inches(1.5)
    tbl_width = Inches(3.5)
    tbl_height = Inches(1.0 + 0.3 * rows)

    table_shape = slide.shapes.add_table(rows + 1, cols, tbl_left, tbl_top, tbl_width, tbl_height)
    table = table_shape.table

    # Header row
    for j, col in enumerate(metrics_tbl.columns):
        table.cell(0, j).text = col

    # Data rows
    for i in range(rows):
        for j, col in enumerate(metrics_tbl.columns):
            val = metrics_tbl.iloc[i, j]
            if isinstance(val, float):
                txt = "" if np.isnan(val) else f"{val:.3f}"
            else:
                txt = str(val)
            table.cell(i + 1, j).text = txt

    # Column widths: Metric slightly wider; Comments much wider
    metric_col_width = Inches(1.0)
    other_col_width = Inches(0.8)
    comments_col_width = Inches(2.0)

    for j, col in enumerate(metrics_tbl.columns):
        if j == 0:  # Metric
            table.columns[j].width = metric_col_width
        elif "Comments" in col:
            table.columns[j].width = comments_col_width
        else:
            table.columns[j].width = other_col_width

    return slide

def add_labor_manpower_slides(prs, program, labor_df, manpower_df, page_size=15):
    """
    Subteam Labor & Manpower slides, 15 subteams per slide.
    Program Manpower table sits lower so it doesn't overlap.
    """
    n = len(labor_df)
    pages = max(1, int(np.ceil(n / page_size)))

    for p in range(pages):
        slide = prs.slides.add_slide(get_layout(prs, ["Title"]))
        title = slide.shapes.title
        label = "" if pages == 1 else f" (Page {p+1})"
        title.text = f"{program} EVMS Detail – Sub Team Labor & Manpower{label}"

        chunk = labor_df.iloc[p*page_size : (p+1)*page_size]

        # Main subteam table
        rows, cols = chunk.shape
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9.0)
        height = Inches(0.3 * (rows + 1))

        tshape = slide.shapes.add_table(rows + 1, cols, left, top, width, height)
        table = tshape.table

        # Headers
        for j, col in enumerate(chunk.columns):
            table.cell(0, j).text = col

        # Data
        for i in range(rows):
            for j, col in enumerate(chunk.columns):
                val = chunk.iloc[i, j]
                if isinstance(val, float):
                    if "Comments" in col:
                        txt = ""
                    elif "BAC" in col or "EAC" in col or "VAC" in col:
                        txt = f"{val:,.1f}"
                    else:
                        txt = f"{val:.3f}"
                else:
                    txt = str(val)
                table.cell(i + 1, j).text = txt

        # Column widths – make comments column wider
        for j, col in enumerate(chunk.columns):
            if "Comments" in col:
                table.columns[j].width = Inches(3.0)
            else:
                table.columns[j].width = Inches(1.2)

        # Program manpower table – pushed down to avoid overlap
        mp_rows, mp_cols = manpower_df.shape
        mp_left = Inches(0.5)
        mp_top = top + height + Inches(0.4)  # this extra 0.4 moves it lower
        mp_width = Inches(9.0)
        mp_height = Inches(0.8)

        mp_shape = slide.shapes.add_table(mp_rows + 1, mp_cols, mp_left, mp_top, mp_width, mp_height)
        mp_tbl = mp_shape.table

        # Headers
        for j, col in enumerate(manpower_df.columns):
            mp_tbl.cell(0, j).text = col

        # Data
        for i in range(mp_rows):
            for j, col in enumerate(manpower_df.columns):
                val = manpower_df.iloc[i, j]
                if isinstance(val, float):
                    if "Comments" in col:
                        txt = ""
                    elif "% Var" in col:
                        txt = f"{val:.2%}"
                    else:
                        txt = f"{val:,.1f}"
                else:
                    txt = str(val)
                mp_tbl.cell(i + 1, j).text = txt

        # Column widths – comments widest
        for j, col in enumerate(manpower_df.columns):
            if "Comments" in col:
                mp_tbl.columns[j].width = Inches(3.0)
            else:
                mp_tbl.columns[j].width = Inches(1.2)

# ---------------------------------------------------------
# Main program handler
# ---------------------------------------------------------
def process_program(program_name, cobra_file):
    cobra_path = os.path.join(DATA_DIR, cobra_file)
    if not os.path.exists(cobra_path):
        print(f">> Skipping {program_name} – file not found: {cobra_path}")
        return

    print(f"\n=== Processing {program_name} from {os.path.basename(cobra_path)} ===")
    df_raw = load_cobra(cobra_path)

    try:
        cobra = normalize_cobra(df_raw)
    except ValueError as e:
        print(f"!! Error for {program_name}: {e}")
        return

    # EV timeseries
    evdf = compute_ev_timeseries(cobra)
    curr_date, prev_date = get_status_dates(evdf)
    metrics_tbl = extract_program_metrics(evdf, curr_date, prev_date)
    labor_df, manpower_df = build_subteam_labor_table(cobra)

    print(f"CTD date: {curr_date.date()}, LSD date: {prev_date.date()}")
    print(metrics_tbl[["Metric", "CTD", "LSD"]])

    # Build deck
    prs = load_template()
    add_ev_plot_slide(prs, program_name, evdf, metrics_tbl)
    add_labor_manpower_slides(prs, program_name, labor_df, manpower_df, page_size=15)

    # Save outputs
    tables_xlsx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Tables.xlsx")
    with pd.ExcelWriter(tables_xlsx, engine="xlsxwriter") as writer:
        evdf.to_excel(writer, sheet_name="EV_Series", index=False)
        metrics_tbl.to_excel(writer, sheet_name="Program_Metrics", index=False)
        labor_df.to_excel(writer, sheet_name="Subteam_Labor", index=False)
        manpower_df.to_excel(writer, sheet_name="Program_Manpower", index=False)

    out_pptx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Deck.pptx")
    prs.save(out_pptx)

    print(f"✓ Saved tables: {tables_xlsx}")
    print(f"✓ Saved deck:   {out_pptx}")

# ---------------------------------------------------------
# Run for all configured programs
# ---------------------------------------------------------
for program, cobra_file in PROGRAM_CONFIG.items():
    try:
        process_program(program, cobra_file)
    except Exception as e:
        print(f"!! Error for {program}: {e}")

print("\nALL PROGRAM EVMS DECKS COMPLETE ✓")