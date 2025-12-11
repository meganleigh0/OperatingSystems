# =========================================================
# EVMS Deck Generator – Standard-format Cobra files only
# (Abrams_STS_2022, Abrams_STS, XM30, ARV, ARV30,
#  Stryker_Bulgaria_150, Stryker_CSISR_F0010, Stryker_SES_F0010)
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
DATA_DIR   = "data"          # Folder with Cobra exports
OUTPUT_DIR = "EVMS_Output"   # Folder for tables + decks
THEME_PATH = os.path.join(DATA_DIR, "Theme.pptx")  # your GDLS template, optional

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ONLY programs whose Cobra files already match the “standard” format
PROGRAM_CONFIG = {
    "Abrams_STS_2022":          "Cobra-Abrams STS 2022.xlsx",
    "Abrams_STS":               "Cobra-Abrams STS.xlsx",
    "ARV":                      "Cobra-ARV.xlsx",
    "ARV30":                    "Cobra-ARV30.xlsx",
    "Stryker_Bulgaria_150":     "Cobra-Stryker Bulgaria 150.xlsx",
    "Stryker_CSISR_F0010":      "Cobra-Stryker CSISR - F0010.xlsx",
    "Stryker_SES_F0010":        "Cobra-Stryker SES - F0010.xlsx",
    "XM30":                     "Cobra-XM30.xlsx",
}

# EVMS cost sets (must match Cobra exports)
COSTSET_BCWS = "BCWS"
COSTSET_BCWP = "BCWP"
COSTSET_ACWP = "ACWP"

# Y-axis for EV plot
YMIN, YMAX = 0.75, 1.25

# Colors for EV index thresholds (CPI / SPI)
BLUE  = RGBColor(0, 112, 192)
GREEN = RGBColor(0, 176, 80)
YELLOW= RGBColor(255, 192, 0)
RED   = RGBColor(192, 0, 0)

# ---------------------------------------------------------
# Cobra loading + normalization (standard format only)
# ---------------------------------------------------------
def load_cobra(path):
    return pd.read_excel(path)

def normalize_cobra_standard(df_raw):
    """
    Normalize to columns: SUBTEAM, COSTSET, DATE, HOURS
    for the 'standard' Cobra schema we already know works.
    If it doesn't match, raise ValueError so we can log + skip.
    """
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    colmap = {}
    for c in df.columns:
        cu = c.upper().replace(" ", "").replace("_", "")
        if cu in ["SUBTEAM", "SUB_TEAM"]:
            colmap[c] = "SUBTEAM"
        elif cu in ["COSTSET", "COST-SET", "COST_SET"]:
            colmap[c] = "COSTSET"
        elif cu.startswith("DATE"):
            colmap[c] = "DATE"
        elif cu in ["HOURS", "HRS"]:
            colmap[c] = "HOURS"

    df = df.rename(columns=colmap)
    required = ["SUBTEAM", "COSTSET", "DATE", "HOURS"]
    if not all(c in df.columns for c in required):
        missing = [c for c in required if c not in df.columns]
        raise ValueError(f"Standard-format columns not found (need {required}, missing {missing})")

    df = df[required].copy()
    df["DATE"] = pd.to_datetime(df["DATE"])
    df["HOURS"] = pd.to_numeric(df["HOURS"], errors="coerce").fillna(0.0)
    df["SUBTEAM"] = df["SUBTEAM"].astype(str).str.strip()

    return df

# ---------------------------------------------------------
# EVMS calculation helpers
# ---------------------------------------------------------
def compute_ev_timeseries(df_norm):
    """
    Aggregate BCWS, BCWP, ACWP by week and compute
    Monthly & Cumulative SPI/CPI series.
    """
    df = df_norm[df_norm["COSTSET"].isin([COSTSET_BCWS, COSTSET_BCWP, COSTSET_ACWP])].copy()
    if df.empty:
        raise ValueError("No BCWS/BCWP/ACWP rows found for this dataset")

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

    pivot = pivot.set_index("DATE").resample("W-MON").sum()

    bcws = pivot[COSTSET_BCWS]
    bcwp = pivot[COSTSET_BCWP]
    acwp = pivot[COSTSET_ACWP]

    # Monthly indices
    cpi_m = np.where(acwp > 0, bcwp / acwp, np.nan)
    spi_m = np.where(bcws > 0, bcwp / bcws, np.nan)

    # Cumulative
    bcws_c = bcws.cumsum()
    bcwp_c = bcwp.cumsum()
    acwp_c = acwp.cumsum()

    cpi_c = np.where(acwp_c > 0, bcwp_c / acwp_c, np.nan)
    spi_c = np.where(bcws_c > 0, bcwp_c / bcws_c, np.nan)

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
    return evdf

def get_status_dates(evdf):
    dates = sorted(evdf["DATE"].unique())
    if not dates:
        raise ValueError("No EV dates available")
    curr = dates[-1]
    prev = dates[-2] if len(dates) > 1 else dates[-1]
    return curr, prev

def program_metric_table(evdf, curr_date, prev_date):
    """
    Program-level EVMS metrics table – SPI & CPI only (no BEI).
    Rows: SPI, CPI
    Cols: Metric, CTD, LSD, Comments...
    """
    row_ctd = evdf.loc[evdf["DATE"] == curr_date].iloc[-1]
    row_lsd = evdf.loc[evdf["DATE"] == prev_date].iloc[-1]

    metrics = pd.DataFrame(
        {
            "Metric": ["SPI", "CPI"],
            "CTD": [row_ctd["SPI_C"], row_ctd["CPI_C"]],
            "LSD": [row_lsd["SPI_C"], row_lsd["CPI_C"]],
            "Comments / Root Cause & Corrective Actions": ["", ""],
        }
    )
    # Put SPI row first, then CPI – already that way
    return metrics

def subteam_metric_table(df_norm):
    """
    Subteam-level SPI/CPI CTD & LSD (cumulative) for each SUBTEAM.
    No BEI.
    """
    out_rows = []
    for st in sorted(df_norm["SUBTEAM"].unique()):
        sub = df_norm[df_norm["SUBTEAM"] == st]
        try:
            ev_sub = compute_ev_timeseries(sub)
        except ValueError:
            continue

        dates = sorted(ev_sub["DATE"].unique())
        if not dates:
            continue
        curr = dates[-1]
        prev = dates[-2] if len(dates) > 1 else dates[-1]

        row_ctd = ev_sub.loc[ev_sub["DATE"] == curr].iloc[-1]
        row_lsd = ev_sub.loc[ev_sub["DATE"] == prev].iloc[-1]

        out_rows.append(
            {
                "Sub Team": st,
                "SPI LSD": row_lsd["SPI_C"],
                "SPI CTD": row_ctd["SPI_C"],
                "CPI LSD": row_lsd["CPI_C"],
                "CPI CTD": row_ctd["CPI_C"],
                "Comments / Root Cause & Corrective Actions": "",
            }
        )

    if not out_rows:
        return pd.DataFrame(
            columns=[
                "Sub Team", "SPI LSD", "SPI CTD",
                "CPI LSD", "CPI CTD",
                "Comments / Root Cause & Corrective Actions",
            ]
        )
    return pd.DataFrame(out_rows)

def labor_and_manpower_tables(df_norm):
    """
    Subteam Labor & VAC + Program Manpower summary.
    BAC ≈ BCWS sum, EAC ≈ ACWP sum (can refine later).
    """
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

    for col in [COSTSET_BCWS, COSTSET_ACWP]:
        if col not in pivot.columns:
            pivot[col] = 0.0

    pivot["BAC"] = pivot[COSTSET_BCWS]
    pivot["EAC"] = pivot[COSTSET_ACWP]
    pivot["VAC"] = pivot["BAC"] - pivot["EAC"]

    labor_df = pivot[["SUBTEAM", "BAC", "EAC", "VAC"]].rename(columns={"SUBTEAM": "Sub Team"})
    labor_df = labor_df.sort_values("Sub Team").reset_index(drop=True)
    labor_df["Comments / Root Cause & Corrective Actions"] = ""

    demand = labor_df["BAC"].sum()
    actual = labor_df["EAC"].sum()
    pct_var = actual / demand if demand > 0 else np.nan

    manpower_df = pd.DataFrame(
        {
            "Demand Hours": [demand],
            "Actual Hours": [actual],
            "% Var": [pct_var],
            "Next Mo BCWS Hours": [0.0],
            "Next Mo ETC Hours": [0.0],
            "Comments / Root Cause & Corrective Actions": [""],
        }
    )

    return labor_df, manpower_df

# ---------------------------------------------------------
# Color helpers
# ---------------------------------------------------------
def ev_index_color(val):
    if pd.isna(val):
        return None
    if val >= 1.05:
        return BLUE
    if val >= 0.98:
        return GREEN
    if val >= 0.95:
        return YELLOW
    return RED

def vac_color(vac, bac):
    if bac <= 0 or pd.isna(vac):
        return None
    pct = vac / bac
    if pct >= 0.05:
        return BLUE
    if pct >= 0.01:
        return GREEN
    if pct >= -0.01:
        return YELLOW
    return RED

def manpower_var_color(pct):
    if pd.isna(pct):
        return None
    # pct is ratio, e.g., 1.10 for 110%
    if 0.90 <= pct <= 1.05:
        return GREEN
    if 0.85 <= pct < 0.90 or 1.05 < pct <= 1.10:
        return YELLOW
    return RED

# ---------------------------------------------------------
# PowerPoint helpers (using same layout indices as the “working” version)
# ---------------------------------------------------------
def load_template():
    if os.path.exists(THEME_PATH):
        return Presentation(THEME_PATH)
    return Presentation()

# use a fixed layout index that worked before (typically "Title and Content")
TREND_LAYOUT_IDX = 1   # for EV trend slide
DETAIL_LAYOUT_IDX = 1  # for detail slides as well

def add_ev_trend_slide(prs, program, evdf, metrics_df):
    slide = prs.slides.add_slide(prs.slide_layouts[TREND_LAYOUT_IDX])
    slide.shapes.title.text = f"{program} EVMS Trend Overview"

    # ----- Plot -----
    fig, ax = plt.subplots(figsize=(7, 4))

    ax.axhspan(YMIN, 0.9,  facecolor="#ffcccc", alpha=0.5)
    ax.axhspan(0.9,  0.95, facecolor="#fff2cc", alpha=0.5)
    ax.axhspan(0.95, 1.05, facecolor="#c6efce", alpha=0.5)
    ax.axhspan(1.05, YMAX, facecolor="#cfe2ff", alpha=0.5)

    ax.scatter(evdf["DATE"], evdf["CPI_M"], s=10, label="Monthly CPI", color="gold")
    ax.scatter(evdf["DATE"], evdf["SPI_M"], s=10, label="Monthly SPI", color="black")
    ax.plot(evdf["DATE"], evdf["CPI_C"], linewidth=2, label="Cumulative CPI", color="blue")
    ax.plot(evdf["DATE"], evdf["SPI_C"], linewidth=2, label="Cumulative SPI", color="gray")

    ax.set_ylim(YMIN, YMAX)
    ax.set_xlabel("Month")
    ax.set_ylabel("EV Indices")
    ax.legend(fontsize=8)
    ax.grid(True, axis="y", alpha=0.3)
    fig.tight_layout()

    img_path = os.path.join(OUTPUT_DIR, f"{program}_EV_plot.png")
    fig.savefig(img_path, dpi=200)
    plt.close(fig)

    chart_left = Inches(0.5)
    chart_top = Inches(1.5)
    slide.shapes.add_picture(img_path, chart_left, chart_top, height=Inches(3.5))

    # ----- Metric table -----
    rows, cols = metrics_df.shape
    t_left = Inches(6.0)
    t_top = Inches(1.5)
    t_width = Inches(3.5)
    t_height = Inches(1.0 + 0.3 * rows)

    t_shape = slide.shapes.add_table(rows + 1, cols, t_left, t_top, t_width, t_height)
    table = t_shape.table

    # headers
    for j, col in enumerate(metrics_df.columns):
        table.cell(0, j).text = col

    # data
    for i in range(rows):
        for j, col in enumerate(metrics_df.columns):
            val = metrics_df.iloc[i, j]
            if isinstance(val, float):
                txt = "" if np.isnan(val) else f"{val:.3f}"
            else:
                txt = str(val)
            table.cell(i + 1, j).text = txt

    # column widths: Metric wider, Comments widest
    for j, col in enumerate(metrics_df.columns):
        if j == 0:  # Metric
            table.columns[j].width = Inches(1.0)
        elif "Comments" in col:
            table.columns[j].width = Inches(2.0)
        else:
            table.columns[j].width = Inches(0.75)

    # color-code CTD / LSD cells
    for i in range(rows):
        for j, col in enumerate(["CTD", "LSD"]):
            val = metrics_df.iloc[i][col]
            rgb = ev_index_color(val)
            if rgb is None:
                continue
            cell = table.cell(i + 1, metrics_df.columns.get_loc(col))
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb

def add_subteam_metric_slides(prs, program, metrics_sub_df, page_size=15):
    if metrics_sub_df.empty:
        return

    n = len(metrics_sub_df)
    pages = int(np.ceil(n / page_size))

    for p in range(pages):
        slide = prs.slides.add_slide(prs.slide_layouts[DETAIL_LAYOUT_IDX])
        label = "" if pages == 1 else f" (Page {p+1})"
        slide.shapes.title.text = f"{program} EVMS Detail – Sub Team CPI / SPI Metrics{label}"

        chunk = metrics_sub_df.iloc[p*page_size:(p+1)*page_size]

        rows, cols = chunk.shape
        left = Inches(0.5)
        top = Inches(1.3)
        width = Inches(9.0)
        height = Inches(0.3 * (rows + 1))

        t_shape = slide.shapes.add_table(rows + 1, cols, left, top, width, height)
        table = t_shape.table

        # headers
        for j, col in enumerate(chunk.columns):
            table.cell(0, j).text = col

        # data + color coding
        for i in range(rows):
            for j, col in enumerate(chunk.columns):
                val = chunk.iloc[i, j]
                cell = table.cell(i + 1, j)

                if isinstance(val, float):
                    txt = "" if np.isnan(val) else f"{val:.3f}"
                else:
                    txt = str(val)
                cell.text = txt

                if col in ["SPI LSD", "SPI CTD", "CPI LSD", "CPI CTD"]:
                    rgb = ev_index_color(chunk.iloc[i][col])
                    if rgb is not None:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = rgb

        # column widths – Comments wider
        for j, col in enumerate(chunk.columns):
            if col == "Sub Team":
                table.columns[j].width = Inches(1.2)
            elif "Comments" in col:
                table.columns[j].width = Inches(3.0)
            else:
                table.columns[j].width = Inches(1.0)

def add_labor_manpower_slides(prs, program, labor_df, manpower_df, page_size=15):
    n = len(labor_df)
    pages = int(np.ceil(max(1, n) / page_size))

    for p in range(pages):
        slide = prs.slides.add_slide(prs.slide_layouts[DETAIL_LAYOUT_IDX])
        label = "" if pages == 1 else f" (Page {p+1})"
        slide.shapes.title.text = f"{program} EVMS Detail – Sub Team Labor & Manpower{label}"

        chunk = labor_df.iloc[p*page_size:(p+1)*page_size]

        rows, cols = chunk.shape
        left = Inches(0.5)
        top = Inches(1.3)
        width = Inches(9.0)
        height = Inches(0.3 * (rows + 1))

        t_shape = slide.shapes.add_table(rows + 1, cols, left, top, width, height)
        table = t_shape.table

        # headers
        for j, col in enumerate(chunk.columns):
            table.cell(0, j).text = col

        # data + VAC color coding
        for i in range(rows):
            for j, col in enumerate(chunk.columns):
                val = chunk.iloc[i, j]
                cell = table.cell(i + 1, j)

                if isinstance(val, float):
                    if "Comments" in col:
                        txt = ""
                    else:
                        txt = f"{val:,.1f}"
                else:
                    txt = str(val)
                cell.text = txt

            # VAC color
            bac = chunk.iloc[i]["BAC"]
            vac = chunk.iloc[i]["VAC"]
            rgb = vac_color(vac, bac)
            if rgb is not None:
                vac_idx = chunk.columns.get_loc("VAC")
                cell = table.cell(i + 1, vac_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb

        # column widths – Comments wide
        for j, col in enumerate(chunk.columns):
            if col == "Sub Team":
                table.columns[j].width = Inches(1.3)
            elif "Comments" in col:
                table.columns[j].width = Inches(3.0)
            else:
                table.columns[j].width = Inches(1.2)

        # Program manpower table – intentionally lower so it doesn’t overlap
        mp_rows, mp_cols = manpower_df.shape
        mp_left = Inches(0.5)
        mp_top  = top + height + Inches(0.4)
        mp_width = Inches(9.0)
        mp_height = Inches(0.8)

        mp_shape = slide.shapes.add_table(mp_rows + 1, mp_cols, mp_left, mp_top, mp_width, mp_height)
        mp_tbl = mp_shape.table

        for j, col in enumerate(manpower_df.columns):
            mp_tbl.cell(0, j).text = col

        for i in range(mp_rows):
            for j, col in enumerate(manpower_df.columns):
                val = manpower_df.iloc[i, j]
                cell = mp_tbl.cell(i + 1, j)

                if isinstance(val, float):
                    if col == "% Var":
                        txt = f"{val:.2%}" if not np.isnan(val) else ""
                    elif "Comments" in col:
                        txt = ""
                    else:
                        txt = f"{val:,.1f}"
                else:
                    txt = str(val)
                cell.text = txt

        # Color for % Var
        pct = manpower_df.iloc[0]["% Var"]
        rgb = manpower_var_color(pct)
        if rgb is not None:
            idx = manpower_df.columns.get_loc("% Var")
            cell = mp_tbl.cell(1, idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb

        # Column widths – Comments wide
        for j, col in enumerate(manpower_df.columns):
            if "Comments" in col:
                mp_tbl.columns[j].width = Inches(3.0)
            else:
                mp_tbl.columns[j].width = Inches(1.3)

# ---------------------------------------------------------
# Main program processor
# ---------------------------------------------------------
skipped_programs = []

def process_program(program_name, cobra_file):
    cobra_path = os.path.join(DATA_DIR, cobra_file)
    if not os.path.exists(cobra_path):
        reason = f"File not found: {cobra_path}"
        print(f">> Skipping {program_name} – {reason}")
        skipped_programs.append((program_name, reason))
        return

    print(f"\n=== Processing {program_name} from {os.path.basename(cobra_path)} ===")
    df_raw = load_cobra(cobra_path)

    try:
        cobra = normalize_cobra_standard(df_raw)
    except ValueError as e:
        print(f">> Skipping {program_name} – {e}")
        skipped_programs.append((program_name, str(e)))
        return

    # EV series
    evdf = compute_ev_timeseries(cobra)
    curr_date, prev_date = get_status_dates(evdf)
    metrics_prog = program_metric_table(evdf, curr_date, prev_date)
    metrics_sub  = subteam_metric_table(cobra)
    labor_df, manpower_df = labor_and_manpower_tables(cobra)

    print(f"CTD date: {curr_date.date()}, LSD date: {prev_date.date()}")
    print(metrics_prog[["Metric", "CTD", "LSD"]])

    prs = load_template()

    # 1) Trend slide
    add_ev_trend_slide(prs, program_name, evdf, metrics_prog)

    # 2) Subteam CPI / SPI metrics
    add_subteam_metric_slides(prs, program_name, metrics_sub, page_size=15)

    # 3) Subteam Labor & Manpower
    add_labor_manpower_slides(prs, program_name, labor_df, manpower_df, page_size=15)

    # Save outputs
    tables_xlsx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Tables.xlsx")
    with pd.ExcelWriter(tables_xlsx, engine="xlsxwriter") as writer:
        evdf.to_excel(writer, sheet_name="EV_Series", index=False)
        metrics_prog.to_excel(writer, sheet_name="Program_Metrics", index=False)
        metrics_sub.to_excel(writer, sheet_name="Subteam_Metrics", index=False)
        labor_df.to_excel(writer, sheet_name="Subteam_Labor", index=False)
        manpower_df.to_excel(writer, sheet_name="Program_Manpower", index=False)

    out_pptx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Deck.pptx")
    prs.save(out_pptx)

    print(f"✓ Saved tables: {tables_xlsx}")
    print(f"✓ Saved deck:   {out_pptx}")

# ---------------------------------------------------------
# Run for all standard-format programs
# ---------------------------------------------------------
for program, cobra_file in PROGRAM_CONFIG.items():
    try:
        process_program(program, cobra_file)
    except Exception as e:
        print(f"!! Unexpected error for {program}: {e}")
        skipped_programs.append((program, f"Unexpected error: {e}"))

print("\nALL STANDARD-FORMAT PROGRAM EVMS DECKS COMPLETE ✓")

if skipped_programs:
    print("\nPrograms needing re-export / clarification (not processed):")
    for prog, reason in skipped_programs:
        print(f" - {prog}: {reason}")