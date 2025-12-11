# ============================================================
# EVMS Pipeline – Standard-format Cobra files, 2-slide deck
# ============================================================

import os
import math
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.dml.color import RGBColor

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

DATA_DIR   = "data"
OUTPUT_DIR = "EVMS_Output"
PPTX_TEMPLATE = None   # or path to a .pptx template if you have one

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ONLY include programs whose Cobra export has the standard EVMS layout
# (SUBTEAM, COSTSET, DATE, HOURS etc.). You can add/remove here.
PROGRAM_CONFIG = {
    "Abrams_STS_2022": "Cobra-Abrams STS 2022.xlsx",
    "Abrams_STS"     : "Cobra-Abrams STS.xlsx",
    "ARV"            : "Cobra-ARV.xlsx",
    "ARV30"          : "Cobra-ARV30.xlsx",
    "Stryker_Bulgaria_150": "Cobra-Stryker Bulgaria 150.xlsx",
    "XM30"           : "Cobra-XM30.xlsx",
}

# Name mapping from raw Cobra columns to logical names we’ll use
# (adjust these if your actual headers differ)
COBRA_COLUMN_MAP = {
    "SUB_TEAM": "SUBTEAM",
    "SUBTEAM": "SUBTEAM",
    "COST-SET": "COSTSET",
    "COSTSET": "COSTSET",
    "DATE": "DATE",
    "HOURS": "HOURS",
}

# EVMS cost sets we expect in standard-format Cobra files
ACWP_CODE = "ACWP"
BCWP_CODE = "BCWP"
BCWS_CODE = "BCWS"
ETC_CODE  = "ETC"   # used for EAC = ACWP + ETC

# Number of subteams per slide on the labor/manpower slide(s)
SUBTEAMS_PER_SLIDE = 15

# EV index plot range & clipping
YMIN, YMAX = 0.75, 1.25

# ------------------------------------------------------------
# Utility helpers
# ------------------------------------------------------------

def normalize_cobra_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardise key columns to SUBTEAM / COSTSET / DATE / HOURS.
    Raises ValueError if any logical column is missing.
    """
    col_map = {}
    for raw, logical in COBRA_COLUMN_MAP.items():
        if raw in df.columns:
            col_map[raw] = logical

    df = df.rename(columns=col_map)

    required = ["SUBTEAM", "COSTSET", "DATE", "HOURS"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Could not normalise Cobra file – missing logical columns: {missing}")

    # Clean types
    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
    df["HOURS"] = pd.to_numeric(df["HOURS"], errors="coerce")
    df = df.dropna(subset=["DATE", "HOURS"])

    # Normalise SUBTEAM as string
    df["SUBTEAM"] = df["SUBTEAM"].astype(str).str.strip()

    return df


def compute_ev_timeseries(cobra: pd.DataFrame) -> pd.DataFrame:
    """
    From a normalised Cobra DataFrame, compute monthly & cumulative
    CPI/SPI series at month-end. Returns DataFrame with:
    DATE (month-end), CPI_M, SPI_M, CPI_CUM, SPI_CUM.
    """
    mask = cobra["COSTSET"].isin([ACWP_CODE, BCWP_CODE, BCWS_CODE])
    ev = cobra.loc[mask].copy()

    if ev.empty:
        raise ValueError("No BCWS/BCWP/ACWP rows found after filtering COSTSET")

    # Pivot to daily sums
    pivot = (
        ev.pivot_table(
            index="DATE",
            columns="COSTSET",
            values="HOURS",
            aggfunc="sum"
        )
        .sort_index()
    )

    # Fill missing cost-set columns if needed
    for cs in (ACWP_CODE, BCWP_CODE, BCWS_CODE):
        if cs not in pivot.columns:
            pivot[cs] = 0.0

    # Monthly totals at month-end ('ME' avoids the deprecated 'M' alias)
    monthly = pivot.resample("ME").sum()

    # Monthly indices
    acwp = monthly[ACWP_CODE].replace(0, np.nan)
    bcwp = monthly[BCWP_CODE]
    bcws = monthly[BCWS_CODE].replace(0, np.nan)

    monthly_cpi = bcwp / acwp
    monthly_spi = bcwp / bcws

    # Cumulative hours & indices
    cum = pivot.cumsum().resample("ME").last()
    acwp_c = cum[ACWP_CODE].replace(0, np.nan)
    bcwp_c = cum[BCWP_CODE]
    bcws_c = cum[BCWS_CODE].replace(0, np.nan)

    cum_cpi = bcwp_c / acwp_c
    cum_spi = bcwp_c / bcws_c

    evdf = pd.DataFrame(
        {
            "DATE": monthly.index,
            "CPI_M": monthly_cpi.values,
            "SPI_M": monthly_spi.values,
            "CPI_CUM": cum_cpi.reindex(monthly.index).values,
            "SPI_CUM": cum_spi.reindex(monthly.index).values,
        }
    ).dropna(subset=["CPI_M", "SPI_M"], how="all")

    return evdf


def get_ctd_lsd(evdf: pd.DataFrame):
    """
    For now, treat the last month-end as both CTD and LSD.
    (You can later plug in explicit LSD dates if you have them.)
    """
    if evdf.empty:
        raise ValueError("EV time series is empty")
    ctd_date = evdf["DATE"].max()
    lsd_date = ctd_date
    return ctd_date, lsd_date


def build_program_metric_table(evdf: pd.DataFrame) -> pd.DataFrame:
    """
    Build top-level metrics table with rows for SPI and CPI only.
    Columns: Metric, CTD, LSD, Comments / Root Cause & Corrective Actions
    """
    ctd_date, lsd_date = get_ctd_lsd(evdf)

    # CTD values = latest cumulative; LSD values = latest monthly
    ev_ctd = evdf.loc[evdf["DATE"] == ctd_date].iloc[-1]
    ev_lsd = evdf.loc[evdf["DATE"] == lsd_date].iloc[-1]

    data = []
    for metric, ctd_val, lsd_val in [
        ("SPI", ev_ctd["SPI_CUM"], ev_lsd["SPI_M"]),
        ("CPI", ev_ctd["CPI_CUM"], ev_lsd["CPI_M"]),
    ]:
        data.append(
            {
                "Metric": metric,
                "CTD": float(ctd_val) if pd.notna(ctd_val) else np.nan,
                "LSD": float(lsd_val) if pd.notna(lsd_val) else np.nan,
                "Comments / Root Cause & Corrective Actions": "",
            }
        )

    return pd.DataFrame(data), ctd_date, lsd_date


def build_labor_table(cobra: pd.DataFrame) -> pd.DataFrame:
    """
    Subteam BAC/EAC/VAC table.
    Assumes:
      - BCWS hours ~ budget (BAC in hours)
      - EAC hours = ACWP + ETC
    Adjust the logic if your org uses different cost-set codes for BAC/EAC.
    """
    mask = cobra["COSTSET"].isin([ACWP_CODE, BCWP_CODE, BCWS_CODE, ETC_CODE])
    df = cobra.loc[mask].copy()

    # Pivot per SUBTEAM
    pivot = (
        df.pivot_table(
            index="SUBTEAM",
            columns="COSTSET",
            values="HOURS",
            aggfunc="sum"
        )
        .fillna(0.0)
    )

    # Derived metrics
    # BAC (hours) ~ total BCWS
    pivot["BAC"] = pivot.get(BCWS_CODE, 0.0)
    # EAC (hours) ~ ACWP + ETC
    pivot["EAC"] = pivot.get(ACWP_CODE, 0.0) + pivot.get(ETC_CODE, 0.0)
    pivot["VAC"] = pivot["BAC"] - pivot["EAC"]

    out = (
        pivot[["BAC", "EAC", "VAC"]]
        .reset_index()
        .rename(columns={"SUBTEAM": "Sub Team"})
    )

    out["Comments / Root Cause & Corrective Actions"] = ""

    return out


def build_manpower_table(cobra: pd.DataFrame) -> pd.DataFrame:
    """
    Program-level manpower summary.
    This is a simple version: total budget vs actual,
    plus placeholders for next-month BCWS/ETC.
    You can plug in your exact 9/80 logic later.
    """
    # Total BCWS, ACWP, ETC (hours)
    agg = (
        cobra[cobra["COSTSET"].isin([BCWS_CODE, ACWP_CODE, ETC_CODE])]
        .pivot_table(
            index="COSTSET",
            values="HOURS",
            aggfunc="sum"
        )
        .to_dict()["HOURS"]
    )

    demand_hours = float(agg.get(BCWS_CODE, 0.0))
    actual_hours = float(agg.get(ACWP_CODE, 0.0))

    pct_var = (actual_hours / demand_hours) if demand_hours else np.nan

    df = pd.DataFrame(
        [
            {
                "Demand Hours": demand_hours,
                "Actual Hours": actual_hours,
                "% Var": pct_var,
                "Next Mo BCWS Hours": 0.0,
                "Next Mo ETC Hours": 0.0,
                "Comments / Root Cause & Corrective Actions": "",
            }
        ]
    )

    return df


# ------------------------------------------------------------
# Plotting – EV trend with colored bands and clipping
# ------------------------------------------------------------

def add_color_band(ax, y0, y1, color):
    ax.axhspan(y0, y1, facecolor=color, alpha=0.3, linewidth=0)


def make_ev_plot(evdf: pd.DataFrame, program_name: str, out_png: str):
    """
    Create EVMS trend plot with CPI/SPI monthly + cumulative,
    clipped to [YMIN, YMAX] for display, with colored bands.
    """
    plot_df = evdf.copy()

    # Clip EV indices for plotting ONLY (metrics use raw values)
    for col in ["CPI_M", "SPI_M", "CPI_CUM", "SPI_CUM"]:
        plot_df[col] = plot_df[col].clip(lower=YMIN, upper=YMAX)

    plt.close("all")
    fig, ax = plt.subplots(figsize=(7, 4))

    # Color bands – red <0.95, yellow 0.95–0.98, green 0.98–1.05, blue >1.05
    add_color_band(ax, YMIN, 0.95, "red")
    add_color_band(ax, 0.95, 0.98, "yellow")
    add_color_band(ax, 0.98, 1.05, "green")
    add_color_band(ax, 1.05, YMAX, "lightblue")

    # Monthly points
    ax.scatter(plot_df["DATE"], plot_df["CPI_M"], s=15, label="Monthly CPI")
    ax.scatter(plot_df["DATE"], plot_df["SPI_M"], s=15, label="Monthly SPI")

    # Cumulative lines
    ax.plot(plot_df["DATE"], plot_df["CPI_CUM"], linewidth=2, label="Cumulative CPI")
    ax.plot(plot_df["DATE"], plot_df["SPI_CUM"], linewidth=2, label="Cumulative SPI")

    ax.set_ylim(YMIN, YMAX)
    ax.set_xlabel("Month")
    ax.set_ylabel("EV Indices")
    ax.set_title(f"{program_name} EVMS Trend Overview")
    ax.legend(loc="upper left", fontsize=8)

    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(out_png, dpi=200)
    plt.close(fig)


# ------------------------------------------------------------
# PowerPoint helpers – remove placeholders & build slides
# ------------------------------------------------------------

def remove_body_placeholders(slide):
    """
    Remove default content placeholders, including any 'Click to add text'
    boxes that might sit behind tables or charts.
    """
    for shape in list(slide.shapes):
        if not shape.is_placeholder:
            # Also strip freeform text boxes with 'Click to add text' text
            if getattr(shape, "has_text_frame", False):
                txt = shape.text_frame.text.strip()
                if txt.lower().startswith("click to add"):
                    el = shape._element
                    el.getparent().remove(el)
            continue

        phf = shape.placeholder_format
        if phf.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.CONTENT):
            el = shape._element
            el.getparent().remove(el)
        else:
            if getattr(shape, "has_text_frame", False):
                txt = shape.text_frame.text.strip()
                if txt.lower().startswith("click to add"):
                    el = shape._element
                    el.getparent().remove(el)


def add_overview_slide(prs, program_name, ev_plot_png, metrics_df, ctd_date, lsd_date):
    """
    Slide 1: EVMS trend plot + SPI/CPI overview table.
    BEI is intentionally omitted.
    """
    layout = prs.slide_layouts[1]  # title + content
    slide = prs.slides.add_slide(layout)
    remove_body_placeholders(slide)

    # Title
    slide.shapes.title.text = f"{program_name} EVMS Trend Overview"

    # Plot on the left
    plot_left = Inches(0.5)
    plot_top = Inches(1.4)
    plot_height = Inches(4.0)
    slide.shapes.add_picture(ev_plot_png, plot_left, plot_top, height=plot_height)

    # SPI/CPI metrics (only) – ensure order SPI, CPI
    m = metrics_df.copy()
    m = m[m["Metric"].isin(["SPI", "CPI"])]
    m["Metric"] = pd.Categorical(m["Metric"], categories=["SPI", "CPI"], ordered=True)
    m = m.sort_values("Metric")

    rows = len(m) + 1
    cols = 4  # Metric, CTD, LSD, Comments

    tbl_left = Inches(6.1)
    tbl_top = Inches(1.4)
    tbl_width = Inches(4.0)
    tbl_height = Inches(1.6)

    tbl_shape = slide.shapes.add_table(rows, cols, tbl_left, tbl_top,
                                       tbl_width, tbl_height)
    tbl = tbl_shape.table

    # Column widths – metric nice, comments widest
    tbl.columns[0].width = Inches(0.9)  # Metric
    tbl.columns[1].width = Inches(0.9)  # CTD
    tbl.columns[2].width = Inches(0.9)  # LSD
    tbl.columns[3].width = Inches(1.3)  # Comments

    headers = ["Metric", "CTD", "LSD", "Comments / Root Cause & Corrective Actions"]
    for j, h in enumerate(headers):
        tbl.cell(0, j).text = h

    for i, (_, row) in enumerate(m.iterrows(), start=1):
        tbl.cell(i, 0).text = str(row["Metric"])
        tbl.cell(i, 1).text = f"{row['CTD']:.3f}" if pd.notna(row["CTD"]) else ""
        tbl.cell(i, 2).text = f"{row['LSD']:.3f}" if pd.notna(row["LSD"]) else ""
        tbl.cell(i, 3).text = ""  # user fills in comments


def add_labor_manpower_slide(
    prs,
    program_name,
    labor_df,
    manpower_df,
    page_idx: int,
    total_pages: int
):
    """
    Slide 2+ : Sub Team Labor & Manpower (paged; 15 subteams per slide).
    """
    layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)
    remove_body_placeholders(slide)

    # Title with page index if multiple pages
    if total_pages > 1:
        title_txt = f"{program_name} EVMS Detail – Sub Team Labor & Manpower (Page {page_idx+1})"
    else:
        title_txt = f"{program_name} EVMS Detail – Sub Team Labor & Manpower"
    slide.shapes.title.text = title_txt

    # Slice labor_df for this page
    start = page_idx * SUBTEAMS_PER_SLIDE
    end   = (page_idx + 1) * SUBTEAMS_PER_SLIDE
    ldf   = labor_df.iloc[start:end].reset_index(drop=True)

    labor_cols = [
        "Sub Team",
        "BAC",
        "EAC",
        "VAC",
        "Comments / Root Cause & Corrective Actions",
    ]

    n_rows = len(ldf) + 1
    n_cols = len(labor_cols)

    top_left = Inches(0.5)
    top_top = Inches(1.4)
    top_width = Inches(9.0)
    top_height = Inches(3.6)

    labor_shape = slide.shapes.add_table(n_rows, n_cols,
                                         top_left, top_top,
                                         top_width, top_height)
    labor_tbl = labor_shape.table

    # Column widths – comments column wide
    labor_tbl.columns[0].width = Inches(1.0)
    labor_tbl.columns[1].width = Inches(1.3)
    labor_tbl.columns[2].width = Inches(1.3)
    labor_tbl.columns[3].width = Inches(1.3)
    labor_tbl.columns[4].width = Inches(4.1)

    # Header
    for j, col in enumerate(labor_cols):
        labor_tbl.cell(0, j).text = col

    # Data
    for i, (_, row) in enumerate(ldf.iterrows(), start=1):
        labor_tbl.cell(i, 0).text = str(row["Sub Team"])
        for j, col in enumerate(labor_cols[1:-1], start=1):
            val = row[col]
            if pd.isna(val):
                txt = ""
            elif isinstance(val, (int, float)):
                txt = f"{val:,.1f}"
            else:
                txt = str(val)
            labor_tbl.cell(i, j).text = txt
        # comments – blank
        labor_tbl.cell(i, n_cols-1).text = ""

    # Program Manpower table – always on same slide, pushed down
    pm_cols = [
        "Demand Hours",
        "Actual Hours",
        "% Var",
        "Next Mo BCWS Hours",
        "Next Mo ETC Hours",
        "Comments / Root Cause & Corrective Actions",
    ]
    pm_rows = len(manpower_df) + 1
    pm_cols_n = len(pm_cols)

    pm_left = Inches(0.5)
    pm_top  = top_top + top_height + Inches(0.3)
    pm_width = Inches(9.0)
    pm_height = Inches(1.1)

    pm_shape = slide.shapes.add_table(pm_rows, pm_cols_n,
                                      pm_left, pm_top,
                                      pm_width, pm_height)
    pm_tbl = pm_shape.table

    # Column widths – comments wider
    for j in range(pm_cols_n - 1):
        pm_tbl.columns[j].width = Inches(1.2)
    pm_tbl.columns[pm_cols_n - 1].width = Inches(2.4)

    # Headers
    for j, col in enumerate(pm_cols):
        pm_tbl.cell(0, j).text = col

    # Data
    for i, (_, row) in enumerate(manpower_df.iterrows(), start=1):
        for j, col in enumerate(pm_cols):
            val = row[col]
            if "Comments" in col:
                pm_tbl.cell(i, j).text = ""
                continue
            if pd.isna(val):
                txt = ""
            elif isinstance(val, (int, float)):
                # % Var as percentage
                if col == "% Var":
                    txt = f"{val*100:.2f}%"
                else:
                    txt = f"{val:,.1f}"
            else:
                txt = str(val)
            pm_tbl.cell(i, j).text = txt


# ------------------------------------------------------------
# Core per-program processing
# ------------------------------------------------------------

def process_program(program_name: str, cobra_filename: str):
    cobra_path = os.path.join(DATA_DIR, cobra_filename)
    if not os.path.exists(cobra_path):
        raise FileNotFoundError(f"Cobra file not found: {cobra_path}")

    print(f"\n=== Processing {program_name} from {cobra_filename} ===")

    raw = pd.read_excel(cobra_path)
    cobra = normalize_cobra_columns(raw)

    # EV time series & metrics
    evdf = compute_ev_timeseries(cobra)
    metrics_df, ctd_date, lsd_date = build_program_metric_table(evdf)

    # Subteam labor & program manpower
    labor_df = build_labor_table(cobra)
    manpower_df = build_manpower_table(cobra)

    # Make EV plot
    ev_plot_png = os.path.join(OUTPUT_DIR, f"{program_name}_EV_Plot.png")
    make_ev_plot(evdf, program_name, ev_plot_png)

    # Write tables workbook
    tables_xlsx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Tables.xlsx")
    with pd.ExcelWriter(tables_xlsx, engine="xlsxwriter") as writer:
        evdf.to_excel(writer, sheet_name="EV_Series", index=False)
        metrics_df.to_excel(writer, sheet_name="Program_Metrics", index=False)
        labor_df.to_excel(writer, sheet_name="Subteam_Labor", index=False)
        manpower_df.to_excel(writer, sheet_name="Program_Manpower", index=False)

    # Build deck
    if PPTX_TEMPLATE and os.path.exists(PPTX_TEMPLATE):
        prs = Presentation(PPTX_TEMPLATE)
    else:
        prs = Presentation()

    # Slide 1 – overview
    add_overview_slide(prs, program_name, ev_plot_png, metrics_df, ctd_date, lsd_date)

    # Slide 2+ – labor & manpower (paged)
    n_pages = max(1, math.ceil(len(labor_df) / SUBTEAMS_PER_SLIDE))
    for page_idx in range(n_pages):
        add_labor_manpower_slide(prs, program_name, labor_df, manpower_df,
                                 page_idx, n_pages)

    out_pptx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Deck.pptx")
    prs.save(out_pptx)

    print(f"✓ CTD date: {ctd_date.date()}, LSD date: {lsd_date.date()}")
    print(f"✓ Saved tables: {tables_xlsx}")
    print(f"✓ Saved deck:   {out_pptx}")


# ------------------------------------------------------------
# Run pipeline for all configured standard-format programs
# ------------------------------------------------------------

program_errors = {}

for program, cobra_file in PROGRAM_CONFIG.items():
    try:
        process_program(program, cobra_file)
    except Exception as e:
        print(f"!! Error for {program}: {e}")
        program_errors[program] = str(e)

print("\nALL STANDARD-FORMAT PROGRAM EVMS DECKS COMPLETE ✓")

if program_errors:
    print("\nPrograms needing re-export / clarification (not processed):")
    for prog, msg in program_errors.items():
        print(f"- {prog}: {msg}")