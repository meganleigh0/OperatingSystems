# ============================================================
# EVMS Pipeline – Standard-format Cobra files (with theme + subteam EVMS detail)
# ============================================================

import os
import math
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Inches, Pt

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

DATA_DIR   = "data"
OUTPUT_DIR = "EVMS_Output"
THEME_PATH = os.path.join(DATA_DIR, "Theme.pptx")   # your GDLS theme file

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Programs whose Cobra exports use the standard layout (SUBTEAM/COSTSET/DATE/HOURS)
PROGRAM_CONFIG = {
    "Abrams_STS_2022": "Cobra-Abrams STS 2022.xlsx",
    "Abrams_STS"     : "Cobra-Abrams STS.xlsx",
    "ARV"            : "Cobra-ARV.xlsx",
    "ARV30"          : "Cobra-ARV30.xlsx",
    "Stryker_Bulgaria_150": "Cobra-Stryker Bulgaria 150.xlsx",
    "XM30"           : "Cobra-XM30.xlsx",
}

# Column name normalization
COBRA_COLUMN_MAP = {
    "SUB_TEAM": "SUBTEAM",
    "SUBTEAM": "SUBTEAM",
    "COST-SET": "COSTSET",
    "COSTSET": "COSTSET",
    "DATE": "DATE",
    "HOURS": "HOURS",
}

# Logical cost-set labels
ACWP_CODE = "ACWP"
BCWP_CODE = "BCWP"
BCWS_CODE = "BCWS"
ETC_CODE  = "ETC"

SUBTEAMS_PER_SLIDE = 15

# EV plot y-range
YMIN, YMAX = 0.75, 1.25

# ------------------------------------------------------------
# Normalisation & cost-set mapping
# ------------------------------------------------------------

def normalize_cobra_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Standardise key columns and add COSTSET_LOGIC based on fuzzy mapping."""
    # Rename to logical names
    col_map = {}
    for raw, logical in COBRA_COLUMN_MAP.items():
        if raw in df.columns:
            col_map[raw] = logical
    df = df.rename(columns=col_map)

    required = ["SUBTEAM", "COSTSET", "DATE", "HOURS"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Could not normalise Cobra file – missing logical columns: {missing}")

    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
    df["HOURS"] = pd.to_numeric(df["HOURS"], errors="coerce")
    df = df.dropna(subset=["DATE", "HOURS"])

    df["SUBTEAM"] = df["SUBTEAM"].astype(str).str.strip()
    df["COSTSET"] = df["COSTSET"].astype(str).str.strip()

    # Fuzzy mapping from raw COSTSET → logical COSTSET_LOGIC
    def logical_from_raw(val: str):
        s = str(val).upper()
        if "ACWP" in s:
            return ACWP_CODE
        if "BCWP" in s:
            return BCWP_CODE
        # Treat “BCWS”, “BUDGET”, “BAC” as budget (BCWS)
        if ("BCWS" in s) or ("BUDG" in s) or ("BAC" in s):
            return BCWS_CODE
        if "ETC" in s or "REMAIN" in s:
            return ETC_CODE
        return None

    df["COSTSET_LOGIC"] = df["COSTSET"].map(logical_from_raw)

    return df


# ------------------------------------------------------------
# EV time series & metrics
# ------------------------------------------------------------

def compute_ev_timeseries(cobra: pd.DataFrame) -> pd.DataFrame:
    """Compute monthly & cumulative CPI/SPI at month-end."""
    mask = cobra["COSTSET_LOGIC"].isin([ACWP_CODE, BCWP_CODE, BCWS_CODE])
    ev = cobra.loc[mask].copy()
    if ev.empty:
        raise ValueError("No BCWS/BCWP/ACWP rows found after mapping COSTSET")

    pivot = (
        ev.pivot_table(
            index="DATE",
            columns="COSTSET_LOGIC",
            values="HOURS",
            aggfunc="sum"
        )
        .sort_index()
    )

    for cs in (ACWP_CODE, BCWP_CODE, BCWS_CODE):
        if cs not in pivot.columns:
            pivot[cs] = 0.0

    # Monthly totals at month-end (ME is non-deprecated alias for M)
    monthly = pivot.resample("ME").sum()

    acwp = monthly[ACWP_CODE].replace(0, np.nan)
    bcwp = monthly[BCWP_CODE]
    bcws = monthly[BCWS_CODE].replace(0, np.nan)

    monthly_cpi = bcwp / acwp
    monthly_spi = bcwp / bcws

    # Cumulative
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
    """For now, use latest month-end as both CTD and LSD."""
    if evdf.empty:
        raise ValueError("EV time series is empty")
    ctd_date = evdf["DATE"].max()
    lsd_date = ctd_date
    return ctd_date, lsd_date


def build_program_metric_table(evdf: pd.DataFrame) -> tuple[pd.DataFrame, datetime, datetime]:
    """Build top-level SPI/CPI CTD/LSD table (no BEI)."""
    ctd_date, lsd_date = get_ctd_lsd(evdf)

    ev_ctd = evdf.loc[evdf["DATE"] == ctd_date].iloc[-1]
    ev_lsd = evdf.loc[evdf["DATE"] == lsd_date].iloc[-1]

    rows = []
    for metric, ctd_val, lsd_val in [
        ("SPI", ev_ctd["SPI_CUM"], ev_lsd["SPI_M"]),
        ("CPI", ev_ctd["CPI_CUM"], ev_lsd["CPI_M"]),
    ]:
        rows.append(
            {
                "Metric": metric,
                "CTD": float(ctd_val) if pd.notna(ctd_val) else np.nan,
                "LSD": float(lsd_val) if pd.notna(lsd_val) else np.nan,
                "Comments / Root Cause & Corrective Actions": "",
            }
        )

    return pd.DataFrame(rows), ctd_date, lsd_date


def build_subteam_metric_table(cobra: pd.DataFrame,
                               ctd_date: datetime,
                               lsd_date: datetime) -> pd.DataFrame:
    """
    Sub-team EVMS metrics (SPI/CPI, CTD/LSD).
    CTD = cumulative SPI/CPI at/through CTD date.
    LSD = monthly SPI/CPI at LSD date.
    """
    subteams = sorted(cobra["SUBTEAM"].dropna().astype(str).unique())
    rows = []

    for st in subteams:
        sub_df = cobra[cobra["SUBTEAM"] == st]
        try:
            ev_sub = compute_ev_timeseries(sub_df)
        except ValueError:
            # If we can't compute EV for this subteam, leave as NaN
            rows.append(
                {
                    "Sub Team": st,
                    "SPI CTD": np.nan,
                    "SPI LSD": np.nan,
                    "CPI CTD": np.nan,
                    "CPI LSD": np.nan,
                    "Comments / Root Cause & Corrective Actions": "",
                }
            )
            continue

        ev_ctd = ev_sub[ev_sub["DATE"] <= ctd_date]
        ev_lsd = ev_sub[ev_sub["DATE"] <= lsd_date]

        if ev_ctd.empty or ev_lsd.empty:
            rows.append(
                {
                    "Sub Team": st,
                    "SPI CTD": np.nan,
                    "SPI LSD": np.nan,
                    "CPI CTD": np.nan,
                    "CPI LSD": np.nan,
                    "Comments / Root Cause & Corrective Actions": "",
                }
            )
            continue

        ctd_row = ev_ctd.iloc[-1]
        lsd_row = ev_lsd.iloc[-1]

        rows.append(
            {
                "Sub Team": st,
                "SPI CTD": float(ctd_row["SPI_CUM"]) if pd.notna(ctd_row["SPI_CUM"]) else np.nan,
                "SPI LSD": float(lsd_row["SPI_M"]) if pd.notna(lsd_row["SPI_M"]) else np.nan,
                "CPI CTD": float(ctd_row["CPI_CUM"]) if pd.notna(ctd_row["CPI_CUM"]) else np.nan,
                "CPI LSD": float(lsd_row["CPI_M"]) if pd.notna(lsd_row["CPI_M"]) else np.nan,
                "Comments / Root Cause & Corrective Actions": "",
            }
        )

    return pd.DataFrame(rows)


# ------------------------------------------------------------
# Labor & manpower tables
# ------------------------------------------------------------

def build_labor_table(cobra: pd.DataFrame) -> pd.DataFrame:
    """Subteam BAC/EAC/VAC table using logical cost-sets."""
    mask = cobra["COSTSET_LOGIC"].isin([ACWP_CODE, BCWP_CODE, BCWS_CODE, ETC_CODE])
    df = cobra.loc[mask].copy()

    pivot = (
        df.pivot_table(
            index="SUBTEAM",
            columns="COSTSET_LOGIC",
            values="HOURS",
            aggfunc="sum"
        )
        .fillna(0.0)
    )

    pivot["BAC"] = pivot.get(BCWS_CODE, 0.0)
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
    """Program-level manpower summary (simple version)."""
    mask = cobra["COSTSET_LOGIC"].isin([BCWS_CODE, ACWP_CODE, ETC_CODE])
    df = cobra.loc[mask].copy()

    agg = (
        df.pivot_table(
            index="COSTSET_LOGIC",
            values="HOURS",
            aggfunc="sum"
        )
        .to_dict()
        .get("HOURS", {})
    )

    demand_hours = float(agg.get(BCWS_CODE, 0.0))
    actual_hours = float(agg.get(ACWP_CODE, 0.0))
    pct_var = (actual_hours / demand_hours) if demand_hours else np.nan

    return pd.DataFrame(
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


# ------------------------------------------------------------
# Plotting
# ------------------------------------------------------------

def add_color_band(ax, y0, y1, color):
    ax.axhspan(y0, y1, facecolor=color, alpha=0.3, linewidth=0)


def make_ev_plot(evdf: pd.DataFrame, program_name: str, out_png: str):
    """Create EVMS trend plot with clipped indices and color bands."""
    plot_df = evdf.copy()
    for col in ["CPI_M", "SPI_M", "CPI_CUM", "SPI_CUM"]:
        plot_df[col] = plot_df[col].clip(lower=YMIN, upper=YMAX)

    plt.close("all")
    fig, ax = plt.subplots(figsize=(7, 4))

    add_color_band(ax, YMIN, 0.95, "red")
    add_color_band(ax, 0.95, 0.98, "yellow")
    add_color_band(ax, 0.98, 1.05, "green")
    add_color_band(ax, 1.05, YMAX, "lightblue")

    ax.scatter(plot_df["DATE"], plot_df["CPI_M"], s=15, label="Monthly CPI")
    ax.scatter(plot_df["DATE"], plot_df["SPI_M"], s=15, label="Monthly SPI")

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
# PowerPoint helpers
# ------------------------------------------------------------

def remove_click_to_add_text_boxes(slide):
    """Remove any shape whose text is 'Click to add text' etc."""
    for shape in list(slide.shapes):
        if getattr(shape, "has_text_frame", False):
            txt = (shape.text_frame.text or "").strip().lower()
            if txt.startswith("click to add"):
                el = shape._element
                el.getparent().remove(el)


def add_overview_slide(prs, program_name, ev_plot_png, metrics_df, ctd_date, lsd_date):
    """Slide 1: trend plot + SPI/CPI overview table."""
    layout = prs.slide_layouts[1]  # title + content from Theme.pptx
    slide = prs.slides.add_slide(layout)
    remove_click_to_add_text_boxes(slide)

    slide.shapes.title.text = f"{program_name} EVMS Trend Overview"

    # Plot on left
    plot_left = Inches(0.5)
    plot_top = Inches(1.4)
    plot_height = Inches(4.0)
    slide.shapes.add_picture(ev_plot_png, plot_left, plot_top, height=plot_height)

    # SPI/CPI table on right
    m = metrics_df.copy()
    m = m[m["Metric"].isin(["SPI", "CPI"])]
    m["Metric"] = pd.Categorical(m["Metric"], categories=["SPI", "CPI"], ordered=True)
    m = m.sort_values("Metric")

    rows = len(m) + 1
    cols = 4

    tbl_left = Inches(6.0)
    tbl_top = Inches(1.4)
    tbl_width = Inches(4.8)
    tbl_height = Inches(1.8)

    tbl_shape = slide.shapes.add_table(rows, cols, tbl_left, tbl_top,
                                       tbl_width, tbl_height)
    tbl = tbl_shape.table

    # Column widths (Metric wider, Comments widest)
    tbl.columns[0].width = Inches(1.3)  # Metric
    tbl.columns[1].width = Inches(1.1)  # CTD
    tbl.columns[2].width = Inches(1.1)  # LSD
    tbl.columns[3].width = Inches(1.3)  # Comments

    headers = ["Metric", "CTD", "LSD", "Comments / Root Cause & Corrective Actions"]
    for j, h in enumerate(headers):
        tbl.cell(0, j).text = h

    for i, (_, row) in enumerate(m.iterrows(), start=1):
        tbl.cell(i, 0).text = str(row["Metric"])
        tbl.cell(i, 1).text = f"{row['CTD']:.3f}" if pd.notna(row["CTD"]) else ""
        tbl.cell(i, 2).text = f"{row['LSD']:.3f}" if pd.notna(row["LSD"]) else ""
        tbl.cell(i, 3).text = ""  # comments


def add_subteam_evms_slide(prs, program_name, sub_metrics_df,
                           page_idx: int, total_pages: int):
    """Slides: sub-team EVMS metrics (SPI/CPI CTD/LSD)."""
    layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)
    remove_click_to_add_text_boxes(slide)

    if total_pages > 1:
        title = f"{program_name} EVMS Detail – Sub Team EVMS Metrics (Page {page_idx+1})"
    else:
        title = f"{program_name} EVMS Detail – Sub Team EVMS Metrics"
    slide.shapes.title.text = title

    start = page_idx * SUBTEAMS_PER_SLIDE
    end = (page_idx + 1) * SUBTEAMS_PER_SLIDE
    sdf = sub_metrics_df.iloc[start:end].reset_index(drop=True)

    cols = [
        "Sub Team",
        "SPI CTD",
        "SPI LSD",
        "CPI CTD",
        "CPI LSD",
        "Comments / Root Cause & Corrective Actions",
    ]

    n_rows = len(sdf) + 1
    n_cols = len(cols)

    left = Inches(0.5)
    top = Inches(1.4)
    width = Inches(9.5)
    height = Inches(4.0)

    shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    tbl = shape.table

    # Column widths – comments very wide
    tbl.columns[0].width = Inches(1.0)
    tbl.columns[1].width = Inches(1.0)
    tbl.columns[2].width = Inches(1.0)
    tbl.columns[3].width = Inches(1.0)
    tbl.columns[4].width = Inches(1.0)
    tbl.columns[5].width = Inches(4.5)

    for j, col in enumerate(cols):
        tbl.cell(0, j).text = col

    for i, (_, row) in enumerate(sdf.iterrows(), start=1):
        tbl.cell(i, 0).text = str(row["Sub Team"])
        for j, col in enumerate(cols[1:-1], start=1):
            val = row[col]
            if pd.isna(val):
                txt = ""
            else:
                txt = f"{val:.3f}"
            tbl.cell(i, j).text = txt
        tbl.cell(i, n_cols - 1).text = ""  # comments


def add_labor_manpower_slide(prs, program_name, labor_df, manpower_df,
                             page_idx: int, total_pages: int):
    """Slides: subteam labor & manpower (BAC/EAC/VAC + program manpower)."""
    layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)
    remove_click_to_add_text_boxes(slide)

    if total_pages > 1:
        title = f"{program_name} EVMS Detail – Sub Team Labor & Manpower (Page {page_idx+1})"
    else:
        title = f"{program_name} EVMS Detail – Sub Team Labor & Manpower"
    slide.shapes.title.text = title

    # Slice subteams for this page
    start = page_idx * SUBTEAMS_PER_SLIDE
    end = (page_idx + 1) * SUBTEAMS_PER_SLIDE
    ldf = labor_df.iloc[start:end].reset_index(drop=True)

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
    top_width = Inches(9.5)
    top_height = Inches(3.5)

    labor_shape = slide.shapes.add_table(n_rows, n_cols,
                                         top_left, top_top,
                                         top_width, top_height)
    labor_tbl = labor_shape.table

    # Column widths – comments very wide
    labor_tbl.columns[0].width = Inches(1.0)
    labor_tbl.columns[1].width = Inches(1.4)
    labor_tbl.columns[2].width = Inches(1.4)
    labor_tbl.columns[3].width = Inches(1.4)
    labor_tbl.columns[4].width = Inches(4.3)

    for j, col in enumerate(labor_cols):
        labor_tbl.cell(0, j).text = col

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
        labor_tbl.cell(i, n_cols - 1).text = ""

    # Program manpower table, pushed down
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
    pm_top = top_top + top_height + Inches(0.3)
    pm_width = Inches(9.5)
    pm_height = Inches(1.1)

    pm_shape = slide.shapes.add_table(pm_rows, pm_cols_n,
                                      pm_left, pm_top,
                                      pm_width, pm_height)
    pm_tbl = pm_shape.table

    for j in range(pm_cols_n - 1):
        pm_tbl.columns[j].width = Inches(1.3)
    pm_tbl.columns[pm_cols_n - 1].width = Inches(2.3)

    for j, col in enumerate(pm_cols):
        pm_tbl.cell(0, j).text = col

    for i, (_, row) in enumerate(manpower_df.iterrows(), start=1):
        for j, col in enumerate(pm_cols):
            val = row[col]
            if "Comments" in col:
                pm_tbl.cell(i, j).text = ""
                continue
            if pd.isna(val):
                txt = ""
            elif isinstance(val, (int, float)):
                if col == "% Var":
                    txt = f"{val*100:.2f}%"
                else:
                    txt = f"{val:,.1f}"
            else:
                txt = str(val)
            pm_tbl.cell(i, j).text = txt


# ------------------------------------------------------------
# Per-program processing
# ------------------------------------------------------------

def process_program(program_name: str, cobra_filename: str):
    cobra_path = os.path.join(DATA_DIR, cobra_filename)
    if not os.path.exists(cobra_path):
        raise FileNotFoundError(f"Cobra file not found: {cobra_path}")

    print(f"\n=== Processing {program_name} from {cobra_filename} ===")

    raw = pd.read_excel(cobra_path)
    cobra = normalize_cobra_columns(raw)

    evdf = compute_ev_timeseries(cobra)
    metrics_df, ctd_date, lsd_date = build_program_metric_table(evdf)
    sub_metrics_df = build_subteam_metric_table(cobra, ctd_date, lsd_date)
    labor_df = build_labor_table(cobra)
    manpower_df = build_manpower_table(cobra)

    # Plot
    ev_plot_png = os.path.join(OUTPUT_DIR, f"{program_name}_EV_Plot.png")
    make_ev_plot(evdf, program_name, ev_plot_png)

    # Tables workbook
    tables_xlsx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Tables.xlsx")
    with pd.ExcelWriter(tables_xlsx, engine="xlsxwriter") as writer:
        evdf.to_excel(writer, sheet_name="EV_Series", index=False)
        metrics_df.to_excel(writer, sheet_name="Program_Metrics", index=False)
        sub_metrics_df.to_excel(writer, sheet_name="Subteam_EVMS", index=False)
        labor_df.to_excel(writer, sheet_name="Subteam_Labor", index=False)
        manpower_df.to_excel(writer, sheet_name="Program_Manpower", index=False)

    # Deck using Theme.pptx
    if os.path.exists(THEME_PATH):
        prs = Presentation(THEME_PATH)
    else:
        prs = Presentation()

    # 1) Overview slide
    add_overview_slide(prs, program_name, ev_plot_png, metrics_df, ctd_date, lsd_date)

    # 2) Sub-team EVMS detail slides
    n_ev_pages = max(1, math.ceil(len(sub_metrics_df) / SUBTEAMS_PER_SLIDE))
    for page_idx in range(n_ev_pages):
        add_subteam_evms_slide(prs, program_name, sub_metrics_df,
                               page_idx, n_ev_pages)

    # 3) Sub-team labor & manpower slides
    n_labor_pages = max(1, math.ceil(len(labor_df) / SUBTEAMS_PER_SLIDE))
    for page_idx in range(n_labor_pages):
        add_labor_manpower_slide(prs, program_name, labor_df, manpower_df,
                                 page_idx, n_labor_pages)

    out_pptx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Deck.pptx")
    prs.save(out_pptx)

    print(f"✓ CTD date: {ctd_date.date()}, LSD date: {lsd_date.date()}")
    print(f"✓ Saved tables: {tables_xlsx}")
    print(f"✓ Saved deck:   {out_pptx}")


# ------------------------------------------------------------
# Run all standard-format programs
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