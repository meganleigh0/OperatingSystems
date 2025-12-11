# ============================================================
# EVMS PowerPoint Generator – All Cobra Programs (Updated)
# ============================================================

import os
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# --------------------------
# CONFIG
# --------------------------
DATA_DIR   = "data"
OUTPUT_DIR = "EVMS_Output"
TEMPLATE_PPTX = os.path.join(DATA_DIR, "Theme.pptx")  # your GDLS template, if present

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ALL Cobra files you want to process
COBRA_FILES = [
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

# Derive a program name from file name automatically
def program_name_from_file(fname: str) -> str:
    base = os.path.splitext(fname)[0]
    base = base.replace("Cobra-", "").strip()
    base = base.replace(" ", "_")
    return base

# Build PROGRAM_CONFIG for ALL files
PROGRAM_CONFIG = {
    program_name_from_file(f): f
    for f in COBRA_FILES
}

# EV index thresholds (roughly your bands)
EV_BLUE_MIN   = 1.05
EV_GREEN_MIN  = 0.98
EV_YELLOW_MIN = 0.95

# Manpower % Var thresholds
MP_GREEN_MIN  = 0.90
MP_YELLOW_MIN = 0.80

# VAC thresholds as % of BAC
VAC_BLUE_MIN   = 0.05
VAC_GREEN_MIN  = 0.01
VAC_YELLOW_MIN = -0.01


# ============================================================
# Helpers – Cobra normalization
# ============================================================

def _find_column(df, candidates, contains_any=None):
    cols = list(df.columns)
    # exact candidates
    for cand in candidates:
        if cand in cols:
            return cand
    # substring search
    if contains_any:
        for col in cols:
            low = str(col).lower()
            if any(sub.lower() in low for sub in contains_any):
                return col
    return None


def standardize_costset(value):
    """
    Map a variety of Cobra COST-SET strings into standard codes:
      BCWS, BCWP, ACWP, ETC, other.
    """
    s = str(value).upper()
    if "BCWS" in s:
        return "BCWS"
    if "BCWP" in s:
        return "BCWP"
    if "ACWP" in s:
        return "ACWP"
    if "ETC" in s:
        return "ETC"
    return s  # leave others as-is


def normalize_cobra(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize different Cobra export formats to:
        SUBTEAM, COSTSET, DATE, HOURS
    """
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    subteam_col = _find_column(
        df,
        candidates=["SUB_TEAM", "SubTeam", "SUBTEAM", "SUB TEAM"],
        contains_any=["sub", "team"]
    )
    costset_col = _find_column(
        df,
        candidates=["COST-SET", "COSTSET", "PLUG COST-SET"],
        contains_any=["cost", "plug"]
    )
    date_col = _find_column(
        df,
        candidates=["DATE", "DATE_HOURS", "DATE HOURS"],
        contains_any=["date"]
    )
    hours_col = _find_column(
        df,
        candidates=["HOURS", "HRS", "ACWP_HRS", "ACWP_HOURS"],
        contains_any=["hour"]
    )

    if any(c is None for c in [subteam_col, costset_col, date_col, hours_col]):
        missing = [name for name, val in [
            ("SUBTEAM", subteam_col),
            ("COSTSET", costset_col),
            ("DATE", date_col),
            ("HOURS", hours_col),
        ] if val is None]
        raise ValueError(f"Could not normalize Cobra file – missing logical columns: {missing}")

    # Rename to standard
    rename_map = {
        subteam_col: "SUBTEAM",
        costset_col: "COSTSET",
        date_col: "DATE",
        hours_col: "HOURS",
    }
    df = df.rename(columns=rename_map)

    # Standardize COSTSET
    df["COSTSET"] = df["COSTSET"].map(standardize_costset)

    # Types
    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
    df["HOURS"] = pd.to_numeric(df["HOURS"], errors="coerce")

    # Drop unusable rows
    df = df.dropna(subset=["DATE", "HOURS", "COSTSET"])

    return df[["SUBTEAM", "COSTSET", "DATE", "HOURS"]]


# ============================================================
# EV time-series from Cobra
# ============================================================

def compute_ev_timeseries(cobra_df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute EV time-series from normalized Cobra data.
    Returns a DataFrame indexed by month-end with:
        BCWS, BCWP, ACWP,
        CPI_month, SPI_month,
        CPI_cum, SPI_cum
    """
    df = cobra_df.copy()

    df = df[df["COSTSET"].isin(["BCWS", "BCWP", "ACWP"])]
    if df.empty:
        raise ValueError("No BCWS/BCWP/ACWP rows found after filtering COSTSET")

    df = df.set_index(pd.to_datetime(df["DATE"]))
    monthly = (
        df.groupby("COSTSET")
          .resample("ME")["HOURS"]  # 'ME' = month-end, avoids deprecated 'M'
          .sum()
          .unstack(0)
          .sort_index()
    )

    for col in ["BCWS", "BCWP", "ACWP"]:
        if col not in monthly.columns:
            monthly[col] = 0.0

    monthly = monthly[["BCWS", "BCWP", "ACWP"]].fillna(0.0)

    monthly["CPI_month"] = np.where(
        monthly["ACWP"] > 0, monthly["BCWP"] / monthly["ACWP"], np.nan
    )
    monthly["SPI_month"] = np.where(
        monthly["BCWS"] > 0, monthly["BCWP"] / monthly["BCWS"], np.nan
    )

    monthly["BCWS_cum"] = monthly["BCWS"].cumsum()
    monthly["BCWP_cum"] = monthly["BCWP"].cumsum()
    monthly["ACWP_cum"] = monthly["ACWP"].cumsum()

    monthly["CPI_cum"] = np.where(
        monthly["ACWP_cum"] > 0,
        monthly["BCWP_cum"] / monthly["ACWP_cum"],
        np.nan,
    )
    monthly["SPI_cum"] = np.where(
        monthly["BCWS_cum"] > 0,
        monthly["BCWP_cum"] / monthly["BCWS_cum"],
        np.nan,
    )

    return monthly


# ============================================================
# Plotting
# ============================================================

def clipped(series, lo=0.75, hi=1.25):
    return series.where((series >= lo) & (series <= hi), np.nan)


def create_evms_plot(ev_df, program_name, out_png):
    """
    EVMS trend plot with bands, 0.75–1.25 range, outliers removed.
    """
    if ev_df.empty:
        raise ValueError("EV time-series is empty")

    x = ev_df.index.to_timestamp() if isinstance(ev_df.index, pd.PeriodIndex) else ev_df.index

    CPI_m = clipped(ev_df["CPI_month"])
    SPI_m = clipped(ev_df["SPI_month"])
    CPI_c = clipped(ev_df["CPI_cum"])
    SPI_c = clipped(ev_df["SPI_cum"])

    fig, ax = plt.subplots(figsize=(8, 5))

    # bands
    ax.axhspan(0.75, 0.95, facecolor="#ffcccc", alpha=0.6)   # red
    ax.axhspan(0.95, 0.98, facecolor="#fff2cc", alpha=0.6)   # yellow
    ax.axhspan(0.98, 1.05, facecolor="#c6efce", alpha=0.6)   # green
    ax.axhspan(1.05, 1.25, facecolor="#c9daf8", alpha=0.6)   # blue

    ax.scatter(x, CPI_m, s=20, label="Monthly CPI", color="black")
    ax.scatter(x, SPI_m, s=20, label="Monthly SPI", color="gold")

    ax.plot(x, CPI_c, label="Cumulative CPI", linewidth=2, color="blue")
    ax.plot(x, SPI_c, label="Cumulative SPI", linewidth=2, color="dimgray")

    ax.set_ylim(0.75, 1.25)
    ax.set_ylabel("EV Indices")
    ax.set_xlabel("Month")
    ax.set_title(f"{program_name} EVMS Trend Overview")
    ax.legend(loc="upper left", fontsize=8)
    ax.grid(True, axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_png, dpi=150)
    plt.close(fig)


# ============================================================
# Color helpers
# ============================================================

def ev_index_color(value):
    if pd.isna(value):
        return None
    if value >= EV_BLUE_MIN:
        return RGBColor(0, 112, 192)
    if value >= EV_GREEN_MIN:
        return RGBColor(0, 176, 80)
    if value >= EV_YELLOW_MIN:
        return RGBColor(255, 192, 0)
    return RGBColor(192, 0, 0)


def vac_color(vac, bac):
    if pd.isna(vac) or pd.isna(bac) or bac == 0:
        return None
    pct = vac / bac
    if pct >= VAC_BLUE_MIN:
        return RGBColor(0, 112, 192)
    if pct >= VAC_GREEN_MIN:
        return RGBColor(0, 176, 80)
    if pct >= VAC_YELLOW_MIN:
        return RGBColor(255, 192, 0)
    return RGBColor(192, 0, 0)


def manpower_var_color(var_ratio):
    if pd.isna(var_ratio):
        return None
    if var_ratio >= MP_GREEN_MIN:
        return RGBColor(0, 176, 80)
    if var_ratio >= MP_YELLOW_MIN:
        return RGBColor(255, 192, 0)
    return RGBColor(192, 0, 0)


# ============================================================
# Subteam tables
# ============================================================

def build_subteam_metric_table(cobra_df, ev_df, curr_date):
    dates = ev_df.index[ev_df.index <= curr_date]
    if len(dates) == 0:
        raise ValueError("No EV dates <= CTD")
    ctd_date = dates.max()
    lsd_date = dates[dates < ctd_date].max() if len(dates) > 1 else ctd_date

    row_ctd = ev_df.loc[ctd_date]
    row_lsd = ev_df.loc[lsd_date]

    spi_ctd = row_ctd["SPI_cum"]
    cpi_ctd = row_ctd["CPI_cum"]
    spi_lsd = row_lsd["SPI_cum"]
    cpi_lsd = row_lsd["CPI_cum"]

    subteams = sorted(cobra_df["SUBTEAM"].dropna().unique())
    rows = []
    for st in subteams:
        rows.append(
            {
                "Sub Team": st,
                "SPI LSD": spi_lsd,
                "SPI CTD": spi_ctd,
                "CPI LSD": cpi_lsd,
                "CPI CTD": cpi_ctd,
                "Comments / Root Cause & Corrective Actions": "",
            }
        )

    return pd.DataFrame(rows)


def build_labor_manpower_tables(cobra_df):
    df = cobra_df.copy()
    agg = df.pivot_table(
        index="SUBTEAM",
        columns="COSTSET",
        values="HOURS",
        aggfunc="sum",
        fill_value=0.0,
    )

    bac  = agg.get("BCWS", pd.Series(0.0, index=agg.index))
    acwp = agg.get("ACWP", pd.Series(0.0, index=agg.index))
    etc  = agg.get("ETC",  pd.Series(0.0, index=agg.index))

    eac = acwp + etc
    vac = bac - eac

    labor_rows = []
    for st in agg.index:
        labor_rows.append(
            {
                "Sub Team": st,
                "BAC": float(bac.loc[st]),
                "EAC": float(eac.loc[st]),
                "VAC": float(vac.loc[st]),
                "Comments / Root Cause & Corrective Actions": "",
            }
        )
    labor_df = pd.DataFrame(labor_rows)

    num_cols = ["BAC", "EAC", "VAC"]
    mask_nonzero = (labor_df[num_cols].abs().sum(axis=1) > 0)
    labor_df = labor_df[mask_nonzero].reset_index(drop=True)

    demand_hours = bac.sum()
    actual_hours = acwp.sum()
    pct_var = actual_hours / demand_hours if demand_hours > 0 else np.nan

    manpower_df = pd.DataFrame(
        [
            {
                "Demand Hours": float(demand_hours),
                "Actual Hours": float(actual_hours),
                "% Var": pct_var,
                "Next Mo BCWS Hours": 0.0,
                "Next Mo ETC Hours": 0.0,
                "Comments / Root Cause & Corrective Actions": "",
            }
        ]
    )

    return labor_df, manpower_df


# ============================================================
# PPTX helpers (with robust layout + formatting tweaks)
# ============================================================

def get_title_layout_index(prs: Presentation) -> int:
    """
    Try to find a simple title-only or title slide layout.
    Fallback to 0 if we can't guess.
    """
    for idx, layout in enumerate(prs.slide_layouts):
        name = layout.name.lower()
        if "title" in name and "content" not in name:
            return idx
    return 0  # safe default


def add_title_slide(prs: Presentation, text: str):
    layout_idx = get_title_layout_index(prs)
    slide_layout = prs.slide_layouts[layout_idx]
    slide = prs.slides.add_slide(slide_layout)
    if slide.shapes.title:
        slide.shapes.title.text = text
    else:
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
        p = tb.text_frame.paragraphs[0]
        p.text = text
        p.font.size = Pt(24)
    return slide


def autofit_table_font(tbl, size=11):
    for row in tbl.rows:
        for cell in row.cells:
            if not cell.text_frame:
                continue
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(size)


def add_metric_table(slide, metrics_df):
    rows, cols = metrics_df.shape[0] + 1, metrics_df.shape[1]

    left = Inches(6.2)
    top = Inches(1.5)
    width = Inches(3.3)   # a bit wider
    height = Inches(1.6)

    shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    tbl = shape.table

    # headers
    for j, col in enumerate(metrics_df.columns):
        tbl.cell(0, j).text = col

    # rows
    for i, (_, row) in enumerate(metrics_df.iterrows(), start=1):
        for j, col in enumerate(metrics_df.columns):
            val = row[col]
            if isinstance(val, (float, int)):
                tbl.cell(i, j).text = f"{val:.3f}"
            else:
                tbl.cell(i, j).text = "" if pd.isna(val) else str(val)

    autofit_table_font(tbl, size=10)

    # widen Metric column more for readability
    tbl.columns[0].width = Inches(1.4)

    # color CTD/LSD
    ctd_col = metrics_df.columns.get_loc("CTD")
    lsd_col = metrics_df.columns.get_loc("LSD")
    for i in range(1, rows):
        for j in [ctd_col, lsd_col]:
            txt = tbl.cell(i, j).text
            try:
                val = float(txt)
            except ValueError:
                val = np.nan
            rgb = ev_index_color(val)
            if rgb is not None:
                cell = tbl.cell(i, j)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb


def add_evms_overview_slide(prs, program_name, ev_df, plot_png):
    dates = ev_df.index
    ctd_date = dates.max()
    lsd_date = dates[dates < ctd_date].max() if len(dates) > 1 else ctd_date

    row_ctd = ev_df.loc[ctd_date]
    row_lsd = ev_df.loc[lsd_date]

    # Ignore BEI for now (leave blank, to be wired to Penske later)
    metrics_df = pd.DataFrame(
        [
            {"Metric": "SPI", "CTD": row_ctd["SPI_cum"], "LSD": row_lsd["SPI_cum"], "Comments / Root Cause & Corrective Actions": ""},
            {"Metric": "CPI", "CTD": row_ctd["CPI_cum"], "LSD": row_lsd["CPI_cum"], "Comments / Root Cause & Corrective Actions": ""},
            {"Metric": "BEI", "CTD": np.nan,            "LSD": np.nan,            "Comments / Root Cause & Corrective Actions": ""},
        ]
    )

    slide = add_title_slide(prs, f"{program_name} EVMS Trend Overview")

    # plot image
    left = Inches(0.6)
    top = Inches(1.3)
    slide.shapes.add_picture(plot_png, left, top, height=Inches(3.5))

    add_metric_table(slide, metrics_df[["Metric", "CTD", "LSD", "Comments / Root Cause & Corrective Actions"]])


def chunk_list(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i: i + n]


def add_subteam_metric_slides(prs, program_name, metrics_df):
    for chunk in chunk_list(metrics_df, 15):  # 15 per slide
        slide = add_title_slide(prs, f"{program_name} EVMS Detail – Sub Team CPI / SPI Metrics")

        rows, cols = chunk.shape[0] + 1, chunk.shape[1]
        left = Inches(0.6)
        top  = Inches(1.3)
        width = Inches(9.0)
        height = Inches(4.0)

        shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        tbl = shape.table

        for j, col in enumerate(chunk.columns):
            tbl.cell(0, j).text = col

        for i, (_, row) in enumerate(chunk.iterrows(), start=1):
            for j, col in enumerate(chunk.columns):
                val = row[col]
                if isinstance(val, (float, int)):
                    tbl.cell(i, j).text = f"{val:.3f}"
                else:
                    tbl.cell(i, j).text = "" if pd.isna(val) else str(val)

        autofit_table_font(tbl, size=9)

        for i in range(1, rows):
            for col_name in ["SPI LSD", "SPI CTD", "CPI LSD", "CPI CTD"]:
                j = chunk.columns.get_loc(col_name)
                txt = tbl.cell(i, j).text
                try:
                    v = float(txt)
                except ValueError:
                    v = np.nan
                rgb = ev_index_color(v)
                if rgb is not None:
                    cell = tbl.cell(i, j)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = rgb


def add_labor_manpower_slides(prs, program_name, labor_df, manpower_df):
    for chunk in chunk_list(labor_df, 15):  # 15 subteams per slide
        slide = add_title_slide(prs, f"{program_name} EVMS Detail – Sub Team Labor & Manpower")

        # main labor table
        rows, cols = chunk.shape[0] + 1, chunk.shape[1]
        left = Inches(0.6)
        top  = Inches(1.3)
        width = Inches(9.0)
        height = Inches(4.0)

        shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        tbl = shape.table

        for j, col in enumerate(chunk.columns):
            tbl.cell(0, j).text = col

        for i, (_, row) in enumerate(chunk.iterrows(), start=1):
            for j, col in enumerate(chunk.columns):
                val = row[col]
                if isinstance(val, (float, int)):
                    if col in ["BAC", "EAC", "VAC"]:
                        tbl.cell(i, j).text = f"{val:,.1f}"
                    else:
                        tbl.cell(i, j).text = f"{val:.3f}"
                else:
                    tbl.cell(i, j).text = "" if pd.isna(val) else str(val)

        autofit_table_font(tbl, size=9)

        # VAC colors
        bac_idx = chunk.columns.get_loc("BAC")
        vac_idx = chunk.columns.get_loc("VAC")
        for i in range(1, rows):
            try:
                bac_val = float(tbl.cell(i, bac_idx).text.replace(",", ""))
            except ValueError:
                bac_val = np.nan
            try:
                vac_val = float(tbl.cell(i, vac_idx).text.replace(",", ""))
            except ValueError:
                vac_val = np.nan
            rgb = vac_color(vac_val, bac_val)
            if rgb is None and not pd.isna(vac_val):
                rgb = RGBColor(221, 235, 247)
            if rgb is not None:
                cell = tbl.cell(i, vac_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb

        # Program Manpower table – LOWER so it doesn't overlap
        mp_left = Inches(0.6)
        mp_top  = Inches(5.4)   # <<< slightly lower than before
        mp_width = Inches(9.0)
        mp_height = Inches(1.0)

        mp_rows, mp_cols = manpower_df.shape[0] + 1, manpower_df.shape[1]
        mp_shape = slide.shapes.add_table(mp_rows, mp_cols, mp_left, mp_top, mp_width, mp_height)
        mp_tbl = mp_shape.table

        for j, col in enumerate(manpower_df.columns):
            mp_tbl.cell(0, j).text = col

        for i, (_, row) in enumerate(manpower_df.iterrows(), start=1):
            for j, col in enumerate(manpower_df.columns):
                val = row[col]
                if isinstance(val, (float, int)):
                    if col in ["Demand Hours", "Actual Hours"]:
                        mp_tbl.cell(i, j).text = f"{val:,.1f}"
                    elif col == "% Var":
                        mp_tbl.cell(i, j).text = f"{val:.2%}"
                    else:
                        mp_tbl.cell(i, j).text = f"{val:,.1f}"
                else:
                    mp_tbl.cell(i, j).text = "" if pd.isna(val) else str(val)

        autofit_table_font(mp_tbl, size=9)

        # % Var colors
        var_idx = manpower_df.columns.get_loc("% Var")
        for i in range(1, mp_rows):
            txt = mp_tbl.cell(i, var_idx).text.replace("%", "").strip()
            try:
                val = float(txt) / 100.0 if "%" in mp_tbl.cell(i, var_idx).text else float(txt)
            except ValueError:
                val = np.nan
            rgb = manpower_var_color(val)
            if rgb is not None:
                cell = mp_tbl.cell(i, var_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb


# ============================================================
# Main driver
# ============================================================

def process_program(program_name, cobra_file):
    print(f"\n=== Processing {program_name} from {cobra_file} ===")
    path = os.path.join(DATA_DIR, cobra_file)
    if not os.path.exists(path):
        print(f"  >> Skipping – file not found: {path}")
        return

    raw = pd.read_excel(path)
    cobra = normalize_cobra(raw)

    ev = compute_ev_timeseries(cobra)

    plot_png = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Plot.png")
    create_evms_plot(ev, program_name, plot_png)

    curr_date = ev.index.max()
    metrics_sub = build_subteam_metric_table(cobra, ev, curr_date)
    labor_df, manpower_df = build_labor_manpower_tables(cobra)

    # start presentation
    if os.path.exists(TEMPLATE_PPTX):
        prs = Presentation(TEMPLATE_PPTX)
    else:
        prs = Presentation()

    add_evms_overview_slide(prs, program_name, ev, plot_png)
    add_subteam_metric_slides(prs, program_name, metrics_sub)
    add_labor_manpower_slides(prs, program_name, labor_df, manpower_df)

    # Save tables
    tables_xlsx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Tables.xlsx")
    with pd.ExcelWriter(tables_xlsx, engine="xlsxwriter") as writer:
        ev.to_excel(writer, sheet_name="EV_Series")
        metrics_sub.to_excel(writer, sheet_name="Subteam_Metrics", index=False)
        labor_df.to_excel(writer, sheet_name="Subteam_Labor", index=False)
        manpower_df.to_excel(writer, sheet_name="Program_Manpower", index=False)

    # Save deck
    out_pptx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Deck.pptx")
    prs.save(out_pptx)

    print(f"  ✓ Saved tables: {tables_xlsx}")
    print(f"  ✓ Saved deck:   {out_pptx}")


# Run for all config’d programs (ALL Cobra files)
for program, cobra_file in PROGRAM_CONFIG.items():
    try:
        process_program(program, cobra_file)
    except Exception as e:
        print(f"!! Error for {program}: {e}")

print("\nALL PROGRAM EVMS DECKS COMPLETE ✓")