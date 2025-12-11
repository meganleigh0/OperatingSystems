import os
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER

# =========================================================
# CONFIG
# =========================================================

COBRA_DIR   = "data"
PENSKE_PATH = os.path.join("data", "OpenPlan_Activity-Penske.xlsx")
THEME_PATH  = os.path.join("data", "Theme.pptx")
OUTPUT_DIR  = "EVMS_Output"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Only programs that share the “standard” Cobra format
PROGRAM_CONFIG = {
    "Abrams_STS_2022": "Cobra-Abrams STS 2022.xlsx",
    "Abrams_STS"     : "Cobra-Abrams STS.xlsx",
    "ARV"            : "Cobra-ARV.xlsx",
    "ARV30"          : "Cobra-ARV30.xlsx",
    "Stryker_Bulgaria_150": "Cobra-Stryker Bulgaria 150.xlsx",
    "XM30"           : "Cobra-XM30.xlsx",
}

# EV index bands and colors
YMIN, YMAX = 0.75, 1.25
RED_MIN, RED_MAX       = 0.90, 0.95
YELLOW_MIN, YELLOW_MAX = 0.95, 0.98
GREEN_MIN, GREEN_MAX   = 0.98, 1.05
BLUE_MIN, BLUE_MAX     = 1.05, 1.25

# Color constants for tables (RGB)
COLOR_RED   = RGBColor(192,   0,   0)
COLOR_YELLOW= RGBColor(255, 192,   0)
COLOR_GREEN = RGBColor(  0, 176,  80)
COLOR_BLUE  = RGBColor( 68, 114, 196)
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_GRAY  = RGBColor(242, 242, 242)
COLOR_HEADER_BG = RGBColor(31,  73, 125)
COLOR_HEADER_TX = RGBColor(255,255,255)

# =========================================================
# GENERAL HELPERS
# =========================================================

def find_col(cols, *keywords):
    """
    Find first column whose name contains all keywords (case/space insensitive).
    """
    for c in cols:
        name = c.replace(" ", "").upper()
        if all(k.upper() in name for k in keywords):
            return c
    return None

def normalize_cobra_standard(path):
    """
    Load a Cobra file that follows the 'standard' format used by XM30 / Abrams / etc.
    Returns a DataFrame with columns: SUBTEAM, COSTSET, DATE, HOURS.
    Raises ValueError if required columns not found.
    """
    df = pd.read_excel(path)

    subteam_col = find_col(df.columns, "SUB", "TEAM")
    costset_col = find_col(df.columns, "COST", "SET")
    date_col    = find_col(df.columns, "DATE")
    hours_col   = find_col(df.columns, "HOUR")

    missing = [n for n, c in
               [("SUBTEAM",subteam_col),("COSTSET",costset_col),
                ("DATE",date_col),("HOURS",hours_col)]
               if c is None]
    if missing:
        raise ValueError(f"Missing logical columns in Cobra file: {missing}")

    df = df[[subteam_col, costset_col, date_col, hours_col]].copy()
    df.columns = ["SUBTEAM", "COSTSET", "DATE", "HOURS"]

    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
    df = df.dropna(subset=["SUBTEAM","COSTSET","DATE","HOURS"])

    # Ensure numeric hours
    df["HOURS"] = pd.to_numeric(df["HOURS"], errors="coerce")
    df = df.dropna(subset=["HOURS"])

    # Standardize cost-set labels
    df["COSTSET"] = df["COSTSET"].astype(str).str.strip().str.upper()

    return df

# =========================================================
# EVMS CALCULATIONS
# =========================================================

def compute_ev_timeseries(cobra_df):
    """
    Compute monthly & cumulative SPI/CPI timeseries from normalized Cobra df.
    Returns DataFrame indexed by month (datetime), with columns:
    BCWS, BCWP, ACWP, SPI_M, CPI_M, SPI_CUM, CPI_CUM
    """
    ev = cobra_df[cobra_df["COSTSET"].isin(["BCWS","BCWP","ACWP"])].copy()
    if ev.empty:
        raise ValueError("No BCWS/BCWP/ACWP rows found for this dataset")

    ev = ev.groupby(["DATE","COSTSET"])["HOURS"].sum().unstack("COSTSET")
    # Resample to month-start frequency and sum within month
    ev = ev.resample("MS").sum()

    # Cumulative sums
    cum = ev.cumsum()

    bcws = ev.get("BCWS")
    bcwp = ev.get("BCWP")
    acwp = ev.get("ACWP")

    # Monthly indices
    spi_m = bcwp / bcws.replace(0, np.nan)
    cpi_m = bcwp / acwp.replace(0, np.nan)

    # Cumulative indices
    spi_c = cum.get("BCWP") / cum.get("BCWS").replace(0, np.nan)
    cpi_c = cum.get("BCWP") / cum.get("ACWP").replace(0, np.nan)

    evdf = pd.DataFrame({
        "BCWS": bcws,
        "BCWP": bcwp,
        "ACWP": acwp,
        "SPI_M": spi_m,
        "CPI_M": cpi_m,
        "SPI_CUM": spi_c,
        "CPI_CUM": cpi_c,
    }).dropna(how="all")

    return evdf

def get_status_date(evdf):
    """
    Simple CTD/LSD status date using latest non-null SPI_CUM.
    For now CTD and LSD are the same (can be refined when Penske mapping is added).
    """
    valid = evdf["SPI_CUM"].dropna()
    if valid.empty:
        raise ValueError("No valid SPI_CUM values to determine status date")
    ctd_date = valid.index[-1]
    lsd_date = ctd_date
    return ctd_date, lsd_date

def build_program_metric_table(evdf, ctd_date, lsd_date):
    """
    Program-level SPI/CPI CTD/LSD table (BEI omitted for now).
    """
    def get_val(series, dt):
        try:
            return float(series.loc[dt])
        except Exception:
            return np.nan

    spi_ctd = get_val(evdf["SPI_CUM"], ctd_date)
    cpi_ctd = get_val(evdf["CPI_CUM"], ctd_date)
    spi_lsd = get_val(evdf["SPI_CUM"], lsd_date)
    cpi_lsd = get_val(evdf["CPI_CUM"], lsd_date)

    metric_df = pd.DataFrame({
        "Metric": ["SPI","CPI"],
        "CTD": [spi_ctd, cpi_ctd],
        "LSD": [spi_lsd, cpi_lsd],
        "Comments / Root Cause & Corrective Actions": ["",""],
    })

    return metric_df

def build_subteam_metric_table(cobra_df, ctd_date, lsd_date):
    """
    Sub-team level SPI/CPI CTD/LSD (no BEI right now).
    Returns DataFrame: SUBTEAM, SPI_CTD, SPI_LSD, CPI_CTD, CPI_LSD, Comments...
    """
    ev = cobra_df[cobra_df["COSTSET"].isin(["BCWS","BCWP","ACWP"])].copy()
    if ev.empty:
        return pd.DataFrame(columns=[
            "Sub Team","SPI CTD","SPI LSD","CPI CTD","CPI LSD",
            "Comments / Root Cause & Corrective Actions"
        ])

    ev["DATE"] = pd.to_datetime(ev["DATE"])

    ctd_mask = ev["DATE"] <= ctd_date
    lsd_mask = ev["DATE"] <= lsd_date

    grp_ctd = ev[ctd_mask].groupby(["SUBTEAM","COSTSET"])["HOURS"].sum().unstack("COSTSET")
    grp_lsd = ev[lsd_mask].groupby(["SUBTEAM","COSTSET"])["HOURS"].sum().unstack("COSTSET")

    subteams = sorted(ev["SUBTEAM"].unique())
    rows = []
    for st in subteams:
        g_ctd = grp_ctd.loc[st] if st in grp_ctd.index else pd.Series(dtype=float)
        g_lsd = grp_lsd.loc[st] if st in grp_lsd.index else pd.Series(dtype=float)

        bcwp_ctd = g_ctd.get("BCWP", np.nan)
        bcws_ctd = g_ctd.get("BCWS", np.nan)
        acwp_ctd = g_ctd.get("ACWP", np.nan)

        bcwp_lsd = g_lsd.get("BCWP", np.nan)
        bcws_lsd = g_lsd.get("BCWS", np.nan)
        acwp_lsd = g_lsd.get("ACWP", np.nan)

        spi_ctd = bcwp_ctd / bcws_ctd if bcws_ctd not in (0, np.nan) else np.nan
        cpi_ctd = bcwp_ctd / acwp_ctd if acwp_ctd not in (0, np.nan) else np.nan
        spi_lsd = bcwp_lsd / bcws_lsd if bcws_lsd not in (0, np.nan) else np.nan
        cpi_lsd = bcwp_lsd / acwp_lsd if acwp_lsd not in (0, np.nan) else np.nan

        if all(pd.isna([spi_ctd, cpi_ctd, spi_lsd, cpi_lsd])):
            # Skip totally empty subteams
            continue

        rows.append([st, spi_ctd, spi_lsd, cpi_ctd, cpi_lsd, ""])

    sub_df = pd.DataFrame(rows, columns=[
        "Sub Team","SPI CTD","SPI LSD","CPI CTD","CPI LSD",
        "Comments / Root Cause & Corrective Actions"
    ])

    return sub_df

def build_labor_table(cobra_df):
    """
    Sub-team labor table with BAC, EAC, VAC:
    EAC = ACWP + ETC, VAC = BAC - EAC
    """
    labor = cobra_df[cobra_df["COSTSET"].isin(["BAC","ACWP","ETC"])].copy()
    if labor.empty:
        return pd.DataFrame(columns=["Sub Team","BAC","EAC","VAC",
                                     "Comments / Root Cause & Corrective Actions"])

    pivot = labor.pivot_table(index=["SUBTEAM","COSTSET"], values="HOURS", aggfunc="sum").unstack("COSTSET")
    pivot.columns = pivot.columns.get_level_values(1)

    bac = pivot.get("BAC", pd.Series(dtype=float))
    acwp = pivot.get("ACWP", pd.Series(dtype=float)).fillna(0.0)
    etc  = pivot.get("ETC",  pd.Series(dtype=float)).fillna(0.0)

    eac = acwp + etc
    vac = bac - eac

    df = pd.DataFrame({
        "Sub Team": pivot.index,
        "BAC": bac,
        "EAC": eac,
        "VAC": vac,
        "Comments / Root Cause & Corrective Actions": ""
    })

    return df

def build_manpower_table(penske_df, program_name):
    """
    Program Manpower summary from Penske data.
    If Penske file missing or no rows for program, returns NaNs but still
    produces one-row table so layout doesn't break.
    """
    if penske_df is None:
        return pd.DataFrame([{
            "Demand Hours": np.nan,
            "Actual Hours": np.nan,
            "% Var": np.nan,
            "Next Mo BCWS Hours": np.nan,
            "Next Mo ETC Hours": np.nan,
            "Comments / Root Cause & Corrective Actions": ""
        }])

    df = penske_df.copy()
    if "Program" in df.columns:
        prog_col = "Program"
    elif "PROGRAM" in df.columns:
        prog_col = "PROGRAM"
    else:
        prog_col = None

    if prog_col is not None:
        pdf = df[df[prog_col].astype(str).str.upper().str.contains(program_name.upper())]
    else:
        pdf = df

    if pdf.empty:
        return pd.DataFrame([{
            "Demand Hours": np.nan,
            "Actual Hours": np.nan,
            "% Var": np.nan,
            "Next Mo BCWS Hours": np.nan,
            "Next Mo ETC Hours": np.nan,
            "Comments / Root Cause & Corrective Actions": ""
        }])

    # These column names depend on your Penske export; adjust if needed.
    # For now assume:
    #   'Demand Hours' (planned), 'Hours' (actual), 'Next_Month_BCWS', 'Next_Month_ETC'
    demand_col = find_col(pdf.columns, "DEMAND","HOUR")
    actual_col = find_col(pdf.columns, "HOUR")  # actual
    bcws_col   = find_col(pdf.columns, "BCWS")
    etc_col    = find_col(pdf.columns, "ETC")

    demand = pdf[demand_col].sum() if demand_col else np.nan
    actual = pdf[actual_col].sum() if actual_col else np.nan
    bcws   = pdf[bcws_col].sum()   if bcws_col   else np.nan
    etc    = pdf[etc_col].sum()    if etc_col    else np.nan

    pct_var = actual / demand if demand not in (0, np.nan) else np.nan

    mp = pd.DataFrame([{
        "Demand Hours": demand,
        "Actual Hours": actual,
        "% Var": pct_var,
        "Next Mo BCWS Hours": bcws,
        "Next Mo ETC Hours": etc,
        "Comments / Root Cause & Corrective Actions": ""
    }])

    return mp

# =========================================================
# COLOR CODING HELPERS
# =========================================================

def ev_index_color(val):
    """Color for CPI/SPI based on 0.90 / 0.95 / 0.98 / 1.05 thresholds."""
    if pd.isna(val):
        return None
    if val >= BLUE_MIN:
        return COLOR_BLUE
    elif val >= GREEN_MIN:
        return COLOR_GREEN
    elif val >= YELLOW_MIN:
        return COLOR_YELLOW
    elif val >= RED_MIN:
        return COLOR_RED
    else:
        return COLOR_RED

def vac_color(vac, bac):
    """
    Color for VAC/BAC:
        Blue  >= +5%
        Green +1% to +5%
        Yellow -1% to +1%
        Red   < -1%
    """
    if pd.isna(vac) or pd.isna(bac) or bac == 0:
        return None
    ratio = vac / bac
    if ratio >= 0.05:
        return COLOR_BLUE
    elif ratio >= 0.01:
        return COLOR_GREEN
    elif ratio >= -0.01:
        return COLOR_YELLOW
    else:
        return COLOR_RED

def manpower_var_color(pct):
    """
    Color for manpower % Var:
        Green  90–105%
        Yellow 85–90% or 105–110%
        Red    <85% or >110%
    """
    if pd.isna(pct):
        return None
    if 0.90 <= pct <= 1.05:
        return COLOR_GREEN
    elif 0.85 <= pct < 0.90 or 1.05 < pct <= 1.10:
        return COLOR_YELLOW
    else:
        return COLOR_RED

# =========================================================
# PLOTTING
# =========================================================

def make_evms_plot(evdf, program_name, out_png):
    """
    Create EVMS trend plot with color bands and SPI/CPI monthly + cumulative.
    """
    plot_df = evdf.copy()
    # Clip for plotting only (keeps cumulative smooth but avoids crazy spikes)
    for col in ["SPI_M","CPI_M","SPI_CUM","CPI_CUM"]:
        plot_df[col] = plot_df[col].clip(lower=YMIN, upper=YMAX)

    fig, ax = plt.subplots(figsize=(7, 4))

    # Color bands
    ax.axhspan(RED_MIN, RED_MAX,   color="#F4CCCC", zorder=0)   # red band
    ax.axhspan(YELLOW_MIN, YELLOW_MAX, color="#FFF2CC", zorder=0)
    ax.axhspan(GREEN_MIN, GREEN_MAX,   color="#D9EAD3", zorder=0)
    ax.axhspan(BLUE_MIN, YMAX,         color="#D0E0E3", zorder=0)

    # Monthly scatter
    ax.scatter(plot_df.index, plot_df["CPI_M"], s=18, color="black", label="Monthly CPI")
    ax.scatter(plot_df.index, plot_df["SPI_M"], s=18, color="gold",  label="Monthly SPI")

    # Cumulative lines
    ax.plot(plot_df.index, plot_df["CPI_CUM"], color="blue",  linewidth=1.8, label="Cumulative CPI")
    ax.plot(plot_df.index, plot_df["SPI_CUM"], color="gray",  linewidth=1.8, label="Cumulative SPI")

    ax.set_ylim(YMIN, YMAX)
    ax.set_ylabel("EV Indices")
    ax.set_xlabel("Month")
    ax.set_title(f"{program_name} EVMS Trend Overview")

    ax.grid(True, axis="y", alpha=0.3)
    ax.legend(loc="upper left", fontsize=8)

    fig.tight_layout()
    fig.savefig(out_png, dpi=200)
    plt.close(fig)

# =========================================================
# POWERPOINT HELPERS
# =========================================================

def keep_only_title_placeholder(slide):
    """
    Remove every placeholder except TITLE / CENTER_TITLE so we don't see
    'Click to add text' boxes.
    """
    for shape in list(slide.shapes):
        if not shape.is_placeholder:
            continue
        pht = shape.placeholder_format.type
        if pht not in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            slide.shapes._spTree.remove(shape._element)

def style_header_cell(cell):
    cell.fill.solid()
    cell.fill.fore_color.rgb = COLOR_HEADER_BG
    cell.text_frame.clear()
    p = cell.text_frame.paragraphs[0]
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_HEADER_TX
    p.alignment = PP_ALIGN.CENTER

def style_body_cell(cell):
    cell.text_frame.paragraphs[0].font.size = Pt(10)

def add_program_overview_slide(prs, program_name, ev_plot_png, metric_df):
    """
    First slide: EVMS trend plot + program level SPI/CPI table.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # title + content
    keep_only_title_placeholder(slide)

    title = slide.shapes.title
    title.text = f"{program_name} EVMS Trend Overview"

    # Insert plot
    pic_left = Inches(0.6)
    pic_top  = Inches(1.6)
    pic_height = Inches(3.7)
    slide.shapes.add_picture(ev_plot_png, pic_left, pic_top, height=pic_height)

    # Metric table to the right
    rows = len(metric_df) + 1
    cols = 4
    table_left = Inches(6.0)
    table_top  = Inches(1.6)
    table_width = Inches(3.4)
    table_height = Inches(1.8)

    shape = slide.shapes.add_table(rows, cols, table_left, table_top,
                                   table_width, table_height)
    tbl = shape.table

    # Column widths – Comments column wider
    tbl.columns[0].width = Inches(0.9)
    tbl.columns[1].width = Inches(0.8)
    tbl.columns[2].width = Inches(0.8)
    tbl.columns[3].width = Inches(1.9)

    # Headers
    headers = ["Metric","CTD","LSD","Comments / Root Cause & Corrective Actions"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        style_header_cell(cell)
        cell.text = h

    # Body
    for i, (_, row) in enumerate(metric_df.iterrows(), start=1):
        tbl.cell(i, 0).text = str(row["Metric"])
        tbl.cell(i, 1).text = f"{row['CTD']:.3f}" if pd.notna(row["CTD"]) else ""
        tbl.cell(i, 2).text = f"{row['LSD']:.3f}" if pd.notna(row["LSD"]) else ""
        tbl.cell(i, 3).text = row["Comments / Root Cause & Corrective Actions"]

        for j in range(cols):
            style_body_cell(tbl.cell(i, j))

        # Color coding
        for j, metric_val in enumerate([row["CTD"], row["LSD"]], start=1):
            rgb = ev_index_color(metric_val)
            if rgb is not None:
                cell = tbl.cell(i, j)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb

def chunk_df(df, size):
    for start in range(0, len(df), size):
        yield df.iloc[start:start+size]

def add_subteam_metric_slides(prs, program_name, metrics_df, page_size=15):
    """
    Slides: Sub Team EVMS Metrics (SPI/CPI) – may span multiple pages.
    """
    if metrics_df.empty:
        return

    for page, df_part in enumerate(chunk_df(metrics_df, page_size), start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        keep_only_title_placeholder(slide)
        title = slide.shapes.title
        if len(metrics_df) > page_size:
            title.text = f"{program_name} EVMS Detail – Sub Team EVMS Metrics (Page {page})"
        else:
            title.text = f"{program_name} EVMS Detail – Sub Team EVMS Metrics"

        rows = len(df_part) + 1
        cols = 6
        left = Inches(0.6)
        top  = Inches(1.6)
        width = Inches(9.0)
        height = Inches(3.8)

        shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        tbl = shape.table

        # Column widths – last column (Comments) widest
        tbl.columns[0].width = Inches(1.1)  # Sub Team
        tbl.columns[1].width = Inches(0.8)
        tbl.columns[2].width = Inches(0.8)
        tbl.columns[3].width = Inches(0.8)
        tbl.columns[4].width = Inches(0.8)
        tbl.columns[5].width = Inches(2.7)

        headers = ["Sub Team","SPI CTD","SPI LSD","CPI CTD","CPI LSD",
                   "Comments / Root Cause & Corrective Actions"]
        for j, h in enumerate(headers):
            cell = tbl.cell(0,j)
            style_header_cell(cell)
            cell.text = h

        for i, (_, row) in enumerate(df_part.iterrows(), start=1):
            tbl.cell(i,0).text = str(row["Sub Team"])
            tbl.cell(i,1).text = f"{row['SPI CTD']:.3f}" if pd.notna(row["SPI CTD"]) else ""
            tbl.cell(i,2).text = f"{row['SPI LSD']:.3f}" if pd.notna(row["SPI LSD"]) else ""
            tbl.cell(i,3).text = f"{row['CPI CTD']:.3f}" if pd.notna(row["CPI CTD"]) else ""
            tbl.cell(i,4).text = f"{row['CPI LSD']:.3f}" if pd.notna(row["CPI LSD"]) else ""
            tbl.cell(i,5).text = row["Comments / Root Cause & Corrective Actions"]

            for j in range(cols):
                style_body_cell(tbl.cell(i,j))

            # Color code SPI/CPI cells
            for col_name, col_idx in [("SPI CTD",1),("SPI LSD",2),
                                      ("CPI CTD",3),("CPI LSD",4)]:
                rgb = ev_index_color(row[col_name])
                if rgb is not None:
                    cell = tbl.cell(i,col_idx)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = rgb

def add_labor_manpower_slides(prs, program_name, labor_df, manpower_df, page_size=15):
    """
    Slides: Sub Team Labor & Manpower (with Program Manpower at bottom).
    """
    if labor_df.empty:
        return

    for page, df_part in enumerate(chunk_df(labor_df, page_size), start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        keep_only_title_placeholder(slide)
        title = slide.shapes.title
        if len(labor_df) > page_size:
            title.text = f"{program_name} EVMS Detail – Sub Team Labor & Manpower (Page {page})"
        else:
            title.text = f"{program_name} EVMS Detail – Sub Team Labor & Manpower"

        # Top table: Sub Team BAC/EAC/VAC/Comments
        rows = len(df_part) + 1
        cols = 5
        left = Inches(0.6)
        top  = Inches(1.6)
        width = Inches(9.0)
        height = Inches(3.0)

        shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        tbl = shape.table

        tbl.columns[0].width = Inches(1.1)  # Sub Team
        tbl.columns[1].width = Inches(1.3)
        tbl.columns[2].width = Inches(1.3)
        tbl.columns[3].width = Inches(1.3)
        tbl.columns[4].width = Inches(3.0)  # Comments wider

        headers = ["Sub Team","BAC","EAC","VAC",
                   "Comments / Root Cause & Corrective Actions"]
        for j, h in enumerate(headers):
            cell = tbl.cell(0,j)
            style_header_cell(cell)
            cell.text = h

        for i, (_, row) in enumerate(df_part.iterrows(), start=1):
            tbl.cell(i,0).text = str(row["Sub Team"])
            tbl.cell(i,1).text = f"{row['BAC']:.1f}" if pd.notna(row["BAC"]) else ""
            tbl.cell(i,2).text = f"{row['EAC']:.1f}" if pd.notna(row["EAC"]) else ""
            tbl.cell(i,3).text = f"{row['VAC']:.1f}" if pd.notna(row["VAC"]) else ""
            tbl.cell(i,4).text = row["Comments / Root Cause & Corrective Actions"]

            for j in range(cols):
                style_body_cell(tbl.cell(i,j))

            rgb = vac_color(row["VAC"], row["BAC"])
            if rgb is not None:
                cell = tbl.cell(i,3)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb

        # Bottom Program Manpower table, lower so it doesn't overlap
        mp_rows = len(manpower_df) + 1
        mp_cols = 6
        mp_left = Inches(0.6)
        mp_top  = Inches(4.8)      # << lowered vs previous versions
        mp_width= Inches(9.0)
        mp_height = Inches(1.5)

        shape_mp = slide.shapes.add_table(mp_rows, mp_cols, mp_left, mp_top,
                                          mp_width, mp_height)
        mp_tbl = shape_mp.table

        mp_tbl.columns[0].width = Inches(1.5)
        mp_tbl.columns[1].width = Inches(1.5)
        mp_tbl.columns[2].width = Inches(1.0)
        mp_tbl.columns[3].width = Inches(1.5)
        mp_tbl.columns[4].width = Inches(1.5)
        mp_tbl.columns[5].width = Inches(2.0)

        headers_mp = ["Demand Hours","Actual Hours","% Var",
                      "Next Mo BCWS Hours","Next Mo ETC Hours",
                      "Comments / Root Cause & Corrective Actions"]
        for j, h in enumerate(headers_mp):
            cell = mp_tbl.cell(0,j)
            style_header_cell(cell)
            cell.text = h

        for i, (_, row) in enumerate(manpower_df.iterrows(), start=1):
            mp_tbl.cell(i,0).text = f"{row['Demand Hours']:.1f}" if pd.notna(row["Demand Hours"]) else ""
            mp_tbl.cell(i,1).text = f"{row['Actual Hours']:.1f}" if pd.notna(row["Actual Hours"]) else ""
            mp_tbl.cell(i,2).text = f"{row['% Var']*100:,.2f}%" if pd.notna(row["% Var"]) else ""
            mp_tbl.cell(i,3).text = f"{row['Next Mo BCWS Hours']:.1f}" if pd.notna(row["Next Mo BCWS Hours"]) else ""
            mp_tbl.cell(i,4).text = f"{row['Next Mo ETC Hours']:.1f}" if pd.notna(row["Next Mo ETC Hours"]) else ""
            mp_tbl.cell(i,5).text = row["Comments / Root Cause & Corrective Actions"]

            for j in range(mp_cols):
                style_body_cell(mp_tbl.cell(i,j))

            rgb = manpower_var_color(row["% Var"])
            if rgb is not None:
                cell = mp_tbl.cell(i,2)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb

# =========================================================
# MAIN PROGRAM PROCESSOR
# =========================================================

def load_penske():
    if not os.path.exists(PENSKE_PATH):
        return None
    try:
        return pd.read_excel(PENSKE_PATH)
    except Exception:
        return None

def process_program(program_name, cobra_file, penske_df):
    cobra_path = os.path.join(COBRA_DIR, cobra_file)
    if not os.path.exists(cobra_path):
        raise FileNotFoundError(f"Cobra file not found: {cobra_path}")

    print(f"\n--- Processing {program_name} from {cobra_file} ---")

    cobra_df = normalize_cobra_standard(cobra_path)

    evdf = compute_ev_timeseries(cobra_df)
    ctd_date, lsd_date = get_status_date(evdf)
    metric_df = build_program_metric_table(evdf, ctd_date, lsd_date)
    sub_metrics_df = build_subteam_metric_table(cobra_df, ctd_date, lsd_date)
    labor_df = build_labor_table(cobra_df)
    manpower_df = build_manpower_table(penske_df, program_name)

    print(f"✓ CTD date: {ctd_date.date()}, LSD date: {lsd_date.date()}")
    print(metric_df)

    # Build plot
    ev_plot_png = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Trend.png")
    make_evms_plot(evdf, program_name, ev_plot_png)

    # Build deck from theme
    prs = Presentation(THEME_PATH)

    # Slides in order:
    # 1. Program Overview (EV plot + metrics)
    add_program_overview_slide(prs, program_name, ev_plot_png, metric_df)

    # 2. Sub Team EVMS Metrics (may be multi-page)
    add_subteam_metric_slides(prs, program_name, sub_metrics_df, page_size=15)

    # 3. Sub Team Labor & Manpower (may be multi-page)
    add_labor_manpower_slides(prs, program_name, labor_df, manpower_df, page_size=15)

    # Save Excel tables
    tables_xlsx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Tables.xlsx")
    with pd.ExcelWriter(tables_xlsx, engine="xlsxwriter") as writer:
        evdf.to_excel(writer, sheet_name="EV_Series")
        metric_df.to_excel(writer, sheet_name="Program_Metrics", index=False)
        sub_metrics_df.to_excel(writer, sheet_name="Subteam_Metrics", index=False)
        labor_df.to_excel(writer, sheet_name="Subteam_Labor", index=False)
        manpower_df.to_excel(writer, sheet_name="Program_Manpower", index=False)

    out_pptx = os.path.join(OUTPUT_DIR, f"{program_name}_EVMS_Deck.pptx")
    prs.save(out_pptx)

    print(f"✓ Saved tables: {tables_xlsx}")
    print(f"✓ Saved deck:   {out_pptx}")

# =========================================================
# RUN ALL STANDARD-FORMAT PROGRAMS
# =========================================================

penske_df = load_penske()
program_errors = {}

for program, cobra_file in PROGRAM_CONFIG.items():
    try:
        process_program(program, cobra_file, penske_df)
    except Exception as e:
        print(f"!! Error for {program}: {e}")
        program_errors[program] = str(e)

print("\nALL STANDARD-FORMAT PROGRAM EVMS DECKS COMPLETE ✓")

if program_errors:
    print("\nPrograms needing re-export / clarification (not processed):")
    for prog, msg in program_errors.items():
        print(f" - {prog}: {msg}")