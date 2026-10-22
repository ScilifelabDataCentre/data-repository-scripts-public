#!/usr/bin/env python3
"""
storage_check.py
================
Process a Figshare anonymised user report and produce storage statistics
and plots for both individual users and group projects.

Merges and replaces:
  storage_info.py                   — byte conversion, column cleanup, Excel output
  storage_stats_figshare_per_user.R — per-user statistics and plots
  storage_stats_figshare_per_group.R — per-group statistics and plots

All plotting is done in Python (matplotlib), removing the R dependency.

Usage
-----
    python3 storage_check.py anonymised_scilifelab_2026-04-07.xlsx
    python3 storage_check.py anonymised_scilifelab_2026-04-07.xlsx -o /path/to/output/
    python3 storage_check.py anonymised_scilifelab_2026-04-07.xlsx --no-excel
    python3 storage_check.py --help

Output
------
Per run, the script saves into the output directory:
  storage_<date>.xlsx                          — processed Excel with byte columns added
  users_used_unused_quota_<date>.png           — used vs unused allocated quota (users)
  users_public_private_quota_<date>.png        — public vs private quota breakdown (users)
  users_count_by_storage_range_<date>.png      — user count per storage range (6 bands)
  users_violin_quota_nonzero_<date>.png        — quota distribution (users > 0 GB)
  users_violin_quota_0_50GB_<date>.png         — quota distribution (0–50 GB users)
  users_violin_quota_gt50GB_<date>.png         — quota distribution (>50 GB users)
  groups_public_private_storage_<date>.png     — public vs private storage (groups)
  groups_count_by_storage_range_<date>.png     — group count per storage range
  groups_violin_storage_nonzero_<date>.png     — storage distribution (groups > 0 GB)
  combined_users_groups_storage_<date>.png     — total storage: users + groups combined
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Binary conversion factors (bytes)
KB_TO_B = 1_024
MB_TO_B = 1_048_576
GB_TO_B = 1_073_741_824

# Binary: bytes → GB  (2^-30)
B_TO_GB = 2 ** -30

# GB → TB (binary)
GB_TO_TB = 1 / 1024

# Maximum system storage (GB) — used as a reference line in combined plot.
# Override at runtime with --max-storage.
MAX_STORAGE_GB = 60_000

# SciLifeLab colour palette
COL = {
    "teal":         "#045C64",
    "mid_teal":     "#82AEB2",
    "light_teal":   "#A6CBCF",
    "green":        "#A7C947",
    "light_green":  "#D3E4A3",
    "purple":       "#77577E",
    "grey":         "#A6A6A6",
}

# Six storage-range bands (label, lower_bound_exclusive, upper_bound_inclusive)
STORAGE_RANGES = [
    ("0 GB",       0.0,  0.0),
    (">0–1 GB",    0.0,  1.0),
    (">1–5 GB",    1.0,  5.0),
    (">5–10 GB",   5.0,  10.0),
    (">10–50 GB",  10.0, 50.0),
    (">50 GB",     50.0, float("inf")),
]

# Colours matching the six range bands (mirrors the R scripts)
RANGE_COLS_6 = [
    COL["teal"],
    COL["mid_teal"],
    COL["green"],
    COL["purple"],
    COL["grey"],
    "#3F3F3F",
]

# Columns to drop from the Accounts sheet (mirrors storage_info.py)
_DROP_ACCOUNTS = [
    "ORCID", "Email address", "University", "Group", "Parent group",
    "Public items (number of items including co-authored)",
    "Public items owned (number of owned items)",
    "Private items (number of owned private items including co-authored)",
    "Private items owned (number of owned private items",
    "Embargoed items (whole item)",
    "Public embargoed items (files)",
    "Public items with confidential files",
    "Co-authors for items",
    "Number of projects", "No of owned projects",
    "No of owned individual projects", "No of owned group projects",
    "No of projects where collaborator", "No of projects where viewer",
    "No of collaborators in projects",
    "Collaborating universities (projects)",
    "Public collections including co-authored",
    "Public collections (number of owned collections)",
    "Private owned collection",
    "Total item views including co-authored items",
    "Total item downloads including co-authored items",
    "Total collection views including co-authored collections",
    "Total project views",
    "Total altmetric score including co-authored",
    "Total altmetric score only for owned items",
]

# Accounts storage columns: raw string name → internal GB column name
_ACCT_STORAGE_COLS = {
    "Quota used":         "quota_used_gb",
    "Private quota used": "private_quota_used_gb",
    "Public quota used":  "public_quota_used_gb",
    "Allocated quota":    "allocated_quota_gb",
}

# Projects storage columns: raw string name → internal GB column name
_PROJ_STORAGE_COLS = {
    "Size of private items": "size_private_gb",
    "Size of public items":  "size_public_gb",
}


# ---------------------------------------------------------------------------
# Storage string conversion
# ---------------------------------------------------------------------------

def _parse_storage_str(value: object) -> float:
    """Convert a Figshare storage string ('17.35 GB', '500 MB', …) to bytes.

    Returns 0.0 for empty / non-parseable values so downstream arithmetic
    stays numeric.
    """
    s = str(value).strip()
    try:
        if s.endswith(" GB"):
            return float(s[:-3]) * GB_TO_B
        if s.endswith(" MB"):
            return float(s[:-3]) * MB_TO_B
        if s.endswith(" KB"):
            return float(s[:-3]) * KB_TO_B
        if s.endswith(" B"):
            return float(s[:-2])
        return float(s)
    except (ValueError, AttributeError):
        return 0.0


def _bytes_series_to_gb(series: pd.Series) -> pd.Series:
    """Apply _parse_storage_str then convert bytes → GB, rounded to 3 d.p."""
    return series.apply(_parse_storage_str).apply(lambda b: round(b * B_TO_GB, 3))


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_accounts(path: Path) -> pd.DataFrame:
    """Load the Accounts sheet, drop non-storage columns, add GB columns."""
    df = pd.read_excel(path, sheet_name="Accounts")

    # Drop columns not needed for storage analysis
    existing_drops = [c for c in _DROP_ACCOUNTS if c in df.columns]
    df = df.drop(columns=existing_drops)

    # Convert each storage string column → GB numeric column
    for raw_col, new_col in _ACCT_STORAGE_COLS.items():
        if raw_col in df.columns:
            df[new_col] = _bytes_series_to_gb(df[raw_col])
        else:
            df[new_col] = 0.0

    df["quota_unused_gb"] = (
        df["allocated_quota_gb"] - df["quota_used_gb"]
    ).round(3)

    return df


def load_projects(path: Path) -> pd.DataFrame:
    """Load the Projects sheet, filter to group projects, add GB columns."""
    df = pd.read_excel(path, sheet_name="Projects")

    for raw_col, new_col in _PROJ_STORAGE_COLS.items():
        if raw_col in df.columns:
            df[new_col] = _bytes_series_to_gb(df[raw_col])
        else:
            df[new_col] = 0.0

    df["total_storage_used_gb"] = (
        df["size_private_gb"] + df["size_public_gb"]
    ).round(3)

    # Keep only group-type projects, remove test entries
    type_col  = next((c for c in df.columns if c.lower().replace(" ", "_") == "project_type"), None)
    title_col = next((c for c in df.columns if c.lower() == "title"), None)

    if type_col:
        df = df[df[type_col].str.lower() == "group"]
    if title_col:
        df = df[~df[title_col].str.strip().str.lower().isin(["test"])]

    return df.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Statistics helpers
# ---------------------------------------------------------------------------

def _range_counts(series: pd.Series) -> dict[str, int]:
    """Return {range_label: count} for each of the six storage bands."""
    counts: dict[str, int] = {}
    for label, lo, hi in STORAGE_RANGES:
        if lo == hi == 0.0:
            counts[label] = int((series == 0.0).sum())
        elif hi == float("inf"):
            counts[label] = int((series > lo).sum())
        else:
            counts[label] = int(((series > lo) & (series <= hi)).sum())
    return counts


def _gb_to_tb(gb: float) -> str:
    """Format a GB value as TB with one decimal place."""
    return f"{gb * GB_TO_TB:.1f} TB"


def print_quick_summary(df_accounts: pd.DataFrame, df_projects: pd.DataFrame) -> None:
    """Print the compact one-line summary shown at the top of the output."""
    used_tb      = df_accounts["quota_used_gb"].sum()  * GB_TO_TB
    alloc_tb     = df_accounts["allocated_quota_gb"].sum() * GB_TO_TB
    groups_tb    = df_projects["total_storage_used_gb"].sum() * GB_TO_TB
    combined_tb  = used_tb + groups_tb

    print(f"\n{'═' * 58}")
    print(f"  User quota used (allocated) : {used_tb:.1f}  ({alloc_tb:.1f}) TB")
    print(f"  Group projects quota used   : {groups_tb:.1f} TB")
    print(f"  Combined total              : {combined_tb:.1f} TB")
    print(f"{'═' * 58}")


def print_summary(header: str, df: pd.DataFrame, col: str) -> None:
    """Print a detailed storage summary (totals in TB, range counts unchanged)."""
    counts = _range_counts(df[col])
    total  = df[col].sum()
    print(f"\n{'─' * 58}")
    print(f"  {header}  (n = {len(df)})")
    print(f"{'─' * 58}")
    print(f"  Total storage used   : {_gb_to_tb(total):>10}")
    if "allocated_quota_gb" in df.columns:
        alloc = df["allocated_quota_gb"].sum()
        print(f"  Total allocated quota: {_gb_to_tb(alloc):>10}")
    print(f"\n  Distribution by used storage (GB ranges, count of users/groups):")
    for label, n in counts.items():
        bar = "█" * min(n, 40)
        print(f"    {label:>12}  {n:>5}  {bar}")


# ---------------------------------------------------------------------------
# Plotting helpers
# ---------------------------------------------------------------------------

def _apply_theme(ax: plt.Axes) -> None:
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(colors="black")
    ax.xaxis.label.set_color("black")
    ax.yaxis.label.set_color("black")


def _savefig(fig: plt.Figure, output_dir: Path, filename: str) -> None:
    path = output_dir / filename
    fig.savefig(path, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"    {path.name}")


def _stacked_bar(
    ax: plt.Axes,
    x_label: str,
    segments: list[tuple[str, float, str]],  # (label, value, colour)
    ylabel: str,
    title: str,
) -> None:
    """Draw a single stacked bar and label each segment."""
    bottom = 0.0
    for label, value, colour in segments:
        ax.bar([x_label], [value], bottom=[bottom], color=colour, label=label, zorder=3)
        if value > 0:
            ax.text(
                0, bottom + value / 2,
                f"{value:.0f} GB",
                ha="center", va="center",
                fontsize=9, fontweight="bold", color="black",
            )
        bottom += value
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    _apply_theme(ax)


def _range_bar(
    ax: plt.Axes,
    x_label: str,
    counts: dict[str, int],
    ylabel: str,
    title: str,
) -> None:
    """Draw a stacked bar showing the count in each storage range."""
    bottom = 0
    for (label, *_), colour in zip(STORAGE_RANGES, RANGE_COLS_6):
        n = counts[label]
        ax.bar([x_label], [n], bottom=[bottom], color=colour, label=label, zorder=3)
        if n > 0:
            ax.text(
                0, bottom + n / 2,
                str(n),
                ha="center", va="center",
                fontsize=8.5, fontweight="bold", color="white",
            )
        bottom += n
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.legend(fontsize=8, bbox_to_anchor=(1.02, 1), loc="upper left")
    _apply_theme(ax)


def _violin(
    ax: plt.Axes,
    values: np.ndarray,
    x_label: str,
    ylabel: str,
    title: str,
    colour: str = COL["green"],
) -> None:
    """Draw a violin + boxplot for a 1-D array of values."""
    parts = ax.violinplot(values, positions=[0], widths=0.6,
                          showmedians=True, showextrema=True)
    for pc in parts["bodies"]:
        pc.set_facecolor(colour)
        pc.set_alpha(0.6)
    parts["cmedians"].set_color(COL["teal"])
    ax.boxplot(values, positions=[0], widths=0.08, patch_artist=True,
               boxprops=dict(facecolor="white", linewidth=0.8),
               medianprops=dict(color=COL["teal"], linewidth=1.5),
               whiskerprops=dict(linewidth=0.8),
               capprops=dict(linewidth=0.8),
               flierprops=dict(marker="o", markersize=3, alpha=0.5))
    median_val = float(np.median(values))
    x_nudge = 0.38
    ax.text(x_nudge, median_val, f"median\n{median_val:.1f} GB",
            va="center", fontsize=8, color=COL["teal"])
    ax.set_xticks([0])
    ax.set_xticklabels([x_label])
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    _apply_theme(ax)


# ---------------------------------------------------------------------------
# Per-user plots  (mirror of storage_stats_figshare_per_user.R)
# ---------------------------------------------------------------------------

def plot_users(df: pd.DataFrame, date: str, output_dir: Path) -> None:
    print("\n  Per-user plots:")

    used   = round(df["quota_used_gb"].sum(),         1)
    unused = round(df["quota_unused_gb"].sum(),        1)
    pub    = round(df["public_quota_used_gb"].sum(),   1)
    priv   = round(df["private_quota_used_gb"].sum(),  1)

    # 1. Used vs unused allocated quota
    fig, ax = plt.subplots(figsize=(5, 4.5))
    _stacked_bar(ax, date,
                 [(f"Used ({used} GB)",   used,   COL["green"]),
                  (f"Unused ({unused} GB)", unused, COL["grey"])],
                 "Storage (GB)",
                 "Allocated quota: used vs unused (users)")
    ax.axhline(MAX_STORAGE_GB, ls="--", color=COL["teal"], lw=0.8, label="Max system storage")
    ax.legend(fontsize=8)
    _savefig(fig, output_dir, f"users_used_unused_quota_{date}.png")

    # 2. Public vs private used quota
    fig, ax = plt.subplots(figsize=(5, 4.5))
    _stacked_bar(ax, date,
                 [(f"Public ({pub} GB)",   pub,  COL["light_green"]),
                  (f"Private ({priv} GB)", priv, COL["light_teal"])],
                 "Storage (GB)",
                 "Public vs private used quota (users)")
    ax.legend(fontsize=8)
    _savefig(fig, output_dir, f"users_public_private_quota_{date}.png")

    # 3. User count by storage range (6 bands)
    fig, ax = plt.subplots(figsize=(5.5, 4.5))
    _range_bar(ax, date, _range_counts(df["quota_used_gb"]),
               "Number of users", "Users by used quota range")
    _savefig(fig, output_dir, f"users_count_by_storage_range_{date}.png")

    # 4. Violin: all users > 0 GB
    df_nz = df[df["quota_used_gb"] > 0]["quota_used_gb"].values
    if len(df_nz) > 1:
        fig, ax = plt.subplots(figsize=(5, 4.5))
        _violin(ax, df_nz, date,
                "Used quota per user (GB)",
                "Used quota distribution\n(users > 0 GB)", COL["green"])
        _savefig(fig, output_dir, f"users_violin_quota_nonzero_{date}.png")

    # 5. Violin: 0–50 GB users
    df_0_50 = df[(df["quota_used_gb"] > 0) & (df["quota_used_gb"] <= 50)]["quota_used_gb"].values
    if len(df_0_50) > 1:
        fig, ax = plt.subplots(figsize=(5, 4.5))
        _violin(ax, df_0_50, date,
                "Used quota per user (GB)",
                "Used quota distribution\n(0–50 GB users)", COL["green"])
        _savefig(fig, output_dir, f"users_violin_quota_0_50GB_{date}.png")

    # 6. Violin: >50 GB users
    df_gt50 = df[df["quota_used_gb"] > 50]["quota_used_gb"].values
    if len(df_gt50) > 1:
        fig, ax = plt.subplots(figsize=(5, 4.5))
        _violin(ax, df_gt50, date,
                "Used quota per user (GB)",
                "Used quota distribution\n(>50 GB users)", COL["teal"])
        _savefig(fig, output_dir, f"users_violin_quota_gt50GB_{date}.png")


# ---------------------------------------------------------------------------
# Per-group plots  (mirror of storage_stats_figshare_per_group.R)
# ---------------------------------------------------------------------------

def plot_groups(df: pd.DataFrame, date: str, output_dir: Path) -> None:
    print("\n  Per-group plots:")

    pub  = round(df["size_public_gb"].sum(),  1)
    priv = round(df["size_private_gb"].sum(), 1)

    # 1. Public vs private storage
    fig, ax = plt.subplots(figsize=(5, 4.5))
    _stacked_bar(ax, date,
                 [(f"Public ({pub} GB)",   pub,  COL["light_green"]),
                  (f"Private ({priv} GB)", priv, COL["light_teal"])],
                 "Storage (GB)",
                 "Public vs private storage (group projects)")
    ax.legend(fontsize=8)
    _savefig(fig, output_dir, f"groups_public_private_storage_{date}.png")

    # 2. Group count by storage range
    fig, ax = plt.subplots(figsize=(5.5, 4.5))
    _range_bar(ax, date, _range_counts(df["total_storage_used_gb"]),
               "Number of group projects", "Group projects by storage range")
    _savefig(fig, output_dir, f"groups_count_by_storage_range_{date}.png")

    # 3. Violin: groups > 0 GB
    df_nz = df[df["total_storage_used_gb"] > 0]["total_storage_used_gb"].values
    if len(df_nz) > 1:
        fig, ax = plt.subplots(figsize=(5, 4.5))
        _violin(ax, df_nz, date,
                "Used storage per group project (GB)",
                "Storage distribution per group project\n(groups > 0 GB)", COL["green"])
        _savefig(fig, output_dir, f"groups_violin_storage_nonzero_{date}.png")


# ---------------------------------------------------------------------------
# Combined users + groups plot
# ---------------------------------------------------------------------------

def plot_combined(df_users: pd.DataFrame, df_groups: pd.DataFrame,
                  date: str, output_dir: Path) -> None:
    print("\n  Combined plot:")

    user_total  = round(df_users["quota_used_gb"].sum(),           1)
    group_total = round(df_groups["total_storage_used_gb"].sum(),  1)
    grand_total = round(user_total + group_total,                  1)
    pct         = round(100 * grand_total / MAX_STORAGE_GB,        1)

    fig, ax = plt.subplots(figsize=(5.5, 5))
    ax.bar([date], [user_total],  color=COL["green"],       label=f"Users ({user_total} GB)",          zorder=3)
    ax.bar([date], [group_total], bottom=[user_total],
           color=COL["light_green"], label=f"Group projects ({group_total} GB)", zorder=3)

    ax.axhline(MAX_STORAGE_GB,        ls="--", color=COL["teal"],     lw=0.9, label="Max system storage")
    ax.axhline(MAX_STORAGE_GB * 0.5,  ls="--", color=COL["mid_teal"], lw=0.9, label="50% of system storage")

    ax.text(0, user_total / 2,
            f"{user_total:.0f} GB", ha="center", va="center", fontsize=9, fontweight="bold")
    ax.text(0, user_total + group_total / 2,
            f"{group_total:.0f} GB", ha="center", va="center", fontsize=9, fontweight="bold")
    ax.text(0, grand_total + MAX_STORAGE_GB * 0.015,
            f"Total: {grand_total:.0f} GB ({pct}% of max)",
            ha="center", va="bottom", fontsize=8.5, color=COL["teal"])

    ax.set_ylabel("Storage (GB)")
    ax.set_title("Total storage: users + group projects")
    ax.legend(fontsize=8, bbox_to_anchor=(1.02, 1), loc="upper left")
    _apply_theme(ax)
    _savefig(fig, output_dir, f"combined_users_groups_storage_{date}.png")


# ---------------------------------------------------------------------------
# Summary tables  (replaces the R df_totals objects saved as .rds files)
# ---------------------------------------------------------------------------

def build_user_summary(df: pd.DataFrame, date: str) -> pd.DataFrame:
    """Return a single-row DataFrame mirroring the R df_totals for users.

    Columns match what storage_stats_figshare_per_user.R stored in
    df_total_storage_per_user_6_time_points.binary.GB.rds.
    """
    counts = _range_counts(df["quota_used_gb"])
    row = {
        "date":                 date,
        "n_users":              len(df),
        "allocated_quota_GB":   round(df["allocated_quota_gb"].sum(), 1),
        "quota_used_GB":        round(df["quota_used_gb"].sum(), 1),
        "quota_unused_GB":      round(df["quota_unused_gb"].sum(), 1),
        "public_quota_used_GB": round(df["public_quota_used_gb"].sum(), 1),
        "private_quota_used_GB":round(df["private_quota_used_gb"].sum(), 1),
        "used_0GB":             counts["0 GB"],
        "used_0_1GB":           counts[">0–1 GB"],
        "used_1_5GB":           counts[">1–5 GB"],
        "used_5_10GB":          counts[">5–10 GB"],
        "used_10_50GB":         counts[">10–50 GB"],
        "used_0_50GB":          sum(counts[r] for r in [">0–1 GB", ">1–5 GB",
                                                         ">5–10 GB", ">10–50 GB"]),
        "used_gt50GB":          counts[">50 GB"],
    }
    return pd.DataFrame([row])


def build_group_summary(df: pd.DataFrame, date: str) -> pd.DataFrame:
    """Return a single-row DataFrame mirroring the R df_totals for groups.

    Columns match what storage_stats_figshare_per_group.R stored in
    df_total_storage_per_group_projects_6_time_points.binary.GB.rds.
    """
    counts = _range_counts(df["total_storage_used_gb"])
    row = {
        "date":                  date,
        "n_groups":              len(df),
        "size_public_items_GB":  round(df["size_public_gb"].sum(), 1),
        "size_private_items_GB": round(df["size_private_gb"].sum(), 1),
        "total_storage_used_GB": round(df["total_storage_used_gb"].sum(), 1),
        "used_0GB":              counts["0 GB"],
        "used_0_1GB":            counts[">0–1 GB"],
        "used_1_5GB":            counts[">1–5 GB"],
        "used_5_10GB":           counts[">5–10 GB"],
        "used_10_50GB":          counts[">10–50 GB"],
        "used_0_50GB":           sum(counts[r] for r in [">0–1 GB", ">1–5 GB",
                                                          ">5–10 GB", ">10–50 GB"]),
        "used_gt50GB":           counts[">50 GB"],
    }
    return pd.DataFrame([row])


# ---------------------------------------------------------------------------
# Processed Excel output  (mirrors storage_info.py)
# ---------------------------------------------------------------------------

def save_processed_excel(path: Path, date: str, output_dir: Path,
                          df_user_summary: pd.DataFrame,
                          df_group_summary: pd.DataFrame) -> None:
    """Save a cleaned Excel with byte columns inserted after each string column."""
    accounts = pd.read_excel(path, sheet_name="Accounts")
    projects = pd.read_excel(path, sheet_name="Projects")

    # Accounts: insert byte column immediately after each storage string column.
    # get_loc() already reflects the shifted position after each insert, so no
    # separate offset counter is needed — adding one caused the IndexError.
    for raw_col in ["Quota used", "Private quota used", "Public quota used", "Allocated quota"]:
        if raw_col in accounts.columns:
            idx = accounts.columns.get_loc(raw_col) + 1
            byte_vals = accounts[raw_col].apply(_parse_storage_str)
            accounts.insert(idx, f"{raw_col} (B)", byte_vals)

    # Projects: same treatment
    for raw_col in ["Size of private items", "Size of public items"]:
        if raw_col in projects.columns:
            idx = projects.columns.get_loc(raw_col) + 1
            byte_vals = projects[raw_col].apply(_parse_storage_str)
            projects.insert(idx, f"{raw_col} (B)", byte_vals)

    out_path = output_dir / f"storage_{date}.xlsx"
    with pd.ExcelWriter(out_path) as writer:
        accounts.to_excel(writer, sheet_name="Accounts",      index=False)
        projects.to_excel(writer, sheet_name="Projects",      index=False)
        df_user_summary.to_excel(writer,  sheet_name="Summary_users",  index=False)
        df_group_summary.to_excel(writer, sheet_name="Summary_groups", index=False)
    print(f"\n  Processed Excel (4 sheets) → {out_path.name}")

    # Also save the summaries as standalone CSV files — one row per run,
    # so they can be appended across reports to build a time-series table.
    user_csv  = output_dir / f"summary_users_{date}.csv"
    group_csv = output_dir / f"summary_groups_{date}.csv"
    df_user_summary.to_csv(user_csv,  index=False)
    df_group_summary.to_csv(group_csv, index=False)
    print(f"  User summary CSV        → {user_csv.name}")
    print(f"  Group summary CSV       → {group_csv.name}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _find_latest_xlsx(directory: Path, pattern: str) -> Path | None:
    """Return the most recently modified file matching pattern in directory."""
    candidates = sorted(
        directory.glob(pattern),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def main() -> None:
    from datetime import date as _date

    parser = argparse.ArgumentParser(
        description=(
            "Process a Figshare anonymised user report and produce "
            "storage statistics and plots for users and group projects."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python3 storage_check.py                                    # auto-detect latest xlsx\n"
            "  python3 storage_check.py anonymised_scilifelab_2026-04-07.xlsx\n"
            "  python3 storage_check.py report.xlsx -o results/\n"
            "  python3 storage_check.py report.xlsx --no-excel\n"
        ),
    )
    parser.add_argument(
        "input", type=Path, nargs="?", default=None,
        help="Path to the anonymised user report (.xlsx). "
             "If omitted, the latest file matching --file-pattern in the "
             "script's directory is used automatically.",
    )
    parser.add_argument(
        "--output", "-o", type=Path, default=None,
        help="Output directory for plots and processed Excel "
             "(default: check_YYYY_MM_DD subfolder next to the input file).",
    )
    parser.add_argument(
        "--no-excel", action="store_true",
        help="Skip saving the processed Excel file with byte columns.",
    )
    parser.add_argument(
        "--max-storage", type=float, default=MAX_STORAGE_GB,
        metavar="GB",
        help=f"Total system storage capacity in GB, used as a reference line "
             f"in the combined plot (default: {MAX_STORAGE_GB:,}).",
    )
    parser.add_argument(
        "--file-pattern", default="anonymised_scilifelab_*.xlsx",
        metavar="GLOB",
        help="Glob pattern used when auto-detecting the input file "
             "(default: anonymised_scilifelab_*.xlsx).",
    )
    args = parser.parse_args()

    # Allow --max-storage to override the module-level constant for this run
    global MAX_STORAGE_GB
    MAX_STORAGE_GB = args.max_storage

    # Resolve input file — explicit arg or auto-detect latest xlsx
    if args.input is not None:
        path = args.input.resolve()
        if not path.is_file():
            sys.exit(f"Error: file not found: {path}")
    else:
        script_dir = Path(__file__).resolve().parent
        path = _find_latest_xlsx(script_dir, args.file_pattern)
        if path is None:
            sys.exit(
                f"Error: no file matching '{args.file_pattern}' found in "
                f"{script_dir}\nProvide a path explicitly or adjust --file-pattern."
            )
        print(f"Auto-detected input: {path.name}")

    # Extract date from filename; fall back to today's date
    m = re.search(r"(\d{4}-\d{2}-\d{2})", path.name)
    date = m.group(1) if m else _date.today().strftime("%Y-%m-%d")

    # Default output: a dated subfolder next to the input file.
    # Override with --output if a different location is needed.
    folder_name = "check_" + date.replace("-", "_")   # e.g. check_2026_10_22
    output_dir  = (args.output or (path.parent / folder_name)).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nStorage check — {date}")
    print(f"  Input : {path}")
    print(f"  Output: {output_dir}")

    # Load both sheets
    print("\nLoading data …")
    df_accounts = load_accounts(path)
    df_projects = load_projects(path)
    print(f"  Accounts : {len(df_accounts):>5} users")
    print(f"  Projects : {len(df_projects):>5} group projects")

    # Build summary tables (equivalent to the R df_totals objects)
    df_user_summary  = build_user_summary(df_accounts, date)
    df_group_summary = build_group_summary(df_projects, date)

    # Compact top-line summary
    print_quick_summary(df_accounts, df_projects)

    # Detailed console summary tables
    print_summary("Users",          df_accounts, col="quota_used_gb")
    print_summary("Group projects", df_projects, col="total_storage_used_gb")

    # Plots
    print("\nGenerating plots …")
    plot_users(df_accounts, date, output_dir)
    plot_groups(df_projects, date, output_dir)
    plot_combined(df_accounts, df_projects, date, output_dir)

    # Processed Excel + summary CSVs
    if not args.no_excel:
        save_processed_excel(path, date, output_dir, df_user_summary, df_group_summary)

    print("\nDone.\n")


if __name__ == "__main__":
    main()
