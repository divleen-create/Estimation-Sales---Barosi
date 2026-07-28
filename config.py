"""Central configuration for the Daily Sales Reporting one-pager.

Nothing here writes anywhere. The Google Sheet IDs are recorded for the
read-only reader; the pipeline defaults to the local .xlsx snapshots until
read-only Sheets access is wired.
"""
from __future__ import annotations
import calendar
import datetime as _dt
from pathlib import Path

# --- Reporting period (auto-rolling) ----------------------------------------
# The month is resolved automatically at run time: the latest month tab that
# is populated in BOTH workbooks is used (see data_source.resolve_period).
# This makes the report roll from July -> August by itself as soon as the
# August tab starts getting data, and it falls back to the last populated
# month if the newer month is not updated yet. No clock dependence.
#
# To pin a specific month (debugging / re-issuing a past month), set e.g.
# FORCE_PERIOD = (2026, 7); leave None for auto.
FORCE_PERIOD = None

# --- Source workbooks (local snapshots) -------------------------------------
BASE_DIR = Path(__file__).resolve().parent.parent  # ...\Sales & Estimation
WORKBOOK_QC = BASE_DIR / "Daily Sales Reporting.xlsx"
WORKBOOK_MKT = BASE_DIR / "Daily Sales Reporting-(New -Platform + D2C+ Amazon).xlsx"

# --- Google Sheets (read-only; wired later, never written) ------------------
GSHEET_QC_ID = "12dc5vEy4D1ws9aAjbFEravibLX3CYQO8uZuTKZAS4Lo"
GSHEET_MKT_ID = "174gdhqGFrCKj0EY3WdbbtObM46fp4n7wCFyPtPp8krg"
# Read-only scope only. The pipeline must never request write scopes.
GSHEET_READONLY_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly"

# --- Header canonicalisation ------------------------------------------------
# Source field headers (row 2) are inconsistent across tabs; map to one name.
FIELD_ALIASES = {
    "units": "units",
    "value": "gmv",
    "lm gmv": "lm_gmv",
    "ad sales": "ad_spend",
    "dolchi": "dolchi",
    "total sale": "total_sale",
}

# Channel group headers (row 1) -> clean display name.
CHANNEL_ALIASES = {
    "blinkit": "Blinkit",
    "big basket": "Big Basket",
    "swiggy": "Swiggy",
    "zepto": "Zepto",
    "shopify": "Shopify",
    "amazon": "Amazon",
    "amazon core": "Amazon Core",
    "amazon now + fresh": "Amazon NOW+Fresh",
    "flipkart": "Flipkart",
    "minutes": "Minutes",
    "first club": "First Club",
}
# Groups that are not real channels for the leadership view.
CHANNEL_SKIP = {"total", "xx", ""}

# --- Business grouping & cadence --------------------------------------------
# Section -> channels. Daily channels get the date grid; monthly channels
# are shown as a single MTD figure (data may be entered daily but is reviewed
# monthly, so aggregating is honest, not bluffing).
# Each section lists its channels in display order. Every channel gets a
# day-wise table, EXCEPT sub-channels (see SUBCHANNELS) which are components of
# a parent and are shown as a contribution split on the parent's table instead.
SECTIONS = {
    "Quick Commerce": {
        "channels": ["Blinkit", "Big Basket", "Swiggy", "Zepto"],
    },
    "Marketplace & D2C": {
        # Shopify: Value = gross sales; Total Sale = net (cancellations/discounts).
        # Amazon GMV (and Blinkit ad) arrive T-2. Amazon Core + Amazon NOW+Fresh
        # are the components of Amazon (Amazon ≈ Core + NOW+Fresh).
        "channels": [
            "Shopify", "Amazon", "Amazon Core", "Amazon NOW+Fresh",
            "Flipkart", "Minutes", "First Club",
        ],
    },
}

# Sub-channels roll up into a parent: shown in the summary and as a contribution
# split at the top of the parent's day-wise table, NOT as their own daily table.
SUBCHANNELS = {"Amazon": ["Amazon Core", "Amazon NOW+Fresh"]}

# Nothing excluded now (Shopify is included).
EXCLUDED_CHANNELS: set[str] = set()

# GMV data that lands with a lag; used to label per-channel "data as of" and to
# avoid reading the pending tail days as a drop. (Amazon GMV is T-2.)
LAG_DAYS = {"Amazon": 2}

# Ad-sales specifically that lands T-2 (independent of GMV cadence). Blinkit's
# GMV is same-day, but its Ad Sales column lags by 2 days. Recorded here so the
# ad-spend column is handled correctly if/when it is surfaced in the report.
AD_LAG_DAYS = {"Amazon": 2, "Blinkit": 2}

# Channels that are EXPECTED to carry an Ad Sales column. If one of these ever
# loses its ad column (header renamed / block moved), the report would silently
# show Ad Sales 0 — QC layer E fails on that. A channel gaining ads is only an
# advisory (add it here to make it expected).
AD_CHANNELS = {"Blinkit", "Big Basket", "Swiggy", "Zepto", "Amazon"}

# Source columns deliberately parsed-and-dropped (not part of the leadership
# view). Shopify: 'Total Sale' is net of cancellations/discounts and 'Dolchi' is
# a sub-brand split — GMV must stay the gross 'Value' column. QC layer E proves
# these never reach the model.
DROPPED_FIELDS = {"total_sale", "dolchi"}

# Tolerance for the sub-channel roll-up identity (parent ~ sum of its parts).
# Amazon = Core + NOW+Fresh; small residual is expected from rounding/other
# sub-lines, a large one means the split no longer describes the parent.
ROLLUP_TOL = 0.05          # 5%

# --- Data-freshness note ----------------------------------------------------
# Expected data lag per platform (days behind the reference date) for the
# freshness note. Anything staler than this is flagged as behind/missing.
FRESHNESS_LAG_DAYS = {"_default": 1, "Amazon": 2}
# Platforms reviewed monthly — never flagged for missing daily data.
MONTHLY_PLATFORMS = {"Flipkart"}

# --- Output -----------------------------------------------------------------
OUTPUT_DIR = Path(__file__).resolve().parent / "output"
