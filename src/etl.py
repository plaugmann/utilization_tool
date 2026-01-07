import pandas as pd
import numpy as np
import yaml
from pathlib import Path
from typing import Tuple, Dict, Union

# -------------------------------------------------
# Paths
# -------------------------------------------------
BASE_PATH = Path(__file__).parents[1]
CONFIG_PATH = BASE_PATH / "config"

# -------------------------------------------------
# Rank mapping (raw -> report bucket)
# -------------------------------------------------
RANK_MAP = {
    "Executive Director": "Associate Partner",
    "Manager": "Manager",
    "Partner": "Partner",
    "Senior": "Senior Consultant",
    "Senior Manager": "Senior Manager",  # may later be overridden to Director
    "Staff": "Consultant",
}

# -------------------------------------------------
# Fix names to 'Firstname Lastname' format
# -------------------------------------------------
def normalize_display_name(pbi_name: str) -> str:
    """
    Convert 'Lastname, Firstname(s)' -> 'Firstname(s) Lastname'
    Falls back safely if format is unexpected.
    """
    if not isinstance(pbi_name, str):
        return ""

    name = pbi_name.strip()

    if "," in name:
        last, first = name.split(",", 1)
        return f"{first.strip()} {last.strip()}"

    # Fallback: return as-is
    return name


# -------------------------------------------------
# Config loaders
# -------------------------------------------------
def load_org_config() -> Dict:
    with open(CONFIG_PATH / "org.yaml", encoding="utf-8") as f:
        return yaml.safe_load(f)

def load_week_config() -> pd.DataFrame:
    path = CONFIG_PATH / "weeks_FY26.csv"
    if path.exists():
        return pd.read_csv(path)
    return pd.DataFrame()

# -------------------------------------------------
# PowerBI export loader
# -------------------------------------------------
def load_powerbi_export(
    source: Union[str, Path, object]
) -> pd.DataFrame:
    """
    Load PowerBI export.
    `source` can be a file path or a Streamlit UploadedFile.

    IMPORTANT:
    PowerBI exports often include a 'Total' row (and possibly 'Applied filters...' lines)
    at the bottom. We truncate the dataset at the first 'Total' occurrence.
    """
    df = pd.read_excel(source, sheet_name=0)

    # Normalize column names early
    df.columns = [c.strip() for c in df.columns]

    # --- Truncate at 'Total' row (before cleaning types) ---
    # Most reliable is Employee Name == 'Total' (as in your screenshot).
    if "Employee Name" in df.columns:
        total_mask = df["Employee Name"].astype(str).str.strip().str.lower().eq("total")
        if total_mask.any():
            first_total_idx = total_mask.idxmax()  # first True index
            df = df.loc[: first_total_idx - 1].copy()

    # Also drop any footer lines like "Applied filters..."
    # These sometimes appear in the "Week" column or another text column.
    # We keep it conservative: remove rows where GPN is empty AND Employee Name is empty-ish.
    if "GPN" in df.columns:
        df = df[df["GPN"].notna()].copy()

    # Normalize key fields
    df["GPN"] = df["GPN"].astype(str).str.strip()
    df["Employee Name"] = df["Employee Name"].astype(str).str.strip()
    df["display_name_auto"] = df["Employee Name"].apply(normalize_display_name)
    df["Competency"] = df["Competency"].astype(str).str.strip()

    # Rank
    if "Rank Description" in df.columns:
        df["Rank Description"] = df["Rank Description"].astype(str).str.strip()
    else:
        df["Rank Description"] = ""

    # Missing timesheets (boolean 0/1)
    if "Missing Timesheets" in df.columns:
        df["Missing Timesheets"] = df["Missing Timesheets"].fillna(0).astype(int)
    else:
        df["Missing Timesheets"] = 0

    # Employee status (inactive = on leave)
    if "Employee Status" not in df.columns:
        df["Employee Status"] = ""

    # Hours (defensive defaults)
    for col in ["Effective Available Hours", "Chargeable Hours"]:
        if col not in df.columns:
            df[col] = 0.0
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    return df

# -------------------------------------------------
# Mapping: Competency Area -> BU / SSL
# -------------------------------------------------
def map_competency_to_ssl(df: pd.DataFrame, org_cfg: Dict) -> pd.DataFrame:
    lookup = {}

    for bu, bu_data in org_cfg["business_units"].items():
        for ssl, ssl_data in bu_data["ssls"].items():
            for ca in ssl_data["competency_areas"]:
                lookup[ca] = (bu, ssl)

    mapped = df["Competency"].map(lookup)

    df["BU"] = mapped.apply(lambda x: x[0] if isinstance(x, tuple) else None)
    df["SSL"] = mapped.apply(lambda x: x[1] if isinstance(x, tuple) else None)

    return df

# -------------------------------------------------
# Enrichment & flags
# -------------------------------------------------
def enrich_flags_and_util(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    # Rank buckets
    out["rank_bucket_raw"] = (
        out["Rank Description"]
        .map(RANK_MAP)
        .fillna(out["Rank Description"])
    )

    # This column is REQUIRED by Streamlit
    out["rank_bucket"] = out["rank_bucket_raw"]

    # Flags
    out["missing_timesheet"] = out["Missing Timesheets"] == 1
    out["vacation"] = out["Effective Available Hours"] <= 0
    out["inactive_leave"] = (
        out["Employee Status"].astype(str).str.lower() == "inactive"
    )

    # Utilization (row-level)
    available = out["Effective Available Hours"].replace(0, np.nan)
    chargeable = out["Chargeable Hours"]

    out["util"] = chargeable / available

    return out

# -------------------------------------------------
# Aggregations
# -------------------------------------------------
def compute_aggregates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Computes totals per BU/SSL.

    Rules:
    - Vacation (available <= 0) is excluded from totals
    - Missing timesheets ARE included in 'util_all'
    - Missing timesheets are excluded from 'util_excl_missing'
    """
    base = df[~df["vacation"]].copy()

    grp = base.groupby(["BU", "SSL"], dropna=False).agg(
        total_chargeable=("Chargeable Hours", "sum"),
        total_available=("Effective Available Hours", "sum"),
        headcount=("GPN", "nunique"),
        missing_timesheets=("missing_timesheet", "sum"),
    ).reset_index()

    excl = (
        base[~base["missing_timesheet"]]
        .groupby(["BU", "SSL"], dropna=False)
        .agg(
            total_chargeable_excl_missing=("Chargeable Hours", "sum"),
            total_available_excl_missing=("Effective Available Hours", "sum"),
        )
        .reset_index()
    )

    out = grp.merge(excl, on=["BU", "SSL"], how="left")

    out["util_all"] = (
        out["total_chargeable"] /
        out["total_available"].replace(0, np.nan)
    )

    out["util_excl_missing"] = (
        out["total_chargeable_excl_missing"] /
        out["total_available_excl_missing"].replace(0, np.nan)
    )

    return out

# -------------------------------------------------
# Main ETL entrypoint
# -------------------------------------------------
def run_etl(
    export_source: Union[str, Path, object]
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Main ETL pipeline.

    Returns:
      raw_df    -> row-level enriched dataframe
      totals_df -> aggregated utilization per BU/SSL
    """
    org_cfg = load_org_config()

    df = load_powerbi_export(export_source)
    df = map_competency_to_ssl(df, org_cfg)
    df = enrich_flags_and_util(df)

    totals = compute_aggregates(df)

    return df, totals

# -------------------------------------------------
# CLI helper (optional)
# -------------------------------------------------
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python src/etl.py <powerbi_export.xlsx>")
        sys.exit(1)

    raw_df, totals_df = run_etl(sys.argv[1])

    print("Rows:", len(raw_df))
    print("\nAggregated totals:")
    print(totals_df)
