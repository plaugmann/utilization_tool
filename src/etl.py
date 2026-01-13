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
def _norm_week(s: object) -> str:
    # Collapse any whitespace (double spaces, tabs, NBSP) to single spaces
    return " ".join(str(s).split()) if s is not None else ""

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
# Week metadata resolution
# -------------------------------------------------
def resolve_week_meta(week_str: str, week_cfg: pd.DataFrame) -> Dict:
    """
    Returns dict with keys:
      short_key, export_format, long_format, start_date, end_date, real_week
    Raises ValueError if not found.
    """
    if week_cfg is None or week_cfg.empty:
        raise ValueError(f"Week config is empty; cannot resolve week '{week_str}'")

    week_val = _norm_week(week_str)
    if not week_val:
        raise ValueError("Week value is empty; cannot resolve week metadata")

    cfg = week_cfg.copy()
    cfg["export_format_norm"] = cfg["export_format"].apply(_norm_week)

    matches = cfg[cfg["export_format_norm"] == week_val]

    if matches.empty:
        raise ValueError(f"Week '{week_val}' not found in week config")

    row = matches.iloc[0]
    return {
        "short_key": str(row["short_key"]).strip(),
        "export_format": str(row["export_format"]).strip(),
        "long_format": str(row["long_format"]).strip(),
        "real_week": str(row["real_week"]).strip(),
        "start_date": str(row["start_date"]).strip(),
        "end_date": str(row["end_date"]).strip(),
    }

# -------------------------------------------------
# Output paths
# -------------------------------------------------
def get_output_dir(short_key: str, bu: str, ssl: str, base_dir: Path) -> Path:
    """
    Returns outputs/<short_key>/<BU>_<SSL> and ensures it exists.
    """
    def sanitize(part: str) -> str:
        if part is None or pd.isna(part):
            return "UNKNOWN"
        cleaned = str(part).strip().replace(" ", "_")
        return cleaned.replace("/", "").replace("\\", "")

    safe_bu = sanitize(bu)
    safe_ssl = sanitize(ssl)
    output_dir = base_dir / short_key / f"{safe_bu}_{safe_ssl}"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir

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
        total_positions = np.where(total_mask.to_numpy())[0]
        if total_positions.size > 0:
            df = df.iloc[: total_positions[0]].copy()

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
    df["unmapped_competency"] = df["BU"].isna() | df["SSL"].isna()

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
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict]:
    """
    Main ETL pipeline.

    Returns:
      raw_df    -> row-level enriched dataframe
      totals_df -> aggregated utilization per BU/SSL
      week_meta -> week metadata dict
    """
    org_cfg = load_org_config()

    df = load_powerbi_export(export_source)

    week_cfg = load_week_config()
    week_val = None
    if "Week" in df.columns:
        week_series = df["Week"].dropna()
        if len(week_series) > 0:
            week_val = _norm_week(week_series.iloc[0])
    if not week_val:
        raise ValueError("Week column missing or empty in PowerBI export")
    week_meta = resolve_week_meta(week_val, week_cfg)

    df = map_competency_to_ssl(df, org_cfg)
    df = enrich_flags_and_util(df)

    totals = compute_aggregates(df)

    return df, totals, week_meta

# -------------------------------------------------
# CLI helper (optional)
# -------------------------------------------------
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python src/etl.py <powerbi_export.xlsx>")
        sys.exit(1)

    raw_df, totals_df, week_meta = run_etl(sys.argv[1])

    print("Rows:", len(raw_df))
    print("Week:", week_meta)
    print("\nAggregated totals:")
    print(totals_df)
