import pandas as pd
import yaml
from pathlib import Path
from typing import Tuple, Dict

CONFIG_PATH = Path(__file__).parents[1] / "config"

def load_config() -> Dict:
    """Load yaml config for BU/SSL/CA mapping."""
    with open(CONFIG_PATH / "org.yaml", encoding="utf-8") as f:
        return yaml.safe_load(f)

def load_powerbi_export(path: str) -> pd.DataFrame:
    """Load PowerBI export and normalize column types."""
    df = pd.read_excel(path, sheet_name=0)
    df["GPN"] = df["GPN"].astype(str).str.strip()
    df["Employee Name"] = df["Employee Name"].astype(str).str.strip()
    df["Competency"] = df["Competency"].astype(str).str.strip()
    df["Rank Description"] = df["Rank Description"].astype(str).str.strip()
    df["Missing Timesheets"] = df["Missing Timesheets"].fillna(0).astype(int)
    return df

def map_ssl_competency(df: pd.DataFrame, config: Dict) -> pd.DataFrame:
    """Map each competency to its SSL and BU."""
    lookup = {}
    for bu, bu_data in config["business_units"].items():
        for ssl, ssl_data in bu_data["ssls"].items():
            for ca in ssl_data["competency_areas"]:
                lookup[ca] = (bu, ssl)
    df["mapping"] = df["Competency"].map(lookup)
    df[["BU", "SSL"]] = pd.DataFrame(df["mapping"].tolist(), index=df.index)
    df = df.drop(columns=["mapping"])
    return df

def enrich_util_flags(df: pd.DataFrame) -> pd.DataFrame:
    """Add flags: missing_timesheet, vacation, util calculation."""
    df["missing_timesheet"] = df["Missing Timesheets"] == 1
    df["vacation"] = df["Effective Available Hours"].fillna(0) <= 0
    df["util"] = (
        df["Chargeable Hours"].fillna(0) /
        df["Effective Available Hours"].replace({0: pd.NA})
    )
    return df

def compute_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Compute aggregated totals per BU/SSL/week."""
    agg = df.groupby(["BU", "SSL"]).agg(
        total_chargeable=("Chargeable Hours", "sum"),
        total_available=("Effective Available Hours", "sum"),
        total_chargeable_excl_missing=("Chargeable Hours", lambda x: x[df["missing_timesheet"] == False].sum()),
        total_available_excl_missing=("Effective Available Hours", lambda x: x[df["missing_timesheet"] == False].sum())
    ).reset_index()
    
    # Compute utilization ratios
    agg["util_all"] = agg["total_chargeable"] / agg["total_available"]
    agg["util_excl_missing"] = agg["total_chargeable_excl_missing"] / agg["total_available_excl_missing"]
    return agg

def run_etl(
    export_path: str
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Main ETL entrypoint.
    Returns:
      - enriched DataFrame (row-level)
      - aggregated totals DataFrame
    """
    config = load_config()
    df = load_powerbi_export(export_path)
    df = map_ssl_competency(df, config)
    df = enrich_util_flags(df)
    totals = compute_totals(df)
    return df, totals

if __name__ == "__main__":
    import sys
    p = sys.argv[1] if len(sys.argv) > 1 else None
    if not p:
        print("Usage: python etl.py path/to/export.xlsx")
        sys.exit(1)
    raw, totals = run_etl(p)
    print("Row-level rows:", len(raw))
    print("Aggregated totals:", totals)
