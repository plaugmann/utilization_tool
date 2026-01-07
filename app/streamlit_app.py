import sys
from pathlib import Path
from datetime import datetime
import streamlit as st
import pandas as pd

# -------------------------------------------------
# Ensure repo root is on PYTHONPATH so "src" can be imported
# -------------------------------------------------
BASE = Path(__file__).resolve().parents[1]
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

from src.etl import run_etl  # noqa: E402

# -------------------------------------------------
# Paths
# -------------------------------------------------
MASTER_PATH = BASE / "master" / "roster.csv"
AUDIT_PATH = BASE / "master" / "audit_log.csv"
PHOTOS_DIR = BASE / "photos"

# -------------------------------------------------
# Streamlit setup
# -------------------------------------------------
st.set_page_config(page_title="Utilization QC", layout="wide")
st.title("Utilization – QC & Approval")

st.caption(
    "Workflow: Upload PowerBI-export → QC → Update master roster → (next step) Render Excel/PDF."
)

# -------------------------------------------------
# Helpers: Master roster & audit log
# -------------------------------------------------
MASTER_COLUMNS = [
    "gpn",
    "display_name",
    "bu",
    "ssl",
    "rank_bucket",
    "on_leave",
    "active",
    "notes",
]

AUDIT_COLUMNS = [
    "timestamp",
    "user",
    "week",
    "action",
    "gpn",
    "field",
    "old",
    "new",
    "comment",
]

def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

def load_master() -> pd.DataFrame:
    ensure_parent(MASTER_PATH)
    if MASTER_PATH.exists() and MASTER_PATH.stat().st_size > 0:
        df = pd.read_csv(MASTER_PATH, dtype=str)
        # Ensure all expected columns exist
        for c in MASTER_COLUMNS:
            if c not in df.columns:
                df[c] = ""
        return df[MASTER_COLUMNS]
    # Create empty
    df = pd.DataFrame(columns=MASTER_COLUMNS)
    df.to_csv(MASTER_PATH, index=False)
    return df

def save_master(df: pd.DataFrame) -> None:
    ensure_parent(MASTER_PATH)
    # Normalize booleans to "True"/"False" strings for CSV stability
    for col in ["on_leave", "active"]:
        if col in df.columns:
            df[col] = df[col].astype(str)
    df[MASTER_COLUMNS].to_csv(MASTER_PATH, index=False)

def ensure_audit_header() -> None:
    ensure_parent(AUDIT_PATH)
    if not AUDIT_PATH.exists() or AUDIT_PATH.stat().st_size == 0:
        pd.DataFrame(columns=AUDIT_COLUMNS).to_csv(AUDIT_PATH, index=False)

def log_action(user: str, week: str, action: str, gpn: str, field: str, old, new, comment: str = "") -> None:
    ensure_audit_header()
    entry = pd.DataFrame([{
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "user": user,
        "week": week,
        "action": action,
        "gpn": gpn,
        "field": field,
        "old": "" if old is None else str(old),
        "new": "" if new is None else str(new),
        "comment": comment,
    }])
    entry.to_csv(AUDIT_PATH, mode="a", header=False, index=False)

# -------------------------------------------------
# Sidebar controls
# -------------------------------------------------
with st.sidebar:
    st.header("Settings")
    user = st.text_input("Audit user", value="admin")
    st.markdown("---")
    st.caption("Master files")
    st.write("Roster:", str(MASTER_PATH))
    st.write("Audit:", str(AUDIT_PATH))
    st.write("Photos:", str(PHOTOS_DIR))

# -------------------------------------------------
# Upload PowerBI export
# -------------------------------------------------
uploaded = st.file_uploader("Upload PowerBI export (xlsx)", type=["xlsx"])

if not uploaded:
    st.info("Upload a PowerBI export to begin.")
    st.stop()

# Run ETL
raw, totals = run_etl(uploaded)

# Defensive: ensure required columns exist
if "rank_bucket" not in raw.columns:
    raw["rank_bucket"] = raw.get("Rank Description", "")

if "Week" in raw.columns and len(raw) > 0:
    week = str(raw["Week"].iloc[0])
else:
    week = "UNKNOWN"

# Load master roster
master = load_master()

# Normalize master types
if len(master) > 0:
    master["gpn"] = master["gpn"].astype(str).str.strip()

raw["GPN"] = raw["GPN"].astype(str).str.strip()

# -------------------------------------------------
# Summary / totals
# -------------------------------------------------
st.subheader("📊 Totals (per BU/SSL)")
if totals is not None and len(totals) > 0:
    display_totals = totals.copy()
    # Pretty formatting
    for col in ["util_all", "util_excl_missing"]:
        if col in display_totals.columns:
            display_totals[col] = (display_totals[col] * 100).round(1)
    st.dataframe(display_totals)
else:
    st.warning("No totals computed (check mapping / data).")

st.markdown("---")

# -------------------------------------------------
# QC calculations
# -------------------------------------------------
raw_gpns = set(raw["GPN"].dropna())
master_gpns = set(master["gpn"].dropna()) if len(master) > 0 else set()

new_gpns = raw_gpns - master_gpns
missing_gpns = master_gpns - raw_gpns

# New employees
st.subheader("🆕 New employees (in PowerBI, not in master)")
cols = ["GPN", "Employee Name", "BU", "SSL", "rank_bucket"]
cols = [c for c in cols if c in raw.columns]
new_df = raw[raw["GPN"].isin(new_gpns)][cols].sort_values(["SSL", "Employee Name"], na_position="last")

if len(new_df) == 0:
    st.success("No new employees found.")
else:
    for _, row in new_df.iterrows():
        gpn = row["GPN"]
        name = row.get("Employee Name", "")
        bu = row.get("BU", "Denmark")
        ssl = row.get("SSL", "")
        rank_bucket = row.get("rank_bucket", "")
        with st.expander(f"{name} ({gpn})"):
            c1, c2 = st.columns([2, 1])
            with c1:
                display_name = st.text_input("Display name", value=name, key=f"dn_{gpn}")
                notes = st.text_input("Notes", value="", key=f"notes_{gpn}")
                active = st.checkbox("Active", value=True, key=f"active_{gpn}")
                on_leave = st.checkbox("On leave", value=False, key=f"ol_{gpn}")
            with c2:
                st.write(f"BU: **{bu}**")
                st.write(f"SSL: **{ssl}**")
                st.write(f"Rank: **{rank_bucket}**")

            if st.button("Add to master", key=f"add_{gpn}"):
                new_row = pd.DataFrame([{
                    "gpn": gpn,
                    "display_name": display_name,
                    "bu": bu,
                    "ssl": ssl,
                    "rank_bucket": rank_bucket,
                    "on_leave": str(on_leave),
                    "active": str(active),
                    "notes": notes,
                }])
                master = pd.concat([master, new_row], ignore_index=True)
                save_master(master)
                log_action(user, week, "ADD", gpn, "ALL", "", "added", comment="Added from PowerBI export")
                st.success("Added to master roster. Reload page to see QC update.")
                st.stop()

# Missing employees (assumed on leave)
st.subheader("🚫 Missing employees (in master, not in PowerBI) → assumed On leave")
missing_df = master[master["gpn"].isin(missing_gpns)].copy() if len(master) > 0 else pd.DataFrame()

if len(missing_df) == 0:
    st.success("No missing employees found.")
else:
    missing_df = missing_df.sort_values(["ssl", "display_name"], na_position="last")
    for _, row in missing_df.iterrows():
        gpn = row["gpn"]
        name = row.get("display_name", "")
        with st.expander(f"{name} ({gpn})"):
            current = row.get("on_leave", "False")
            st.write(f"Current on_leave: **{current}**")
            if st.button("Mark on leave = True", key=f"markleave_{gpn}"):
                master.loc[master["gpn"] == gpn, "on_leave"] = "True"
                save_master(master)
                log_action(user, week, "ON_LEAVE", gpn, "on_leave", current, "True", comment="Missing from export => on leave")
                st.warning("Marked on leave. Reload page to see QC update.")
                st.stop()

# SSL mismatch (employee present in both, but SSL differs)
st.subheader("🔁 SSL mismatch (PowerBI vs Master)")
if len(master) == 0:
    st.info("Master roster is empty; no SSL mismatches possible.")
else:
    merged = raw.merge(master, left_on="GPN", right_on="gpn", how="inner", suffixes=("_pbi", "_master"))
    if "SSL" in merged.columns and "ssl" in merged.columns:
        ssl_mismatch = merged[merged["SSL"] != merged["ssl"]].copy()
    else:
        ssl_mismatch = pd.DataFrame()

    if len(ssl_mismatch) == 0:
        st.success("No SSL mismatches found.")
    else:
        for _, row in ssl_mismatch.iterrows():
            gpn = row["GPN"]
            name = row.get("Employee Name", row.get("display_name", ""))
            pbi_ssl = row.get("SSL", "")
            master_ssl = row.get("ssl", "")
            with st.expander(f"{name} ({gpn})"):
                st.write(f"PowerBI SSL: **{pbi_ssl}**")
                st.write(f"Master SSL: **{master_ssl}**")
                new_ssl = st.selectbox(
                    "Set master SSL to",
                    options=["TC", "BC", "RC"],
                    index=["TC", "BC", "RC"].index(pbi_ssl) if pbi_ssl in ["TC", "BC", "RC"] else 0,
                    key=f"sslpick_{gpn}"
                )
                if st.button("Update master SSL", key=f"updssl_{gpn}"):
                    master.loc[master["gpn"] == gpn, "ssl"] = new_ssl
                    save_master(master)
                    log_action(user, week, "UPDATE", gpn, "ssl", master_ssl, new_ssl, comment="Resolved SSL mismatch")
                    st.success("Updated master SSL. Reload page to see QC update.")
                    st.stop()

# Missing photos
st.subheader("📸 Missing photos (master gpn without photos/<GPN>.jpg)")
PHOTOS_DIR.mkdir(parents=True, exist_ok=True)

if len(master) == 0:
    st.info("Master roster is empty; no photos to check yet.")
else:
    def has_photo(gpn: str) -> bool:
        if not gpn:
            return False
        return (PHOTOS_DIR / f"{gpn}.jpg").exists()

    photo_missing = master[~master["gpn"].apply(has_photo)].copy()

    if len(photo_missing) == 0:
        st.success("All master employees have photos.")
    else:
        st.dataframe(photo_missing[["gpn", "display_name", "ssl", "rank_bucket", "on_leave", "active"]])

        st.caption("Tip: Add photos as photos/<GPN>.jpg (JPG). A placeholder will be used in PDF if missing.")

st.markdown("---")
st.subheader("🔎 Debug (optional)")
with st.expander("Show ETL columns and preview"):
    st.write("Raw columns:", list(raw.columns))
    st.dataframe(raw.head(25))
