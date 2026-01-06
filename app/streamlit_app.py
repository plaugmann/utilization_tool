import sys
from pathlib import Path

# Ensure repo root is on PYTHONPATH so "src" can be imported
BASE = Path(__file__).resolve().parents[1]
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

import streamlit as st
import pandas as pd
from datetime import datetime

from src.etl import run_etl

BASE = Path(__file__).parents[1]
MASTER_PATH = BASE / "master" / "roster.csv"
AUDIT_PATH = BASE / "master" / "audit_log.csv"

st.set_page_config(page_title="Utilization QC", layout="wide")
st.title("Utilization – QC & Approval")

# ---------- Helpers ----------

def load_master():
    if MASTER_PATH.exists() and MASTER_PATH.stat().st_size > 0:
        return pd.read_csv(MASTER_PATH, dtype=str)
    return pd.DataFrame(columns=[
        "gpn","display_name","bu","ssl","rank_bucket","on_leave","active","notes"
    ])

def save_master(df):
    df.to_csv(MASTER_PATH, index=False)

def log_action(user, week, action, gpn, field, old, new):
    entry = pd.DataFrame([{
        "timestamp": datetime.now().isoformat(),
        "user": user,
        "week": week,
        "action": action,
        "gpn": gpn,
        "field": field,
        "old": old,
        "new": new
    }])
    if AUDIT_PATH.exists() and AUDIT_PATH.stat().st_size > 0:
        entry.to_csv(AUDIT_PATH, mode="a", header=False, index=False)
    else:
        entry.to_csv(AUDIT_PATH, index=False)

# ---------- Upload ----------

uploaded = st.file_uploader("Upload PowerBI export", type=["xlsx"])
user = st.text_input("Your name (for audit log)", value="admin")

if not uploaded:
    st.stop()

raw, totals = run_etl(uploaded)
week = raw["Week"].iloc[0] if "Week" in raw.columns else "UNKNOWN"

master = load_master()

# ---------- QC: New & Missing ----------

raw_gpns = set(raw["GPN"])
master_gpns = set(master["gpn"])

new_gpns = raw_gpns - master_gpns
missing_gpns = master_gpns - raw_gpns

st.subheader("🆕 New employees")
new_df = raw[raw["GPN"].isin(new_gpns)][
    ["GPN","Employee Name","BU","SSL","rank_bucket"]
]

for _, row in new_df.iterrows():
    with st.expander(f"{row['Employee Name']} ({row['GPN']})"):
        if st.button("Add to master", key=f"add_{row['GPN']}"):
            master = pd.concat([master, pd.DataFrame([{
                "gpn": row["GPN"],
                "display_name": row["Employee Name"],
                "bu": row["BU"],
                "ssl": row["SSL"],
                "rank_bucket": row["rank_bucket"],
                "on_leave": False,
                "active": True,
                "notes": ""
            }])])
            save_master(master)
            log_action(user, week, "ADD", row["GPN"], "ALL", "", "added")
            st.success("Added")

st.subheader("🚫 Missing employees (assumed on leave)")
missing_df = master[master["gpn"].isin(missing_gpns)]

for _, row in missing_df.iterrows():
    with st.expander(f"{row['display_name']} ({row['gpn']})"):
        if st.button("Mark on leave", key=f"leave_{row['gpn']}"):
            master.loc[master["gpn"] == row["gpn"], "on_leave"] = True
            save_master(master)
            log_action(user, week, "ON_LEAVE", row["gpn"], "on_leave", "", "True")
            st.warning("Marked on leave")

# ---------- QC: SSL mismatch ----------

st.subheader("🔁 SSL mismatch")
merged = raw.merge(master, left_on="GPN", right_on="gpn", how="inner")
ssl_mismatch = merged[merged["SSL"] != merged["ssl"]]

for _, row in ssl_mismatch.iterrows():
    with st.expander(f"{row['Employee Name']} ({row['GPN']})"):
        st.write(f"PowerBI SSL: {row['SSL']} | Master SSL: {row['ssl']}")
        new_ssl = st.selectbox(
            "Select correct SSL",
            options=["TC","BC","RC"],
            index=["TC","BC","RC"].index(row["SSL"])
        )
        if st.button("Update SSL", key=f"ssl_{row['GPN']}"):
            master.loc[master["gpn"] == row["GPN"], "ssl"] = new_ssl
            save_master(master)
            log_action(user, week, "UPDATE", row["GPN"], "ssl", row["ssl"], new_ssl)
            st.success("Updated")

# ---------- QC: Missing photos ----------

st.subheader("📸 Missing photos")
PHOTO_PATH = BASE / "photos"
missing_photos = master[
    ~master["gpn"].apply(lambda g: (PHOTO_PATH / f"{g}.jpg").exists())
]

st.dataframe(missing_photos[["gpn","display_name","ssl","rank_bucket"]])

st.success("QC completed – you may now proceed to rendering.")
