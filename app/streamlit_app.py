import sys
from pathlib import Path
from datetime import datetime
import streamlit as st
import pandas as pd
import io
import urllib.request
from PIL import Image
from streamlit_cropper import st_cropper

# -------------------------------------------------
# Ensure repo root is on PYTHONPATH so "src" can be imported
# -------------------------------------------------
BASE = Path(__file__).resolve().parents[1]
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

from src.etl import run_etl, load_org_config, get_output_dir  # noqa: E402
from src.render_excel import build_ssl_dataset, build_ssl_summary, write_ssl_excel  # noqa: E402

# -------------------------------------------------
# Paths
# -------------------------------------------------
MASTER_PATH = BASE / "master" / "roster.csv"
AUDIT_PATH = BASE / "master" / "audit_log.csv"
PHOTOS_DIR = BASE / "photos"
ASSETS_DIR = BASE / "assets"
DEFAULT_PHOTO = ASSETS_DIR / "placeholder_profile_400.png"
DEFAULT_PHOTO_SIZE = 400
CONFIG_PATH = BASE / "config"

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

RANK_OPTIONS = [
    "Partner",
    "Associate Partner",
    "Director",
    "Senior Manager",
    "Manager",
    "Senior Consultant",
    "Consultant",
    "Junior Consultant",
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

def normalize_bool_str(value: object) -> str:
    if value is None:
        return "False"
    normalized = str(value).strip().lower()
    if normalized in {"true", "1", "yes", "y", "t"}:
        return "True"
    if normalized in {"false", "0", "no", "n", "f"}:
        return "False"
    return "False"

def load_master() -> pd.DataFrame:
    ensure_parent(MASTER_PATH)
    if MASTER_PATH.exists() and MASTER_PATH.stat().st_size > 0:
        df = pd.read_csv(MASTER_PATH, dtype=str)
        # Ensure all expected columns exist
        for c in MASTER_COLUMNS:
            if c not in df.columns:
                df[c] = ""
        for col in ["on_leave", "active"]:
            if col in df.columns:
                df[col] = df[col].apply(normalize_bool_str)
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
            df[col] = df[col].apply(normalize_bool_str)
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

def ensure_default_photo() -> None:
    if not DEFAULT_PHOTO.exists():
        st.error(f"Missing placeholder image: {DEFAULT_PHOTO}")
        st.stop()



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

# Upload PowerBI export
# -------------------------------------------------
uploaded = st.file_uploader("Upload PowerBI export (xlsx)", type=["xlsx"])

if not uploaded:
    st.info("Upload a PowerBI export to begin.")
    st.stop()

# Run ETL
try:
    raw, totals, week_meta = run_etl(uploaded)
except ValueError as exc:
    st.error(str(exc))
    st.stop()

org_cfg = load_org_config()

# Defensive: ensure required columns exist
if "rank_bucket" not in raw.columns:
    raw["rank_bucket"] = raw.get("Rank Description", "")

week = week_meta.get("export_format", "UNKNOWN")

# Load master roster
master = load_master()

# Normalize master types
if len(master) > 0:
    master["gpn"] = master["gpn"].astype(str).str.strip()

raw["GPN"] = raw["GPN"].astype(str).str.strip()

# -------------------------------------------------
# Week metadata
# -------------------------------------------------
st.subheader("Week")
st.write(f"Week: {week_meta.get('short_key', 'UNKNOWN')} ({week_meta.get('export_format', 'UNKNOWN')})")
st.write(f"Period: {week_meta.get('start_date', 'UNKNOWN')} - {week_meta.get('end_date', 'UNKNOWN')}")

# Summary / totals
# -------------------------------------------------
st.subheader("Totals (per BU/SSL)")
if totals is not None and len(totals) > 0:
    display_totals = totals.copy()
    # Pretty formatting
    for col in ["util_all", "util_excl_missing"]:
        if col in display_totals.columns:
            display_totals[col] = (display_totals[col] * 100).round(1)
    st.dataframe(display_totals)
else:
    st.warning("No totals computed (check mapping / data).")

st.subheader("Planned output folders")
outputs_base = BASE / "outputs"
if totals is not None and len(totals) > 0 and "BU" in totals.columns and "SSL" in totals.columns:
    unique_pairs = totals[["BU", "SSL"]].drop_duplicates()
    for _, row in unique_pairs.iterrows():
        out_dir = get_output_dir(week_meta.get('short_key', 'UNKNOWN'), row["BU"], row["SSL"], outputs_base)
        st.write(str(out_dir))
else:
    st.info("No BU/SSL totals available for output folder resolution.")

st.subheader("📄 Generate Excel outputs")
if st.button("Generate Excel for all SSLs"):
    created_paths = []
    bu = "Denmark"
    week_ctx = dict(week_meta)
    week_ctx["master_df"] = master
    for ssl in ["TC", "BC", "RC"]:
        detail_df = build_ssl_dataset(master, raw, bu, ssl)
        summary_df = build_ssl_summary(totals, raw, bu, ssl, week_ctx)
        out_dir = get_output_dir(week_meta.get("short_key", "UNKNOWN"), bu, ssl, outputs_base)
        filename = f"{week_meta.get('short_key', 'UNKNOWN')} - {week_meta.get('export_format', 'UNKNOWN')}_{ssl}.xlsx"
        out_path = out_dir / filename
        write_ssl_excel(detail_df, summary_df, out_path)
        created_paths.append(out_path)
    st.success(f"Created {len(created_paths)} Excel files.")
    for path in created_paths:
        st.write(str(path))

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
cols = ["GPN", "Employee Name", "display_name_auto", "BU", "SSL", "rank_bucket"]
cols = [c for c in cols if c in raw.columns]
new_df = raw[raw["GPN"].isin(new_gpns)][cols].sort_values(["SSL", "Employee Name"], na_position="last")

if len(new_df) == 0:
    st.success("No new employees found.")
else:
    for _, row in new_df.iterrows():
        gpn = row["GPN"]
        name = row.get("display_name_auto", row.get("Employee Name", ""))
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
                bu_options = list(org_cfg.get("business_units", {}).keys())
                if bu and bu not in bu_options:
                    bu_options = [bu] + bu_options
                if not bu_options:
                    bu_options = [bu or "Unknown"]
                bu_idx = bu_options.index(bu) if bu in bu_options else 0
                selected_bu = st.selectbox("BU", options=bu_options, index=bu_idx, key=f"bu_{gpn}")

                ssl_options = list(org_cfg.get("business_units", {}).get(selected_bu, {}).get("ssls", {}).keys())
                if ssl and ssl not in ssl_options:
                    ssl_options = [ssl] + ssl_options
                if not ssl_options:
                    ssl_options = [ssl or ""]
                ssl_idx = ssl_options.index(ssl) if ssl in ssl_options else 0
                selected_ssl = st.selectbox("SSL", options=ssl_options, index=ssl_idx, key=f"ssl_{gpn}")

                rank_options = list(RANK_OPTIONS)
                if rank_bucket and rank_bucket not in rank_options:
                    rank_options = [rank_bucket] + rank_options
                rank_idx = rank_options.index(rank_bucket) if rank_bucket in rank_options else 0
                selected_rank = st.selectbox("Rank", options=rank_options, index=rank_idx, key=f"rank_{gpn}")

            if st.button("Add to master", key=f"add_{gpn}"):
                new_row = pd.DataFrame([{
                    "gpn": gpn,
                    "display_name": display_name,
                    "bu": selected_bu,
                    "ssl": selected_ssl,
                    "rank_bucket": selected_rank,
                    "on_leave": str(on_leave),
                    "active": str(active),
                    "notes": notes,
                }])
                master = pd.concat([master, new_row], ignore_index=True)
                save_master(master)
                log_action(user, week, "ADD", gpn, "ALL", "", "added", comment="Added from PowerBI export")
                st.success("Added to master roster. Reload page to see QC update.")
                st.stop()

    if st.button("Add all new employees to master"):
        new_rows = []
        for _, row in new_df.iterrows():
            gpn = row["GPN"]
            if gpn in master["gpn"].astype(str).values:
                continue
            name = row.get("display_name_auto", row.get("Employee Name", ""))
            bu = row.get("BU", "Denmark")
            ssl = row.get("SSL", "")
            rank_bucket = row.get("rank_bucket", "")

            display_name = st.session_state.get(f"dn_{gpn}", name)
            selected_bu = st.session_state.get(f"bu_{gpn}", bu)
            selected_ssl = st.session_state.get(f"ssl_{gpn}", ssl)
            selected_rank = st.session_state.get(f"rank_{gpn}", rank_bucket)
            notes = st.session_state.get(f"notes_{gpn}", "")
            active = st.session_state.get(f"active_{gpn}", True)
            on_leave = st.session_state.get(f"ol_{gpn}", False)

            new_rows.append({
                "gpn": gpn,
                "display_name": display_name,
                "bu": selected_bu,
                "ssl": selected_ssl,
                "rank_bucket": selected_rank,
                "on_leave": str(on_leave),
                "active": str(active),
                "notes": notes,
            })

        if not new_rows:
            st.info("No new employees to add.")
        else:
            master = pd.concat([master, pd.DataFrame(new_rows)], ignore_index=True)
            save_master(master)
            for row in new_rows:
                log_action(user, week, "ADD", row["gpn"], "ALL", "", "added", comment="Added from PowerBI export (bulk)")
            st.success(f"Added {len(new_rows)} new employees. Reload page to see QC update.")
            st.stop()

# Missing employees (assumed on leave)
st.subheader("🚫 Missing employees (in master, not in PowerBI) → assumed On leave")
missing_df = master[master["gpn"].isin(missing_gpns)].copy() if len(master) > 0 else pd.DataFrame()

if len(missing_df) == 0:
    st.success("No missing employees found.")
else:
    if st.button("Mark ALL missing as on leave"):
        for _, row in missing_df.iterrows():
            gpn = row["gpn"]
            current = row.get("on_leave", "False")
            master.loc[master["gpn"] == gpn, "on_leave"] = "True"
            log_action(user, week, "ON_LEAVE", gpn, "on_leave", current, "True", comment="Bulk: missing from export => on leave")
        save_master(master)
        st.warning(f"Marked {len(missing_df)} employees as on leave. Reload page to see QC update.")
        st.stop()
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
            name = row.get("display_name_auto", row.get("Employee Name", row.get("display_name", "")))
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

# Photos (upload + crop)
st.subheader("📸 Photos (upload + crop)")
ensure_default_photo()
PHOTOS_DIR.mkdir(parents=True, exist_ok=True)

if len(master) == 0:
    st.info("Master roster is empty; no photos to check yet.")
else:
    def has_photo(gpn: str) -> bool:
        if not gpn:
            return False
        return (PHOTOS_DIR / f"{gpn}.jpg").exists()

    show_all_photos = st.checkbox("Show all employees (including existing photos)", value=False)
    photo_missing = master[~master["gpn"].apply(has_photo)].copy()
    target = master.copy() if show_all_photos else photo_missing

    if len(target) == 0:
        st.success("All master employees have photos.")
    else:
        for _, row in target.sort_values(["ssl", "display_name"], na_position="last").iterrows():
            gpn = row.get("gpn", "")
            name = row.get("display_name", "")
            photo_path = PHOTOS_DIR / f"{gpn}.jpg"
            existing = photo_path if photo_path.exists() else DEFAULT_PHOTO

            with st.expander(f"{name} ({gpn})"):
                c1, c2 = st.columns([1, 2])
                with c1:
                    st.image(str(existing), caption="Current", width=250)
                with c2:
                    photo_url = st.text_input(
                        "LinkedIn image URL (direct)",
                        value="",
                        key=f"url_{gpn}",
                        help="Paste a direct image URL. LinkedIn may block some URLs.",
                    )
                    uploaded_photo = st.file_uploader(
                        "Or upload a profile photo (JPG/PNG)",
                        type=["jpg", "jpeg", "png"],
                        key=f"photo_{gpn}",
                    )

                    img = None
                    if photo_url:
                        try:
                            req = urllib.request.Request(
                                photo_url,
                                headers={"User-Agent": "Mozilla/5.0"},
                            )
                            with urllib.request.urlopen(req, timeout=10) as resp:
                                img = Image.open(io.BytesIO(resp.read()))
                        except Exception:
                            st.error("Could not load image from URL.")
                    elif uploaded_photo is not None:
                        img = Image.open(uploaded_photo)

                    if img is not None:
                        st.image(img, caption="Original", width=200)

                        cropped_preview = st_cropper(
                            img,
                            realtime_update=True,
                            box_color="#00A3FF",
                            aspect_ratio=(1, 1),
                        )

                        if st.button("Save cropped photo", key=f"save_{gpn}"):
                            if cropped_preview is None:
                                st.error("Crop the image before saving.")
                            else:
                                cropped = cropped_preview.resize(
                                    (DEFAULT_PHOTO_SIZE, DEFAULT_PHOTO_SIZE), Image.Resampling.LANCZOS
                                )
                                cropped.save(photo_path, format="JPEG", quality=90)
                                log_action(user, week, "PHOTO", gpn, "photo", "", str(photo_path), comment="Uploaded/cropped photo")
                                st.success("Saved photo. Reload page to see it in the list.")
                                st.stop()

        st.caption("Photos are saved as photos/<GPN>.jpg. Missing photos use a default placeholder.")

st.markdown("---")
st.subheader("🔎 Debug (optional)")
with st.expander("Show ETL columns and preview"):
    st.write("Raw columns:", list(raw.columns))
    st.dataframe(raw.head(25))
