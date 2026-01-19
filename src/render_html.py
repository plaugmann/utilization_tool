from __future__ import annotations

from pathlib import Path
from typing import Dict, List
import base64
import io

import pandas as pd
from PIL import Image
from jinja2 import Environment, BaseLoader
from src.etl import load_org_config
from src.utils import load_app_config


def _coerce_bool(value: object) -> bool:
    if value is None:
        return False
    normalized = str(value).strip().lower()
    if normalized in {"true", "1", "yes", "y", "t"}:
        return True
    if normalized in {"false", "0", "no", "n", "f"}:
        return False
    return False


def _status_from_row(row: pd.Series) -> str:
    if "status" in row and isinstance(row["status"], str) and row["status"].strip():
        return row["status"]

    on_leave = (
        _coerce_bool(row.get("on_leave_master"))
        or _coerce_bool(row.get("inactive_leave"))
    )
    if on_leave:
        return "On leave"

    if pd.notna(row.get("effective_available_hours")):
        try:
            if float(row.get("effective_available_hours")) <= 0:
                return "Vacation"
        except Exception:
            pass

    if _coerce_bool(row.get("missing_timesheet")):
        return "Missing timesheet"

    if _coerce_bool(row.get("unmapped_competency")):
        return "Unmapped"

    return "Normal"


def _overlay_label(status: str) -> str:
    if status == "Missing timesheet":
        return "MISSING TIMESHEET"
    if status == "Vacation":
        return "VACATION"
    if status == "On leave":
        return "LEAVE"
    return ""


def _border_class(status: str, util_val: object) -> str:
    if status != "Normal":
        return "border-none"
    try:
        util = float(util_val)
    except Exception:
        return "border-none"
    if util < 0.25:
        return "border-green"
    if util < 0.75:
        return "border-yellow"
    return "border-red"


def _fmt_percent(value: object, decimals: int = 0) -> str:
    try:
        fval = float(value)
    except Exception:
        return "-"
    return f"{fval * 100:.{decimals}f}%"


def _image_to_data_uri(path: Path) -> str:
    try:
        img = Image.open(path)
    except Exception:
        img = Image.new("RGB", (200, 200), color=(50, 50, 50))

    img = img.convert("L").convert("RGB")
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=85)
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


def _ssl_display_name(ssl: str, bu: str) -> str:
    cfg = load_org_config()
    bu_cfg = cfg.get("business_units", {}).get(str(bu), {})
    ssl_cfg = bu_cfg.get("ssls", {}).get(str(ssl), {})
    return ssl_cfg.get("display_name") or str(ssl).strip()


def _file_to_data_uri(path: Path) -> str:
    data = path.read_bytes()
    encoded = base64.b64encode(data).decode("ascii")
    suffix = path.suffix.lower()
    mime = "image/png" if suffix == ".png" else "image/jpeg"
    return f"data:{mime};base64,{encoded}"


def write_ssl_html(
    detail_df: pd.DataFrame,
    summary: Dict,
    out_html_path: Path,
    photos_dir: Path,
    placeholder_path: Path,
    ssl: str,
    bu: str,
) -> None:
    """
    Writes a self-contained HTML report (one page, A4 landscape) for one BU/SSL.
    """
    out_html_path.parent.mkdir(parents=True, exist_ok=True)

    ssl_display = _ssl_display_name(ssl, bu)

    start_date = str(summary.get("start_date", "")).strip()
    end_date = str(summary.get("end_date", "")).strip()
    date_range = f"{start_date} - {end_date}".strip(" -")

    util_val = summary.get("util_excl_missing", summary.get("util_all"))
    util_percent = _fmt_percent(util_val, decimals=0)

    app_cfg = load_app_config()
    bg_rel = app_cfg.get("background_image")
    if not bg_rel:
        raise ValueError("background_image is not configured in app config")
    bg_path = (Path(__file__).resolve().parents[1] / bg_rel).resolve()
    bg_data = _file_to_data_uri(bg_path)

    rank_order = [
        "Manager",
        "Senior Manager",
        "Director",
        "Associate Partner",
        "Partner",
        "Senior Consultant",
        "Consultant",
    ]
    rank_classes = {
        "Manager": "manager",
        "Senior Manager": "senior-manager",
        "Director": "director",
        "Associate Partner": "associate-partner",
        "Partner": "partner",
        "Senior Consultant": "senior-consultant",
        "Consultant": "consultant",
    }

    groups: List[Dict] = []
    for rank in rank_order:
        section = detail_df[detail_df["rank_bucket"] == rank].copy()
        if section.empty:
            continue
        employees: List[Dict] = []
        for _, row in section.sort_values("display_name").iterrows():
            gpn = str(row.get("gpn", "")).strip()
            photo_path = photos_dir / f"{gpn}.jpg" if gpn else placeholder_path
            if not photo_path.exists():
                photo_path = placeholder_path
            photo_data = _image_to_data_uri(photo_path)

            status = _status_from_row(row)
            employees.append(
                {
                    "display_name": str(row.get("display_name", "")).strip(),
                    "photo_data": photo_data,
                    "overlay_label": _overlay_label(status),
                    "border_class": _border_class(status, row.get("util")),
                }
            )
        groups.append(
            {
                "rank": rank,
                "css_class": rank_classes.get(rank, ""),
                "employees": employees,
            }
        )

    template = """
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>{{ title }}</title>
    <style>
      :root {
        --bg: #2f3136;
        --yellow: #f3c300;
        --red: #ff3b30;
        --tile: 48px;
        --card: 72px;
        --gap: 10px;
        --border: 4px;
        --borderOffset: 2px;
      }
      * { box-sizing: border-box; }
      html, body { height: 100%; }
      body {
        margin: 0;
        font-family: "Segoe UI", "Calibri", "Trebuchet MS", sans-serif;
        background: #1f2124;
      }
      .page {
        position: relative;
        background-image: url("{{ bg_data }}");
        background-size: cover;
        background-repeat: no-repeat;
        background-position: center;
        color: #f5f5f5;
        padding: 24px;
      }
      .page::before {
        content: "";
        position: absolute;
        inset: 0;
        background: rgba(0, 0, 0, 0.25);
        pointer-events: none;
      }
      .page > * {
        position: relative;
        z-index: 1;
      }
      .topbar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 6mm;
      }
      .top-left {
        display: flex;
        flex-direction: column;
        gap: 4px;
      }
      .ssl-title {
        font-size: 20px;
        font-weight: 800;
        color: #ffffff;
        letter-spacing: 1px;
        text-transform: uppercase;
      }
      .daterange {
        font-size: 26px;
        font-weight: 700;
        color: var(--red);
        letter-spacing: 0.5px;
      }
      .legend {
        display: flex;
        gap: 8px;
        align-items: center;
      }
      .legend-item {
        display: flex;
        align-items: center;
        gap: 6px;
        font-size: 11px;
        color: #d8d8d8;
      }
      .legend-box {
        width: 18px;
        height: 12px;
        border: 3px solid transparent;
      }
      .legend-box.green { border-color: #00b050; }
      .legend-box.yellow { border-color: #ffc000; }
      .legend-box.red { border-color: #ff0000; }
      .grid {
        position: relative;
        height: calc(100% - 24mm);
        display: grid;
        grid-template-rows: auto auto;
        gap: 0;
      }
      .top-row, .bottom-row {
        display: grid;
        gap: 0;
      }
      .top-row {
        grid-template-columns: 1.6fr 1fr 1fr 1fr 1fr;
        padding-bottom: 0;
      }
      .bottom-row {
        grid-template-columns: repeat(2, 1fr);
        padding-top: 0;
        position: relative;
        border-top: 2px dashed var(--yellow);
        padding-top: 18px;
      }
      .bottom-row .box {
        padding-top: 14px;
      }
      .box {
        padding: 10px 10px;
      }
      .top-row .box + .box {
        border-left: 2px dashed var(--yellow);
      }
      .bottom-row .box + .box {
        border-left: 2px dashed var(--yellow);
      }
      .rank-title {
        font-size: 12px;
        font-weight: 700;
        color: var(--yellow);
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
      }
      .employees {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(var(--card), 1fr));
        gap: var(--gap);
        justify-items: center;
      }
      .emp {
        width: var(--card);
        text-align: center;
        font-size: 9px;
        color: #e6e6e6;
      }
      .photo {
        position: relative;
        width: calc(var(--tile) + 2 * (var(--border) + var(--borderOffset)));
        height: calc(var(--tile) + 2 * (var(--border) + var(--borderOffset)));
        margin: 0 auto 6px auto;
        padding: calc(var(--border) + var(--borderOffset));
      }
      .photo img {
        width: var(--tile);
        height: var(--tile);
        object-fit: cover;
        display: block;
      }
      .border-green img { outline: var(--border) solid #00b050; outline-offset: var(--borderOffset); }
      .border-yellow img { outline: var(--border) solid #ffc000; outline-offset: var(--borderOffset); }
      .border-red img { outline: var(--border) solid #ff0000; outline-offset: var(--borderOffset); }
      .border-none img { outline: none; }
      .overlay {
        position: absolute;
        left: 4px;
        right: 4px;
        top: 50%;
        transform: translateY(-50%);
        background: rgba(0, 0, 0, 0.75);
        color: #ffffff;
        font-weight: 800;
        font-size: 9px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        padding: 4px 2px;
        border-radius: 3px;
      }
      .name {
        font-size: 8.5px;
        line-height: 1.2;
        margin-top: 3px;
        min-height: 16px;
      }
      .center-circle {
        position: absolute;
        top: 0;
        left: 50%;
        width: 64px;
        height: 64px;
        background: var(--yellow);
        color: #111;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 18px;
        font-weight: 800;
        transform: translate(-50%, -55%);
        box-shadow: 0 0 0 4px rgba(0,0,0,0.4);
        z-index: 10;
      }
      @media screen {
        html, body { width: 100%; height: 100%; }
        body { margin: 0; background: #111; }
        :root {
          --tile: 56px;
          --card: 84px;
          --gap: 12px;
          --border: 4px;
        }
        .page {
          width: 100vw;
          height: 100vh;
          padding: 24px;
        }
      }
      @media print {
        @page { size: A4 landscape; margin: 0; }
        html, body { width: 297mm; height: 210mm; }
        body { margin: 0; background: white; }
        .page {
          width: 297mm;
          height: 210mm;
          padding: 8mm;
        }
      }
    </style>
  </head>
  <body>
    <div class="page">
      <div class="topbar">
        <div class="top-left">
          <div class="ssl-title">{{ ssl_name }}</div>
          <div class="daterange">{{ date_range }}</div>
        </div>
        <div class="legend">
          <div class="legend-item"><span class="legend-box green"></span><span>Less than 25% utilized</span></div>
          <div class="legend-item"><span class="legend-box yellow"></span><span>Less than 75% utilized</span></div>
          <div class="legend-item"><span class="legend-box red"></span><span>More than 75% utilized</span></div>
        </div>
      </div>

      <div class="grid">
        <div class="top-row">
          {% for group in groups %}
            {% if group.css_class in ["manager", "senior-manager", "director", "associate-partner", "partner"] %}
              <div class="box {{ group.css_class }}">
                <div class="rank-title">{{ group.rank }}</div>
                <div class="employees">
                  {% for emp in group.employees %}
                    <div class="emp {{ emp.border_class }}">
                      <div class="photo">
                        <img src="{{ emp.photo_data }}" alt="{{ emp.display_name }}">
                        {% if emp.overlay_label %}
                          <div class="overlay">{{ emp.overlay_label }}</div>
                        {% endif %}
                      </div>
                      <div class="name">{{ emp.display_name }}</div>
                    </div>
                  {% endfor %}
                </div>
              </div>
            {% endif %}
          {% endfor %}
        </div>

        <div class="bottom-row">
          <div class="center-circle">{{ util_percent }}</div>
          {% for group in groups %}
            {% if group.css_class in ["senior-consultant", "consultant"] %}
              <div class="box {{ group.css_class }}">
                <div class="rank-title">{{ group.rank }}</div>
                <div class="employees">
                  {% for emp in group.employees %}
                    <div class="emp {{ emp.border_class }}">
                      <div class="photo">
                        <img src="{{ emp.photo_data }}" alt="{{ emp.display_name }}">
                        {% if emp.overlay_label %}
                          <div class="overlay">{{ emp.overlay_label }}</div>
                        {% endif %}
                      </div>
                      <div class="name">{{ emp.display_name }}</div>
                    </div>
                  {% endfor %}
                </div>
              </div>
            {% endif %}
          {% endfor %}
        </div>

      </div>
    </div>
  </body>
</html>
"""

    env = Environment(loader=BaseLoader(), autoescape=True)
    html = env.from_string(template).render(
        title=f"{bu} {ssl_display} photo wall",
        date_range=date_range,
        util_percent=util_percent,
        groups=groups,
        bg_data=bg_data,
        ssl_name=ssl_display,
        ssl_label=f"{bu} {ssl_display}",
    )
    out_html_path.write_text(html, encoding="utf-8")
