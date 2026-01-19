from pathlib import Path
from typing import Dict, Optional

import pandas as pd
from PIL import Image
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


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


def _util_color(util_val: object) -> tuple:
    try:
        util = float(util_val)
    except Exception:
        return (0.6, 0.6, 0.6)

    if util < 0.25:
        return (0, 176 / 255, 80 / 255)
    if util <= 0.75:
        return (1, 192 / 255, 0)
    return (1, 0, 0)


def _fmt_percent(val: object) -> str:
    try:
        fval = float(val)
        return f"{fval * 100:.1f}%"
    except Exception:
        return "-"


def try_register_ey_font() -> Optional[str]:
    candidates = [
        r"C:\Windows\Fonts\EYInterstate-Light.ttf",
        r"C:\Windows\Fonts\EYInterstateLight.ttf",
        r"C:\Windows\Fonts\EYInterstate Light.ttf",
        r"C:\Windows\Fonts\EYInterstateLight.otf",
        r"C:\Windows\Fonts\EYInterstate-Light.otf",
    ]
    for p in candidates:
        try:
            pdfmetrics.registerFont(TTFont("EYInterstateLight", p))
            return "EYInterstateLight"
        except Exception:
            continue
    return None


def write_ssl_pdf(
    detail_df: pd.DataFrame,
    summary: Dict,
    out_path: Path,
    photos_dir: Path,
    placeholder_path: Path,
) -> None:
    """
    Writes a single PDF photo wall for one BU/SSL to out_path.
    Uses ReportLab + Pillow.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)

    page_w, page_h = landscape(A4)
    c = canvas.Canvas(str(out_path), pagesize=(page_w, page_h))

    margin = 30
    header_h = 26
    summary_h = 55
    gap_after_summary = 12

    ey_font = try_register_ey_font()
    base_font = ey_font or "Helvetica"
    base_font_bold = ey_font or "Helvetica-Bold"

    def draw_header_and_summary():
        y = page_h - margin
        c.setFont(base_font_bold, 14)
        header = f"{summary.get('short_key', '')} - {summary.get('export_format', '')}".strip(" -")
        c.drawString(margin, y, header)
        c.setFont(base_font, 10)
        period = f"{summary.get('start_date', '')} - {summary.get('end_date', '')}".strip(" -")
        c.drawString(margin, y - 16, period)

        box_top = y - header_h
        box_bottom = box_top - summary_h
        c.rect(margin, box_bottom, page_w - 2 * margin, summary_h, stroke=1, fill=0)

        c.setFont(base_font_bold, 10)
        c.drawString(margin + 8, box_top - 16, "Summary")

        c.setFont(base_font, 9)
        left_x = margin + 8
        mid_x = margin + (page_w - 2 * margin) / 2
        y1 = box_top - 30

        c.drawString(left_x, y1, f"Util (all): {_fmt_percent(summary.get('util_all'))}")
        c.drawString(left_x, y1 - 14, f"Util (excl missing): {_fmt_percent(summary.get('util_excl_missing'))}")
        c.drawString(left_x, y1 - 28, f"Headcount: {summary.get('headcount', 0)}")

        c.drawString(mid_x, y1, f"Missing timesheets: {summary.get('missing_timesheets', 0)}")
        c.drawString(mid_x, y1 - 14, f"Vacation: {summary.get('vacation', 0)}")
        c.drawString(mid_x, y1 - 28, f"On leave: {summary.get('on_leave', 0)}")

        return box_bottom - gap_after_summary

    def draw_tile(x, y_top, row: pd.Series, photo_size: int, show_util: bool, show_names: bool):
        gpn = str(row.get("gpn", "")).strip()
        photo_path = photos_dir / f"{gpn}.jpg"
        img_path = photo_path if photo_path.exists() else placeholder_path

        try:
            img = Image.open(img_path)
        except Exception:
            img = Image.open(placeholder_path)

        status = _status_from_row(row)
        util_val = row.get("util")

        img = img.convert("L").convert("RGB")
        if status == "On leave":
            img = Image.eval(img, lambda px: min(255, int(px * 1.15)))

        img = img.resize((photo_size, photo_size))
        c.drawImage(ImageReader(img), x, y_top - photo_size, width=photo_size, height=photo_size, mask="auto")

        if status in {"Normal", "Unmapped"}:
            r, g, b = _util_color(util_val)
            c.setStrokeColorRGB(r, g, b)
            border_w = 4 if photo_size >= 70 else 3
            c.setLineWidth(border_w)
            c.rect(x - 1, y_top - photo_size - 1, photo_size + 2, photo_size + 2, stroke=1, fill=0)
            c.setLineWidth(1)

        overlay = None
        if status == "Missing timesheet":
            overlay = "MISSING TIMESHEET"
        elif status == "Vacation":
            overlay = "VACATION"
        elif status == "On leave":
            overlay = "ON LEAVE"
        elif status == "Unmapped":
            overlay = "UNMAPPED"

        if overlay:
            c.saveState()
            try:
                c.setFillAlpha(0.65)
            except Exception:
                pass
            c.setFillColorRGB(0, 0, 0)
            box_h = max(14, int(photo_size * 0.18))
            c.rect(x, y_top - photo_size / 2 - box_h / 2, photo_size, box_h, fill=1, stroke=0)
            try:
                c.setFillAlpha(1)
            except Exception:
                pass
            c.setFillColorRGB(1, 1, 1)
            c.setFont(base_font_bold, 12 if photo_size >= 70 else 10)
            c.drawCentredString(x + photo_size / 2, y_top - photo_size / 2 - 4, overlay)
            c.restoreState()

        name = str(row.get("display_name", "")).strip()
        if show_names:
            if photo_size <= 42 and len(name) > 18:
                name = f"{name[:17]}..."
            name_font = 6 if photo_size <= 42 else (7 if photo_size < 70 else 8)
            name_y = y_top - photo_size - (10 if photo_size <= 42 else 12)
            c.setFont(base_font, name_font)
            c.drawCentredString(x + photo_size / 2, name_y, name)
        elif photo_size <= 42:
            parts = [p for p in name.split() if p]
            if parts:
                if len(parts) == 1:
                    ini = parts[0][:2].upper()
                else:
                    ini = (parts[0][0] + parts[-1][0]).upper()
                c.setFont(base_font_bold, 8)
                c.setFillColorRGB(1, 1, 1)
                c.drawString(x + 2, y_top - photo_size + 2, ini)
                c.setFillColorRGB(0, 0, 0)

        if show_util and status == "Normal":
            c.setFont(base_font, 7)
            c.drawCentredString(x + photo_size / 2, y_top - photo_size - 22, _fmt_percent(util_val))

    ranks = [
        "Partner",
        "Associate Partner",
        "Director",
        "Senior Manager",
        "Manager",
        "Senior Consultant",
        "Consultant",
    ]
    split_rank = "Manager"

    col_gap = 6
    row_gap = 8
    title_h = 12
    section_gap = 4

    available_w = page_w - 2 * margin
    y_start = draw_header_and_summary()
    available_h = y_start - margin
    def required_height(photo_size: int) -> int:
        tile_w_local = photo_size + 12
        show_names_local = photo_size > 42
        name_h = 9 if show_names_local else 0
        show_util_local = photo_size >= 70
        util_h_local = 8 if show_util_local else 0
        tile_h_local = photo_size + name_h + util_h_local + 6
        cols_local = max(1, int((available_w + col_gap) // (tile_w_local + col_gap)))

        total = 0
        for rank in ranks:
            n = int((detail_df["rank_bucket"] == rank).sum())
            if n == 0:
                continue

            rows = (n + cols_local - 1) // cols_local
            total += (title_h + 3)
            total += rows * tile_h_local + (rows - 1) * row_gap
            total += 10
            if rank == split_rank:
                total += 10
            total += section_gap

        return total

    chosen = None
    for size in [90, 80, 75, 70, 65, 60, 55, 50, 46, 42, 38, 34]:
        if required_height(size) <= available_h:
            chosen = size
            break
    if chosen is None:
        chosen = 34

    photo_size = chosen
    tile_w = photo_size + 12
    show_names = photo_size > 42
    name_h = 9 if show_names else 0
    show_util = photo_size >= 70
    util_h = 8 if show_util else 0
    tile_h = photo_size + name_h + util_h + 6
    cols = max(1, int((available_w + col_gap) // (tile_w + col_gap)))
    total_w = cols * tile_w + (cols - 1) * col_gap
    start_x = margin + (available_w - total_w) / 2

    y_cursor = y_start

    overflow = False
    for rank in ranks:
        section = detail_df[detail_df["rank_bucket"] == rank].copy()
        if section.empty:
            continue

        c.setFont(base_font_bold, 11)
        c.drawString(margin, y_cursor, rank)
        y_cursor -= title_h + 3

        col = 0
        row_top = y_cursor
        for _, row in section.sort_values("display_name").iterrows():
            if col >= cols:
                col = 0
                row_top -= tile_h + row_gap
            if row_top - tile_h < margin:
                overflow = True
                break

            x = start_x + col * (tile_w + col_gap)
            draw_tile(x, row_top, row, photo_size, show_util, show_names)
            col += 1

        if overflow:
            break

        y_cursor = row_top - tile_h - 10

        if rank == split_rank:
            c.setStrokeColorRGB(0.6, 0.6, 0.6)
            c.line(margin, y_cursor, page_w - margin, y_cursor)
            y_cursor -= 10

    if overflow:
        c.setFont(base_font_bold, 12)
        c.setFillColorRGB(0.8, 0, 0)
        c.drawCentredString(page_w / 2, margin + 20, "TOO MANY EMPLOYEES TO FIT ON ONE PAGE")
        c.setFillColorRGB(0, 0, 0)

    c.save()
