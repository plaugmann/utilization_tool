from pathlib import Path
import asyncio
import sys

from playwright.sync_api import sync_playwright


def html_to_pdf(html_path: Path, pdf_path: Path) -> None:
    """
    Uses Playwright Chromium to print a single-page A4 landscape PDF.
    """
    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        except AttributeError:
            pass
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    html_url = html_path.resolve().as_uri()
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1919, "height": 1079})
        page.goto(html_url, wait_until="networkidle")
        page.emulate_media(media="print")
        page.pdf(
            path=str(pdf_path),
            format="A4",
            landscape=True,
            print_background=True,
            prefer_css_page_size=True,
            margin={"top": "0mm", "bottom": "0mm", "left": "0mm", "right": "0mm"},
            page_ranges="1",
        )
        browser.close()
