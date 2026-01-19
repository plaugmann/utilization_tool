from pathlib import Path
import asyncio
import sys

from playwright.sync_api import sync_playwright


def html_to_png(html_path: Path, png_path: Path, width: int = 1919, height: int = 1079) -> None:
    """
    Renders an HTML file to a PNG image for email embedding.
    """
    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        except AttributeError:
            pass

    png_path.parent.mkdir(parents=True, exist_ok=True)
    html_url = html_path.resolve().as_uri()

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": width, "height": height})
        page.goto(html_url, wait_until="networkidle")
        page.emulate_media(media="screen")
        page.screenshot(path=str(png_path), full_page=True)
        browser.close()
