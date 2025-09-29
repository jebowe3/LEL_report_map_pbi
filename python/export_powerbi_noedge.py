# ============================================
# Power BI tabs → screenshots → single PDF
# --------------------------------------------
# Install once:
#   pip install playwright Pillow reportlab
#   python -m playwright install chromium
#
# Run:
#   python export_powerbi_tabs_simple.py
# ============================================

import time
from datetime import datetime
from pathlib import Path
from typing import List, Tuple, Optional

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.utils import ImageReader

# ========= USER CONFIG =========
WORKSPACE_ID = "34206745-f731-4cf0-abb3-60f8c6063a17"
REPORT_ID    = "9b1ef4dc-67e0-4fb0-b7c0-a60b13d56d9b"

# Put your tab (section) IDs here in the order to export:
SECTIONS: List[str] = [
    "aee1ca53a294adcd19d6",  # Summary
    "f3068a3490443d692480",  # When & Where
    "bcef587a2687c468c590",  # Why
    "1d8509acae2a7c233b94",  # HCC Map
    "cf3c2bc02309c11cf0d6", # Summary - Bertie
    "a295e311cdc72150e15e", # When & Where - Bertie
    "fa92017d1ee22557b834", # HCC Map - Bertie
    "7f44c7593fdc0788af2b", # Summary - Camden
    "c5b8d2e47f0c7ddcbb53", # When & Where - Camden
    "5df08783c0b6d7388b5f", # HCC Map - Camden
]

# If you want a cleaner look, hide some chrome (works in many cases):
CHROMELESS = True

# Per-page waits (tune if your visuals load slower/faster)
AFTER_NAV_WAIT_SEC = 4.0     # quiet time after navigation
RENDER_BUFFER_SEC  = 3.0     # extra buffer for custom visuals

OUTPUT_DIR = Path("powerbi_tab_exports")
# Date-stamped PDF name
OUTPUT_PDF = Path(f"PowerBI_Report_Tabs_{datetime.now():%Y%m%d_%H%M}.pdf")

# Viewport (fixed to 1920x1200 per your request)
VIEWPORT_W, VIEWPORT_H = 1920, 1200
DEVICE_SCALE_FACTOR = 2  # “Retina” crispness

# ---- Border removal options ----
# 1) Preferred: try to screenshot ONLY the report canvas element.
USE_ELEMENT_CAPTURE = True

# Candidate selectors for the report canvas in the Power BI Service.
# We’ll try these in order; if none are found we fall back to full-page shot.
CANVAS_SELECTORS = [
    'div.reportCanvas',                             # classic
    'div[aria-label="Report canvas"]',              # ARIA
    '[data-testid="report-view-container"]',        # newer test id
    'div.visual-container-host',                    # sometimes the host wraps canvas
]

# 2) Fallback: crop fixed margins (in pixels) from a full-page screenshot.
# Adjust if you still see a sliver of left nav or bottom tabs.
CROP_MARGINS = dict(left=140, right=0, top=0, bottom=145)
# ===============================


def tab_url(section_id: str) -> str:
    base = f"https://app.powerbi.com/groups/{WORKSPACE_ID}/reports/{REPORT_ID}/{section_id}"
    qs = "experience=power-bi&clientSideAuth=0"
    if CHROMELESS:
        qs += "&chromeless=1&filterPaneEnabled=false&navContentPaneEnabled=false"
    return f"{base}?{qs}"


def wait_fixed() -> None:
    """Simple fixed waits; no DOM probing to avoid false negatives."""
    time.sleep(AFTER_NAV_WAIT_SEC)
    time.sleep(RENDER_BUFFER_SEC)


def find_report_canvas(page) -> Optional[object]:
    """
    Try a few selectors that commonly match the report canvas.
    Returns a Locator or None.
    """
    for sel in CANVAS_SELECTORS:
        loc = page.locator(sel).first
        try:
            loc.wait_for(state="visible", timeout=2000)
            return loc
        except PWTimeout:
            continue
        except Exception:
            continue
    return None


def screenshot_clean(page, path: Path) -> None:
    """
    Try to capture only the report canvas. If not found, capture full page and crop margins.
    """
    if USE_ELEMENT_CAPTURE:
        canvas_loc = find_report_canvas(page)
        if canvas_loc:
            canvas_loc.screenshot(path=path.as_posix())
            return

    # Fallback: full-page then crop
    tmp_full = path.with_suffix(".full.png")
    page.screenshot(path=tmp_full.as_posix(), full_page=True)

    if any(CROP_MARGINS.values()):
        with Image.open(tmp_full) as im:
            w, h = im.size
            left   = max(0, CROP_MARGINS["left"])
            right  = max(0, CROP_MARGINS["right"])
            top    = max(0, CROP_MARGINS["top"])
            bottom = max(0, CROP_MARGINS["bottom"])
            box = (left, top, max(left, w - right), max(top, h - bottom))
            im.crop(box).save(path.as_posix())
    else:
        # no cropping requested—just keep the full image
        Path(tmp_full).rename(path)


def merge_pngs_to_pdf(images: List[Path], out_pdf: Path) -> None:
    """
    Create a PDF where each page is the same size as the image placed on it.
    We convert pixels -> points using a DPI hint (default 96). If your screenshots
    are effectively 2x (device_scale_factor=2), you can pass dpi_hint=192.
    """
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
    from PIL import Image

    dpi_hint = 96  # typical CSS pixel density. Use 192 if you want 2x pixels -> 1 inch.

    c = canvas.Canvas(out_pdf.as_posix())
    for img_path in images:
        with Image.open(img_path) as im:
            w_px, h_px = im.size
            # If the image carries a DPI, prefer that; otherwise use the hint.
            img_dpi = im.info.get("dpi", (dpi_hint, dpi_hint))[0] or dpi_hint
            px_to_pt = 72.0 / float(img_dpi)

            page_w = w_px * px_to_pt
            page_h = h_px * px_to_pt

            c.setPageSize((page_w, page_h))
            c.drawImage(ImageReader(im), 0, 0, width=page_w, height=page_h)
            c.showPage()
    c.save()



def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as pw:
        user_dir = str((Path.home() / ".pbi_auth_powerbi").resolve())

        # Persistent context keeps you signed in between runs
        browser = pw.chromium.launch_persistent_context(
            user_data_dir=user_dir,
            headless=False,  # visible so you can log in if needed
            viewport={"width": VIEWPORT_W, "height": VIEWPORT_H},
            device_scale_factor=DEVICE_SCALE_FACTOR,
        )
        page = browser.new_page()

        # Open the first tab to trigger sign-in if needed
        print("Opening first tab to let you sign in (if prompted)…")
        page.goto(tab_url(SECTIONS[0]), wait_until="domcontentloaded")

        resp = input(
            "\n🔐 If Power BI prompts for login/MFA, complete it in the browser.\n"
            "When the report looks loaded, type Y then ENTER to begin export: "
        ).strip().lower()
        if resp != "y":
            browser.close()
            print("Cancelled.")
            return

        shots: List[Path] = []
        for i, sec in enumerate(SECTIONS, 1):
            url = tab_url(sec)
            print(f"\n→ Loading tab {i}/{len(SECTIONS)}: {url}")
            page.goto(url, wait_until="domcontentloaded")
            wait_fixed()

            img_path = OUTPUT_DIR / f"tab_{i:02d}_{sec}.png"
            screenshot_clean(page, img_path)
            print(f"  Saved {img_path.name}")
            shots.append(img_path)

        browser.close()

    merge_pngs_to_pdf(shots, OUTPUT_PDF)
    print(f"\n✅ Done. Combined PDF: {OUTPUT_PDF.resolve()}")


if __name__ == "__main__":
    main()
