"""Render the HTML one-pager to a WhatsApp-ready PNG.

Uses the Chrome/Edge already installed on Windows in headless mode — no pip
install, no browser download. Captures at 2x scale for crisp text, then
auto-crops the trailing whitespace with Pillow.
"""
from __future__ import annotations
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image

# Prefer Chrome, fall back to Edge; both ship a compatible --screenshot flag.
_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]


def _find_browser() -> str | None:
    for p in _CANDIDATES:
        if Path(p).exists():
            return p
    return shutil.which("chrome") or shutil.which("msedge")


def _autocrop(path: Path, pad: int = 28, bg=(244, 246, 248)) -> None:
    """Trim the uniform page background around the content, keeping a small
    even pad on all sides. Requires the browser fill colour to match `bg`."""
    img = Image.open(path).convert("RGB")
    from PIL import ImageChops
    bgimg = Image.new("RGB", img.size, bg)
    diff = ImageChops.difference(img, bgimg)
    box = diff.getbbox()
    if box:
        l, t, r, b = box
        l = max(0, l - pad); t = max(0, t - pad)
        r = min(img.width, r + pad); b = min(img.height, b + pad)
        img.crop((l, t, r, b)).save(path)


def html_to_png(html_path: Path, png_path: Path, width: int = 1000,
                height: int = 9000, scale: int = 2) -> Path:
    browser = _find_browser()
    if not browser:
        raise RuntimeError("No Chrome/Edge found for PNG rendering.")
    url = html_path.resolve().as_uri()
    with tempfile.TemporaryDirectory() as tmp:
        cmd = [
            browser, "--headless", "--disable-gpu", "--no-sandbox",
            "--hide-scrollbars", f"--user-data-dir={tmp}",
            "--default-background-color=F4F6F8FF",
            f"--force-device-scale-factor={scale}",
            f"--window-size={width},{height}",
            f"--screenshot={png_path}", url,
        ]
        subprocess.run(cmd, check=True, capture_output=True, timeout=120)
    _autocrop(png_path)
    return png_path


if __name__ == "__main__":
    import config
    src = config.OUTPUT_DIR / "index.html"
    if not src.exists():
        raise SystemExit("Run render_html.py first.")
    out = src.with_suffix(".png")
    print("rendering", src.name, "->", out.name)
    html_to_png(src, out)
    print("wrote", out, "size", Image.open(out).size)
