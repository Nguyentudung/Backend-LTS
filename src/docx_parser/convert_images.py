from pathlib import Path
import os
import subprocess
import tempfile
from PIL import Image, ImageChops

BASE_DIR = Path(__file__).resolve().parent.parent

try:
    from src.imagemagick import resolve_imagemagick
except ModuleNotFoundError:
    from imagemagick import resolve_imagemagick


def crop_white_margins(png_path: Path, padding: int = 10):
    im = Image.open(png_path).convert("RGB")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    diff = ImageChops.difference(im, bg)
    bbox = diff.getbbox()
    if not bbox:
        return

    left, top, right, bottom = bbox
    left = max(left - padding, 0)
    top = max(top - padding, 0)
    right = min(right + padding, im.width)
    bottom = min(bottom + padding, im.height)

    im.crop((left, top, right, bottom)).save(png_path)


def convert_wmf_to_png(image_dir: str):
    image_dir = Path(image_dir)
    wmf_files = list(image_dir.glob("*.wmf"))
    if not wmf_files:
        return

    # Higher density => higher pixel resolution (clearer formulas) without adding extra whitespace,
    # because we crop margins after rasterization.
    density = int(os.getenv("WMF_RASTER_DENSITY", "300"))

    magick = resolve_imagemagick(BASE_DIR)
    if magick is None:
        raise RuntimeError("ImageMagick not found")

    for wmf in wmf_files:
        png = wmf.with_suffix(".png")

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            pdf = tmpdir / f"{wmf.stem}.pdf"

            subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(tmpdir),
                    str(wmf),
                ],
                check=True,
                capture_output=True,
                text=True,
            )

            if not pdf.exists():
                raise RuntimeError("LibreOffice did not produce PDF from WMF")

            subprocess.run(
                magick.command + [
                    "-density", str(density),
                    str(pdf),
                    str(png),
                ],
                check=True,
                env=magick.env,
                capture_output=True,
                text=True,
            )

            # ✅ FIX DỨT ĐIỂM Ở ĐÂY
            crop_white_margins(png, padding=10)

            if png.exists():
                try:
                    wmf.unlink()
                except Exception:
                    pass
