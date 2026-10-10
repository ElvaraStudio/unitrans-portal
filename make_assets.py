"""Generate ``assets/app.ico`` from ``UniTrans Logo.png``.

Run automatically by ``build.ps1`` before PyInstaller, and runnable by hand:

    python make_assets.py

Why this exists instead of a one-off conversion
-----------------------------------------------
Three problems make a naive ``Logo.png -> app.ico`` conversion unusable at the
sizes Windows actually renders:

1. The source is 2237x1176 (about 1.9:1, not square). Handing a wide bitmap to
   PyInstaller's ``--icon`` produces a squashed glyph at the 16x16 size used in
   the taskbar and Explorer's small-icon view. This script scales the artwork to
   fit a square canvas *without* distorting it.
2. The export carries a large transparent margin, so the artwork would sit
   off-centre. The margin is trimmed to the real content bounds.
3. The logo includes the tagline *"We can't spell out success without 'u'"*
   below the wordmark. Keeping it costs roughly half the canvas height, which
   leaves "UNiTRANS" an illegible smear at small sizes. The tagline is dropped
   from the icon so the wordmark keeps the space - it stays intact everywhere
   else (`UniTrans Logo.png` is still bundled and shown in the app UI).

The tagline is found by row profile rather than a hardcoded pixel offset, so a
re-exported logo at a different size still works.

The output is committed to the repository so a logo change shows up as a
reviewable diff on ``assets/app.ico`` rather than appearing only inside a build
artifact.
"""

from __future__ import annotations

import argparse
from pathlib import Path

#: Every size Windows asks for. 16 is the taskbar, 256 the large-icon view.
ICON_SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)

#: The artwork sits slightly proud of the canvas edge. Inset it so the logo
#: does not touch the icon border at small sizes.
CONTENT_SCALE = 0.94


def _find_bands(alpha, width: int, height: int, threshold: int = 8) -> list[tuple[int, int]]:
    """Return ``(y_start, y_end_exclusive)`` for each run of non-empty rows."""
    bands: list[tuple[int, int]] = []
    start: int | None = None
    for y in range(height):
        row = alpha.crop((0, y, width, y + 1))
        filled = any(value > threshold for value in row.get_flattened_data())
        if filled and start is None:
            start = y
        elif not filled and start is not None:
            bands.append((start, y))
            start = None
    if start is not None:
        bands.append((start, height))
    return bands


def _trim_tagline(image):
    """Drop any content block sitting below a wide gap - i.e. the tagline.

    The logo is one solid mark followed by a separated strapline. Splitting on
    the largest vertical gap and keeping everything above it removes the
    strapline without hardcoding pixel coordinates.
    """
    alpha = image.getchannel("A")
    bands = _find_bands(alpha, image.width, image.height)
    if len(bands) < 2:
        return image, False

    # Largest gap between consecutive bands.
    gaps = [
        (bands[index + 1][0] - bands[index][1], index)
        for index in range(len(bands) - 1)
    ]
    gap_size, index = max(gaps)

    # Only treat it as a strapline if the gap is a real separation, not the
    # natural spacing inside a single glyph.
    if gap_size < image.height * 0.02:
        return image, False

    return image.crop((0, 0, image.width, bands[index][1])), True


def _trim_margin(image):
    """Trim transparent border so the artwork is centred on its own content."""
    bounds = image.getchannel("A").getbbox()
    return image.crop(bounds) if bounds else image


def build_icon(source: Path, destination: Path, sizes: tuple[int, ...] = ICON_SIZES) -> list[int]:
    """Write a multi-resolution .ico built from ``source``. Returns the sizes written."""
    try:
        from PIL import Image
    except ImportError as exc:  # pragma: no cover - only when run outside the build venv
        raise SystemExit(
            "Pillow is required to generate the icon.\n"
            "Run this through the build venv (build.ps1 does this automatically), or:\n"
            "  python -m pip install Pillow"
        ) from exc

    if not source.exists():
        raise SystemExit(f"Logo not found: {source}")

    logo = Image.open(source).convert("RGBA")
    original_size = logo.size

    logo = _trim_margin(logo)
    logo, tagline_dropped = _trim_tagline(logo)
    # Trimming the tagline can expose a margin of its own on the remaining art.
    logo = _trim_margin(logo)

    largest = max(sizes)
    inner = max(1, int(round(largest * CONTENT_SCALE)))

    # Fit inside the square without distorting: scale on the constraining axis.
    ratio = min(inner / logo.width, inner / logo.height)
    scaled_size = (
        max(1, int(round(logo.width * ratio))),
        max(1, int(round(logo.height * ratio))),
    )
    # LANCZOS is the right filter when shrinking artwork with fine detail.
    scaled = logo.resize(scaled_size, Image.LANCZOS)

    canvas = Image.new("RGBA", (largest, largest), (0, 0, 0, 0))
    canvas.paste(
        scaled,
        ((largest - scaled_size[0]) // 2, (largest - scaled_size[1]) // 2),
        scaled,
    )

    destination.parent.mkdir(parents=True, exist_ok=True)
    # Pillow generates every requested size from this single master image.
    canvas.save(destination, format="ICO", sizes=[(size, size) for size in sizes])

    written = sorted(sizes)
    print(f"Wrote {destination}  ({destination.stat().st_size:,} bytes)")
    print(f"  source      : {source.name} ({original_size[0]}x{original_size[1]})")
    print(f"  tagline     : {'dropped' if tagline_dropped else 'kept (no separate block found)'}")
    print(f"  artwork     : {logo.width}x{logo.height} -> {scaled_size[0]}x{scaled_size[1]} on a {largest}x{largest} canvas")
    print(f"  resolutions : {', '.join(str(size) for size in written)}")
    return written


def main(argv: list[str] | None = None) -> int:
    root = Path(__file__).resolve().parent

    parser = argparse.ArgumentParser(description="Generate the application icon from the source logo.")
    parser.add_argument("--source", default=str(root / "UniTrans Logo.png"), help="Source logo (PNG).")
    parser.add_argument("--output", default=str(root / "assets" / "app.ico"), help="Destination .ico path.")
    args = parser.parse_args(argv)

    build_icon(Path(args.source), Path(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
