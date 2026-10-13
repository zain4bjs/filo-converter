"""Build the Filo watermark images from the FILO wordmark (run once, commit the output).

    python watermark/make_assets.py

Input:  watermark/assets/logo-source.png, the FILO wordmark on a solid
        white background, 1420 px wide.
Output: watermark/assets/watermark-filo.png       the wordmark on a transparent
                                                  background, full opacity
        watermark/assets/watermark-filo-mark.png  the same wordmark with a soft
                                                  white halo, at the mark's 60 %
                                                  opacity: what the converters
                                                  embed (PDF, DOCX, PPTX), so
                                                  no output format needs its
                                                  own transparency support

The white background is knocked out by "colour to alpha": a pixel's alpha is
how far its darkest channel is from white, and its colour is un-blended from
white, so anti-aliased edges stay smooth on any background.
Needs Pillow only.
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageChops, ImageFilter

ASSETS = Path(__file__).resolve().parent / "assets"
SOURCE = ASSETS / "logo-source.png"
PLAIN = ASSETS / "watermark-filo.png"
MARK = ASSETS / "watermark-filo-mark.png"

OUTPUT_WIDTH = 600
OPACITY = 0.6
# The logo's lightest ink (the blue gradient) has a darkest channel near 59;
# anything this far from white or further is fully opaque.
FULL_ALPHA_DISTANCE = 150


def knock_out_white(image: Image.Image) -> Image.Image:
    rgb = image.convert("RGB")
    out = Image.new("RGBA", rgb.size)
    src = rgb.load()
    dst = out.load()
    width, height = rgb.size
    for y in range(height):
        for x in range(width):
            r, g, b = src[x, y]
            distance = 255 - min(r, g, b)
            alpha = min(1.0, distance / FULL_ALPHA_DISTANCE)
            if alpha <= 0.02:
                dst[x, y] = (255, 255, 255, 0)
                continue
            # pixel = alpha * colour + (1 - alpha) * white  ->  solve for colour
            colour = tuple(
                max(0, min(255, round((c - (1 - alpha) * 255) / alpha))) for c in (r, g, b)
            )
            dst[x, y] = (*colour, round(alpha * 255))
    return out


def with_halo(logo: Image.Image, radius: int) -> Image.Image:
    pad = radius * 3
    canvas = Image.new("RGBA", (logo.width + 2 * pad, logo.height + 2 * pad), (255, 255, 255, 0))
    canvas.alpha_composite(logo, (pad, pad))
    alpha = canvas.getchannel("A")
    spread = alpha.filter(ImageFilter.MaxFilter(radius * 2 + 1)).filter(ImageFilter.GaussianBlur(radius))
    spread = spread.point(lambda v: min(255, int(v * 0.75)))
    halo = Image.new("RGBA", canvas.size, (255, 255, 255, 0))
    halo.putalpha(spread)
    halo.alpha_composite(canvas)
    return halo


def fade(image: Image.Image, opacity: float) -> Image.Image:
    alpha = image.getchannel("A").point(lambda v: round(v * opacity))
    out = image.copy()
    out.putalpha(alpha)
    return out


def main() -> None:
    plain = knock_out_white(Image.open(SOURCE))
    plain = plain.crop(plain.getchannel("A").getbbox())
    plain = plain.resize((OUTPUT_WIDTH, round(plain.height * OUTPUT_WIDTH / plain.width)), Image.LANCZOS)
    plain.save(PLAIN, optimize=True)
    mark = fade(with_halo(plain, radius=max(4, OUTPUT_WIDTH // 80)), OPACITY)
    # Trim the fully transparent border the halo padding left.
    bbox = ImageChops.difference(mark.getchannel("A"), Image.new("L", mark.size, 0)).getbbox()
    mark.crop(bbox).save(MARK, optimize=True)
    print(f"{PLAIN.name} {plain.size}  {MARK.name} {mark.crop(bbox).size}")


if __name__ == "__main__":
    main()
