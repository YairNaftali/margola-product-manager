"""Core watermarking logic. Extracted verbatim from the original Aline-session
build (Aug 2026), only LOGO_PATH updated. Locked production settings for the
2026-08-12 full-catalog run: position="center", scale=0.35, opacity=0.4.
"""
from PIL import Image

LOGO_PATH = "/Users/yair/Downloads/margola_logo_upscaled.png"


def apply_watermark(image_path, out_path, *,
                     scale=0.35, opacity=0.4,
                     position="center", margin=0.03):
    base = Image.open(image_path).convert("RGBA")
    logo = Image.open(LOGO_PATH).convert("RGBA")
    target_w = int(base.width * scale)
    target_h = int(logo.height * (target_w / logo.width))
    logo = logo.resize((target_w, target_h), Image.LANCZOS)
    if opacity < 1.0:
        alpha = logo.getchannel("A").point(lambda a: int(a * opacity))
        logo.putalpha(alpha)
    m = int(base.width * margin)
    positions = {
        "bottom-right": (base.width - target_w - m, base.height - target_h - m),
        "bottom-center": ((base.width - target_w) // 2, base.height - target_h - m),
        "bottom-left": (m, base.height - target_h - m),
        "center": ((base.width - target_w) // 2, (base.height - target_h) // 2),
        "top-right": (base.width - target_w - m, m),
    }
    pos = positions[position]
    out = base.copy()
    out.alpha_composite(logo, dest=pos)
    out.convert("RGB").save(out_path, "JPEG", quality=90)
