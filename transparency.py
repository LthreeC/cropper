# -*- coding: utf-8 -*-
"""
Background-to-alpha helpers.
"""

import re


def parse_hex_color(value):
    """Parse #RRGGBB or RRGGBB into an RGB tuple."""
    text = (value or "").strip()
    match = re.fullmatch(r"#?([0-9a-fA-F]{6})", text)
    if not match:
        raise ValueError("颜色必须是 #RRGGBB 格式")
    hex_value = match.group(1)
    return (
        int(hex_value[0:2], 16),
        int(hex_value[2:4], 16),
        int(hex_value[4:6], 16),
    )


def rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)


def sample_target_color(img, mode="corners", custom_color=None):
    """Pick the color that should become transparent."""
    if mode == "custom":
        return parse_hex_color(custom_color)

    rgba = img.convert("RGBA")
    width, height = rgba.size

    if mode == "top_left":
        samples = [rgba.getpixel((0, 0))]
    else:
        samples = [
            rgba.getpixel((0, 0)),
            rgba.getpixel((width - 1, 0)),
            rgba.getpixel((0, height - 1)),
            rgba.getpixel((width - 1, height - 1)),
        ]

    visible = [pixel for pixel in samples if pixel[3] > 0]
    if not visible:
        visible = samples

    count = len(visible)
    return tuple(int(round(sum(pixel[i] for pixel in visible) / count)) for i in range(3))


def make_transparent(
    img,
    target_color=None,
    color_mode="corners",
    custom_color=None,
    tolerance=18,
    edge_only=True,
    feather=1,
):
    """Return (RGBA image, target_color, removed_ratio)."""
    import numpy as np
    from PIL import Image, ImageFilter

    rgba = img.convert("RGBA")
    if target_color is None:
        target_color = sample_target_color(rgba, color_mode, custom_color)

    arr = np.array(rgba)
    rgb = arr[:, :, :3].astype(np.int16)
    alpha = arr[:, :, 3]
    target = np.array(target_color, dtype=np.int16)
    diff = np.max(np.abs(rgb - target), axis=2)
    matched = (diff <= max(0, int(tolerance))) & (alpha > 0)

    if edge_only:
        matched = _edge_connected_mask(matched)

    removed = int(matched.sum())
    if removed == 0:
        return rgba, target_color, 0.0

    if feather > 0:
        mask_img = Image.fromarray(matched.astype("uint8") * 255)
        blurred = mask_img.filter(ImageFilter.GaussianBlur(radius=float(feather)))
        remove_alpha = np.asarray(blurred, dtype=np.uint8)
        arr[:, :, 3] = np.minimum(alpha, 255 - remove_alpha)
    else:
        arr[matched, 3] = 0

    return Image.fromarray(arr), target_color, removed / matched.size


def _edge_connected_mask(mask):
    """Keep only mask pixels connected to the image edge."""
    import numpy as np
    from PIL import Image, ImageDraw, ImageOps

    if not mask.any():
        return mask

    height, width = mask.shape
    mask_img = Image.fromarray(mask.astype("uint8") * 255)
    padded = ImageOps.expand(mask_img, border=1, fill=255)
    ImageDraw.floodfill(padded, (0, 0), 128, thresh=0)
    connected = padded.crop((1, 1, width + 1, height + 1))
    return np.asarray(connected) == 128
