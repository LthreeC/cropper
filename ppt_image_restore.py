# -*- coding: utf-8 -*-
"""恢复 PowerPoint 导出 PDF 时被降采样的原始位图。"""

import io
import math
import os
import zipfile
from concurrent.futures import ThreadPoolExecutor

import pymupdf
from PIL import Image, ImageChops, ImageStat

from units import points_to_pixels


SIGNATURE_SIZE = (64, 64)
MAX_PDF_IMAGE_DPI = 300
MAX_MATCH_SCORE = 0.10
MIN_MATCH_MARGIN = 0.05
MIN_PIXEL_AREA_GAIN = 1.5
ALPHA_THRESHOLD = 8
# Pillow 解码和 Lanczos 缩放会释放 GIL；限制并发数以控制峰值内存。
MAX_IMAGE_WORKERS = 3


def restore_pptx_images(
    source_pdf,
    source_pptx,
    target_pdf,
    max_image_dpi=MAX_PDF_IMAGE_DPI,
):
    """恢复位图并保存独立 PDF，供测试和独立调用使用。"""
    candidates = _load_pptx_images(source_pptx)
    if not candidates:
        return []

    doc = pymupdf.open(source_pdf)
    try:
        replacements = restore_pptx_images_in_document(
            doc,
            source_pptx,
            max_image_dpi=max_image_dpi,
            candidates=candidates,
        )
        if not replacements:
            return []
        doc.save(target_pdf, garbage=4, deflate=True)
        return replacements
    finally:
        doc.close()


def restore_pptx_images_in_document(
    doc,
    source_pptx,
    max_image_dpi=MAX_PDF_IMAGE_DPI,
    candidates=None,
):
    """在已打开文档中恢复位图；PyMuPDF 读取和写入始终保持单线程。"""
    if candidates is None:
        candidates = _load_pptx_images(source_pptx)
    if not candidates:
        return []

    work_items = []
    for pdf_image in _collect_pdf_images(doc):
        try:
            rendered = _pdf_image(
                doc,
                pdf_image["xref"],
                pdf_image["smask"],
            )
            signature = _signature(rendered, premultiplied=True)
        except Exception:
            continue

        ranked = [
            (
                _signature_score(signature, candidate["signature"]),
                candidate,
            )
            for candidate in candidates
        ]
        ranked.sort(key=lambda pair: pair[0])
        if not ranked or ranked[0][0] > MAX_MATCH_SCORE:
            continue
        if (
            len(ranked) > 1
            and ranked[1][0] - ranked[0][0] < MIN_MATCH_MARGIN
        ):
            continue

        score, candidate = ranked[0]
        work_items.append(
            (
                pdf_image,
                candidate,
                rendered,
                score,
                max_image_dpi,
            )
        )

    worker_count = _image_worker_count(len(work_items))
    if worker_count > 1:
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            prepared = list(executor.map(_prepare_replacement, work_items))
    else:
        prepared = [_prepare_replacement(item) for item in work_items]

    replacements = [item for item in prepared if item is not None]
    for item in replacements:
        doc[item["page_index"]].replace_image(
            item["xref"],
            stream=item["stream"],
        )

    return [
        {key: value for key, value in item.items() if key != "stream"}
        for item in replacements
    ]


def _prepare_replacement(work_item):
    pdf_image, candidate, rendered, score, max_image_dpi = work_item
    stream, replacement_size = _replacement_stream(
        candidate,
        rendered,
        pdf_image["display_size"],
        max_image_dpi,
    )
    current_area = rendered.width * rendered.height
    replacement_area = replacement_size[0] * replacement_size[1]
    if replacement_area <= current_area * MIN_PIXEL_AREA_GAIN:
        return None

    return {
        "xref": pdf_image["xref"],
        "page_index": pdf_image["page_index"],
        "media_name": candidate["name"],
        "score": score,
        "old_size": rendered.size,
        "new_size": replacement_size,
        "stream": stream,
    }


def _image_worker_count(item_count):
    cpu_count = max(1, os.cpu_count() or 1)
    return max(1, min(MAX_IMAGE_WORKERS, item_count, cpu_count))


def _load_pptx_images(source_pptx):
    media_items = []
    with zipfile.ZipFile(source_pptx) as archive:
        for member in archive.namelist():
            if member.startswith("ppt/media/") and not member.endswith("/"):
                media_items.append((member, archive.read(member)))

    worker_count = _image_worker_count(len(media_items))
    if worker_count > 1:
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            loaded = list(executor.map(_load_pptx_candidate, media_items))
    else:
        loaded = [_load_pptx_candidate(item) for item in media_items]
    return [candidate for candidate in loaded if candidate is not None]


def _load_pptx_candidate(media_item):
    member, data = media_item
    try:
        with Image.open(io.BytesIO(data)) as opened:
            image = opened.convert("RGBA")
            image.load()
        signature = _signature(image, premultiplied=False)
    except Exception:
        return None

    return {
        "name": os.path.basename(member),
        "data": data,
        "signature": signature,
    }


def _collect_pdf_images(doc):
    images = {}
    for page_index, page in enumerate(doc):
        for info in page.get_images(full=True):
            xref, smask = info[:2]
            if xref and xref not in images:
                images[xref] = {
                    "xref": xref,
                    "smask": smask,
                    "page_index": page_index,
                    "display_size": (0.0, 0.0),
                }
            if not xref:
                continue
            try:
                placements = page.get_image_rects(xref, transform=True)
            except Exception:
                placements = []
            width, height = images[xref]["display_size"]
            for _, transform in placements:
                width = max(width, math.hypot(transform.a, transform.b))
                height = max(height, math.hypot(transform.c, transform.d))
            images[xref]["display_size"] = (width, height)
    return list(images.values())


def _pdf_image(doc, xref, smask):
    base = pymupdf.Pixmap(pymupdf.Pixmap(doc, xref), 0)
    if smask:
        mask = pymupdf.Pixmap(doc, smask)
        pixmap = pymupdf.Pixmap(base, mask)
        mode = "RGBA"
    else:
        pixmap = base
        mode = "RGB"
    return Image.frombytes(mode, (pixmap.width, pixmap.height), pixmap.samples).convert("RGBA")


def _visible_bbox(image):
    alpha = image.getchannel("A")
    alpha_min, alpha_max = alpha.getextrema()
    if alpha_min > ALPHA_THRESHOLD:
        return (0, 0, image.width, image.height)
    if alpha_max <= ALPHA_THRESHOLD:
        return None
    mask = alpha.point(lambda value: 255 if value > ALPHA_THRESHOLD else 0)
    return mask.getbbox()


def _signature(image, premultiplied):
    bbox = _visible_bbox(image)
    if bbox:
        image = image.crop(bbox)
    image = image.resize(SIGNATURE_SIZE, Image.Resampling.LANCZOS)
    red, green, blue, alpha = image.split()

    if premultiplied:
        black = Image.merge("RGB", (red, green, blue))
    else:
        black = Image.merge(
            "RGB",
            tuple(ImageChops.multiply(channel, alpha) for channel in (red, green, blue)),
        )

    white = Image.merge(
        "RGB",
        tuple(
            ImageChops.add(channel, ImageChops.invert(alpha))
            for channel in black.split()
        ),
    )
    return black, white, alpha


def _signature_score(left, right):
    return (
        _normalized_mae(left[0], right[0]) * 0.4
        + _normalized_mae(left[1], right[1]) * 0.4
        + _normalized_mae(left[2], right[2]) * 0.2
    )


def _normalized_mae(left, right):
    difference = ImageChops.difference(left, right)
    means = ImageStat.Stat(difference).mean
    return sum(means) / (255 * len(means))


def _replacement_stream(candidate, pdf_image, display_size, max_image_dpi):
    with Image.open(io.BytesIO(candidate["data"])) as opened:
        source = opened.convert("RGBA")
        source.load()
    source_bbox = _visible_bbox(source)
    pdf_bbox = _visible_bbox(pdf_image)
    modified = False

    if source_bbox and pdf_bbox and source_bbox != (0, 0, source.width, source.height):
        source = _match_transparent_margins(source, source_bbox, pdf_image.size, pdf_bbox)
        modified = True

    limited = _limit_image_resolution(source, display_size, max_image_dpi)
    if limited.size != source.size:
        source = limited
        modified = True

    if not modified:
        return candidate["data"], source.size

    buffer = io.BytesIO()
    # PNG 仅供 PyMuPDF 解码；最终 PDF 会统一重新压缩。
    if source.getchannel("A").getextrema()[0] < 255:
        source.save(buffer, "PNG", compress_level=1)
    elif os.path.splitext(candidate["name"])[1].lower() in (".jpg", ".jpeg"):
        source.convert("RGB").save(
            buffer,
            "JPEG",
            quality=95,
            subsampling=0,
            optimize=True,
        )
    else:
        source.save(buffer, "PNG", compress_level=1)
    return buffer.getvalue(), source.size


def _limit_image_resolution(source, display_size, max_image_dpi):
    display_width, display_height = display_size
    if not display_width or not display_height or not max_image_dpi:
        return source

    max_width = max(1, round(points_to_pixels(display_width, max_image_dpi)))
    max_height = max(1, round(points_to_pixels(display_height, max_image_dpi)))
    target_size = (
        min(source.width, max_width),
        min(source.height, max_height),
    )
    if target_size == source.size:
        return source
    return source.resize(target_size, Image.Resampling.LANCZOS)



def _match_transparent_margins(source, source_bbox, pdf_size, pdf_bbox):
    source_width = source_bbox[2] - source_bbox[0]
    source_height = source_bbox[3] - source_bbox[1]
    pdf_width = pdf_bbox[2] - pdf_bbox[0]
    pdf_height = pdf_bbox[3] - pdf_bbox[1]
    if not source_width or not source_height or not pdf_width or not pdf_height:
        return source

    scale_x = source_width / pdf_width
    scale_y = source_height / pdf_height
    crop_box = (
        round(source_bbox[0] - pdf_bbox[0] * scale_x),
        round(source_bbox[1] - pdf_bbox[1] * scale_y),
        round(source_bbox[2] + (pdf_size[0] - pdf_bbox[2]) * scale_x),
        round(source_bbox[3] + (pdf_size[1] - pdf_bbox[3]) * scale_y),
    )
    return source.crop(crop_box)
