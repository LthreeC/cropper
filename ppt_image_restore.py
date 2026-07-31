# -*- coding: utf-8 -*-
"""恢复 PowerPoint 导出 PDF 时被降采样的原始位图。"""

import io
import math
import os
import posixpath
import zipfile
from concurrent.futures import ThreadPoolExecutor
from xml.etree import ElementTree

import pymupdf
from PIL import Image, ImageChops, ImageOps, ImageStat

from units import points_to_pixels


SIGNATURE_SIZE = (64, 64)
MAX_PDF_IMAGE_DPI = 300
MAX_MATCH_SCORE = 0.10
MIN_MATCH_MARGIN = 0.05
MIN_PIXEL_AREA_GAIN = 1.5
ALPHA_THRESHOLD = 8
# Pillow 解码和 Lanczos 缩放会释放 GIL；限制并发数以控制峰值内存。
MAX_IMAGE_WORKERS = 3
EXIF_ORIENTATION_TAG = 274
RELATIONSHIP_NAMESPACES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
)


class RestoreResult(list):
    """兼容原列表返回值，同时通过 stats 提供每个 PDF 位图的处理结果。

    restored 为已替换，unmatched 为无可靠候选或处理失败，ambiguous 为
    多个不同视觉候选无法安全区分，sufficient 为候选不会带来明显像素提升。
    """

    def __init__(self, items=(), stats=None):
        super().__init__(items)
        self.stats = stats or {
            "total": 0,
            "restored": 0,
            "unmatched": 0,
            "ambiguous": 0,
            "sufficient": 0,
        }


class _PptxCandidates(list):
    def __init__(
        self,
        items=(),
        slide_count=0,
        order_reliable=False,
        visible_slide_indices=None,
    ):
        super().__init__(items)
        self.slide_count = slide_count
        self.order_reliable = order_reliable
        self.visible_slide_indices = visible_slide_indices


def pptx_visible_slide_indices(source_pptx):
    """返回可靠的 0-based 可见幻灯片页序，不解析媒体内容。"""
    with zipfile.ZipFile(source_pptx) as archive:
        slide_members, order_reliable = _ordered_presentation_slides(
            archive,
            {},
        )
        if not order_reliable:
            raise ValueError("无法从 PPTX 确定可靠的幻灯片页序")

        visible_indices = []
        for slide_index, slide_member in enumerate(slide_members):
            visible = _slide_is_visible(archive, slide_member)
            if visible is None:
                raise ValueError("无法从 PPTX 确定隐藏幻灯片状态")
            if visible:
                visible_indices.append(slide_index)
        return tuple(visible_indices)


def restore_pptx_images(
    source_pdf,
    source_pptx,
    target_pdf,
    max_image_dpi=MAX_PDF_IMAGE_DPI,
    slide_indices=None,
):
    """恢复位图并保存独立 PDF，供测试和独立调用使用。"""
    candidates = _load_pptx_images(source_pptx)

    doc = pymupdf.open(source_pdf)
    try:
        replacements = restore_pptx_images_in_document(
            doc,
            source_pptx,
            max_image_dpi=max_image_dpi,
            candidates=candidates,
            slide_indices=slide_indices,
        )
        if not replacements:
            return replacements
        doc.save(target_pdf, garbage=4, deflate=True)
        return replacements
    finally:
        doc.close()


def restore_pptx_images_in_document(
    doc,
    source_pptx,
    max_image_dpi=MAX_PDF_IMAGE_DPI,
    candidates=None,
    slide_indices=None,
):
    """在已打开文档中恢复位图；slide_indices 为各 PDF 页对应的 0-based PPT 页。"""
    if candidates is None:
        candidates = _load_pptx_images(source_pptx)

    pdf_images = _collect_pdf_images(doc)
    stats = {
        "total": len(pdf_images),
        "restored": 0,
        "unmatched": 0,
        "ambiguous": 0,
        "sufficient": 0,
    }
    if not candidates:
        stats["unmatched"] = len(pdf_images)
        return RestoreResult(stats=stats)

    slide_count = getattr(candidates, "slide_count", 0)
    order_reliable = getattr(candidates, "order_reliable", True)
    visible_slide_indices = getattr(candidates, "visible_slide_indices", None)
    page_to_slide = None
    if slide_indices is not None and order_reliable:
        supplied_indices = tuple(slide_indices)
        if len(supplied_indices) == doc.page_count:
            page_to_slide = supplied_indices
    elif order_reliable and slide_count > 0:
        if slide_count == doc.page_count:
            page_to_slide = tuple(range(doc.page_count))
        elif (
            visible_slide_indices is not None
            and len(visible_slide_indices) == doc.page_count
        ):
            page_to_slide = visible_slide_indices
    work_items = []
    for pdf_image in pdf_images:
        if pdf_image["multiple_placements_on_page"]:
            stats["unmatched"] += 1
            continue
        try:
            rendered = _pdf_image(
                doc,
                pdf_image["xref"],
                pdf_image["smask"],
            )
            signature = _signature(rendered, premultiplied=True)
        except Exception:
            stats["unmatched"] += 1
            continue

        eligible = candidates
        slide_indices_for_image = None
        if page_to_slide is not None:
            slide_indices_for_image = {
                page_to_slide[page_index]
                for page_index in pdf_image["page_indices"]
            }
            eligible = [
                candidate
                for candidate in candidates
                if not candidate["pages"]
                or slide_indices_for_image.intersection(candidate["pages"])
            ]
        ranked = [
            (
                _signature_score(signature, candidate["signature"]),
                candidate,
            )
            for candidate in eligible
        ]
        ranked.sort(
            key=lambda pair: (
                pair[0],
                -pair[1]["pixel_area"],
                pair[1]["name"],
            )
        )
        if not ranked or ranked[0][0] > MAX_MATCH_SCORE:
            stats["unmatched"] += 1
            continue

        best_score, best_candidate = ranked[0]
        if (
            slide_indices_for_image is not None
            and len(pdf_image["page_indices"]) > 1
            and best_candidate["pages"]
            and not slide_indices_for_image.issubset(
                best_candidate["pages"]
            )
        ):
            stats["unmatched"] += 1
            continue
        if (
            len(ranked) > 1
            and ranked[1][0] - best_score < MIN_MATCH_MARGIN
        ):
            stats["ambiguous"] += 1
            continue

        work_items.append(
            (
                pdf_image,
                best_candidate,
                rendered,
                best_score,
                max_image_dpi,
            )
        )

    worker_count = _image_worker_count(len(work_items))
    if worker_count > 1:
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            prepared = list(executor.map(_prepare_replacement, work_items))
    else:
        prepared = [_prepare_replacement(item) for item in work_items]

    replacements = []
    for status, item in prepared:
        stats[status] += 1
        if item is not None:
            replacements.append(item)
    for item in replacements:
        doc[item["page_index"]].replace_image(
            item["xref"],
            stream=item["stream"],
        )

    return RestoreResult(
        (
            {key: value for key, value in item.items() if key != "stream"}
            for item in replacements
        ),
        stats=stats,
    )


def _prepare_replacement(work_item):
    pdf_image, candidate, rendered, score, max_image_dpi = work_item
    try:
        stream, replacement_size = _replacement_stream(
            candidate,
            rendered,
            pdf_image["display_size"],
            max_image_dpi,
        )
    except Exception:
        return "unmatched", None
    current_area = rendered.width * rendered.height
    replacement_area = replacement_size[0] * replacement_size[1]
    if replacement_area <= current_area * MIN_PIXEL_AREA_GAIN:
        return "sufficient", None

    return (
        "restored",
        {
            "xref": pdf_image["xref"],
            "page_index": pdf_image["page_index"],
            "media_name": candidate["name"],
            "score": score,
            "old_size": rendered.size,
            "new_size": replacement_size,
            "stream": stream,
        },
    )


def _image_worker_count(item_count):
    cpu_count = max(1, os.cpu_count() or 1)
    return max(1, min(MAX_IMAGE_WORKERS, item_count, cpu_count))


def _load_pptx_images(source_pptx):
    with zipfile.ZipFile(source_pptx) as archive:
        (
            references,
            slide_count,
            order_reliable,
            visible_slide_indices,
        ) = _pptx_image_references(archive)
        media_groups = {}
        for member in archive.namelist():
            if member.startswith("ppt/media/") and not member.endswith("/"):
                data = archive.read(member)
                group = media_groups.setdefault(
                    data,
                    {"members": [], "data": data},
                )
                group["members"].append(member)

    media_items = []
    for group in media_groups.values():
        members = sorted(group["members"])
        variants = {}
        for member in members:
            for page_index, crop in references.get(member, ()):
                variants.setdefault(crop, set()).add(page_index)
        if not variants and slide_count == 0:
            variants[None] = set()
        for crop, pages in variants.items():
            media_items.append(
                (
                    members[0],
                    group["data"],
                    crop,
                    frozenset(pages),
                )
            )

    worker_count = _image_worker_count(len(media_items))
    if worker_count > 1:
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            loaded = list(executor.map(_load_pptx_candidate, media_items))
    else:
        loaded = [_load_pptx_candidate(item) for item in media_items]
    return _PptxCandidates(
        (candidate for candidate in loaded if candidate is not None),
        slide_count=slide_count,
        order_reliable=order_reliable,
        visible_slide_indices=visible_slide_indices,
    )


def _load_pptx_candidate(media_item):
    member, data, crop, pages = media_item
    try:
        with Image.open(io.BytesIO(data)) as opened:
            orientation = opened.getexif().get(EXIF_ORIENTATION_TAG, 1)
            normalized = ImageOps.exif_transpose(opened)
            image = normalized.convert("RGBA")
            image.load()
        modified = orientation not in (None, 1)
        if crop is not None:
            crop_box = _src_rect_crop_box(image.size, crop)
            if crop_box is None:
                return None
            image = image.crop(crop_box)
            modified = True
        if modified:
            data = _encode_candidate_image(
                image,
                os.path.splitext(member)[1].lower(),
            )
        signature = _signature(image, premultiplied=False)
    except Exception:
        return None

    return {
        "name": os.path.basename(member),
        "data": data,
        "signature": signature,
        "size": image.size,
        "pixel_area": image.width * image.height,
        "pages": pages,
    }


def _pptx_image_references(archive):
    references = {}
    part_cache = {}
    slide_members, order_reliable = _ordered_presentation_slides(
        archive,
        part_cache,
    )
    if not order_reliable:
        slide_members = tuple(sorted(
            (
                member
                for member in archive.namelist()
                if _slide_member_number(member) is not None
            ),
            key=_slide_member_number,
        ))

    visible_slide_indices = []
    visibility_reliable = True
    for slide_index, slide_member in enumerate(slide_members):
        visible = _slide_is_visible(archive, slide_member)
        if visible is None:
            visibility_reliable = False
        elif visible:
            visible_slide_indices.append(slide_index)

        parts = [slide_member]
        slide_relationships = _part_relationships(
            archive,
            slide_member,
            part_cache,
        )
        layout_member = _related_part(
            slide_relationships,
            "slideLayout",
        )
        if layout_member:
            parts.append(layout_member)
            layout_relationships = _part_relationships(
                archive,
                layout_member,
                part_cache,
            )
            master_member = _related_part(
                layout_relationships,
                "slideMaster",
            )
            if master_member:
                parts.append(master_member)

        for part_member in parts:
            for media_member, crop in _part_image_references(
                archive,
                part_member,
                part_cache,
            ):
                references.setdefault(media_member, []).append(
                    (slide_index, crop)
                )

    visible_indices = (
        tuple(visible_slide_indices)
        if visibility_reliable
        else None
    )
    return (
        references,
        len(slide_members),
        order_reliable,
        visible_indices,
    )


def _ordered_presentation_slides(archive, cache):
    presentation_member = "ppt/presentation.xml"
    try:
        root = ElementTree.fromstring(archive.read(presentation_member))
    except (KeyError, ElementTree.ParseError):
        return (), False

    relationships = _part_relationships(
        archive,
        presentation_member,
        cache,
    )
    slide_ids = list(_elements_by_local_name(root, "sldId"))
    if not slide_ids:
        has_slide_parts = any(
            _slide_member_number(member) is not None
            for member in archive.namelist()
        )
        return (), not has_slide_parts

    ordered = []
    seen = set()
    archive_members = set(archive.namelist())
    for slide_id in slide_ids:
        relationship_id = _relationship_attribute(slide_id, "id")
        relationship = relationships.get(relationship_id)
        if (
            not relationship
            or not relationship[1].endswith("/slide")
            or relationship[0] not in archive_members
            or relationship[0] in seen
        ):
            return (), False
        ordered.append(relationship[0])
        seen.add(relationship[0])
    return tuple(ordered), True


def _slide_member_number(member):
    prefix = "ppt/slides/slide"
    suffix = ".xml"
    if not member.startswith(prefix) or not member.endswith(suffix):
        return None
    number = member[len(prefix) : -len(suffix)]
    return int(number) if number.isdigit() else None


def _slide_is_visible(archive, slide_member):
    try:
        root = ElementTree.fromstring(archive.read(slide_member))
    except (KeyError, ElementTree.ParseError):
        return None
    return root.attrib.get("show", "1").strip().lower() not in (
        "0",
        "false",
        "off",
        "no",
    )


def _part_relationships(archive, part_member, cache):
    cache_key = ("relationships", part_member)
    if cache_key in cache:
        return cache[cache_key]

    rels_member = posixpath.join(
        posixpath.dirname(part_member),
        "_rels",
        posixpath.basename(part_member) + ".rels",
    )
    relationships = {}
    try:
        rels_root = ElementTree.fromstring(archive.read(rels_member))
    except (KeyError, ElementTree.ParseError):
        cache[cache_key] = relationships
        return relationships

    for relationship in rels_root:
        if relationship.attrib.get("TargetMode") == "External":
            continue
        relationship_id = relationship.attrib.get("Id")
        target = relationship.attrib.get("Target")
        relationship_type = relationship.attrib.get("Type", "")
        if not relationship_id or not target:
            continue
        target_member = posixpath.normpath(
            posixpath.join(posixpath.dirname(part_member), target)
        ).lstrip("/")
        relationships[relationship_id] = (target_member, relationship_type)

    cache[cache_key] = relationships
    return relationships


def _related_part(relationships, relationship_name):
    suffix = "/" + relationship_name
    for target_member, relationship_type in relationships.values():
        if relationship_type.endswith(suffix):
            return target_member
    return None


def _part_image_references(archive, part_member, cache):
    cache_key = ("images", part_member)
    if cache_key in cache:
        return cache[cache_key]

    relationships = _part_relationships(archive, part_member, cache)
    try:
        root = ElementTree.fromstring(archive.read(part_member))
    except (KeyError, ElementTree.ParseError):
        cache[cache_key] = ()
        return ()

    image_references = []
    for blip_fill in _elements_by_local_name(root, "blipFill"):
        blip = next(_elements_by_local_name(blip_fill, "blip"), None)
        if blip is None:
            continue
        relationship = relationships.get(
            _relationship_attribute(blip, "embed")
        )
        if not relationship or not relationship[1].endswith("/image"):
            continue
        media_member = relationship[0]
        if not media_member.startswith("ppt/media/"):
            continue
        src_rect = next(
            _elements_by_local_name(blip_fill, "srcRect"),
            None,
        )
        crop = None
        if src_rect is not None:
            crop = tuple(
                _parse_src_rect_value(src_rect.attrib.get(side))
                for side in ("l", "t", "r", "b")
            )
            if crop == (0, 0, 0, 0):
                crop = None
        image_references.append((media_member, crop))

    result = tuple(image_references)
    cache[cache_key] = result
    return result


def _elements_by_local_name(root, name):
    suffix = "}" + name
    return (
        element
        for element in root.iter()
        if element.tag == name or element.tag.endswith(suffix)
    )


def _relationship_attribute(element, name):
    for namespace in RELATIONSHIP_NAMESPACES:
        value = element.attrib.get(f"{{{namespace}}}{name}")
        if value:
            return value
    return None


def _parse_src_rect_value(value):
    try:
        return max(0, min(100000, int(value or 0)))
    except (TypeError, ValueError):
        return 0


def _src_rect_crop_box(size, crop):
    width, height = size
    left, top, right, bottom = crop
    box = (
        round(width * left / 100000),
        round(height * top / 100000),
        width - round(width * right / 100000),
        height - round(height * bottom / 100000),
    )
    if box[0] >= box[2] or box[1] >= box[3]:
        return None
    return box


def _encode_candidate_image(image, extension):
    buffer = io.BytesIO()
    if image.getchannel("A").getextrema()[0] < 255:
        image.save(buffer, "PNG", compress_level=1)
    elif extension in (".jpg", ".jpeg"):
        image.convert("RGB").save(
            buffer,
            "JPEG",
            quality=95,
            subsampling=0,
            optimize=True,
        )
    else:
        image.save(buffer, "PNG", compress_level=1)
    return buffer.getvalue()


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
                    "page_indices": set(),
                    "placement_pages": set(),
                    "multiple_placements_on_page": False,
                    "display_size": (0.0, 0.0),
                }
            if not xref:
                continue
            images[xref]["page_indices"].add(page_index)
            if page_index in images[xref]["placement_pages"]:
                continue
            images[xref]["placement_pages"].add(page_index)
            try:
                placements = page.get_image_rects(xref, transform=True)
            except Exception:
                placements = []
            if len(placements) > 1:
                images[xref]["multiple_placements_on_page"] = True
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
