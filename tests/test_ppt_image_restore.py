from contextlib import contextmanager
import io
import os
import unittest
import zipfile

import pymupdf
from PIL import Image, ImageDraw, ImageOps

from ppt_image_restore import (
    _load_pptx_images,
    pptx_visible_slide_indices,
    restore_pptx_images,
)


def _png_bytes(image):
    buffer = io.BytesIO()
    image.save(buffer, "PNG")
    return buffer.getvalue()


def _jpeg_bytes(image, orientation=None):
    buffer = io.BytesIO()
    exif = Image.Exif()
    if orientation is not None:
        exif[274] = orientation
    image.save(
        buffer,
        "JPEG",
        quality=98,
        subsampling=0,
        exif=exif,
    )
    return buffer.getvalue()


def _slide_xml(relationship_id, crop=None):
    src_rect = ""
    if crop is not None:
        left, top, right, bottom = crop
        src_rect = (
            f'<a:srcRect l="{left}" t="{top}" '
            f'r="{right}" b="{bottom}"/>'
        )
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
 <p:cSld><p:spTree><p:pic><p:blipFill>
  <a:blip r:embed="{relationship_id}"/>{src_rect}
 </p:blipFill></p:pic></p:spTree></p:cSld>
</p:sld>'''


def _slide_rels(relationship_id, media_name):
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="{relationship_id}"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  Target="../media/{media_name}"/>
</Relationships>'''


def _presentation_xml(relationship_ids):
    slide_ids = "".join(
        f'<p:sldId id="{256 + index}" r:id="{relationship_id}"/>'
        for index, relationship_id in enumerate(relationship_ids)
    )
    return (
        '<p:presentation '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<p:sldIdLst>{slide_ids}</p:sldIdLst></p:presentation>"
    )


def _write_ordered_pptx(path, media, slides):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for media_name, data in media.items():
            archive.writestr(f"ppt/media/{media_name}", data)
        presentation_relationships = []
        for slide_number, slide in enumerate(slides, 1):
            slide_member, media_name, crop, hidden = slide
            relationship_id = f"rId{slide_number}"
            slide_xml = _slide_xml(relationship_id, crop)
            if hidden:
                slide_xml = slide_xml.replace(
                    "<p:sld ",
                    '<p:sld show="0" ',
                    1,
                )
            archive.writestr(
                f"ppt/slides/{slide_member}",
                slide_xml,
            )
            archive.writestr(
                f"ppt/slides/_rels/{slide_member}.rels",
                _slide_rels(relationship_id, media_name),
            )
            presentation_relationship_id = f"rIdSlide{slide_number}"
            presentation_relationships.append((
                presentation_relationship_id,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                f"slides/{slide_member}",
            ))

        archive.writestr(
            "ppt/presentation.xml",
            _presentation_xml(
                relationship[0]
                for relationship in presentation_relationships
            ),
        )
        archive.writestr(
            "ppt/_rels/presentation.xml.rels",
            _relationships_xml(presentation_relationships),
        )


def _write_pptx(path, media, slides=()):
    ordered_slides = tuple(
        (f"slide{index}.xml", media_name, crop, False)
        for index, (media_name, crop) in enumerate(slides, 1)
    )
    _write_ordered_pptx(path, media, ordered_slides)


def _relationships_xml(relationships):
    items = "".join(
        f'<Relationship Id="{relationship_id}" Type="{relationship_type}" '
        f'Target="{target}"/>'
        for relationship_id, relationship_type, target in relationships
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships '
        'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{items}</Relationships>"
    )


def _picture_part_xml(relationship_id=None):
    picture = ""
    if relationship_id:
        picture = (
            "<p:pic><p:blipFill>"
            f'<a:blip r:embed="{relationship_id}"/>'
            "</p:blipFill></p:pic>"
        )
    return (
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        f"<p:cSld><p:spTree>{picture}</p:spTree></p:cSld></p:sld>"
    )


@contextmanager
def _workspace_temp_files():
    directory = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "tmp", "pdfs")
    )
    os.makedirs(directory, exist_ok=True)
    paths = [
        os.path.join(directory, name)
        for name in ("powerpoint.pdf", "source.pptx", "restored.pdf")
    ]
    for path in paths:
        if os.path.exists(path):
            os.remove(path)
    try:
        yield directory
    finally:
        for path in paths:
            if os.path.exists(path):
                os.remove(path)


class PptImageRestoreTests(unittest.TestCase):
    def test_deduplicates_identical_media_bytes(self):
        source = Image.new("RGB", (400, 300), "navy")
        data = _png_bytes(source)

        with _workspace_temp_files() as directory:
            source_pptx = os.path.join(directory, "source.pptx")
            _write_pptx(
                source_pptx,
                {"image1.png": data, "image2.png": data},
            )

            candidates = _load_pptx_images(source_pptx)

            self.assertEqual(len(candidates), 1)
            self.assertEqual(candidates[0]["name"], "image1.png")

    def test_exif_orientation_is_applied_and_reencoded(self):
        stored = Image.new("RGB", (320, 160), "red")
        draw = ImageDraw.Draw(stored)
        draw.rectangle((170, 10, 310, 150), fill="blue")
        oriented_data = _jpeg_bytes(stored, orientation=6)
        with Image.open(io.BytesIO(oriented_data)) as opened:
            visual = ImageOps.exif_transpose(opened).convert("RGB")
        low_resolution = visual.resize((20, 40), Image.Resampling.LANCZOS)

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            page.insert_image(
                pymupdf.Rect(20, 20, 56, 92),
                stream=_png_bytes(low_resolution),
            )
            doc.save(source_pdf)
            doc.close()
            _write_pptx(source_pptx, {"oriented.jpg": oriented_data})

            candidates = _load_pptx_images(source_pptx)
            with Image.open(io.BytesIO(candidates[0]["data"])) as normalized:
                self.assertEqual(normalized.size, (160, 320))
                self.assertNotIn(274, normalized.getexif())

            matches = restore_pptx_images(source_pdf, source_pptx, restored_pdf)

            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0]["new_size"], (150, 300))
            self.assertEqual(matches.stats["restored"], 1)

    def test_skips_different_media_when_visual_signatures_are_ambiguous(self):
        source = Image.new("RGB", (800, 800), "white")
        draw = ImageDraw.Draw(source)
        draw.rectangle((80, 120, 720, 680), fill="darkgreen")
        draw.ellipse((250, 250, 550, 550), fill="gold")
        smaller = source.resize((200, 200), Image.Resampling.LANCZOS)
        low_resolution = source.resize((40, 40), Image.Resampling.LANCZOS)

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(low_resolution),
            )
            doc.save(source_pdf)
            doc.close()
            _write_pptx(
                source_pptx,
                {
                    "image_small.png": _png_bytes(smaller),
                    "image_large.png": _png_bytes(source),
                },
            )

            matches = restore_pptx_images(
                source_pdf,
                source_pptx,
                restored_pdf,
                max_image_dpi=600,
            )

            self.assertEqual(matches, [])
            self.assertEqual(matches.stats["ambiguous"], 1)
            self.assertFalse(os.path.exists(restored_pdf))

    def test_never_replaces_small_local_detail_with_larger_blank_media(self):
        correct = Image.new("RGB", (400, 400), "white")
        ImageDraw.Draw(correct).rectangle((184, 184, 215, 215), fill="red")
        unrelated = Image.new("RGB", (800, 800), "white")
        low_resolution = correct.resize((40, 40), Image.Resampling.LANCZOS)

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(low_resolution),
            )
            doc.save(source_pdf)
            doc.close()
            _write_pptx(
                source_pptx,
                {
                    "correct.png": _png_bytes(correct),
                    "blank.png": _png_bytes(unrelated),
                },
            )

            matches = restore_pptx_images(
                source_pdf,
                source_pptx,
                restored_pdf,
                max_image_dpi=600,
            )

            self.assertEqual(matches, [])
            self.assertEqual(matches.stats["ambiguous"], 1)
            self.assertFalse(os.path.exists(restored_pdf))

    def test_restores_src_rect_cropped_picture(self):
        source = Image.new("RGB", (800, 400), "red")
        draw = ImageDraw.Draw(source)
        draw.rectangle((400, 0, 799, 399), fill="blue")
        draw.ellipse((500, 100, 700, 300), fill="white")
        cropped = source.crop((400, 0, 800, 400))
        low_resolution = cropped.resize((40, 40), Image.Resampling.LANCZOS)

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(low_resolution),
            )
            doc.save(source_pdf)
            doc.close()
            _write_pptx(
                source_pptx,
                {"image1.png": _png_bytes(source)},
                slides=(("image1.png", (50000, 0, 0, 0)),),
            )

            matches = restore_pptx_images(source_pdf, source_pptx, restored_pdf)

            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0]["new_size"], (300, 300))

    def test_current_slide_mapping_limits_candidates(self):
        large = Image.new("RGB", (800, 800), "purple")
        small = large.resize((400, 400), Image.Resampling.LANCZOS)
        low_resolution = large.resize((40, 40), Image.Resampling.LANCZOS)

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(low_resolution),
            )
            doc.save(source_pdf)
            doc.close()
            _write_pptx(
                source_pptx,
                {
                    "slide1.png": _png_bytes(large),
                    "slide2.png": _png_bytes(small),
                },
                slides=(
                    ("slide1.png", None),
                    ("slide2.png", None),
                ),
            )

            matches = restore_pptx_images(
                source_pdf,
                source_pptx,
                restored_pdf,
                max_image_dpi=600,
                slide_indices=[1],
            )

            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0]["media_name"], "slide2.png")
            self.assertEqual(matches[0]["new_size"], (400, 400))

    def test_presentation_order_overrides_noncontiguous_slide_filenames(self):
        first = Image.new("RGB", (400, 400), "purple")
        second = Image.new("RGB", (800, 800), "purple")
        low_resolution = first.resize((40, 40), Image.Resampling.LANCZOS)

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(low_resolution),
            )
            doc.save(source_pdf)
            doc.close()
            _write_ordered_pptx(
                source_pptx,
                {
                    "first.png": _png_bytes(first),
                    "second.png": _png_bytes(second),
                },
                (
                    ("slide9.xml", "first.png", None, False),
                    ("slide2.xml", "second.png", None, False),
                ),
            )

            candidates = _load_pptx_images(source_pptx)
            self.assertTrue(candidates.order_reliable)
            self.assertEqual(
                {
                    candidate["name"]: candidate["pages"]
                    for candidate in candidates
                },
                {
                    "first.png": frozenset({0}),
                    "second.png": frozenset({1}),
                },
            )

            matches = restore_pptx_images(
                source_pdf,
                source_pptx,
                restored_pdf,
                max_image_dpi=600,
                slide_indices=[0],
            )

            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0]["media_name"], "first.png")
            self.assertEqual(matches[0]["new_size"], (400, 400))

    def test_strict_ooxml_relationship_attributes_preserve_real_order(self):
        strict_rel = "http://purl.oclc.org/ooxml/officeDocument/relationships"
        strict_p = "http://purl.oclc.org/ooxml/presentationml/main"
        strict_a = "http://purl.oclc.org/ooxml/drawingml/main"
        first = _png_bytes(Image.new("RGB", (400, 400), "purple"))
        second = _png_bytes(Image.new("RGB", (500, 500), "orange"))

        with _workspace_temp_files() as directory:
            source_pptx = os.path.join(directory, "source.pptx")
            with zipfile.ZipFile(source_pptx, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("ppt/media/first.png", first)
                archive.writestr("ppt/media/second.png", second)
                archive.writestr(
                    "ppt/presentation.xml",
                    f'<p:presentation xmlns:p="{strict_p}" xmlns:r="{strict_rel}">'
                    '<p:sldIdLst><p:sldId id="256" r:id="rIdFirst"/>'
                    '<p:sldId id="257" r:id="rIdSecond"/>'
                    '</p:sldIdLst></p:presentation>',
                )
                archive.writestr(
                    "ppt/_rels/presentation.xml.rels",
                    _relationships_xml((
                        ("rIdFirst", f"{strict_rel}/slide", "slides/slide9.xml"),
                        ("rIdSecond", f"{strict_rel}/slide", "slides/slide2.xml"),
                    )),
                )
                for slide_member, relationship_id, media_name, hidden in (
                    ("slide9.xml", "rIdImage1", "first.png", False),
                    ("slide2.xml", "rIdImage2", "second.png", True),
                ):
                    show = ' show="0"' if hidden else ""
                    archive.writestr(
                        f"ppt/slides/{slide_member}",
                        f'<p:sld xmlns:p="{strict_p}" xmlns:a="{strict_a}" '
                        f'xmlns:r="{strict_rel}"{show}><p:cSld><p:spTree>'
                        '<p:pic><p:blipFill>'
                        f'<a:blip r:embed="{relationship_id}"/>'
                        '</p:blipFill></p:pic></p:spTree></p:cSld></p:sld>',
                    )
                    archive.writestr(
                        f"ppt/slides/_rels/{slide_member}.rels",
                        _relationships_xml((
                            (
                                relationship_id,
                                f"{strict_rel}/image",
                                f"../media/{media_name}",
                            ),
                        )),
                    )

            candidates = _load_pptx_images(source_pptx)

            self.assertTrue(candidates.order_reliable)
            self.assertEqual(candidates.visible_slide_indices, (0,))
            self.assertEqual(pptx_visible_slide_indices(source_pptx), (0,))
            self.assertEqual(
                {
                    candidate["name"]: candidate["pages"]
                    for candidate in candidates
                },
                {
                    "first.png": frozenset({0}),
                    "second.png": frozenset({1}),
                },
            )

    def test_hidden_slides_are_excluded_from_visible_pdf_auto_mapping(self):
        visible_first = Image.new("RGB", (300, 300), "indigo")
        hidden = Image.new("RGB", (900, 900), "indigo")
        visible_last = Image.new("RGB", (500, 500), "indigo")

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            first_page = doc.new_page()
            first_page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(visible_first.resize((40, 40))),
            )
            last_page = doc.new_page()
            last_page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(visible_last.resize((41, 41))),
            )
            doc.save(source_pdf)
            doc.close()
            _write_ordered_pptx(
                source_pptx,
                {
                    "visible_first.png": _png_bytes(visible_first),
                    "hidden.png": _png_bytes(hidden),
                    "visible_last.png": _png_bytes(visible_last),
                },
                (
                    ("slide7.xml", "visible_first.png", None, False),
                    ("slide2.xml", "hidden.png", None, True),
                    ("slide11.xml", "visible_last.png", None, False),
                ),
            )

            candidates = _load_pptx_images(source_pptx)
            self.assertEqual(candidates.slide_count, 3)
            self.assertEqual(candidates.visible_slide_indices, (0, 2))
            self.assertEqual(pptx_visible_slide_indices(source_pptx), (0, 2))

            matches = restore_pptx_images(
                source_pdf,
                source_pptx,
                restored_pdf,
                max_image_dpi=600,
            )

            self.assertEqual(len(matches), 2)
            self.assertEqual(
                {
                    match["page_index"]: (
                        match["media_name"],
                        match["new_size"],
                    )
                    for match in matches
                },
                {
                    0: ("visible_first.png", (300, 300)),
                    1: ("visible_last.png", (500, 500)),
                },
            )

    def test_inherited_layout_and_master_images_are_scoped_to_using_slide(self):
        layout_image = _png_bytes(Image.new("RGB", (300, 300), "green"))
        master_image = _png_bytes(Image.new("RGB", (400, 400), "blue"))
        unused_image = _png_bytes(Image.new("RGB", (500, 500), "red"))

        with _workspace_temp_files() as directory:
            source_pptx = os.path.join(directory, "source.pptx")
            with zipfile.ZipFile(
                source_pptx,
                "w",
                zipfile.ZIP_DEFLATED,
            ) as archive:
                archive.writestr("ppt/media/layout.png", layout_image)
                archive.writestr("ppt/media/master.png", master_image)
                archive.writestr("ppt/media/unused.png", unused_image)

                archive.writestr(
                    "ppt/slides/slide1.xml",
                    _picture_part_xml(),
                )
                archive.writestr(
                    "ppt/slides/_rels/slide1.xml.rels",
                    _relationships_xml((
                        (
                            "rIdLayout",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                            "../slideLayouts/slideLayout1.xml",
                        ),
                    )),
                )
                archive.writestr(
                    "ppt/slideLayouts/slideLayout1.xml",
                    _picture_part_xml("rIdLayoutImage"),
                )
                archive.writestr(
                    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                    _relationships_xml((
                        (
                            "rIdLayoutImage",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                            "../media/layout.png",
                        ),
                        (
                            "rIdMaster",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
                            "../slideMasters/slideMaster1.xml",
                        ),
                    )),
                )
                archive.writestr(
                    "ppt/slideMasters/slideMaster1.xml",
                    _picture_part_xml("rIdMasterImage"),
                )
                archive.writestr(
                    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                    _relationships_xml((
                        (
                            "rIdMasterImage",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                            "../media/master.png",
                        ),
                    )),
                )

                archive.writestr(
                    "ppt/slides/slide2.xml",
                    _picture_part_xml(),
                )
                archive.writestr(
                    "ppt/slides/_rels/slide2.xml.rels",
                    _relationships_xml((
                        (
                            "rIdLayout",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                            "../slideLayouts/slideLayout2.xml",
                        ),
                    )),
                )
                archive.writestr(
                    "ppt/slideLayouts/slideLayout2.xml",
                    _picture_part_xml(),
                )

            candidates = _load_pptx_images(source_pptx)

            self.assertEqual(
                {candidate["name"] for candidate in candidates},
                {"layout.png", "master.png"},
            )
            self.assertTrue(
                all(candidate["pages"] == frozenset({0}) for candidate in candidates)
            )

    def test_shared_pdf_xref_requires_one_candidate_covering_all_pages(self):
        large = Image.new("RGB", (800, 800), "teal")
        small = large.resize((400, 400), Image.Resampling.LANCZOS)
        low_resolution = large.resize((40, 40), Image.Resampling.LANCZOS)

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            first_page = doc.new_page()
            xref = first_page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(low_resolution),
            )
            second_page = doc.new_page()
            second_page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                xref=xref,
            )
            doc.save(source_pdf)
            doc.close()
            _write_pptx(
                source_pptx,
                {
                    "slide1.png": _png_bytes(small),
                    "slide2.png": _png_bytes(large),
                },
                slides=(
                    ("slide1.png", None),
                    ("slide2.png", None),
                ),
            )

            matches = restore_pptx_images(
                source_pdf,
                source_pptx,
                restored_pdf,
                max_image_dpi=600,
            )

            self.assertEqual(matches, [])
            self.assertEqual(matches.stats["unmatched"], 1)
            self.assertFalse(os.path.exists(restored_pdf))

    def test_shared_xref_with_multiple_placements_on_one_page_is_not_replaced(self):
        source = Image.new("RGB", (800, 800), "teal")
        low_resolution = source.resize((40, 40), Image.Resampling.LANCZOS)

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            xref = page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(low_resolution),
            )
            page.insert_image(
                pymupdf.Rect(120, 20, 156, 92),
                xref=xref,
            )
            doc.save(source_pdf)
            doc.close()
            _write_pptx(source_pptx, {"source.png": _png_bytes(source)})

            matches = restore_pptx_images(
                source_pdf,
                source_pptx,
                restored_pdf,
            )

            self.assertEqual(matches, [])
            self.assertEqual(matches.stats["unmatched"], 1)
            self.assertFalse(os.path.exists(restored_pdf))

    def test_reports_ambiguous_and_already_sufficient_images(self):
        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            page.insert_image(
                pymupdf.Rect(20, 20, 70, 70),
                stream=_png_bytes(Image.new("RGB", (50, 50), (128, 128, 128))),
            )
            doc.save(source_pdf)
            doc.close()
            _write_pptx(
                source_pptx,
                {
                    "dark.png": _png_bytes(
                        Image.new("RGB", (500, 500), (110, 110, 110))
                    ),
                    "light.png": _png_bytes(
                        Image.new("RGB", (500, 500), (146, 146, 146))
                    ),
                },
            )

            ambiguous = restore_pptx_images(
                source_pdf,
                source_pptx,
                restored_pdf,
            )

            self.assertEqual(ambiguous, [])
            self.assertEqual(ambiguous.stats["ambiguous"], 1)
            self.assertEqual(
                sum(
                    ambiguous.stats[key]
                    for key in (
                        "restored",
                        "unmatched",
                        "ambiguous",
                        "sufficient",
                    )
                ),
                ambiguous.stats["total"],
            )

            _write_pptx(
                source_pptx,
                {
                    "same.png": _png_bytes(
                        Image.new("RGB", (50, 50), (128, 128, 128))
                    )
                },
            )
            sufficient = restore_pptx_images(
                source_pdf,
                source_pptx,
                restored_pdf,
            )

            self.assertEqual(sufficient, [])
            self.assertEqual(sufficient.stats["sufficient"], 1)
            self.assertEqual(sufficient.stats["total"], 1)

    def test_restores_original_pixels_without_rasterizing_vector_content(self):
        source = Image.new("RGBA", (600, 600), (0, 0, 0, 0))
        draw = ImageDraw.Draw(source)
        draw.ellipse((140, 80, 460, 520), fill=(30, 120, 230, 255))
        draw.rectangle((275, 160, 325, 440), fill=(255, 255, 255, 255))
        draw.rectangle((205, 255, 395, 305), fill=(255, 255, 255, 255))

        alpha_bbox = source.getchannel("A").getbbox()
        low_resolution = source.crop(alpha_bbox).resize(
            (48, 66), Image.Resampling.LANCZOS
        )

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page(width=300, height=200)
            page.draw_rect(pymupdf.Rect(20, 20, 280, 180), color=(0, 0, 1))
            page.insert_text((30, 40), "Vector text")
            page.insert_image(
                pymupdf.Rect(80, 60, 128, 126),
                stream=_png_bytes(low_resolution),
            )
            doc.save(source_pdf)
            doc.close()

            with zipfile.ZipFile(source_pptx, "w") as archive:
                archive.writestr("ppt/media/image1.png", _png_bytes(source))

            matches = restore_pptx_images(source_pdf, source_pptx, restored_pdf)

            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0]["media_name"], "image1.png")
            self.assertGreater(
                matches[0]["new_size"][0] * matches[0]["new_size"][1],
                matches[0]["old_size"][0] * matches[0]["old_size"][1] * 10,
            )

            before = pymupdf.open(source_pdf)
            after = pymupdf.open(restored_pdf)
            try:
                self.assertEqual(
                    len(before[0].get_drawings()),
                    len(after[0].get_drawings()),
                )
                self.assertEqual(
                    before[0].get_text("words"),
                    after[0].get_text("words"),
                )
                before_bbox = before[0].get_image_info()[0]["bbox"]
                after_bbox = after[0].get_image_info()[0]["bbox"]
                self.assertEqual(before_bbox, after_bbox)
            finally:
                before.close()
                after.close()

    def test_caps_restored_image_at_300_dpi_for_its_display_size(self):
        source = Image.new("RGB", (2400, 2400), (220, 40, 30))
        low_resolution = source.resize((60, 60), Image.Resampling.LANCZOS)

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            page.insert_image(
                pymupdf.Rect(20, 20, 92, 92),
                stream=_png_bytes(low_resolution),
            )
            doc.save(source_pdf)
            doc.close()

            with zipfile.ZipFile(source_pptx, "w") as archive:
                archive.writestr("ppt/media/image1.png", _png_bytes(source))

            matches = restore_pptx_images(source_pdf, source_pptx, restored_pdf)

            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0]["old_size"], (60, 60))
            self.assertEqual(matches[0]["new_size"], (300, 300))

            restored = pymupdf.open(restored_pdf)
            try:
                displayed = restored[0].get_image_info()[0]["bbox"]
                self.assertEqual(displayed, (20.0, 20.0, 92.0, 92.0))
            finally:
                restored.close()

    def test_skips_an_unrelated_pptx_image(self):
        source = Image.new("RGB", (500, 500), (255, 0, 0))
        low_resolution = Image.new("RGB", (50, 50), (0, 0, 255))

        with _workspace_temp_files() as directory:
            source_pdf = os.path.join(directory, "powerpoint.pdf")
            source_pptx = os.path.join(directory, "source.pptx")
            restored_pdf = os.path.join(directory, "restored.pdf")

            doc = pymupdf.open()
            page = doc.new_page()
            page.insert_image(
                pymupdf.Rect(20, 20, 70, 70),
                stream=_png_bytes(low_resolution),
            )
            doc.save(source_pdf)
            doc.close()

            with zipfile.ZipFile(source_pptx, "w") as archive:
                archive.writestr("ppt/media/image1.png", _png_bytes(source))

            matches = restore_pptx_images(source_pdf, source_pptx, restored_pdf)

            self.assertEqual(matches, [])
            self.assertEqual(matches.stats["unmatched"], 1)
            self.assertFalse(os.path.exists(restored_pdf))


if __name__ == "__main__":
    unittest.main()
