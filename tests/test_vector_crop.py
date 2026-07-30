from contextlib import contextmanager
import io
import os
from pathlib import Path
import unittest
from unittest.mock import patch

import numpy as np
import pymupdf
from PIL import Image, ImageDraw

from processor import (
    CropProcessor,
    RASTER_PADDING_COMPENSATION_POINTS,
    RASTER_STRICT_INSET_POINTS,
    VECTOR_PADDING_DPI,
)
from units import pixels_to_points


@contextmanager
def _temporary_pdf_directory():
    root = Path(__file__).resolve().parent.parent / "tmp" / "pdfs"
    root.mkdir(parents=True, exist_ok=True)
    names = [
        "source.pdf",
        "cropped.pdf",
        *[
            f"{case}_{suffix}.pdf"
            for case in (
                "normal",
                "offset",
                "rotated",
                "rotated_180",
                "rotated_270",
                "offset_rotated",
            )
            for suffix in ("source", "cropped")
        ],
        *[f"mode_{mode}.pdf" for mode in ("simple", "smart", "edge")],
        "padding_source.pdf",
        "padding_cropped.pdf",
        "idempotent_first.pdf",
        "idempotent_second.pdf",
        "multipage_source.pdf",
        "multipage_cropped.pdf",
        "interior_vector_source.pdf",
        "interior_image_source.pdf",
        "interior_vector_cropped.pdf",
        "interior_image_cropped.pdf",
        "raster_edge_source.pdf",
        "raster_edge_p0.pdf",
        "raster_edge_p2.pdf",
        "raster_edge_repeat.pdf",
    ]
    paths = [root / name for name in names]
    for path in paths:
        if path.exists():
            path.unlink()
    try:
        yield root
    finally:
        for path in paths:
            if path.exists():
                path.unlink()


def _png_bytes():
    image = Image.new("RGB", (64, 64), (210, 40, 30))
    buffer = io.BytesIO()
    image.save(buffer, "PNG")
    return buffer.getvalue()


def _create_pdf(path, cropbox=None, rotation=0, include_image=False):
    doc = pymupdf.open()
    page = doc.new_page(width=400, height=300)
    page.draw_rect(
        pymupdf.Rect(100, 80, 300, 220),
        color=(0, 0, 0),
        fill=(0.8, 0.8, 0.8),
    )
    page.insert_text((140, 150), "Vector text", fontsize=16)
    if include_image:
        page.insert_image(
            pymupdf.Rect(170, 110, 230, 170),
            stream=_png_bytes(),
        )
    page.insert_link({
        "kind": pymupdf.LINK_URI,
        "from": pymupdf.Rect(140, 130, 230, 155),
        "uri": "https://example.com",
    })
    if cropbox:
        page.set_cropbox(pymupdf.Rect(*cropbox))
    if rotation:
        page.set_rotation(rotation)
    doc.save(path)
    doc.close()


def _create_raster_page_pdf(path):
    image = Image.new("RGB", (400, 300), "white")
    ImageDraw.Draw(image).rectangle((100, 80, 299, 219), fill="black")
    buffer = io.BytesIO()
    image.save(buffer, "PNG")

    document = pymupdf.open()
    page = document.new_page(width=400, height=300)
    page.insert_image(page.rect, stream=buffer.getvalue())
    document.save(path)
    document.close()


def _render_margins(page, dpi=300):
    pixmap = page.get_pixmap(dpi=dpi, alpha=False)
    pixels = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
        pixmap.height,
        pixmap.width,
        pixmap.n,
    )[:, :, :3]
    mask = pixels.mean(axis=2) < 250
    ys, xs = np.where(mask)
    return (
        int(xs.min()),
        int(ys.min()),
        int(pixmap.width - xs.max() - 1),
        int(pixmap.height - ys.max() - 1),
    )


class VectorCropTests(unittest.TestCase):
    def test_crop_preserves_vectors_text_and_image_pixels(self):
        with _temporary_pdf_directory() as directory:
            source = directory / "source.pdf"
            target = directory / "cropped.pdf"
            _create_pdf(source, include_image=True)

            before = pymupdf.open(source)
            try:
                before_page = before[0]
                before_drawings = len(before_page.get_drawings())
                before_words = [word[4] for word in before_page.get_text("words")]
                before_images = [
                    (info["width"], info["height"], info["digest"])
                    for info in before_page.get_image_info(hashes=True)
                ]
                before_links = [
                    (link["kind"], link.get("uri"))
                    for link in before_page.get_links()
                ]
                before_rect = before_page.rect
            finally:
                before.close()

            CropProcessor()._process_vector_crop(
                str(source), "PDF", str(target),
                2, 250, 15, "smart", "source"
            )

            after = pymupdf.open(target)
            try:
                after_page = after[0]
                self.assertEqual(len(after_page.get_drawings()), before_drawings)
                self.assertEqual(
                    [word[4] for word in after_page.get_text("words")],
                    before_words,
                )
                self.assertEqual(
                    [
                        (info["width"], info["height"], info["digest"])
                        for info in after_page.get_image_info(hashes=True)
                    ],
                    before_images,
                )
                self.assertEqual(
                    [
                        (link["kind"], link.get("uri"))
                        for link in after_page.get_links()
                    ],
                    before_links,
                )
                self.assertLess(after_page.rect.width, before_rect.width)
                self.assertLess(after_page.rect.height, before_rect.height)
                self.assertEqual(
                    after.xref_get_key(after_page.xref, "CropBox"),
                    ("null", "null"),
                )
            finally:
                after.close()

    def test_crop_handles_rotation_and_nonzero_cropbox_without_clipping(self):
        cases = [
            ("normal", None, 0),
            ("offset", (40, 30, 360, 270), 0),
            ("rotated", None, 90),
            ("rotated_180", None, 180),
            ("rotated_270", None, 270),
            ("offset_rotated", (40, 30, 360, 270), 90),
        ]

        with _temporary_pdf_directory() as directory:
            for name, cropbox, rotation in cases:
                with self.subTest(name=name):
                    source = directory / f"{name}_source.pdf"
                    target = directory / f"{name}_cropped.pdf"
                    _create_pdf(source, cropbox=cropbox, rotation=rotation)

                    CropProcessor()._process_vector_crop(
                        str(source), "PDF", str(target),
                        2, 250, 15, "smart", name
                    )

                    doc = pymupdf.open(target)
                    try:
                        page = doc[0]
                        self.assertEqual(page.rotation, rotation)
                        self.assertEqual(len(page.get_drawings()), 1)
                        self.assertEqual(
                            [word[4] for word in page.get_text("words")],
                            ["Vector", "text"],
                        )
                        self.assertEqual(
                            doc.xref_get_key(page.xref, "CropBox"),
                            ("null", "null"),
                        )
                        self.assertTrue(all(
                            0 <= margin <= 3
                            for margin in _render_margins(page)
                        ))
                    finally:
                        doc.close()

    def test_zero_padding_removes_all_rendered_border_in_every_mode(self):
        with _temporary_pdf_directory() as directory:
            source = directory / "source.pdf"
            _create_pdf(source)

            for mode in ("simple", "smart", "edge"):
                with self.subTest(mode=mode):
                    target = directory / f"mode_{mode}.pdf"
                    CropProcessor()._process_vector_crop(
                        str(source),
                        "PDF",
                        str(target),
                        0,
                        250,
                        15,
                        mode,
                        mode,
                    )

                    document = pymupdf.open(target)
                    try:
                        self.assertEqual(
                            _render_margins(document[0]),
                            (0, 0, 0, 0),
                        )
                    finally:
                        document.close()

    def test_vector_padding_uses_300_dpi_pixel_units(self):
        with _temporary_pdf_directory() as directory:
            source = directory / "padding_source.pdf"
            target = directory / "padding_cropped.pdf"
            _create_pdf(source)

            CropProcessor()._process_vector_crop(
                str(source),
                "PDF",
                str(target),
                2,
                250,
                15,
                "simple",
                "padding",
            )

            document = pymupdf.open(target)
            try:
                margins = _render_margins(document[0])
                self.assertTrue(all(1 <= margin <= 3 for margin in margins))
            finally:
                document.close()

    def test_interior_image_does_not_trigger_raster_edge_compensation(self):
        with _temporary_pdf_directory() as directory:
            vector_source = directory / "interior_vector_source.pdf"
            image_source = directory / "interior_image_source.pdf"
            vector_target = directory / "interior_vector_cropped.pdf"
            image_target = directory / "interior_image_cropped.pdf"
            _create_pdf(vector_source)
            _create_pdf(image_source, include_image=True)

            for source, target in (
                (vector_source, vector_target),
                (image_source, image_target),
            ):
                CropProcessor()._process_vector_crop(
                    str(source),
                    "PDF",
                    str(target),
                    0,
                    250,
                    15,
                    "smart",
                    target.stem,
                )

            vector_document = pymupdf.open(vector_target)
            image_document = pymupdf.open(image_target)
            try:
                self.assertEqual(
                    tuple(vector_document[0].mediabox),
                    tuple(image_document[0].mediabox),
                )
            finally:
                vector_document.close()
                image_document.close()

    def test_raster_edge_gets_cross_renderer_positive_padding(self):
        with _temporary_pdf_directory() as directory:
            source = directory / "raster_edge_source.pdf"
            zero_target = directory / "raster_edge_p0.pdf"
            padded_target = directory / "raster_edge_p2.pdf"
            repeat_target = directory / "raster_edge_repeat.pdf"
            _create_raster_page_pdf(source)

            for padding, target in ((0, zero_target), (2, padded_target)):
                CropProcessor()._process_vector_crop(
                    str(source),
                    "PDF",
                    str(target),
                    padding,
                    250,
                    15,
                    "smart",
                    target.stem,
                )

            CropProcessor()._process_vector_crop(
                str(padded_target),
                "PDF",
                str(repeat_target),
                2,
                250,
                15,
                "smart",
                repeat_target.stem,
            )

            zero_document = pymupdf.open(zero_target)
            padded_document = pymupdf.open(padded_target)
            repeat_document = pymupdf.open(repeat_target)
            try:
                expected_growth = 2 * (
                    RASTER_STRICT_INSET_POINTS
                    + pixels_to_points(2, VECTOR_PADDING_DPI)
                    + RASTER_PADDING_COMPENSATION_POINTS
                )
                self.assertAlmostEqual(
                    padded_document[0].rect.width
                    - zero_document[0].rect.width,
                    expected_growth,
                    places=2,
                )
                self.assertAlmostEqual(
                    padded_document[0].rect.height
                    - zero_document[0].rect.height,
                    expected_growth,
                    places=2,
                )
                self.assertEqual(
                    tuple(padded_document[0].mediabox),
                    tuple(repeat_document[0].mediabox),
                )
            finally:
                zero_document.close()
                padded_document.close()
                repeat_document.close()

    def test_negative_vector_padding_is_clamped_to_zero(self):
        with _temporary_pdf_directory() as directory:
            source = directory / "padding_source.pdf"
            target = directory / "padding_cropped.pdf"
            _create_pdf(source)

            CropProcessor()._process_vector_crop(
                str(source),
                "PDF",
                str(target),
                -5,
                250,
                15,
                "simple",
                "padding",
            )

            document = pymupdf.open(target)
            try:
                self.assertEqual(_render_margins(document[0]), (0, 0, 0, 0))
            finally:
                document.close()

    def test_vector_crop_is_idempotent_for_zero_and_default_padding(self):
        with _temporary_pdf_directory() as directory:
            source = directory / "source.pdf"
            _create_pdf(source)

            for padding in (0, 2):
                with self.subTest(padding=padding):
                    first = directory / "idempotent_first.pdf"
                    second = directory / "idempotent_second.pdf"
                    CropProcessor()._process_vector_crop(
                        str(source),
                        "PDF",
                        str(first),
                        padding,
                        250,
                        15,
                        "smart",
                        "idempotent_first",
                    )
                    CropProcessor()._process_vector_crop(
                        str(first),
                        "PDF",
                        str(second),
                        padding,
                        250,
                        15,
                        "smart",
                        "idempotent_second",
                    )

                    first_document = pymupdf.open(first)
                    second_document = pymupdf.open(second)
                    try:
                        self.assertEqual(
                            tuple(first_document[0].mediabox),
                            tuple(second_document[0].mediabox),
                        )
                        self.assertEqual(
                            _render_margins(first_document[0]),
                            _render_margins(second_document[0]),
                        )
                        self.assertEqual(
                            first_document[0].get_text("words"),
                            second_document[0].get_text("words"),
                        )
                        self.assertEqual(
                            len(first_document[0].get_drawings()),
                            len(second_document[0].get_drawings()),
                        )
                    finally:
                        first_document.close()
                        second_document.close()

    def test_multi_page_crop_keeps_every_page_and_rotation(self):
        with _temporary_pdf_directory() as directory:
            source = directory / "multipage_source.pdf"
            target = directory / "multipage_cropped.pdf"
            document = pymupdf.open()
            for rotation in (0, 90):
                page = document.new_page(width=400, height=300)
                page.draw_rect(
                    pymupdf.Rect(100, 80, 300, 220),
                    color=(0, 0, 0),
                    fill=(0.8, 0.8, 0.8),
                )
                page.insert_text((140, 150), "Vector text", fontsize=16)
                page.set_rotation(rotation)
            document.save(source)
            document.close()

            CropProcessor()._process_vector_crop(
                str(source),
                "PDF",
                str(target),
                0,
                250,
                15,
                "smart",
                "multipage",
            )

            cropped = pymupdf.open(target)
            try:
                self.assertEqual(len(cropped), 2)
                self.assertEqual([page.rotation for page in cropped], [0, 90])
                for page in cropped:
                    self.assertEqual(
                        [word[4] for word in page.get_text("words")],
                        ["Vector", "text"],
                    )
                    self.assertEqual(_render_margins(page), (0, 0, 0, 0))
                    self.assertEqual(
                        cropped.xref_get_key(page.xref, "CropBox"),
                        ("null", "null"),
                    )
            finally:
                cropped.close()

    def test_single_page_vector_crop_skips_intermediate_pdf_copy(self):
        processor = CropProcessor()
        pdf_path = os.path.abspath("source.pdf")
        output_dir = os.path.abspath("output")
        config = {
            "source_files": [pdf_path],
            "scope": "CURRENT",
            "output_format": "PDF",
            "output_dir": output_dir,
            "page_num": 2,
            "padding": 2,
            "threshold": 250,
            "sensitivity": 15,
            "detect_mode": "smart",
            "dpi": 300,
        }

        with (
            patch(
                "controllers.FileController.get_pdf_page_count",
                return_value=3,
            ),
            patch.object(
                processor,
                "_process_vector_crop",
                return_value="done",
            ) as crop,
            patch.object(processor, "_make_temp_path") as make_temp,
        ):
            result = processor._process_pdf(config, pdf_path)

        self.assertEqual(result, "done")
        make_temp.assert_not_called()
        crop.assert_called_once_with(
            pdf_path,
            "PDF",
            os.path.join(output_dir, "source_p2.pdf"),
            2,
            250,
            15,
            "smart",
            "source",
            page_indices=[1],
        )


if __name__ == "__main__":
    unittest.main()
