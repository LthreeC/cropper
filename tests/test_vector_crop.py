from contextlib import contextmanager
import io
import os
from pathlib import Path
import unittest
from unittest.mock import patch

import pymupdf
from PIL import Image

from processor import CropProcessor


@contextmanager
def _temporary_pdf_directory():
    root = Path(__file__).resolve().parent.parent / "tmp" / "pdfs"
    root.mkdir(parents=True, exist_ok=True)
    names = [
        "source.pdf",
        "cropped.pdf",
        *[
            f"{case}_{suffix}.pdf"
            for case in ("normal", "offset", "rotated", "offset_rotated")
            for suffix in ("source", "cropped")
        ],
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
    if cropbox:
        page.set_cropbox(pymupdf.Rect(*cropbox))
    if rotation:
        page.set_rotation(rotation)
    doc.save(path)
    doc.close()


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
                self.assertLess(after_page.rect.width, before_rect.width)
                self.assertLess(after_page.rect.height, before_rect.height)
            finally:
                after.close()

    def test_crop_handles_rotation_and_nonzero_cropbox_without_clipping(self):
        cases = [
            ("normal", None, 0),
            ("offset", (40, 30, 360, 270), 0),
            ("rotated", None, 90),
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
                        self.assertLessEqual(page.cropbox.x0, 100)
                        self.assertLessEqual(page.cropbox.y0, 80)
                        self.assertGreaterEqual(page.cropbox.x1, 300)
                        self.assertGreaterEqual(page.cropbox.y1, 220)
                    finally:
                        doc.close()

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
