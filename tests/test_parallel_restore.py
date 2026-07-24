from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
import io
from pathlib import Path
import unittest
from unittest.mock import patch
import zipfile

import pymupdf
from PIL import Image, ImageDraw

from ppt_image_restore import (
    _image_worker_count,
    restore_pptx_images_in_document,
)
from processor import CropProcessor


@contextmanager
def _parallel_workspace():
    directory = (
        Path(__file__).resolve().parent.parent
        / "tmp"
        / "pdfs"
        / "parallel_restore"
    )
    directory.mkdir(parents=True, exist_ok=True)
    paths = [
        directory / name
        for name in (
            "source.pdf",
            "source.pptx",
            "cropped.pdf",
        )
    ]
    for path in paths:
        if path.exists():
            path.unlink()
    try:
        yield paths
    finally:
        for path in paths:
            if path.exists():
                path.unlink()


def _png_bytes(image):
    buffer = io.BytesIO()
    image.save(buffer, "PNG")
    return buffer.getvalue()


def _source_image(index):
    colors = (
        (210, 40, 30),
        (40, 180, 70),
        (40, 90, 220),
    )
    image = Image.new("RGB", (400, 400), colors[index])
    draw = ImageDraw.Draw(image)
    draw.rectangle(
        (40 + index * 20, 60, 340, 330 - index * 20),
        outline="white",
        width=20,
    )
    return image


class ParallelRestoreTests(unittest.TestCase):
    def test_worker_count_is_bounded(self):
        with patch("ppt_image_restore.os.cpu_count", return_value=8):
            self.assertEqual(_image_worker_count(1), 1)
            self.assertEqual(_image_worker_count(2), 2)
            self.assertEqual(_image_worker_count(100), 3)
        with patch("ppt_image_restore.os.cpu_count", return_value=1):
            self.assertEqual(_image_worker_count(100), 1)

    def test_parallel_pillow_work_preserves_open_pdf_document(self):
        with _parallel_workspace() as paths:
            source_pdf, source_pptx, cropped_pdf = paths
            sources = [_source_image(index) for index in range(3)]

            doc = pymupdf.open()
            page = doc.new_page(width=320, height=160)
            page.insert_text((20, 20), "Vector text")
            page.draw_rect(
                pymupdf.Rect(15, 30, 305, 145),
                color=(0, 0, 0),
            )
            for index, source in enumerate(sources):
                low_resolution = source.resize(
                    (40, 40),
                    Image.Resampling.LANCZOS,
                )
                page.insert_image(
                    pymupdf.Rect(
                        25 + index * 95,
                        45,
                        97 + index * 95,
                        117,
                    ),
                    stream=_png_bytes(low_resolution),
                )
            doc.save(source_pdf)
            doc.close()

            with zipfile.ZipFile(
                source_pptx,
                "w",
                zipfile.ZIP_DEFLATED,
            ) as archive:
                for index, source in enumerate(sources):
                    archive.writestr(
                        f"ppt/media/image{index + 1}.png",
                        _png_bytes(source),
                    )

            doc = pymupdf.open(source_pdf)
            try:
                with patch(
                    "ppt_image_restore.ThreadPoolExecutor",
                    wraps=ThreadPoolExecutor,
                ) as executor:
                    restored = restore_pptx_images_in_document(
                        doc,
                        str(source_pptx),
                        max_image_dpi=300,
                    )

                self.assertEqual(len(restored), 3)
                self.assertGreaterEqual(executor.call_count, 2)
                self.assertTrue(
                    all(
                        call.kwargs["max_workers"] == 3
                        for call in executor.call_args_list
                    )
                )
                self.assertFalse(doc.is_closed)
                self.assertEqual(
                    [word[4] for word in doc[0].get_text("words")],
                    ["Vector", "text"],
                )
                self.assertEqual(len(doc[0].get_drawings()), 1)

                CropProcessor()._process_vector_crop(
                    doc,
                    "PDF",
                    str(cropped_pdf),
                    2,
                    250,
                    15,
                    "smart",
                    "parallel",
                    pdf_garbage=4,
                )
                self.assertFalse(doc.is_closed)
            finally:
                doc.close()

            cropped = pymupdf.open(cropped_pdf)
            try:
                sizes = {
                    (info["width"], info["height"])
                    for info in cropped[0].get_image_info()
                }
                self.assertEqual(sizes, {(300, 300)})
                self.assertEqual(
                    [word[4] for word in cropped[0].get_text("words")],
                    ["Vector", "text"],
                )
                self.assertEqual(len(cropped[0].get_drawings()), 1)
            finally:
                cropped.close()


if __name__ == "__main__":
    unittest.main()
