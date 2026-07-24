from contextlib import contextmanager
import io
import os
import unittest
import zipfile

import pymupdf
from PIL import Image, ImageDraw

from ppt_image_restore import restore_pptx_images


def _png_bytes(image):
    buffer = io.BytesIO()
    image.save(buffer, "PNG")
    return buffer.getvalue()


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
            self.assertFalse(os.path.exists(restored_pdf))


if __name__ == "__main__":
    unittest.main()
