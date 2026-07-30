from contextlib import contextmanager
import os
from pathlib import Path
import unittest

import pymupdf
from PIL import Image, ImageDraw

from processor import CropProcessor


@contextmanager
def _format_workspace():
    directory = Path(__file__).resolve().parent.parent / "tmp" / "formats"
    directory.mkdir(parents=True, exist_ok=True)
    names = [
        "source.png",
        "source_cropped.pdf",
        "source_cropped.png",
        "export.pdf",
        "export.png",
        "export.jpeg",
        "export.tiff",
        "export.webp",
        "vector_source.pdf",
        "vector_output.svg",
        "blank.pdf",
        "blank_p1.png",
        "converted.jpeg",
        "converted.png",
        "duplicate.png",
        "duplicate.jpg",
        "duplicate_png_cropped.png",
        "duplicate_jpg_cropped.png",
    ]
    paths = [directory / name for name in names]
    for path in paths:
        if path.exists():
            path.unlink()
    try:
        yield directory
    finally:
        for path in paths:
            if path.exists():
                path.unlink()


class FormatExportTests(unittest.TestCase):
    def test_raster_formats_keep_pixel_dimensions(self):
        processor = CropProcessor()
        image = Image.new("RGB", (120, 80), (30, 100, 200))

        with _format_workspace() as directory:
            for fmt, suffix in (
                ("PNG", "png"),
                ("JPEG", "jpeg"),
                ("TIFF", "tiff"),
                ("WebP", "webp"),
            ):
                with self.subTest(fmt=fmt):
                    target = directory / f"export.{suffix}"
                    processor._save_image(image.copy(), target, fmt, 300)
                    with Image.open(target) as saved:
                        self.assertEqual(saved.size, image.size)

    def test_png_and_jpeg_accept_supported_input_color_modes(self):
        processor = CropProcessor()

        with _format_workspace() as directory:
            cmyk = Image.new("CMYK", (40, 30), (20, 80, 120, 10))
            png_target = directory / "converted.png"
            processor._save_image(cmyk, png_target, "PNG", 300)
            with Image.open(png_target) as saved:
                self.assertEqual(saved.mode, "RGB")

            palette = Image.new("P", (40, 30), 0)
            palette.putpalette(
                [255, 255, 255, 20, 80, 120] + [0, 0, 0] * 254
            )
            palette.info["transparency"] = 0
            jpeg_target = directory / "converted.jpeg"
            processor._save_image(palette, jpeg_target, "JPEG", 300)
            with Image.open(jpeg_target) as saved:
                self.assertEqual(saved.mode, "RGB")

    def test_batch_inputs_with_the_same_stem_do_not_overwrite_each_other(self):
        with _format_workspace() as directory:
            for suffix in ("png", "jpg"):
                source = Image.new("RGB", (80, 60), "white")
                ImageDraw.Draw(source).rectangle(
                    (20, 15, 60, 45),
                    fill="black",
                )
                source.save(directory / f"duplicate.{suffix}")

            config = {
                "output_format": "PNG",
                "output_dir": str(directory),
                "padding": 0,
                "threshold": 250,
                "sensitivity": 15,
                "detect_mode": "smart",
                "dpi": 300,
            }
            CropProcessor()._process_images(
                config,
                [
                    str(directory / "duplicate.png"),
                    str(directory / "duplicate.jpg"),
                ],
            )

            self.assertTrue((directory / "duplicate_png_cropped.png").exists())
            self.assertTrue((directory / "duplicate_jpg_cropped.png").exists())

    def test_raster_pdf_embeds_pixels_without_resampling(self):
        processor = CropProcessor()
        image = Image.new("RGBA", (120, 80), (30, 100, 200, 180))

        with _format_workspace() as directory:
            target = directory / "export.pdf"
            processor._save_image(image, target, "PDF", 300)

            doc = pymupdf.open(target)
            try:
                page = doc[0]
                info = page.get_image_info()[0]
                self.assertEqual((info["width"], info["height"]), image.size)
                self.assertAlmostEqual(page.rect.width, 28.8, places=2)
                self.assertAlmostEqual(page.rect.height, 19.2, places=2)
            finally:
                doc.close()

    def test_image_to_pdf_uses_pdf_extension_and_keeps_cropped_pixels(self):
        source = Image.new("RGB", (200, 150), "white")
        ImageDraw.Draw(source).rectangle((40, 30, 160, 120), fill="black")

        with _format_workspace() as directory:
            source_path = directory / "source.png"
            source.save(source_path, dpi=(300, 300))
            config = {
                "output_format": "PDF",
                "output_dir": str(directory),
                "padding": 2,
                "threshold": 250,
                "sensitivity": 15,
                "detect_mode": "smart",
                "dpi": 300,
            }

            result = CropProcessor()._process_images(config, [str(source_path)])

            target = directory / "source_cropped.pdf"
            self.assertEqual(result, str(directory))
            self.assertTrue(target.exists())
            self.assertFalse((directory / "source_cropped.png").exists())
            doc = pymupdf.open(target)
            try:
                info = doc[0].get_image_info()[0]
                self.assertGreater(info["width"], 0)
                self.assertGreater(info["height"], 0)
            finally:
                doc.close()

    def test_pdf_and_other_files_are_not_silently_mixed(self):
        logs = []
        processor = CropProcessor(
            callback=lambda status, log_entry, progress: logs.append(log_entry)
            if log_entry else None
        )
        config = {
            "source_files": [
                os.path.abspath("source.png"),
                os.path.abspath("source.pdf"),
            ]
        }

        result = processor.process_file(config)

        self.assertIsNone(result)
        self.assertTrue(
            any(
                level == "ERROR" and "PDF 文件需单独处理" in message
                for level, message in logs
            )
        )

    def test_blank_pdf_page_does_not_report_a_missing_output_as_success(self):
        with _format_workspace() as directory:
            source = directory / "blank.pdf"
            document = pymupdf.open()
            document.new_page(width=200, height=150)
            document.save(source)
            document.close()
            logs = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: logs.append(
                    log_entry
                ) if log_entry else None
            )

            result = processor._process_pdf(
                {
                    "output_format": "PNG",
                    "output_dir": str(directory),
                    "scope": "CURRENT",
                    "page_num": 1,
                    "padding": 0,
                    "threshold": 250,
                    "sensitivity": 15,
                    "detect_mode": "smart",
                    "dpi": 300,
                },
                str(source),
            )

            self.assertIsNone(result)
            self.assertFalse((directory / "blank_p1.png").exists())
            self.assertTrue(any(
                level == "WARNING" and "未检测到可输出内容" in message
                for level, message in logs
            ))

    def test_pdf_to_svg_remains_vector(self):
        with _format_workspace() as directory:
            source = directory / "vector_source.pdf"
            target = directory / "vector_output.svg"
            doc = pymupdf.open()
            page = doc.new_page(width=300, height=200)
            page.draw_rect(
                pymupdf.Rect(60, 40, 240, 160),
                color=(0, 0, 0),
                fill=(0.2, 0.5, 0.8),
            )
            page.insert_text((90, 105), "Vector text")
            doc.save(source)
            doc.close()

            CropProcessor()._process_vector_crop(
                str(source),
                "SVG",
                str(target),
                2,
                250,
                15,
                "smart",
                "vector",
            )

            svg = target.read_text(encoding="utf-8")
            self.assertIn("<svg", svg)
            self.assertIn("<path", svg)


if __name__ == "__main__":
    unittest.main()
