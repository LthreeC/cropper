import math
import os
from pathlib import Path
import tempfile
import unittest

import pymupdf
from PIL import Image, ImageChops, ImageDraw

from processor import (
    CropProcessor,
    MAX_OUTPUT_DPI,
    validate_dpi,
    validate_padding,
    validate_page_number,
)


def _config(output_dir, output_format="PNG", scope="CURRENT"):
    return {
        "output_format": output_format,
        "output_dir": str(output_dir),
        "scope": scope,
        "page_num": 1,
        "padding": 0,
        "threshold": 250,
        "sensitivity": 15,
        "detect_mode": "smart",
        "dpi": 300,
    }


def _content_image(size=(80, 60), box=(20, 15, 59, 44)):
    image = Image.new("RGB", size, "white")
    ImageDraw.Draw(image).rectangle(box, fill="black")
    return image


def _webp_durations(path):
    durations = []
    with open(path, "rb") as stream:
        riff_header = stream.read(12)
        if riff_header[:4] != b"RIFF" or riff_header[8:] != b"WEBP":
            return durations
        while True:
            header = stream.read(8)
            if len(header) != 8:
                break
            chunk_size = int.from_bytes(header[4:], "little")
            payload_start = stream.tell()
            if header[:4] == b"ANMF" and chunk_size >= 16:
                frame_header = stream.read(16)
                durations.append(int.from_bytes(
                    frame_header[12:15],
                    "little",
                ))
            stream.seek(payload_start + chunk_size + (chunk_size & 1))
    return durations


class InputValidationTests(unittest.TestCase):
    def test_padding_clamps_negative_but_rejects_non_finite_values(self):
        self.assertEqual(validate_padding(-2), 0)
        self.assertEqual(validate_padding("2.5"), 2.5)
        for value in (math.nan, math.inf, -math.inf, "invalid"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                validate_padding(value)

    def test_dpi_and_page_number_require_sensible_positive_integers(self):
        self.assertEqual(validate_dpi(300), 300)
        self.assertEqual(validate_dpi(MAX_OUTPUT_DPI), MAX_OUTPUT_DPI)
        for value in (0, -1, 300.5, math.nan, MAX_OUTPUT_DPI + 1, "invalid"):
            with self.subTest(dpi=value), self.assertRaises(ValueError):
                validate_dpi(value)
        for value in (0, -1, 1.5, math.nan, "invalid"):
            with self.subTest(page=value), self.assertRaises(ValueError):
                validate_page_number(value)

    def test_unwritable_output_target_is_rejected_before_processing(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "source.png"
            _content_image().save(source)
            not_a_directory = directory / "occupied"
            not_a_directory.write_text("file", encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "输出目录不可用"):
                CropProcessor()._process_images(
                    _config(not_a_directory),
                    [str(source)],
                )


class OutputCollisionTests(unittest.TestCase):
    def test_raster_and_transparency_outputs_do_not_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "source.png"
            _content_image().save(source)
            existing_crop = directory / "source_cropped.png"
            existing_crop.write_bytes(b"keep-crop")

            result = CropProcessor()._process_images(
                _config(directory),
                [str(source)],
            )

            self.assertEqual(result, str(directory))
            self.assertEqual(existing_crop.read_bytes(), b"keep-crop")
            self.assertTrue((directory / "source_cropped_2.png").exists())

            existing_transparent = directory / "source_transparent.png"
            existing_transparent.write_bytes(b"keep-transparent")
            transparent_config = {
                "source_files": [str(source)],
                "output_dir": str(directory),
                "output_format": "PNG",
                "color_mode": "corners",
                "tolerance": 18,
                "edge_only": True,
                "feather": 0,
                "dpi": 300,
            }
            CropProcessor().process_transparency(transparent_config)
            self.assertEqual(
                existing_transparent.read_bytes(),
                b"keep-transparent",
            )
            self.assertTrue(
                (directory / "source_transparent_2.png").exists()
            )

    def test_partial_batch_failure_does_not_report_full_progress(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "source.png"
            _content_image().save(source)
            progress_values = []
            logs = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: (
                    logs.append(log_entry)
                    if log_entry is not None
                    else progress_values.append(progress)
                    if progress is not None
                    else None
                )
            )

            result = processor._process_images(
                _config(directory),
                [str(source), str(directory / "missing.png")],
            )

            self.assertEqual(result, str(directory))
            self.assertTrue(processor.had_errors)
            self.assertNotIn(100, progress_values)
            self.assertIn(("WARNING", "部分完成 1 个文件"), logs)
            self.assertFalse(any(level == "SUCCESS" for level, _ in logs))

    def test_transparency_partial_batch_reports_warning_not_success(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "source.png"
            _content_image().save(source)
            logs = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: (
                    logs.append(log_entry) if log_entry else None
                )
            )

            result = processor.process_transparency({
                "source_files": [str(source), str(directory / "missing.png")],
                "output_dir": str(directory),
                "output_format": "PNG",
                "color_mode": "corners",
                "tolerance": 18,
                "edge_only": True,
                "feather": 0,
                "dpi": 300,
            })

            self.assertEqual(result, str(directory))
            self.assertIn(
                ("WARNING", "透明背景部分完成 1 个文件"),
                logs,
            )
            self.assertFalse(any(level == "SUCCESS" for level, _ in logs))

    def test_pdf_single_file_and_multi_page_svg_directory_do_not_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "source.pdf"
            document = pymupdf.open()
            for _ in range(2):
                page = document.new_page(width=100, height=80)
                page.draw_rect(
                    pymupdf.Rect(20, 15, 80, 65),
                    color=(0, 0, 0),
                    fill=(0, 0, 0),
                )
            document.save(source)
            document.close()

            existing_pdf = directory / "output.pdf"
            existing_pdf.write_bytes(b"keep-pdf")
            result_pdf = CropProcessor()._process_vector_crop(
                str(source),
                "PDF",
                str(existing_pdf),
                0,
                250,
                15,
                "smart",
                "output",
                page_indices=[0],
            )
            self.assertEqual(existing_pdf.read_bytes(), b"keep-pdf")
            self.assertEqual(result_pdf, str(directory / "output_2.pdf"))

            existing_svg_dir = directory / "slides"
            existing_svg_dir.mkdir()
            (existing_svg_dir / "keep.txt").write_text(
                "keep",
                encoding="utf-8",
            )
            result_svg = CropProcessor()._process_vector_crop(
                str(source),
                "SVG",
                str(directory / "slides.svg"),
                0,
                250,
                15,
                "smart",
                "slides",
            )
            self.assertEqual(result_svg, str(directory / "slides_2"))
            self.assertEqual(
                (existing_svg_dir / "keep.txt").read_text(encoding="utf-8"),
                "keep",
            )
            self.assertEqual(
                len(list((directory / "slides_2").glob("*.svg"))),
                2,
            )

    def test_pdf_raster_multi_page_directory_does_not_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "paper.pdf"
            document = pymupdf.open()
            for _ in range(2):
                page = document.new_page(width=80, height=60)
                page.draw_rect(
                    pymupdf.Rect(10, 10, 70, 50),
                    color=(0, 0, 0),
                    fill=(0, 0, 0),
                )
            document.save(source)
            document.close()
            existing = directory / "paper_Images"
            existing.mkdir()
            (existing / "keep.txt").write_text("keep", encoding="utf-8")

            result = CropProcessor()._process_pdf(
                _config(directory, "PNG", "ALL"),
                str(source),
            )

            self.assertEqual(result, str(directory / "paper_Images_2"))
            self.assertTrue((existing / "keep.txt").exists())
            self.assertEqual(
                len(list((directory / "paper_Images_2").glob("*.png"))),
                2,
            )

    def test_ppt_single_file_and_multi_page_directory_do_not_overwrite(self):
        class Controller:
            @staticmethod
            def get_page_setup():
                return 100, 80

            @staticmethod
            def export_single_image(path, width, index=1):
                _content_image().save(path)

        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            processor = CropProcessor()
            current_config = _config(directory)
            current_config["scope"] = "CURRENT"
            existing_file = directory / "deck_p1.png"
            existing_file.write_bytes(b"keep-ppt")

            current_result = processor._process_ppt_raster(
                Controller(),
                current_config,
                1,
                str(directory),
                "deck",
                1,
            )

            self.assertEqual(existing_file.read_bytes(), b"keep-ppt")
            self.assertEqual(current_result, str(directory / "deck_p1_2.png"))

            all_config = _config(directory, scope="ALL")
            existing_dir = directory / "deck_Images"
            existing_dir.mkdir()
            (existing_dir / "keep.txt").write_text("keep", encoding="utf-8")
            all_result = processor._process_ppt_raster(
                Controller(),
                all_config,
                1,
                str(directory),
                "deck",
                2,
            )

            self.assertEqual(all_result, str(directory / "deck_Images_2"))
            self.assertTrue((existing_dir / "keep.txt").exists())
            self.assertEqual(
                len(list((directory / "deck_Images_2").glob("*.png"))),
                2,
            )

    def test_ppt_all_continues_after_one_slide_export_failure(self):
        class Controller:
            @staticmethod
            def get_page_setup():
                return 100, 80

            @staticmethod
            def export_single_image(path, width, index=1):
                if index == 2:
                    raise RuntimeError("hidden slide is not in exported PDF")
                _content_image().save(path)

        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            logs = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: (
                    logs.append(log_entry) if log_entry else None
                )
            )

            result = processor._process_ppt_raster(
                Controller(),
                _config(directory, scope="ALL"),
                1,
                str(directory),
                "deck",
                3,
            )

            self.assertEqual(result, str(directory / "deck_Images"))
            self.assertTrue((directory / "deck_Images" / "deck_p001.png").exists())
            self.assertTrue((directory / "deck_Images" / "deck_p003.png").exists())
            self.assertTrue(any(
                level == "ERROR" and "第 2 页失败" in message
                for level, message in logs
            ))
            self.assertIn(
                ("WARNING", f"部分完成: {directory / 'deck_Images'}"),
                logs,
            )
            self.assertFalse(any(level == "SUCCESS" for level, _ in logs))


class MultiFrameTests(unittest.TestCase):
    def test_animated_gif_keeps_frames_timing_loop_and_disposal(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "animation.gif"
            boxes = ((10, 10, 29, 29), (40, 20, 59, 39), (20, 30, 39, 49))
            frames = [_content_image(box=box) for box in boxes]
            durations = [90, 180, 270]
            disposals = [2, 2, 2]
            frames[0].save(
                source,
                "GIF",
                save_all=True,
                append_images=frames[1:],
                duration=durations,
                loop=3,
                disposal=disposals,
            )

            result = CropProcessor()._process_images(
                _config(directory, "GIF"),
                [str(source)],
            )

            self.assertEqual(result, str(directory))
            target = directory / "animation_cropped.gif"
            with Image.open(target) as saved:
                self.assertEqual(saved.n_frames, 3)
                self.assertEqual(saved.size, (50, 40))
                self.assertEqual(saved.info.get("loop"), 3)
                saved_durations = []
                saved_disposals = []
                for index in range(saved.n_frames):
                    saved.seek(index)
                    saved_durations.append(saved.info.get("duration"))
                    saved_disposals.append(saved.disposal_method)
            self.assertEqual(saved_durations, durations)
            self.assertEqual(saved_disposals, disposals)

    def test_animated_gif_mixed_disposal_preserves_composited_frames(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "mixed-disposal.gif"
            frames = []
            for box, color in (
                ((0, 0, 9, 9), (255, 0, 0, 255)),
                ((15, 5, 24, 14), (0, 255, 0, 255)),
                ((5, 20, 14, 29), (0, 0, 255, 255)),
            ):
                frame = Image.new("RGBA", (40, 40), (0, 0, 0, 0))
                ImageDraw.Draw(frame).rectangle(box, fill=color)
                frames.append(frame)
            durations = [80, 140, 200]
            frames[0].save(
                source,
                "GIF",
                save_all=True,
                append_images=frames[1:],
                duration=durations,
                loop=2,
                disposal=[1, 2, 3],
            )

            expected = []
            with Image.open(source) as opened:
                for index in range(opened.n_frames):
                    opened.seek(index)
                    expected.append(opened.convert("RGBA"))
            boxes = [frame.getbbox() for frame in expected if frame.getbbox()]
            union_box = (
                min(box[0] for box in boxes),
                min(box[1] for box in boxes),
                max(box[2] for box in boxes),
                max(box[3] for box in boxes),
            )
            expected = [frame.crop(union_box) for frame in expected]

            CropProcessor()._process_images(
                _config(directory, "GIF"),
                [str(source)],
            )

            actual = []
            saved_disposals = []
            target = directory / "mixed-disposal_cropped.gif"
            with Image.open(target) as saved:
                self.assertEqual(saved.info.get("loop"), 2)
                saved_durations = []
                for index in range(saved.n_frames):
                    saved.seek(index)
                    actual.append(saved.convert("RGBA"))
                    saved_durations.append(saved.info.get("duration"))
                    saved_disposals.append(saved.disposal_method)
            self.assertEqual(saved_durations, durations)
            self.assertEqual(saved_disposals, [2, 2, 2])
            self.assertEqual(len(actual), len(expected))
            for expected_frame, actual_frame in zip(expected, actual):
                self.assertIsNone(
                    ImageChops.difference(expected_frame, actual_frame).getbbox()
                )

    def test_multi_page_tiff_keeps_pages_and_uses_union_crop(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "pages.tiff"
            frames = [
                _content_image(box=(5, 10, 24, 29)),
                _content_image(box=(35, 20, 64, 49)),
            ]
            frames[0].save(
                source,
                "TIFF",
                save_all=True,
                append_images=frames[1:],
                compression="tiff_lzw",
                dpi=(300, 300),
            )

            CropProcessor()._process_images(
                _config(directory, "TIFF"),
                [str(source)],
            )

            with Image.open(directory / "pages_cropped.tiff") as saved:
                self.assertEqual(saved.n_frames, 2)
                self.assertEqual(saved.size, (60, 40))

    def test_multi_page_tiff_preserves_16_bit_pixels(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "sixteen-bit.tiff"
            frames = []
            for box in ((5, 10, 24, 29), (35, 20, 64, 49)):
                frame = Image.new("I;16", (80, 60), 65535)
                ImageDraw.Draw(frame).rectangle(box, fill=0)
                frames.append(frame)
            frames[0].save(
                source,
                "TIFF",
                save_all=True,
                append_images=frames[1:],
                compression="tiff_lzw",
            )

            CropProcessor()._process_images(
                _config(directory, "TIFF"),
                [str(source)],
            )

            with Image.open(directory / "sixteen-bit_cropped.tiff") as saved:
                self.assertEqual(saved.n_frames, 2)
                for index in range(saved.n_frames):
                    saved.seek(index)
                    self.assertIn(saved.mode, ("I;16", "I;16L"))
                    self.assertEqual(saved.getextrema(), (0, 65535))

    def test_animated_webp_keeps_frames_timing_and_loop(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "animation.webp"
            frames = [
                _content_image(box=(10, 10, 29, 29)).convert("RGBA"),
                _content_image(box=(35, 20, 59, 44)).convert("RGBA"),
            ]
            frames[0].save(
                source,
                "WEBP",
                save_all=True,
                append_images=frames[1:],
                duration=[100, 220],
                loop=2,
                lossless=True,
            )

            result = CropProcessor()._process_images(
                _config(directory, "WebP"),
                [str(source)],
            )

            self.assertEqual(result, str(directory))
            with Image.open(directory / "animation_cropped.webp") as saved:
                self.assertEqual(saved.n_frames, 2)
                self.assertEqual(saved.info.get("loop"), 2)
            self.assertEqual(
                _webp_durations(directory / "animation_cropped.webp"),
                [100, 220],
            )

    def test_incompatible_multiframe_output_is_explicit_and_not_partial(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "animation.gif"
            frames = [_content_image(), _content_image(box=(10, 10, 30, 30))]
            frames[0].save(
                source,
                "GIF",
                save_all=True,
                append_images=frames[1:],
                duration=[100, 200],
                loop=0,
                disposal=[2, 2],
            )
            logs = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: logs.append(
                    log_entry
                ) if log_entry else None
            )

            result = processor._process_images(
                _config(directory, "PNG"),
                [str(source)],
            )

            self.assertIsNone(result)
            self.assertFalse((directory / "animation_cropped.png").exists())
            self.assertTrue(any(
                level == "ERROR" and "无法完整保留多帧结构" in message
                for level, message in logs
            ))

    def test_transparency_rejects_multiframe_input_without_partial_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "animation.gif"
            frames = [_content_image(), _content_image(box=(10, 10, 30, 30))]
            frames[0].save(
                source,
                "GIF",
                save_all=True,
                append_images=frames[1:],
                duration=[100, 200],
                loop=0,
                disposal=[2, 2],
            )
            logs = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: logs.append(
                    log_entry
                ) if log_entry else None
            )

            result = processor.process_transparency({
                "source_files": [str(source)],
                "output_dir": str(directory),
                "output_format": "PNG",
                "color_mode": "corners",
                "tolerance": 18,
                "edge_only": True,
                "feather": 0,
                "dpi": 300,
            })

            self.assertIsNone(result)
            self.assertFalse(
                (directory / "animation_transparent.png").exists()
            )
            self.assertTrue(any(
                level == "ERROR" and "暂不支持保留 2 帧" in message
                for level, message in logs
            ))


class PdfSafetyTests(unittest.TestCase):
    def test_orphan_ocmd_object_disables_destructive_hard_crop(self):
        class OptionalContentDocument:
            def get_ocgs(self):
                return {}

            def pdf_catalog(self):
                return 1

            def xref_length(self):
                return 4

            def xref_get_key(self, xref, key):
                if key == "OCProperties":
                    return "null", "null"
                if key == "Type" and xref == 3:
                    return "name", "/OCMD"
                return "null", "null"

        self.assertTrue(CropProcessor._document_has_optional_content(
            OptionalContentDocument()
        ))

    def test_encrypted_pdf_returns_controlled_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "encrypted.pdf"
            document = pymupdf.open()
            document.new_page().insert_text((72, 72), "secret")
            document.save(
                source,
                encryption=pymupdf.PDF_ENCRYPT_AES_256,
                owner_pw="owner",
                user_pw="secret",
            )
            document.close()
            logs = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: logs.append(
                    log_entry
                ) if log_entry else None
            )

            result = processor._process_pdf(
                _config(directory, "PDF"),
                str(source),
            )

            self.assertIsNone(result)
            self.assertTrue(any(
                level == "ERROR" and "密码保护" in message
                for level, message in logs
            ))

    def test_out_of_range_page_is_reported_without_rendering(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "one-page.pdf"
            document = pymupdf.open()
            document.new_page()
            document.save(source)
            document.close()
            logs = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: logs.append(
                    log_entry
                ) if log_entry else None
            )
            config = _config(directory)
            config["page_num"] = 2

            result = processor._process_pdf(config, str(source))

            self.assertIsNone(result)
            self.assertTrue(any(
                level == "ERROR" and "1–1" in message
                for level, message in logs
            ))

    def test_pdf_to_svg_warns_when_interactive_structures_will_be_lost(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "interactive.pdf"
            document = pymupdf.open()
            page = document.new_page(width=200, height=120)
            page.insert_text((30, 60), "OpenAI")
            page.insert_link({
                "kind": pymupdf.LINK_URI,
                "from": pymupdf.Rect(25, 40, 100, 70),
                "uri": "https://openai.com",
            })
            document.set_toc([[1, "Chapter", 1]])
            document.save(source)
            document.close()
            logs = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: logs.append(
                    log_entry
                ) if log_entry else None
            )

            result = processor._process_pdf(
                _config(directory, "SVG"),
                str(source),
            )

            self.assertTrue(result and os.path.exists(result))
            self.assertTrue(any(
                level == "WARNING"
                and "视觉矢量内容" in message
                and "链接" in message
                and "书签" in message
                for level, message in logs
            ))

    def test_failed_blank_export_never_reports_full_progress(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "blank.pdf"
            document = pymupdf.open()
            document.new_page(width=100, height=80)
            document.save(source)
            document.close()
            progress_values = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: (
                    progress_values.append(progress)
                    if progress is not None else None
                )
            )

            result = processor._process_pdf(
                _config(directory, "PNG"),
                str(source),
            )

            self.assertIsNone(result)
            self.assertNotIn(100, progress_values)

    def test_blank_vector_pdf_does_not_create_success_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = directory / "blank-vector.pdf"
            target = directory / "blank-vector-cropped.pdf"
            document = pymupdf.open()
            document.new_page(width=100, height=80)
            document.save(source)
            document.close()
            logs = []
            progress_values = []
            processor = CropProcessor(
                callback=lambda status, log_entry, progress: (
                    logs.append(log_entry) if log_entry else
                    progress_values.append(progress)
                    if progress is not None else None
                )
            )

            result = processor._process_vector_crop(
                str(source),
                "PDF",
                str(target),
                0,
                250,
                15,
                "smart",
                "blank-vector",
            )

            self.assertIsNone(result)
            self.assertFalse(target.exists())
            self.assertNotIn(100, progress_values)
            self.assertTrue(any(
                level == "WARNING" and "未检测到可输出内容" in message
                for level, message in logs
            ))


if __name__ == "__main__":
    unittest.main()
