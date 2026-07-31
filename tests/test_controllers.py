import sys
import re
import types
import unittest
from unittest.mock import patch
from pathlib import Path

from PIL import Image

from controllers import FileController, MacPPTController, WindowsPPTController


class WindowsPPTControllerTests(unittest.TestCase):
    def test_connection_initializes_and_releases_com_for_current_thread(self):
        events = []
        pythoncom = types.ModuleType("pythoncom")
        pythoncom.initialized = False

        def co_initialize():
            pythoncom.initialized = True
            events.append("initialize")

        def co_uninitialize():
            pythoncom.initialized = False
            events.append("uninitialize")

        pythoncom.CoInitialize = co_initialize
        pythoncom.CoUninitialize = co_uninitialize

        app = object()
        client = types.ModuleType("win32com.client")

        def get_active_object(prog_id):
            events.append(f"connect:{prog_id}")
            if not pythoncom.initialized:
                raise RuntimeError("CoInitialize has not been called")
            return app

        client.GetActiveObject = get_active_object
        win32com = types.ModuleType("win32com")
        win32com.client = client

        with patch.dict(
            sys.modules,
            {"pythoncom": pythoncom, "win32com": win32com, "win32com.client": client},
        ):
            controller = WindowsPPTController()
            self.assertTrue(controller.check_connection())
            self.assertIs(controller.app, app)
            controller.close()

        self.assertEqual(
            events,
            ["initialize", "connect:PowerPoint.Application", "uninitialize"],
        )
        self.assertIsNone(controller.app)

    def test_source_copy_uses_openxml_presentation_format(self):
        calls = []
        presentation = types.SimpleNamespace(
            SaveCopyAs=lambda *args: calls.append(args)
        )
        controller = WindowsPPTController()
        controller.app = types.SimpleNamespace(ActivePresentation=presentation)

        controller.export_source_copy("source.pptx")

        self.assertEqual(calls, [("source.pptx", 24)])

    def test_unknown_current_slide_fails_instead_of_defaulting_to_first(self):
        def no_selection(index):
            raise RuntimeError("no slide selection")

        window = types.SimpleNamespace(
            Presentation=types.SimpleNamespace(
                Name="deck.pptx",
                FullName="C:/deck.pptx",
            ),
            View=types.SimpleNamespace(),
            Selection=types.SimpleNamespace(SlideRange=no_selection),
        )
        controller = WindowsPPTController()
        controller.app = types.SimpleNamespace(
            Windows=types.SimpleNamespace(Count=1),
            ActiveWindow=window,
        )

        with self.assertRaisesRegex(RuntimeError, "无法.*活动幻灯片"):
            controller.get_info()

    def test_pdf_export_uses_print_quality_for_all_slides(self):
        calls = []
        presentation = types.SimpleNamespace(
            ExportAsFixedFormat=lambda **kwargs: calls.append(kwargs)
        )
        controller = WindowsPPTController()
        controller.app = types.SimpleNamespace(ActivePresentation=presentation)

        controller.export_temp_pdf("output.pdf", scope="ALL")

        self.assertEqual(
            calls,
            [{
                "Path": "output.pdf",
                "FixedFormatType": 2,
                "Intent": 2,
                "PrintRange": None,
                "RangeType": 1,
            }],
        )

    def test_pdf_export_uses_print_quality_for_current_slide(self):
        calls = []

        presentation = types.SimpleNamespace(
            ExportAsFixedFormat=lambda **kwargs: calls.append(kwargs)
        )
        controller = WindowsPPTController()
        controller.app = types.SimpleNamespace(ActivePresentation=presentation)

        controller.export_temp_pdf("output.pdf", scope="CURRENT", index=3)

        self.assertEqual(
            calls,
            [{
                "Path": "output.pdf",
                "FixedFormatType": 2,
                "Intent": 2,
                "PrintRange": None,
                "RangeType": 3,
            }],
        )

    def test_raster_export_sets_both_pixel_dimensions(self):
        calls = []
        slide = types.SimpleNamespace(
            Export=lambda *args: calls.append(args)
        )
        slides = lambda index: slide
        page_setup = types.SimpleNamespace(SlideWidth=960, SlideHeight=540)
        presentation = types.SimpleNamespace(
            Slides=slides,
            PageSetup=page_setup,
        )
        controller = WindowsPPTController()
        controller.app = types.SimpleNamespace(ActivePresentation=presentation)

        controller.export_single_image("output.png", 4000, index=2)

        self.assertEqual(
            calls,
            [("output.png", "PNG", 4000, 2250)],
        )

    def test_current_export_fails_if_active_slide_changed(self):
        calls = []
        presentation = types.SimpleNamespace(
            ExportAsFixedFormat=lambda **kwargs: calls.append(kwargs)
        )
        controller = WindowsPPTController()
        controller.app_kind = "PowerPoint"
        controller.app = types.SimpleNamespace(
            ActivePresentation=presentation,
            ActiveWindow=types.SimpleNamespace(
                View=types.SimpleNamespace(
                    Slide=types.SimpleNamespace(SlideIndex=1)
                )
            ),
        )

        with self.assertRaisesRegex(Exception, "活动幻灯片.*变化"):
            controller.export_temp_pdf(
                "output.pdf",
                scope="CURRENT",
                index=3,
            )
        self.assertEqual(calls, [])

    def test_wps_current_fallback_extracts_only_requested_slide(self):
        import pymupdf

        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        output = directory / "wps-current.pdf"
        if output.exists():
            output.unlink()

        def fail_export(**kwargs):
            raise RuntimeError("ExportAsFixedFormat unsupported")

        def save_as(path, file_format):
            self.assertEqual(file_format, 32)
            doc = pymupdf.open()
            for page_number in range(1, 4):
                page = doc.new_page(width=72, height=36)
                page.insert_text((8, 20), f"slide-{page_number}")
            doc.save(path)
            doc.close()

        presentation = types.SimpleNamespace(
            ExportAsFixedFormat=fail_export,
            SaveAs=save_as,
            Slides=types.SimpleNamespace(Count=3),
            FullName="Unsaved",
        )
        controller = WindowsPPTController()
        controller.app_kind = "WPS"
        controller.app = types.SimpleNamespace(ActivePresentation=presentation)
        try:
            controller.export_temp_pdf(str(output), scope="CURRENT", index=3)
            with pymupdf.open(output) as doc:
                self.assertEqual(len(doc), 1)
                self.assertIn("slide-3", doc[0].get_text())
        finally:
            if output.exists():
                output.unlink()

    def test_wps_hidden_page_mapping_rejects_unsaved_source_state(self):
        import pymupdf

        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        source = directory / "wps-stale.pptx"
        output = directory / "wps-stale-current.pdf"
        source.write_bytes(b"stale-on-disk")
        if output.exists():
            output.unlink()

        def fail_export(**kwargs):
            raise RuntimeError("ExportAsFixedFormat unsupported")

        def save_as(path, file_format):
            doc = pymupdf.open()
            doc.new_page()
            doc.new_page()
            doc.save(path)
            doc.close()
            presentation.FullName = path
            presentation.Saved = True

        presentation = types.SimpleNamespace(
            ExportAsFixedFormat=fail_export,
            SaveAs=save_as,
            Slides=types.SimpleNamespace(Count=3),
            FullName=str(source),
            Saved=False,
        )
        controller = WindowsPPTController()
        controller.app_kind = "WPS"
        controller.app = types.SimpleNamespace(ActivePresentation=presentation)
        try:
            with self.assertRaisesRegex(Exception, "未保存更改"):
                controller.export_temp_pdf(
                    str(output),
                    scope="CURRENT",
                    index=3,
                )
            self.assertFalse(output.exists())
        finally:
            for path in (source, output):
                if path.exists():
                    path.unlink()


class MacPPTControllerTests(unittest.TestCase):
    def test_info_requests_posix_source_path(self):
        scripts = []
        controller = MacPPTController()

        def run(script):
            scripts.append(script)
            return "2\x1fdeck|v1.pptx\x1f/Users/test/deck|v1.pptx"

        controller._run_applescript = run

        self.assertEqual(
            controller.get_info(),
            (2, "deck|v1.pptx", "/Users/test/deck|v1.pptx"),
        )
        self.assertIn("POSIX path", scripts[0])

    def test_source_copy_uses_local_pptx_without_resaving_presentation(self):
        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        source = directory / "mac-source.pptx"
        target = directory / "mac-copy.pptx"
        source.write_bytes(b"pptx-source")
        if target.exists():
            target.unlink()

        controller = MacPPTController()
        controller.get_info = lambda: (1, source.name, str(source))
        try:
            controller.export_source_copy(str(target))
            self.assertEqual(target.read_bytes(), b"pptx-source")
        finally:
            for path in (source, target):
                if path.exists():
                    path.unlink()

    def test_current_slide_pdf_export_selects_requested_page(self):
        import pymupdf

        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        output = directory / "mac-current.pdf"
        if output.exists():
            output.unlink()

        controller = MacPPTController()

        def export_all(script):
            match = re.search(r'POSIX file "([^"]+)"', script)
            self.assertIsNotNone(match)
            exported = match.group(1).replace("\\\\", "\\")
            doc = pymupdf.open()
            for page_number in range(1, 4):
                page = doc.new_page(width=72, height=36)
                page.insert_text((8, 20), f"slide-{page_number}")
            doc.save(exported)
            doc.close()

        controller._run_applescript = export_all
        controller.get_slide_count = lambda: 3
        try:
            controller.export_temp_pdf(str(output), scope="CURRENT", index=2)
            with pymupdf.open(output) as doc:
                self.assertEqual(len(doc), 1)
                self.assertIn("slide-2", doc[0].get_text())
        finally:
            if output.exists():
                output.unlink()

    def test_current_slide_pdf_maps_visible_pages_around_hidden_slide(self):
        import pymupdf

        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        source = directory / "mac-hidden.pptx"
        output = directory / "mac-visible-current.pdf"
        source.write_bytes(b"placeholder")
        if output.exists():
            output.unlink()

        controller = MacPPTController()

        def export_visible(script):
            match = re.search(r'POSIX file "([^"]+)"', script)
            exported = match.group(1).replace("\\\\", "\\")
            doc = pymupdf.open()
            for slide_number in (1, 3):
                page = doc.new_page(width=72, height=36)
                page.insert_text((8, 20), f"slide-{slide_number}")
            doc.save(exported)
            doc.close()

        controller._run_applescript = export_visible
        controller.get_slide_count = lambda: 3
        controller.get_info = lambda: (3, source.name, str(source))
        controller._source_presentation_is_saved = lambda: True
        try:
            with patch(
                "ppt_image_restore.pptx_visible_slide_indices",
                return_value=(0, 2),
            ):
                controller.export_temp_pdf(
                    str(output),
                    scope="CURRENT",
                    index=3,
                )
            with pymupdf.open(output) as doc:
                self.assertEqual(len(doc), 1)
                self.assertIn("slide-3", doc[0].get_text())
        finally:
            for path in (source, output):
                if path.exists():
                    path.unlink()

    def test_current_hidden_slide_fails_instead_of_returning_another_page(self):
        import pymupdf

        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        source = directory / "mac-hidden.pptx"
        output = directory / "mac-hidden-current.pdf"
        source.write_bytes(b"placeholder")
        if output.exists():
            output.unlink()

        controller = MacPPTController()

        def export_visible(script):
            match = re.search(r'POSIX file "([^"]+)"', script)
            exported = match.group(1).replace("\\\\", "\\")
            doc = pymupdf.open()
            doc.new_page()
            doc.new_page()
            doc.save(exported)
            doc.close()

        controller._run_applescript = export_visible
        controller.get_slide_count = lambda: 3
        controller.get_info = lambda: (2, source.name, str(source))
        controller._source_presentation_is_saved = lambda: True
        try:
            with patch(
                "ppt_image_restore.pptx_visible_slide_indices",
                return_value=(0, 2),
            ), self.assertRaisesRegex(ValueError, "隐藏幻灯片"):
                controller.export_temp_pdf(
                    str(output),
                    scope="CURRENT",
                    index=2,
                )
            self.assertFalse(output.exists())
        finally:
            for path in (source, output):
                if path.exists():
                    path.unlink()

    def test_current_slide_mapping_fails_without_local_pptx(self):
        import pymupdf

        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        output = directory / "mac-unsaved-current.pdf"
        if output.exists():
            output.unlink()

        controller = MacPPTController()

        def export_visible(script):
            match = re.search(r'POSIX file "([^"]+)"', script)
            exported = match.group(1).replace("\\\\", "\\")
            doc = pymupdf.open()
            doc.new_page()
            doc.new_page()
            doc.save(exported)
            doc.close()

        controller._run_applescript = export_visible
        controller.get_slide_count = lambda: 3
        controller.get_info = lambda: (3, "Unsaved", "Unsaved")
        controller._source_presentation_is_saved = lambda: True
        try:
            with self.assertRaisesRegex(RuntimeError, "本地 PPTX"):
                controller.export_temp_pdf(
                    str(output),
                    scope="CURRENT",
                    index=3,
                )
            self.assertFalse(output.exists())
        finally:
            if output.exists():
                output.unlink()

    def test_current_slide_mapping_rejects_stale_local_pptx(self):
        import pymupdf

        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        source = directory / "mac-stale.pptx"
        output = directory / "mac-stale-current.pdf"
        source.write_bytes(b"stale-on-disk")
        if output.exists():
            output.unlink()

        controller = MacPPTController()
        source_state = {"saved": False}

        def export_visible(script):
            match = re.search(r'POSIX file "([^"]+)"', script)
            exported = match.group(1).replace("\\\\", "\\")
            doc = pymupdf.open()
            doc.new_page()
            doc.new_page()
            doc.save(exported)
            doc.close()
            source_state["saved"] = True

        controller._run_applescript = export_visible
        controller.get_slide_count = lambda: 3
        controller.get_info = lambda: (3, source.name, str(source))
        controller._source_presentation_is_saved = lambda: source_state["saved"]
        try:
            with self.assertRaisesRegex(RuntimeError, "未保存更改"):
                controller.export_temp_pdf(
                    str(output),
                    scope="CURRENT",
                    index=3,
                )
            self.assertFalse(output.exists())
        finally:
            for path in (source, output):
                if path.exists():
                    path.unlink()

    def test_raster_export_honors_requested_pixel_width(self):
        import pymupdf

        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        output = directory / "mac-slide.png"
        if output.exists():
            output.unlink()

        controller = MacPPTController()

        def export_current(path, scope="CURRENT", index=1):
            doc = pymupdf.open()
            doc.new_page(width=72, height=36)
            doc.save(path)
            doc.close()

        controller.export_temp_pdf = export_current
        try:
            controller.export_single_image(str(output), 600, index=2)
            with Image.open(output) as rendered:
                self.assertEqual(rendered.size, (600, 300))
        finally:
            if output.exists():
                output.unlink()


class FileControllerTests(unittest.TestCase):
    def test_load_image_applies_exif_orientation(self):
        directory = Path(__file__).resolve().parent.parent / "tmp" / "controllers"
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / "oriented.jpg"
        if path.exists():
            path.unlink()
        try:
            image = Image.new("RGB", (80, 40), "white")
            exif = Image.Exif()
            exif[274] = 6
            image.save(path, exif=exif)

            loaded = FileController().load_image(path)

            self.assertEqual(loaded.size, (40, 80))
            self.assertNotIn(274, loaded.getexif())
        finally:
            if path.exists():
                path.unlink()


if __name__ == "__main__":
    unittest.main()
