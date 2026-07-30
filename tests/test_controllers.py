import sys
import types
import unittest
from unittest.mock import patch
from pathlib import Path

from PIL import Image

from controllers import FileController, WindowsPPTController


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
