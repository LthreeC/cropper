import unittest
from unittest.mock import patch

from processor import CropProcessor


class CropProcessorTests(unittest.TestCase):
    def test_ppt_raster_export_rejects_invalid_dpi(self):
        config = {
            "scope": "CURRENT",
            "output_format": "PNG",
            "padding": 2,
            "threshold": 250,
            "sensitivity": 15,
            "detect_mode": "smart",
            "dpi": 0,
        }

        with self.assertRaisesRegex(ValueError, "输出 DPI"):
            CropProcessor()._process_ppt_raster(
                object(),
                config,
                1,
                ".",
                "deck",
                1,
            )

    def test_process_ppt_closes_controller_when_connection_fails(self):
        class Controller:
            closed = False

            def check_connection(self):
                return False

            def close(self):
                self.closed = True

        controller = Controller()
        with patch("controllers.get_ppt_controller", return_value=controller):
            result = CropProcessor().process_ppt({})

        self.assertIsNone(result)
        self.assertTrue(controller.closed)

    def test_vector_export_passes_selected_pdf_image_dpi(self):
        class Controller:
            def export_temp_pdf(self, path, scope, index):
                pass

            def export_source_copy(self, path):
                pass

        logs = []
        processor = CropProcessor(
            callback=lambda status, log_entry, progress: (
                logs.append(log_entry) if log_entry else None
            )
        )
        config = {
            "scope": "CURRENT",
            "output_format": "PDF",
            "padding": 2,
            "threshold": 250,
            "sensitivity": 15,
            "detect_mode": "smart",
            "pdf_image_dpi": 600,
        }

        class Document:
            closed = False

            def close(self):
                self.closed = True

        document = Document()

        class RestoreOutcome(list):
            stats = {
                "total": 1,
                "restored": 0,
                "unmatched": 1,
                "ambiguous": 0,
                "sufficient": 0,
            }

        with (
            patch.object(
                processor,
                "_make_temp_path",
                side_effect=["temp.pdf", "source.pptx"],
            ),
            patch.object(processor, "_cleanup_temp_file"),
            patch.object(
                processor,
                "_process_vector_crop",
                return_value="done",
            ) as crop,
            patch("pymupdf.open", return_value=document),
            patch(
                "ppt_image_restore.restore_pptx_images_in_document",
                return_value=RestoreOutcome(),
            ) as restore,
        ):
            result = processor._process_ppt_vector(
                Controller(), config, 1, ".", "deck", 1
            )

        self.assertEqual(result, "done")
        restore.assert_called_once_with(
            document,
            "source.pptx",
            max_image_dpi=600,
            slide_indices=[0],
        )
        crop.assert_called_once_with(
            document,
            "PDF",
            ".\\deck_p1.pdf",
            2,
            250,
            15,
            "smart",
            "deck",
            pdf_garbage=4,
        )
        self.assertTrue(document.closed)
        self.assertTrue(any(
            level == "WARNING" and "无法安全恢复高清源图" in message
            for level, message in logs
        ))


if __name__ == "__main__":
    unittest.main()
