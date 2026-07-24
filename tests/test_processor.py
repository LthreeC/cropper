import unittest
from unittest.mock import patch

from processor import CropProcessor


class CropProcessorTests(unittest.TestCase):
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

        processor = CropProcessor()
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
                return_value=[],
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


if __name__ == "__main__":
    unittest.main()
