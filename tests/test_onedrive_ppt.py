import os
import types
import unittest
from unittest.mock import patch

from controllers import WindowsPPTController
from processor import CropProcessor


class OneDrivePptTests(unittest.TestCase):
    def test_current_slide_export_does_not_use_print_range_object(self):
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

    def test_cloud_ppt_defaults_output_to_desktop(self):
        class Controller:
            def check_connection(self):
                return True

            def get_info(self):
                return (
                    1,
                    "演示文稿2.pptx",
                    "https://d.docs.live.net/account/文档/演示文稿2.pptx",
                )

            def close(self):
                pass

        processor = CropProcessor()
        config = {
            "scope": "CURRENT",
            "output_format": "PDF",
            "output_dir": None,
        }

        with (
            patch("controllers.get_ppt_controller", return_value=Controller()),
            patch.object(processor, "_process_ppt_vector", return_value="done") as process,
            patch("processor.os.path.expanduser", return_value=r"C:\Test\Desktop"),
        ):
            result = processor.process_ppt(config)

        self.assertEqual(result, "done")
        self.assertEqual(process.call_args.args[3], r"C:\Test\Desktop")

    def test_locked_temp_file_cleanup_does_not_raise(self):
        logs = []
        processor = CropProcessor(
            callback=lambda status, log_entry, progress: logs.append(log_entry)
            if log_entry else None
        )

        with (
            patch("processor.os.path.exists", return_value=True),
            patch("processor.os.remove", side_effect=PermissionError("locked")),
            patch("processor.time.sleep"),
        ):
            processor._cleanup_temp_file(os.path.join("temp", "locked.pdf"))

        self.assertTrue(any(level == "WARNING" for level, _ in logs))


if __name__ == "__main__":
    unittest.main()
