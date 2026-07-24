import types
import unittest

from controllers import WindowsPPTController


class PdfQualityFailureTests(unittest.TestCase):
    def test_powerpoint_does_not_silently_fallback_to_save_as(self):
        save_as_calls = []

        def fail_export(**kwargs):
            raise TypeError("invalid COM arguments")

        presentation = types.SimpleNamespace(
            ExportAsFixedFormat=fail_export,
            SaveAs=lambda *args: save_as_calls.append(args),
        )
        controller = WindowsPPTController()
        controller.app = types.SimpleNamespace(ActivePresentation=presentation)
        controller.app_kind = "PowerPoint"

        with self.assertRaisesRegex(Exception, "高质量导出 PDF 失败"):
            controller.export_temp_pdf("output.pdf", scope="ALL")

        self.assertEqual(save_as_calls, [])


if __name__ == "__main__":
    unittest.main()
