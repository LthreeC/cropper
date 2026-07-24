import unittest

from ui import get_output_quality_policy


class OutputQualityPolicyTests(unittest.TestCase):
    def assert_policy(self, mode, source_kind, output_format, expected):
        quality, hint = get_output_quality_policy(
            mode,
            source_kind,
            output_format,
        )
        self.assertEqual(quality, expected)
        self.assertTrue(hint)

    def test_ppt_vector_outputs_offer_embedded_image_dpi(self):
        for output_format in ("PDF", "SVG"):
            with self.subTest(output_format=output_format):
                self.assert_policy(
                    "PPT",
                    "ppt",
                    output_format,
                    "pdf_image_dpi",
                )

    def test_ppt_raster_outputs_offer_render_dpi(self):
        for output_format in ("PNG", "TIFF", "JPEG", "WebP"):
            with self.subTest(output_format=output_format):
                self.assert_policy(
                    "PPT",
                    "ppt",
                    output_format,
                    "dpi",
                )

    def test_document_vector_outputs_do_not_offer_dpi(self):
        for output_format in ("PDF", "SVG"):
            with self.subTest(output_format=output_format):
                self.assert_policy(
                    "FILE",
                    "document",
                    output_format,
                    "none",
                )

    def test_document_raster_outputs_offer_render_dpi(self):
        for output_format in ("PNG", "TIFF", "JPEG", "WebP"):
            with self.subTest(output_format=output_format):
                self.assert_policy(
                    "FILE",
                    "document",
                    output_format,
                    "dpi",
                )

    def test_raster_inputs_never_offer_resampling_dpi(self):
        for output_format in ("PDF", "SVG", "PNG", "TIFF", "JPEG", "WebP"):
            with self.subTest(output_format=output_format):
                self.assert_policy(
                    "FILE",
                    "raster",
                    output_format,
                    "none",
                )


if __name__ == "__main__":
    unittest.main()
