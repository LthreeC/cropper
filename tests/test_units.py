import unittest

from PIL import Image

from ppt_image_restore import _limit_image_resolution
from units import POINTS_PER_INCH, points_to_pixels


class UnitConversionTests(unittest.TestCase):
    def test_pdf_point_definition_and_dpi_scaling(self):
        self.assertEqual(POINTS_PER_INCH, 72.0)
        self.assertEqual(points_to_pixels(72, 300), 300)
        self.assertEqual(points_to_pixels(36, 144), 72)

    def test_image_limit_adapts_to_display_size_and_selected_dpi(self):
        source = Image.new("RGB", (2000, 2000), "white")

        limited = _limit_image_resolution(
            source,
            display_size=(144, 36),
            max_image_dpi=450,
        )

        self.assertEqual(limited.size, (900, 225))


if __name__ == "__main__":
    unittest.main()
