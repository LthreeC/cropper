import io
import unittest

import pymupdf
from PIL import Image

from ppt_image_restore import _collect_pdf_images, _limit_image_resolution


class ImageTransformUnitTests(unittest.TestCase):
    def test_rotated_image_uses_its_own_display_axes_for_dpi(self):
        buffer = io.BytesIO()
        Image.new("RGB", (2000, 1000), "red").save(buffer, "PNG")

        doc = pymupdf.open()
        page = doc.new_page(width=300, height=300)
        page.insert_image(
            pymupdf.Rect(20, 20, 120, 220),
            stream=buffer.getvalue(),
            rotate=90,
        )

        try:
            pdf_image = _collect_pdf_images(doc)[0]
            self.assertEqual(pdf_image["display_size"], (200.0, 100.0))

            limited = _limit_image_resolution(
                Image.new("RGB", (2000, 1000), "red"),
                pdf_image["display_size"],
                max_image_dpi=300,
            )
            self.assertEqual(limited.size, (833, 417))
        finally:
            doc.close()


if __name__ == "__main__":
    unittest.main()
