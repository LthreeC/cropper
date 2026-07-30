import unittest

from PIL import Image, ImageDraw

from detector import MAX_DETECTION_SIZE, get_bbox


class WhiteBorderDetectorTests(unittest.TestCase):
    def test_all_modes_detect_the_same_dark_content_region(self):
        rgb = Image.new("RGB", (300, 200), "white")
        ImageDraw.Draw(rgb).rectangle(
            (50, 40, 250, 160),
            fill=(80, 120, 180),
        )

        rgba = Image.new("RGBA", (300, 200), (255, 255, 255, 0))
        ImageDraw.Draw(rgba).rectangle(
            (50, 40, 250, 160),
            fill=(80, 120, 180, 255),
        )

        palette = Image.new("P", (300, 200), 0)
        palette.putpalette(
            [255, 255, 255, 80, 120, 180] + [0, 0, 0] * 254
        )
        palette.info["transparency"] = 0
        ImageDraw.Draw(palette).rectangle((50, 40, 250, 160), fill=1)

        luminance_alpha = Image.new("LA", (300, 200), (255, 0))
        ImageDraw.Draw(luminance_alpha).rectangle(
            (50, 40, 250, 160),
            fill=(100, 255),
        )

        cmyk = Image.new("CMYK", (300, 200), (0, 0, 0, 0))
        ImageDraw.Draw(cmyk).rectangle(
            (50, 40, 250, 160),
            fill=(0, 0, 0, 255),
        )

        for image in (rgb, rgba, palette, luminance_alpha, cmyk):
            for mode in ("simple", "edge", "smart"):
                with self.subTest(image_mode=image.mode, mode=mode):
                    bbox = get_bbox(
                        image,
                        threshold=250,
                        sensitivity=15,
                        mode=mode,
                    )
                    self.assertEqual(bbox, (50, 40, 251, 161))

    def test_native_resolution_detection_keeps_exact_coordinates(self):
        image = Image.new('RGB', (2400, 1200), 'white')
        ImageDraw.Draw(image).rectangle(
            (401, 203, 1999, 999),
            fill=(80, 120, 180),
        )

        bbox = get_bbox(
            image,
            threshold=250,
            sensitivity=15,
            mode='smart',
            max_detection_size=None,
        )

        self.assertEqual(bbox, (401, 203, 2000, 1000))

    def test_large_image_coordinates_are_mapped_back_conservatively(self):
        image = Image.new(
            "RGB",
            (MAX_DETECTION_SIZE * 2, MAX_DETECTION_SIZE),
            "white",
        )
        ImageDraw.Draw(image).rectangle(
            (400, 200, 2000, 1000),
            fill="black",
        )

        bbox = get_bbox(image, mode="smart")

        self.assertLessEqual(bbox[0], 400)
        self.assertLessEqual(bbox[1], 200)
        self.assertGreaterEqual(bbox[2], 2001)
        self.assertGreaterEqual(bbox[3], 1001)


if __name__ == "__main__":
    unittest.main()
