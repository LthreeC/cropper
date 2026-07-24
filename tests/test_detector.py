import unittest

from PIL import Image, ImageDraw

from detector import MAX_DETECTION_SIZE, get_bbox


class WhiteBorderDetectorTests(unittest.TestCase):
    def test_all_modes_detect_the_same_dark_content_region(self):
        image = Image.new("RGB", (300, 200), "white")
        ImageDraw.Draw(image).rectangle(
            (50, 40, 250, 160),
            fill=(80, 120, 180),
        )

        for mode in ("simple", "edge", "smart"):
            with self.subTest(mode=mode):
                bbox = get_bbox(
                    image,
                    threshold=250,
                    sensitivity=15,
                    mode=mode,
                )
                self.assertIsNotNone(bbox)
                self.assertLessEqual(bbox[0], 50)
                self.assertLessEqual(bbox[1], 40)
                self.assertGreaterEqual(bbox[2], 251)
                self.assertGreaterEqual(bbox[3], 161)

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
