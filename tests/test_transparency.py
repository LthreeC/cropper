from pathlib import Path
import unittest

from PIL import Image, ImageDraw

from processor import CropProcessor
from transparency import make_transparent, parse_hex_color


class TransparencyTests(unittest.TestCase):
    def test_edge_only_preserves_enclosed_matching_color(self):
        image = Image.new("RGB", (100, 100), "white")
        draw = ImageDraw.Draw(image)
        draw.rectangle((20, 20, 80, 80), fill="black")
        draw.rectangle((40, 40, 60, 60), fill="white")

        result, _, _ = make_transparent(
            image,
            target_color=(255, 255, 255),
            tolerance=0,
            edge_only=True,
            feather=0,
        )

        alpha = result.getchannel("A")
        self.assertEqual(alpha.getpixel((0, 0)), 0)
        self.assertEqual(alpha.getpixel((30, 30)), 255)
        self.assertEqual(alpha.getpixel((50, 50)), 255)

    def test_global_mode_removes_enclosed_matching_color(self):
        image = Image.new("RGB", (100, 100), "white")
        draw = ImageDraw.Draw(image)
        draw.rectangle((20, 20, 80, 80), fill="black")
        draw.rectangle((40, 40, 60, 60), fill="white")

        result, _, _ = make_transparent(
            image,
            target_color=(255, 255, 255),
            tolerance=0,
            edge_only=False,
            feather=0,
        )

        self.assertEqual(result.getchannel("A").getpixel((50, 50)), 0)

    def test_transparent_png_and_webp_keep_alpha(self):
        directory = (
            Path(__file__).resolve().parent.parent / "tmp" / "transparency"
        )
        directory.mkdir(parents=True, exist_ok=True)
        image = Image.new("RGBA", (40, 30), (20, 80, 120, 0))
        ImageDraw.Draw(image).rectangle(
            (10, 8, 30, 22),
            fill=(20, 80, 120, 255),
        )
        paths = [directory / "alpha.png", directory / "alpha.webp"]
        try:
            processor = CropProcessor()
            processor._save_transparent_image(image, paths[0], "PNG", 300)
            processor._save_transparent_image(image, paths[1], "WEBP", 300)

            for path in paths:
                with Image.open(path) as saved:
                    self.assertIn("A", saved.getbands())
                    self.assertEqual(saved.getchannel("A").getextrema(), (0, 255))
        finally:
            for path in paths:
                if path.exists():
                    path.unlink()

    def test_custom_hex_color_validation(self):
        self.assertEqual(parse_hex_color("#1A2B3C"), (26, 43, 60))
        with self.assertRaises(ValueError):
            parse_hex_color("white")


if __name__ == "__main__":
    unittest.main()
