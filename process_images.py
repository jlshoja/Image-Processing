import sys
from pathlib import Path

from PIL import Image, ImageChops

ROOT = Path(__file__).parent / "Product images"
OUTPUT_DIR_NAME = "processed"
CANVAS = 1200
MAX_BYTES = 100 * 1024
QUALITY_START = 88
QUALITY_FLOOR = 40
TARGET_BOX = (1120, 1056)
BG_THRESHOLD = 248
SUPPORTED = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}


def bounding_box(img: Image.Image) -> tuple[int, int, int, int] | None:
    r, g, b = img.convert("RGB").split()
    min_channel = ImageChops.darker(ImageChops.darker(r, g), b)
    mask = min_channel.point(lambda v: 255 if v < BG_THRESHOLD else 0)
    return mask.getbbox()


def prepare(img: Image.Image) -> Image.Image:
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        white = Image.new("RGBA", img.size, (255, 255, 255, 255))
        img = Image.alpha_composite(white, img)
    img = img.convert("RGB")
    bbox = bounding_box(img)
    if bbox:
        img = img.crop(bbox)
    return img


def compose(img: Image.Image) -> Image.Image:
    img.thumbnail(TARGET_BOX, Image.LANCZOS)
    canvas = Image.new("RGB", (CANVAS, CANVAS), (255, 255, 255))
    top = (CANVAS - img.height) // 2
    left = (CANVAS - img.width) // 2
    canvas.paste(img, (left, top))
    return canvas


def encode(canvas: Image.Image, dest: Path) -> int:
    quality = QUALITY_START
    while quality >= QUALITY_FLOOR:
        canvas.save(dest, format="WEBP", quality=quality, method=6)
        if dest.stat().st_size <= MAX_BYTES:
            return quality
        quality -= 2
    return quality


def processed_files() -> set[str]:
    done = set()
    for dest_dir in ROOT.rglob(OUTPUT_DIR_NAME):
        for f in dest_dir.iterdir():
            if f.is_file():
                done.add(f.stem)
    return done


def process() -> None:
    done = processed_files()
    results = []
    for src in sorted(ROOT.rglob("*")):
        if not src.is_file() or src.suffix.lower() not in SUPPORTED:
            continue
        if OUTPUT_DIR_NAME in src.parts or src.stem in done:
            continue
        dest = src.parent / OUTPUT_DIR_NAME / f"{src.stem}.webp"
        dest.parent.mkdir(exist_ok=True)
        canvas = compose(prepare(Image.open(src)))
        quality = encode(canvas, dest)
        size_kb = dest.stat().st_size / 1024
        results.append(f"{src.relative_to(ROOT)} -> {dest.relative_to(ROOT)}  {size_kb:.1f} KB  q={quality}")
    if results:
        print("\n".join(results))
    print(f"\n{len(results)} new image(s) processed, {len(done)} already done.")


if __name__ == "__main__":
    sys.exit(process())
