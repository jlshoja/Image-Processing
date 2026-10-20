import json
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageChops, ImageFilter
import pillow_heif
from rembg import remove, new_session

pillow_heif.register_heif_opener()

ROOT = Path(__file__).parent / "Product images"
CONVERT_ROOT = Path(__file__).parent / "converted"
OUTPUT_ROOT = Path(__file__).parent / "processed"
CANVAS = 1200
MAX_BYTES = 100 * 1024
DETAIL_MAX_BYTES = 130 * 1024
QUALITY_START = 88
QUALITY_FLOOR = 30
TARGET_BOX = (1120, 1056)
DETAIL_BOX = (CANVAS, CANVAS)
BG_THRESHOLD = 248
SUPPORTED = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".heic", ".heif"}


# --- step 1: convert to webp -------------------------------------------------

def convert_file(src: Path, dest: Path) -> None:
    img = Image.open(src)
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        white = Image.new("RGBA", img.size, (255, 255, 255, 255))
        img = Image.alpha_composite(white, img)
    img.convert("RGB").save(dest, format="WEBP", quality=95, method=6)


def convert(sources: list[Path]) -> list[Path]:
    converted = []
    total = len(sources)
    for i, src in enumerate(sources, 1):
        pct = i / total * 100 if total else 100
        if src.suffix.lower() == ".webp":
            converted.append(src)
            continue
        rel = src.relative_to(ROOT)
        dest = (CONVERT_ROOT / rel.parent / f"{src.stem}.webp").resolve()
        dest.parent.mkdir(parents=True, exist_ok=True)
        if not dest.exists():
            convert_file(src, dest)
            kb = dest.stat().st_size / 1024
            print(f"[{i}/{total} {pct:.0f}%] convert: {rel} -> converted/{rel.parent / (src.stem + '.webp')}  {kb:.1f} KB")
        converted.append(dest)
    print(f"[100%] convert: {total} file(s) checked.\n")
    return converted


# --- step 2: process (center, white canvas, rename, compress) ---------------

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


def detect_color(img: Image.Image) -> str:
    hsv = np.asarray(img.convert("HSV"), dtype=np.float32)
    h = hsv[..., 0] * 360.0 / 255.0
    s = hsv[..., 1] / 255.0
    v = hsv[..., 2] / 255.0
    keep = ~((s < 0.15) & (v > 0.55)) & (v < 0.97)
    if keep.sum() < 2000:
        keep = (v < 0.90) & (~((s < 0.10) & (v > 0.60)))
    h, s, v = h[keep], s[keep], v[keep]
    if len(h) == 0:
        return "white"
    h_mean = np.degrees(np.arctan2(np.sin(np.radians(h)).mean(), np.cos(np.radians(h)).mean())) % 360
    s_med, v_med = float(np.median(s)), float(np.median(v))
    v_dark = float(np.percentile(v, 25))
    if v_dark < 0.35:
        if s_med < 0.13:
            return "black"
        if h_mean >= 340 or h_mean < 55:
            return "brown"
        if 55 <= h_mean < 165:
            return "green" if s_med >= 0.30 else "black"
        if 165 <= h_mean < 200:
            return "teal" if s_med >= 0.30 else "black"
        if 200 <= h_mean < 260:
            return "navy" if s_med >= 0.30 else "black"
        if 260 <= h_mean < 340:
            return "purple" if s_med >= 0.30 else "black"
    if s_med < 0.14:
        return "gray" if v_med < 0.82 else "white"
    if (h_mean >= 340 or h_mean < 12) and s_med > 0.40:
        return "red"
    if 10 <= h_mean < 55 and s_med >= 0.15:
        if v_med < 0.62:
            return "brown"
        return "orange" if s_med > 0.55 else "beige"
    if 55 <= h_mean < 75:
        return "yellow"
    if 75 <= h_mean < 165:
        return "green"
    if 165 <= h_mean < 200:
        return "teal"
    if 200 <= h_mean < 260:
        return "navy" if v_med < 0.35 else "blue"
    if 260 <= h_mean < 300:
        return "purple"
    if 300 <= h_mean < 340:
        return "pink"
    return "red" if s_med > 0.40 else "brown"


def dense_bbox(alpha: np.ndarray) -> tuple[int, int, int, int]:
    cols = (alpha > 64).sum(axis=0)
    rows = (alpha > 64).sum(axis=1)
    xs = np.where(cols >= cols.max() * 0.2)[0]
    ys = np.where(rows >= rows.max() * 0.2)[0]
    return int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1


def compose(img: Image.Image, session) -> tuple[Image.Image, str, bool]:
    img = prepare(img)
    cut = remove(img.convert("RGBA"), session=session)
    alpha = np.asarray(cut.getchannel("A"))
    ys, xs = np.where(alpha > 64)
    if len(xs) == 0:
        raise ValueError("no object detected")
    h, w = alpha.shape
    bbox = (int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1)
    touches_border = (
        bbox[0] <= 1 or bbox[1] <= 1 or bbox[2] >= w - 1 or bbox[3] >= h - 1
    )
    if touches_border:
        side = min(img.size)
        x0 = (img.width - side) // 2
        y0 = (img.height - side) // 2
        crop = img.crop((x0, y0, x0 + side, y0 + side)).resize(
            (CANVAS, CANVAS), Image.LANCZOS
        )
        color_img = Image.alpha_composite(
            Image.new("RGBA", img.size, (255, 255, 255, 255)), cut
        ).convert("RGB")
        return crop, detect_color(color_img), True
    obj_w, obj_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pad = int(0.05 * max(obj_w, obj_h))
    x0 = max(0, bbox[0] - pad)
    y0 = max(0, bbox[1] - pad)
    x1 = min(img.width, bbox[2] + pad)
    y1 = min(img.height, bbox[3] + pad)
    crop = img.crop((x0, y0, x1, y1))
    crop.thumbnail(TARGET_BOX, Image.LANCZOS)
    canvas = Image.new("RGB", (CANVAS, CANVAS), (255, 255, 255))
    canvas.paste(crop, ((CANVAS - crop.width) // 2, (CANVAS - crop.height) // 2))
    color_img = Image.alpha_composite(
        Image.new("RGBA", (x1 - x0, y1 - y0), (255, 255, 255, 255)), cut.crop((x0, y0, x1, y1))
    ).convert("RGB")
    return canvas, detect_color(color_img), False


def encode(canvas: Image.Image, dest: Path, allow_blur: bool = True, max_bytes: int = MAX_BYTES) -> int:
    def attempt(im: Image.Image, start: int) -> int | None:
        quality = start
        while quality >= QUALITY_FLOOR:
            im.save(dest, format="WEBP", quality=quality, method=6)
            if dest.stat().st_size <= max_bytes:
                return quality
            quality -= 2
        return None

    quality = attempt(canvas, QUALITY_START)
    if quality is not None:
        return quality
    if allow_blur:
        for radius in (0.8, 1.1, 1.4, 1.7):
            quality = attempt(canvas.filter(ImageFilter.GaussianBlur(radius)), 40)
            if quality is not None:
                return quality
    return QUALITY_FLOOR - 2


def load_manifest() -> dict:
    path = OUTPUT_ROOT / "manifest.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def save_manifest(manifest: dict) -> None:
    OUTPUT_ROOT.mkdir(exist_ok=True)
    (OUTPUT_ROOT / "manifest.json").write_text(
        json.dumps(manifest, indent=2), encoding="utf-8"
    )


def source_key(src: Path) -> str:
    try:
        return src.relative_to(ROOT).as_posix()
    except ValueError:
        return src.relative_to(Path(__file__).parent).as_posix()


def process(sources: list[Path], session=None) -> None:
    manifest = load_manifest()
    existing_names = set(manifest.values())
    pending = []
    for src in sorted(sources):
        key = source_key(src)
        if key not in manifest:
            pending.append(src)
    if not pending:
        print("process: no new images.")
        return
    session = session or new_session("u2net")
    total = len(pending)
    for i, src in enumerate(pending, 1):
        pct = i / total * 100 if total else 100
        folder = src.parent.name
        canvas, color, is_detail = compose(Image.open(src), session)
        seq = 1
        while (out := f"{folder}-{color}-{seq:02d}") in existing_names or (
            OUTPUT_ROOT / folder / f"{out}.webp"
        ).exists():
            seq += 1
        dest = OUTPUT_ROOT / folder / f"{out}.webp"
        dest.parent.mkdir(parents=True, exist_ok=True)
        quality = encode(
            canvas,
            dest,
            allow_blur=not is_detail,
            max_bytes=DETAIL_MAX_BYTES if is_detail else MAX_BYTES,
        )
        manifest[source_key(src)] = dest.relative_to(OUTPUT_ROOT).as_posix()
        existing_names.add(dest.stem)
        kb = dest.stat().st_size / 1024
        print(f"[{i}/{total} {pct:.0f}%] process: {source_key(src)} -> processed/{dest.relative_to(OUTPUT_ROOT)}  {kb:.1f} KB  q={quality}")
    print("[100%] process: all images done.")
    save_manifest(manifest)


# --- runner ------------------------------------------------------------------

def find_sources() -> list[Path]:
    return [
        p for p in sorted(ROOT.rglob("*"))
        if p.is_file()
        and p.suffix.lower() in SUPPORTED
        and not set(p.parts).intersection({CONVERT_ROOT.name, OUTPUT_ROOT.name})
    ]


def run_convert() -> list[Path]:
    return convert(find_sources())


def run_full() -> None:
    print("Step 1/2: convert")
    converted = run_convert()
    print("Step 2/2: process")
    process(converted)


def run_process() -> None:
    process(find_sources())


def main() -> None:
    mode = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    if mode not in ("convert", "process", "full"):
        print("Usage: python process_images.py [convert|process|full]")
        print("  convert - step 1 only: convert all sources to webp -> converted/<folder>/")
        print("  process - step 2 only: center, white bg, 1200x1200, <=100KB, color name")
        print("            (sources are processed directly, no convert step)")
        print("  full    - step 1 + step 2 (converts, then processes the converted files)")
        return
    if mode == "convert":
        n = len(run_convert())
        print(f"\n{n} file(s) ensured as webp in converted/.")
    elif mode == "process":
        run_process()
    else:
        run_full()
        print("\nFull pipeline complete: outputs in processed/<folder>.")


if __name__ == "__main__":
    main()
