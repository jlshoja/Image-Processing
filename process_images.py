import json
import os
import re
import shutil
import sys
import time
from pathlib import Path

import numpy as np
import openpyxl
from openpyxl import Workbook
from PIL import Image, ImageFilter
import pillow_heif

pillow_heif.register_heif_opener()

BASE = Path(__file__).parent
SETTINGS_FILE = Path(os.environ.get("IMAGE_SETTINGS", BASE / "settings.json"))


# --- settings ----------------------------------------------------------------

def load_settings(path: Path) -> dict:
    if not path.exists():
        raise SystemExit(f"settings file not found: {path}")
    with path.open(encoding="utf-8") as fh:
        data = json.load(fh)
    for section in (
        "paths",
        "files",
        "canvas",
        "encoding",
        "detection",
        "database",
        "color_map",
        "color_rules",
        "color_refs",
        "color_modifiers",
        "color_space",
    ):
        if section not in data:
            raise SystemExit(f"settings file is missing the '{section}' section: {path}")
    return data


SETTINGS = load_settings(SETTINGS_FILE)
PATHS = SETTINGS["paths"]
FILES = SETTINGS["files"]
CANVAS_CFG = SETTINGS["canvas"]
ENCODING = SETTINGS["encoding"]
DETECTION = SETTINGS["detection"]
DATABASE = SETTINGS["database"]
COLOR_MAP_CFG = SETTINGS["color_map"]
RULES = SETTINGS["color_rules"]
MODIFIERS = SETTINGS["color_modifiers"]
FAMILIES = SETTINGS["color_families"]
PRIMARY_COLORS = {c.lower() for c in FAMILIES["primary"]}
FAMILY_MODIFIERS = {m.lower() for m in FAMILIES["modifiers"]}
SYNONYMS = {k.lower(): v.lower() for k, v in FAMILIES.get("synonyms", {}).items()}
COLOR_SPACE = SETTINGS["color_space"]


def resolve_path(value: str) -> Path:
    path = Path(os.path.expandvars(str(value))).expanduser()
    return path if path.is_absolute() else (BASE / path)


ROOT = resolve_path(PATHS["source_dir"])
CONVERT_ROOT = resolve_path(PATHS["converted_dir"])
OUTPUT_ROOT = resolve_path(PATHS["output_dir"])
COLOR_REPORT = resolve_path(PATHS["color_report"])
MANIFEST = resolve_path(PATHS["manifest"])
PRODUCT_DB = resolve_path(PATHS["product_db"])
COLOR_MAP = resolve_path(PATHS["color_map"])

CANVAS = CANVAS_CFG["size"]
TARGET_BOX = tuple(CANVAS_CFG["target_box"])
OBJECT_PADDING = CANVAS_CFG["object_padding"]
MIN_FOREGROUND = CANVAS_CFG["min_foreground_ratio"]
BG_WHITE = CANVAS_CFG["background_threshold"]
BACKGROUND_RATIO_LIMIT = CANVAS_CFG["background_ratio"]
MAX_BYTES = ENCODING["max_bytes"]
QUALITY_START = ENCODING["quality_start"]
QUALITY_FLOOR = ENCODING["quality_floor"]
QUALITY_STEP = ENCODING["quality_step"]
WEBP_METHOD = ENCODING["method"]
CONVERT_MAX_SIDE = ENCODING["convert_max_side"]
BLUR_RADII = ENCODING["blur_radii"]
MODEL = DETECTION["model"]
ALPHA_MIN = DETECTION["alpha_threshold"]
BORDER_TOL = DETECTION["border_tolerance"]
SUPPORTED = {ext.lower() for ext in FILES["supported_extensions"]}
FEATURED_KEYWORDS = [k.lower() for k in FILES.get("featured_keywords", [])]
FEATURED_SUFFIX = FILES.get("featured_suffix", "main")


def hex_to_rgb(value: str) -> tuple[float, float, float]:
    text = str(value).strip().lstrip("#")
    return (
        int(text[0:2], 16),
        int(text[2:4], 16),
        int(text[4:6], 16),
    )


COLOR_REFS = {name: hex_to_rgb(value) for name, value in SETTINGS["color_refs"].items()}


# --- text helpers ------------------------------------------------------------

_PERSIAN_VARIANTS = (
    ("ي", "ی"),
    ("ك", "ک"),
    ("ة", "ه"),
    ("أ", "ا"),
    ("إ", "ا"),
    ("آ", "ا"),
)


def norm_text(text) -> str:
    s = str(text)
    for a, b in _PERSIAN_VARIANTS:
        s = s.replace(a, b)
    s = s.replace("\u200c", " ").replace("\u200f", " ").replace("\u200e", " ")
    s = s.replace("-", " ").replace("_", " ").replace("–", " ").replace("—", " ")
    s = re.sub(r"[\"'’]", "", s)
    return re.sub(r"\s+", " ", s).strip().casefold()


def slug(text) -> str:
    s = str(text).strip().lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", "-", s)
    return re.sub(r"-+", "-", s).strip("-")


def strip_count(text: str) -> str:
    return re.sub(r"[\s\u200c]*[0-9۰-۹٠-٩]+$", "", str(text)).strip()


def is_featured(src: Path) -> bool:
    name = slug(src.stem)
    if not name:
        return False
    tokens = name.split("-")
    return any(keyword in tokens for keyword in FEATURED_KEYWORDS)


# --- color detection ---------------------------------------------------------

def color_stats(img: Image.Image):
    rgb = np.asarray(img.convert("RGB"), dtype=np.float32)
    hsv = np.asarray(img.convert("HSV"), dtype=np.float32)
    h = hsv[..., 0] * 360.0 / 255.0
    s = hsv[..., 1] / 255.0
    v = hsv[..., 2] / 255.0
    keep = ~((s < RULES["background_max_s"]) & (v > RULES["background_min_v"])) & (
        v < RULES["highlight_max_v"]
    )
    if keep.sum() < RULES["min_pixels"]:
        keep = (v < RULES["fallback_max_v"]) & (
            ~((s < RULES["fallback_min_s"]) & (v > RULES["fallback_min_v"]))
        )
    return rgb, h, s, v, keep


def object_rgb(img: Image.Image) -> tuple[float, float, float]:
    rgb, _, _, _, keep = color_stats(img)
    px = rgb[keep]
    if len(px) == 0:
        px = rgb.reshape(-1, 3)
    return tuple(float(x) for x in np.median(px, axis=0))


def rgb_hex(rgb: tuple[float, float, float]) -> str:
    return "#{:02X}{:02X}{:02X}".format(*(max(0, min(255, int(round(c)))) for c in rgb))


def detect_color(img: Image.Image) -> str:
    _, h, s, v, keep = color_stats(img)
    h, s, v = h[keep], s[keep], v[keep]
    if len(h) == 0:
        return "white"
    h_mean = np.degrees(np.arctan2(np.sin(np.radians(h)).mean(), np.cos(np.radians(h)).mean())) % 360
    s_med, v_med = float(np.median(s)), float(np.median(v))
    v_dark = float(np.percentile(v, RULES["dark_percentile"]))
    if v_dark < RULES["dark_value"]:
        if s_med < RULES["black_max_s"]:
            return "black"
        if h_mean >= RULES["hue_red_high"] or h_mean < RULES["hue_brown_high"]:
            return "brown"
        if RULES["hue_brown_high"] <= h_mean < RULES["hue_green_high"]:
            return "green" if s_med >= RULES["colored_min_s"] else "black"
        if RULES["hue_green_high"] <= h_mean < RULES["hue_teal_high"]:
            return "teal" if s_med >= RULES["colored_min_s"] else "black"
        if RULES["hue_teal_high"] <= h_mean < RULES["hue_blue_high"]:
            return "navy" if s_med >= RULES["colored_min_s"] else "black"
        if RULES["hue_blue_high"] <= h_mean < RULES["hue_red_high"]:
            return "purple" if s_med >= RULES["colored_min_s"] else "black"
    if s_med < RULES["gray_max_s"]:
        return "gray" if v_med < RULES["gray_max_v"] else "white"
    if (h_mean >= RULES["hue_red_high"] or h_mean < RULES["hue_red_low"]) and s_med > RULES["red_min_s"]:
        return "red"
    if RULES["hue_red_low"] <= h_mean < RULES["hue_brown_high"] and s_med >= RULES["warm_min_s"]:
        if v_med < RULES["brown_max_v"]:
            return "brown"
        return "orange" if s_med > RULES["orange_min_s"] else "beige"
    if RULES["hue_brown_high"] <= h_mean < RULES["hue_yellow_high"]:
        return "yellow"
    if RULES["hue_yellow_high"] <= h_mean < RULES["hue_green_high"]:
        return "green"
    if RULES["hue_green_high"] <= h_mean < RULES["hue_teal_high"]:
        return "teal"
    if RULES["hue_teal_high"] <= h_mean < RULES["hue_blue_high"]:
        return "navy" if v_med < RULES["navy_max_v"] else "blue"
    if RULES["hue_blue_high"] <= h_mean < RULES["hue_purple_high"]:
        return "purple"
    if RULES["hue_purple_high"] <= h_mean < RULES["hue_pink_high"]:
        return "pink"
    return "red" if s_med > RULES["red_min_s"] else "brown"


# --- allowed colors from the product database --------------------------------

def canonical_color(name: str) -> str:
    text = str(name).strip().lower()
    text = SYNONYMS.get(text, text)
    tokens = [SYNONYMS.get(t, t) for t in text.split("-") if t]
    idx = next((i for i, t in enumerate(tokens) if t in PRIMARY_COLORS), None)
    if idx is None:
        idx = next((i for i, t in enumerate(tokens) if t in COLOR_REFS), None)
        if idx is None:
            return text
    prefix = []
    for token in reversed(tokens[:idx]):
        if token in FAMILY_MODIFIERS:
            prefix.insert(0, token)
        else:
            break
    return "-".join(prefix + [tokens[idx]])


def color_ref(name: str) -> tuple[float, float, float] | None:
    tokens = [t for t in name.split("-") if t]
    refs = [COLOR_REFS[t] for t in tokens if t in COLOR_REFS]
    if not refs:
        return None
    r, g, b = (float(x) for x in refs[0])
    for token, rule in MODIFIERS.items():
        if token in tokens:
            mix = hex_to_rgb(rule["mix"])
            amount = rule["amount"]
            r = r + (mix[0] - r) * amount
            g = g + (mix[1] - g) * amount
            b = b + (mix[2] - b) * amount
    return (r, g, b)


def _to_lab(rgb: tuple[float, float, float]) -> np.ndarray:
    c = np.array(rgb, dtype=np.float32) / 255.0
    c = np.where(c <= 0.04045, c / 12.92, ((c + 0.055) / 1.055) ** 2.4)
    m = np.array(COLOR_SPACE["rgb_to_xyz"], dtype=np.float32)
    xyz = m @ c / np.array(COLOR_SPACE["white_point"], dtype=np.float32)
    f = np.where(xyz > 0.008856, np.cbrt(xyz), 7.787 * xyz + 16 / 116)
    return np.array(
        [116 * f[1] - 16, 500 * (f[0] - f[1]), 200 * (f[1] - f[2])], dtype=np.float32
    )


def nearest_allowed(
    rgb: tuple[float, float, float],
    allowed: list[tuple[str, tuple | None]],
) -> str | None:
    lab = _to_lab(rgb)
    best_name, best_dist = None, None
    for name, ref in allowed:
        if ref is None:
            continue
        dist = float(np.linalg.norm(lab - _to_lab(ref)))
        if best_dist is None or dist < best_dist:
            best_name, best_dist = name, dist
    return best_name


def pick_color(
    img: Image.Image,
    allowed: list[tuple[str, tuple | None]],
) -> tuple[str, str, bool, str]:
    detected = detect_color(img)
    hexcode = rgb_hex(object_rgb(img))
    if allowed:
        for name, _ in allowed:
            if detected == name or detected in name.split("-"):
                return name, detected, True, hexcode
        match = nearest_allowed(object_rgb(img), allowed)
        if match:
            names = ", ".join(name for name, _ in allowed)
            print(f"    color: detected {detected} -> {match} (allowed: {names})")
            return match, detected, True, hexcode
    return detected, detected, False, hexcode


# --- database ----------------------------------------------------------------

def _sheet(workbook, cfg: dict):
    name = cfg.get("sheet")
    if name in (None, ""):
        return workbook[workbook.sheetnames[0]]
    return workbook[name]


def _column_index(header_values, target: str | None, fallback: int) -> int:
    if target:
        wanted = norm_text(target)
        for i, value in enumerate(header_values or ()):
            if value and norm_text(value) == wanted:
                return i
    return fallback - 1


def load_color_map() -> dict[str, str]:
    if not COLOR_MAP.exists():
        print(f"warning: {COLOR_MAP.name} not found, colors are auto-detected only")
        return {}
    wb = openpyxl.load_workbook(COLOR_MAP, read_only=True, data_only=True)
    try:
        ws = _sheet(wb, COLOR_MAP_CFG)
        header_row = COLOR_MAP_CFG["header_row"]
        header = list(
            ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True)
        )
        persian_col = _column_index(
            header[0] if header else (),
            COLOR_MAP_CFG.get("persian_header"),
            COLOR_MAP_CFG["persian_column"],
        )
        english_col = _column_index(
            header[0] if header else (),
            COLOR_MAP_CFG.get("english_header"),
            COLOR_MAP_CFG["english_column"],
        )
        mapping = {}
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            persian = row[persian_col] if len(row) > persian_col else None
            english = row[english_col] if len(row) > english_col else None
            if persian and english:
                mapping.setdefault(norm_text(persian), slug(english))
        return mapping
    finally:
        wb.close()


def resolve_color(raw, color_map: dict[str, str]) -> str | None:
    text = norm_text(raw)
    if not text:
        return None
    tries = [text]
    stripped = norm_text(strip_count(raw))
    if stripped and stripped != text:
        tries.append(stripped)
    variants = []
    for t in tries:
        variants.append(t)
        words = t.split()
        if len(words[-1]) > 2 and words[-1].endswith("ی"):
            variants.append(" ".join(words[:-1] + [words[-1][:-1]]))
        if t.endswith("سورمه"):
            variants.append(t + " ای")
        if "سورمه" in t:
            variants.append(t.replace("سورمه", "سرمه"))
        if t.endswith("نسکافه"):
            variants.append(t + " ای")
    for v in variants:
        if v in color_map:
            return color_map[v]
    best = None
    for key, english in color_map.items():
        if key and key in text and (best is None or len(key) > len(best[0])):
            best = (key, english)
    return best[1] if best else None


def _product_rows():
    if not PRODUCT_DB.exists():
        print(f"warning: {PRODUCT_DB.name} not found, colors are auto-detected only")
        return
    wb = openpyxl.load_workbook(PRODUCT_DB, read_only=True, data_only=True)
    try:
        ws = _sheet(wb, DATABASE)
        header_row = DATABASE["header_row"]
        header = list(
            ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True)
        )
        code_col = _column_index(
            header[0] if header else (),
            DATABASE.get("code_header"),
            DATABASE["code_column"],
        )
        color_col = _column_index(
            header[0] if header else (),
            DATABASE.get("color_header"),
            DATABASE["color_column"],
        )
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            code = row[code_col] if len(row) > code_col else None
            raw = row[color_col] if len(row) > color_col else None
            if code and raw:
                yield re.sub(r"\s+", "", str(code)).upper(), str(raw)
    finally:
        wb.close()


def load_products() -> dict[str, list[tuple[str, tuple | None]]]:
    color_map = load_color_map()
    products: dict[str, list[tuple[str, tuple | None]]] = {}
    for code, raw in _product_rows() or ():
        allowed = products.setdefault(code, [])
        for segment in re.split(r"[|,،;؛\n]", raw):
            english = resolve_color(segment, color_map)
            if not english:
                continue
            english = canonical_color(english)
            if english not in {name for name, _ in allowed}:
                allowed.append((english, color_ref(english)))
    return products


def find_missing_colors() -> list[tuple[str, str]]:
    color_map = load_color_map()
    missing = []
    seen = set()
    for code, raw in _product_rows() or ():
        for segment in re.split(r"[|,،;؛\n]", raw):
            if resolve_color(segment, color_map) is not None:
                continue
            text = norm_text(segment)
            if text and (code, text) not in seen:
                seen.add((code, text))
                missing.append((code, segment.strip()))
    return missing


# --- image processing --------------------------------------------------------

def prepare(img: Image.Image) -> Image.Image:
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        white = Image.new("RGBA", img.size, (255, 255, 255, 255))
        img = Image.alpha_composite(white, img)
    return img.convert("RGB")


def _finish(
    img: Image.Image,
    mask: np.ndarray,
    cut: Image.Image | None,
    allowed: list[tuple[str, tuple | None]],
) -> tuple[Image.Image, str, bool, str, str]:
    ys, xs = np.where(mask)
    if len(xs) == 0:
        raise ValueError("no object detected")
    h, w = mask.shape
    coverage = float(mask.mean())
    bbox = (int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1)
    touches_border = (
        bbox[0] <= BORDER_TOL
        or bbox[1] <= BORDER_TOL
        or bbox[2] >= w - BORDER_TOL
        or bbox[3] >= h - BORDER_TOL
    )
    unreliable = coverage < MIN_FOREGROUND
    if touches_border or unreliable:
        side = min(img.size)
        x0 = (img.width - side) // 2
        y0 = (img.height - side) // 2
        crop = img.crop((x0, y0, x0 + side, y0 + side)).resize(
            (CANVAS, CANVAS), Image.LANCZOS
        )
        color_img = img
        color, detected, _, hexcode = pick_color(color_img, allowed)
        return crop, color, True, detected, hexcode
    obj_w, obj_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pad = int(OBJECT_PADDING * max(obj_w, obj_h))
    x0 = max(0, bbox[0] - pad)
    y0 = max(0, bbox[1] - pad)
    x1 = min(img.width, bbox[2] + pad)
    y1 = min(img.height, bbox[3] + pad)
    crop = img.crop((x0, y0, x1, y1))
    crop.thumbnail(TARGET_BOX, Image.LANCZOS)
    canvas = Image.new("RGB", (CANVAS, CANVAS), (255, 255, 255))
    canvas.paste(crop, ((CANVAS - crop.width) // 2, (CANVAS - crop.height) // 2))
    if cut is None:
        color_img = img.crop((x0, y0, x1, y1))
    else:
        color_img = Image.alpha_composite(
            Image.new("RGBA", (x1 - x0, y1 - y0), (255, 255, 255, 255)), cut.crop((x0, y0, x1, y1))
        ).convert("RGB")
    color, detected, _, hexcode = pick_color(color_img, allowed)
    return canvas, color, False, detected, hexcode


def compose(
    img: Image.Image,
    ctx: dict | None,
    allowed: list[tuple[str, tuple | None]] | None = None,
) -> tuple[Image.Image, str, bool, str, str]:
    img = prepare(img)
    rgb = np.asarray(img, dtype=np.float32)
    foreground = rgb.max(axis=2) < BG_WHITE
    allowed = allowed or []
    if float(foreground.mean()) < BACKGROUND_RATIO_LIMIT:
        # busy background: the threshold mask is not trustworthy, use the AI model
        if ctx is None:
            ctx = {}
        cut = remove(img.convert("RGBA"), session=_get_session(ctx))
        mask = np.asarray(cut.getchannel("A")) > ALPHA_MIN
        return _finish(img, mask, cut, allowed)
    return _finish(img, foreground, None, allowed)


def encode(canvas: Image.Image, dest: Path, quality: int | None = None) -> int:
    quality = quality or QUALITY_START
    canvas.save(dest, format="WEBP", quality=quality, method=WEBP_METHOD)
    return quality


def compress(
    img: Image.Image,
    dest: Path,
    max_bytes: int = MAX_BYTES,
    allow_blur: bool = True,
) -> int:
    def encode_at(im: Image.Image, quality: int) -> int:
        im.save(dest, format="WEBP", quality=quality, method=WEBP_METHOD)
        return dest.stat().st_size

    # one save is enough in the common case: start quality already fits
    if encode_at(img, QUALITY_START) <= max_bytes:
        return QUALITY_START

    def search(im: Image.Image) -> int | None:
        lo, hi = QUALITY_FLOOR, QUALITY_START - QUALITY_STEP
        best = None
        while lo <= hi:
            mid = (lo + hi) // 2
            size = encode_at(im, mid)
            if size <= max_bytes:
                best = mid
                lo = mid + QUALITY_STEP
            else:
                hi = mid - QUALITY_STEP
        return best

    quality = search(img)
    if quality is not None:
        return quality
    if allow_blur:
        for radius in BLUR_RADII:
            quality = search(img.filter(ImageFilter.GaussianBlur(radius)))
            if quality is not None:
                return quality
    return QUALITY_FLOOR - QUALITY_STEP


# --- step 1: convert to webp under the size limit ----------------------------

def convert_file(src: Path, dest: Path) -> tuple[int, str, float]:
    started = time.perf_counter()
    if src.suffix.lower() == ".webp" and src.stat().st_size <= MAX_BYTES:
        shutil.copyfile(src, dest)
        return QUALITY_START, "copied", time.perf_counter() - started
    img = prepare(Image.open(src))
    if CONVERT_MAX_SIDE and max(img.size) > CONVERT_MAX_SIDE:
        img.thumbnail((CONVERT_MAX_SIDE, CONVERT_MAX_SIDE), Image.LANCZOS)
    quality = compress(img, dest, MAX_BYTES)
    return quality, "compressed", time.perf_counter() - started


def convert(sources: list[Path]) -> list[Path]:
    converted = []
    total = len(sources)
    started = time.perf_counter()
    print(
        f"found {total} source file(s) in {ROOT}  ->  {CONVERT_ROOT}", flush=True
    )
    for i, src in enumerate(sources, 1):
        pct = i / total * 100 if total else 100
        rel = src.relative_to(ROOT)
        dest = (CONVERT_ROOT / rel.parent / f"{src.stem}.webp").resolve()
        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.exists():
            converted.append(dest)
            continue
        quality, mode, seconds = convert_file(src, dest)
        converted.append(dest)
        kb = dest.stat().st_size / 1024
        print(
            f"[{i}/{total} {pct:.0f}%] convert: {rel} -> "
            f"converted/{rel.parent / (src.stem + '.webp')}  {kb:.1f} KB  "
            f"q={quality}  {mode}  {seconds:.1f} s"
        )
    print(f"[100%] convert: {total} file(s) checked.\n")
    return converted


# --- manifest & report -------------------------------------------------------

def load_manifest() -> dict:
    path = OUTPUT_ROOT / MANIFEST.name
    if path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
        return {
            key: ({"output": value} if isinstance(value, str) else value)
            for key, value in data.items()
        }
    return {}


def save_manifest(manifest: dict) -> None:
    OUTPUT_ROOT.mkdir(exist_ok=True)
    (OUTPUT_ROOT / MANIFEST.name).write_text(
        json.dumps(manifest, indent=2), encoding="utf-8"
    )


def source_key(src: Path) -> str:
    try:
        return src.relative_to(ROOT).as_posix()
    except ValueError:
        return src.relative_to(BASE).as_posix()


def write_color_report(records: list[dict], missing: list[tuple[str, str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "detected"
    ws.append(
        [
            "source",
            "product code",
            "output file",
            "color (used)",
            "detected",
            "hex",
            "framing",
            "featured",
            "from product list",
            "allowed colors",
        ]
    )
    for rec in records:
        ws.append(
            [
                rec["source"],
                rec["code"],
                rec["output"],
                rec["color"],
                rec["detected"],
                rec["hex"],
                rec["framing"],
                "yes" if rec["featured"] else "no",
                "yes" if rec["matched"] else "no",
                rec["allowed"],
            ]
        )
    ws2 = wb.create_sheet("missing")
    ws2.append(["product code", "color", "problem"])
    for code, raw, problem in missing:
        ws2.append([code, raw, problem])
    for sheet in (ws, ws2):
        for column in sheet.columns:
            width = max(len(str(cell.value or "")) for cell in column) + 2
            sheet.column_dimensions[column[0].column_letter].width = min(width, 60)
    COLOR_REPORT.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(COLOR_REPORT)
    except PermissionError:
        print(
            f"warning: could not write {COLOR_REPORT.name} "
            "(close it in Excel and run again to update it)"
        )
        return False
    return True


# --- step 2: process ---------------------------------------------------------

def process(sources: list[Path], ctx: dict | None = None) -> None:
    ctx = ctx if ctx is not None else {}
    manifest = load_manifest()
    existing_stems = {
        Path(entry["output"]).stem
        for entry in manifest.values()
        if isinstance(entry, dict) and entry.get("output")
    }
    pending = [src for src in sorted(sources) if source_key(src) not in manifest]
    print(
        f"found {len(sources)} source file(s), "
        f"{len(pending)} new image(s) to process "
        f"({len(sources) - len(pending)} already in manifest)", flush=True
    )
    if not pending:
        print("process: no new images.")
        records = [
            dict(entry, source=key)
            for key, entry in manifest.items()
            if isinstance(entry, dict) and entry.get("output")
        ]
        if records:
            missing = [
                (code, raw, "no entry in color mapping")
                for code, raw in find_missing_colors()
            ]
            write_color_report(records, missing)
            print(
                f"color report: {COLOR_REPORT.name} "
                f"({len(records)} image(s), {len(missing)} missing color(s))"
            )
        return
    print("reading product database and color mapping...", flush=True)
    products = load_products()
    total = len(pending)
    announced = set()
    records = []
    no_colors = []
    for i, src in enumerate(pending, 1):
        pct = i / total * 100 if total else 100
        folder = src.parent.name
        code = re.sub(r"\s+", "", folder).upper()
        allowed = products.get(code, [])
        if allowed and folder not in announced:
            names = ", ".join(name for name, _ in allowed)
            print(f"colors for {folder} (from product list): {names}")
            announced.add(folder)
        if not allowed and folder not in announced:
            reason = "color list is empty" if code in products else "code not found in database"
            no_colors.append((folder, "", reason))
            announced.add(folder)
        canvas, color, is_detail, detected, hexcode = compose(
            Image.open(src), ctx, allowed
        )
        seq = 1
        if is_featured(src):
            out = f"{folder}-{color}-{FEATURED_SUFFIX}"
            while out in existing_stems or (OUTPUT_ROOT / folder / f"{out}.webp").exists():
                out = f"{folder}-{color}-{FEATURED_SUFFIX}-{seq}"
                seq += 1
        else:
            while (out := f"{folder}-{color}-{seq:02d}") in existing_stems or (
                OUTPUT_ROOT / folder / f"{out}.webp"
            ).exists():
                seq += 1
        dest = OUTPUT_ROOT / folder / f"{out}.webp"
        dest.parent.mkdir(parents=True, exist_ok=True)
        quality = encode(canvas, dest)
        manifest[source_key(src)] = {
            "output": dest.relative_to(OUTPUT_ROOT).as_posix(),
            "code": folder,
            "color": color,
            "detected": detected,
            "hex": hexcode,
            "framing": "detail" if is_detail else "full",
            "featured": is_featured(src),
            "matched": color in {name for name, _ in allowed},
            "allowed": ", ".join(name for name, _ in allowed),
        }
        existing_stems.add(dest.stem)
        kb = dest.stat().st_size / 1024
        records.append(dict(manifest[source_key(src)], source=source_key(src)))
        print(
            f"[{i}/{total} {pct:.0f}%] process: {source_key(src)} -> "
            f"processed/{dest.relative_to(OUTPUT_ROOT)}  {kb:.1f} KB  q={quality}"
        )
    missing = [(code, raw, "no entry in color mapping") for code, raw in find_missing_colors()]
    missing += no_colors
    write_color_report(records, missing)
    print(
        f"color report: {COLOR_REPORT.name} "
        f"({len(records)} image(s), {len(missing)} missing color(s))"
    )
    print("[100%] process: all images done.")
    save_manifest(manifest)


def _get_session(ctx: dict):
    if ctx.get("session") is None:
        ctx["session"] = _new_session(MODEL)
    return ctx["session"]


def _new_session(model: str):
    from rembg import new_session

    started = time.perf_counter()
    print(f"busy background detected, loading AI model '{model}'...")
    session = new_session(model)
    print(f"model ready in {time.perf_counter() - started:.1f} s")
    return session


def remove(**kwargs):
    from rembg import remove as _remove

    return _remove(**kwargs)


# --- runner ------------------------------------------------------------------

def find_sources() -> list[Path]:
    return [
        p
        for p in sorted(ROOT.rglob("*"))
        if p.is_file()
        and p.suffix.lower() in SUPPORTED
        and not set(p.parts).intersection({CONVERT_ROOT.name, OUTPUT_ROOT.name})
    ]


def run_convert() -> list[Path]:
    return convert(find_sources())


def run_process() -> None:
    process(find_sources())


def run_full() -> None:
    print("Step 1/2: convert")
    converted = run_convert()
    print("Step 2/2: process")
    process(converted)


def main() -> None:
    mode = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    if mode not in ("convert", "process", "full"):
        print("Usage: python process_images.py [convert|process|full]")
        print(f"  settings: {SETTINGS_FILE}")
        print(f"  convert - step 1 only: convert all sources to webp under {MAX_BYTES // 1024} KB, keeping original names/folders")
        print(f"  process - step 2 only: center, white bg, {CANVAS}x{CANVAS}, color name (no size limit)")
        print("            (color is matched against the product color list in the database)")
        print("  full    - step 1 + step 2 (converts, then processes the converted files)")
        return
    print(f"settings: {SETTINGS_FILE}", flush=True)
    if mode == "convert":
        n = len(run_convert())
        print(f"\n{n} file(s) ensured as webp in {CONVERT_ROOT.name}/.")
    elif mode == "process":
        run_process()
    else:
        run_full()
        print(f"\nFull pipeline complete: outputs in {OUTPUT_ROOT.name}/<folder>/.")


if __name__ == "__main__":
    main()
