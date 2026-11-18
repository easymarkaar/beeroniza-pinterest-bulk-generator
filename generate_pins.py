#!/usr/bin/env python3
"""Generate Pinterest graphics from CSV/XLSX rows via Beeroniza's REST API.

No API keys or private configuration are embedded in this file.
Run `python generate_pins.py --help` for options.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import mimetypes
import json
import os
import posixpath
import re
import sys
import time
import zipfile
from pathlib import Path
from typing import Any
from urllib.parse import urlparse
from xml.etree import ElementTree as ET

import requests
from PIL import Image, ImageStat

ROOT = Path(__file__).resolve().parent
IMAGE_FORMATS = {"png", "jpg", "webp"}
MAX_UPLOAD_BYTES = 10 * 1024 * 1024
HEX = re.compile(r"^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$")
XML_S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
XML_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
XML_P = "{http://schemas.openxmlformats.org/package/2006/relationships}"
REQUIRED_COLUMNS = {"background_image", "hook_text"}
DEFAULT_COLOR_AREAS = {
    "hook": [0.08, 0.04, 0.92, 0.24],
    "cta": [0.10, 0.77, 0.90, 0.94],
}


class GeneratorError(RuntimeError):
    """A failure with a useful end-user message."""


def config_from_file(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise GeneratorError(
            f"No config: {path}\n"
            "Copy config.example.json to config.json and fill base_url, api_key, template_id."
        )
    try:
        obj = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        raise GeneratorError(f"Invalid config JSON: {exc}") from exc
    if not isinstance(obj, dict):
        raise GeneratorError("config.json must contain a JSON object")
    return obj


def normalize_hex(value: str, label: str) -> str:
    value = value.strip()
    if not HEX.fullmatch(value):
        raise GeneratorError(f"{label}: use a color like #FFFFFF or #263D59 (got {value!r})")
    if len(value) == 4:
        value = "#" + "".join(c * 2 for c in value[1:])
    return value.upper()


def _column_index(cell_reference: str) -> int:
    match = re.match(r"([A-Z]+)", cell_reference.upper())
    if not match:
        raise GeneratorError(f"Invalid Excel cell address: {cell_reference}")
    n = 0
    for c in match.group(1):
        n = n * 26 + ord(c) - 64
    return n - 1


def _read_xlsx(path: Path) -> list[dict[str, str]]:
    """Read values from the first .xlsx sheet using Python standard library.

    Supports normal Excel shared strings, inline strings, numbers, cached formulas.
    Does not evaluate formulas, execute macros, or interpret formatting.
    """
    try:
        with zipfile.ZipFile(path) as book:
            ns = {"m": XML_S[1:-1]}
            shared_strings: list[str] = []
            if "xl/sharedStrings.xml" in book.namelist():
                root = ET.fromstring(book.read("xl/sharedStrings.xml"))
                for item in root.findall(f"{XML_S}si"):
                    shared_strings.append("".join(t.text or "" for t in item.iter(f"{XML_S}t")))
            wb = ET.fromstring(book.read("xl/workbook.xml"))
            sheets = wb.find(f"{XML_S}sheets")
            if sheets is None or len(sheets) == 0:
                raise GeneratorError("Excel file has no sheets")
            sheet_id = sheets[0].attrib.get(f"{XML_R}id")
            if not sheet_id:
                raise GeneratorError("Could not find first Excel sheet id")
            rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
            paths = {r.attrib.get("Id"): r.attrib.get("Target") for r in rels.findall(f"{XML_P}Relationship")}
            target = paths.get(sheet_id)
            if not target:
                raise GeneratorError("Could not locate first Excel sheet")
            sheet_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            if sheet_path.startswith("xl/xl/"):
                sheet_path = sheet_path[3:]
            root = ET.fromstring(book.read(sheet_path))
            rows: list[list[str]] = []
            for row in root.iter(f"{XML_S}row"):
                cells: dict[int, str] = {}
                for cell in row.findall(f"{XML_S}c"):
                    ref = cell.attrib.get("r")
                    if not ref:
                        continue
                    typ = cell.attrib.get("t", "")
                    val = cell.find(f"{XML_S}v")
                    raw = "" if val is None or val.text is None else val.text
                    if typ == "s":
                        raw = shared_strings[int(raw)] if raw else ""
                    elif typ == "inlineStr":
                        raw = "".join(n.text or "" for n in cell.iter(f"{XML_S}t"))
                    cells[_column_index(ref)] = raw
                if cells:
                    size = max(cells) + 1
                    rows.append([cells.get(i, "") for i in range(size)])
            if not rows:
                return []
            headers = [str(h).strip() for h in rows[0]]
            return [
                {header: (str(values[i]).strip() if i < len(values) else "") for i, header in enumerate(headers) if header}
                for values in rows[1:]
                if any(str(v).strip() for v in values)
            ]
    except (zipfile.BadZipFile, ET.ParseError, KeyError, ValueError, IndexError) as exc:
        raise GeneratorError(f"Could not read Excel file {path.name}: {exc}") from exc


def load_data(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        raise GeneratorError(f"Table does not exist: {path}")
    ext = path.suffix.lower()
    if ext == ".xlsx":
        data = _read_xlsx(path)
    elif ext == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as file:
            sample = file.read(4096)
            file.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            reader = csv.DictReader(file, dialect=dialect)
            data = [
                {str(k).strip(): str(v or "").strip() for k, v in row.items() if k is not None}
                for row in reader
                if any(str(v or "").strip() for v in row.values())
            ]
    else:
        raise GeneratorError("Data file must be .xlsx (Excel) or .csv")
    if not data:
        raise GeneratorError("Table has no data rows")
    headers = set(data[0])
    missing = REQUIRED_COLUMNS - headers
    if missing:
        raise GeneratorError(f"Missing required columns: {', '.join(sorted(missing))}")
    return data


def _relative_file_path(value: str, data_dir: Path) -> Path:
    path = Path(value).expanduser()
    return path if path.is_absolute() else data_dir / path


def contrast_ratio(rgb1: tuple[int, int, int], rgb2: tuple[int, int, int]) -> float:
    def luminance(rgb: tuple[int, int, int]) -> float:
        def lin(x: int) -> float:
            c = x / 255
            return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
        return sum(weight * lin(channel) for weight, channel in zip((.2126, .7152, .0722), rgb))
    a, b = luminance(rgb1), luminance(rgb2)
    return (max(a, b) + 0.05) / (min(a, b) + 0.05)


def hex_rgb(color: str) -> tuple[int, int, int]:
    c = normalize_hex(color, "Color")[1:]
    return tuple(int(c[i:i+2], 16) for i in (0, 2, 4))


def best_contrast(background: tuple[int, int, int], dark: str, light: str) -> str:
    dark = normalize_hex(dark, "dark_text_color")
    light = normalize_hex(light, "light_text_color")
    return light if contrast_ratio(background, hex_rgb(light)) > contrast_ratio(background, hex_rgb(dark)) else dark


def image_average_color(img: Image.Image, relative_box: list[float]) -> tuple[int, int, int]:
    if len(relative_box) != 4 or any(not isinstance(v, (int, float)) or not 0 <= v <= 1 for v in relative_box):
        raise GeneratorError("Text area must be [left, top, right, bottom] with values from 0 to 1")
    w, h = img.size
    l, t, r, b = relative_box
    if l >= r or t >= b:
        raise GeneratorError("Text area must have right > left and bottom > top")
    crop = img.crop((int(w*l), int(h*t), max(int(w*r), int(w*l)+1), max(int(h*b), int(h*t)+1)))
    return tuple(round(v) for v in ImageStat.Stat(crop.convert("RGB")).mean[:3])


def image_for_analysis(session: requests.Session, value: str, data_dir: Path) -> Image.Image:
    if value.startswith("https://") or value.startswith("http://"):
        response = session.get(value, timeout=50)
        response.raise_for_status()
        body = io.BytesIO(response.content)
        return Image.open(body).convert("RGB")
    with Image.open(_relative_file_path(value, data_dir)) as img:
        return img.convert("RGB")


def choose_color(
    value: str, section: str, cfg: dict[str, Any], img: Image.Image | None,
) -> str | None:
    value = (value or "").strip()
    if not value:
        return None  # keep the text color from the Beeroniza template
    if value.lower() != "auto":
        return normalize_hex(value, section + "_color")
    if not cfg.get("auto_text_colors", True):
        raise GeneratorError("Cell value 'auto' requires auto_text_colors: true in config.json")

    dark = cfg.get("dark_text_color", "#222222")
    light = cfg.get("light_text_color", "#FFFFFF")
    # Use fixed overlay color for hook, if the template draws a static panel.
    # The background image alone would give an incorrect result here.
    overlay = cfg.get(section + "_overlay_color")
    if overlay and overlay.lower() != "none":
        bg = hex_rgb(overlay)
    elif img is not None:
        box = (cfg.get("text_areas") or {}).get(section, DEFAULT_COLOR_AREAS[section])
        bg = image_average_color(img, box)
    else:
        raise GeneratorError(f"Cannot choose '{section}_color=auto': no image or overlay color")
    return best_contrast(bg, dark, light)


def is_remote_image(value: str) -> bool:
    return value.startswith("https://") or value.startswith("http://") or value.startswith("data:image/")


def _check_image_payload(payload: bytes, source: str) -> None:
    """Fail before queuing a render if a URL returned HTML or corrupt bytes."""
    if not payload:
        raise GeneratorError(f"Empty background image: {source}")
    if len(payload) > MAX_UPLOAD_BYTES:
        raise GeneratorError(f"Background exceeds the Beeroniza 10 MB upload limit: {source}")
    try:
        with Image.open(io.BytesIO(payload)) as img:
            img.verify()
    except Exception as exc:
        raise GeneratorError(f"Background is not a valid image: {source} ({exc})") from exc


def upload_local_image(
    session: requests.Session, base: str, key: str, source: str,
    data_dir: Path, cache: dict[str, str],
) -> str:
    """Always import background into Beeroniza's media storage.

    Sending the original hosting URL straight to Beeroniza is unreliable:
    its worker silently strips remote images when the host redirects, responds
    slowly, or is rejected by the server's SSRF guard. By uploading from the
    user's computer, the renderer can read the image directly from its own DB.
    """
    if source.startswith("data:image/"):
        return source

    if source in cache:
        return cache[source]

    if source.startswith(("https://", "http://")):
        # Follow image-host redirects on the user's machine (the server's worker
        # deliberately doesn't). Never send the Beeroniza API key to this host.
        try:
            with session.get(source, timeout=60, stream=True, headers={"Accept": "image/*"}) as remote:
                remote.raise_for_status()
                chunks: list[bytes] = []
                total = 0
                for chunk in remote.iter_content(chunk_size=256 * 1024):
                    if chunk:
                        total += len(chunk)
                        if total > MAX_UPLOAD_BYTES:
                            raise GeneratorError(f"Remote background exceeds 10 MB: {source}")
                        chunks.append(chunk)
                image_bytes = b"".join(chunks)
            parsed = urlparse(source)
            remote_name = posixpath.basename(parsed.path) or "background.jpg"
            safe_name = re.sub(r"[^A-Za-z0-9._-]", "_", remote_name)[:100] or "background.jpg"
        except requests.RequestException as exc:
            raise GeneratorError(f"Could not download background from hosting: {source}: {exc}") from exc
    else:
        file = _relative_file_path(source, data_dir).resolve()
        if not file.is_file():
            raise GeneratorError(f"Background image file not found: {file}")
        if file.stat().st_size > MAX_UPLOAD_BYTES:
            raise GeneratorError(f"Background exceeds the Beeroniza 10 MB limit: {file}")
        image_bytes = file.read_bytes()
        safe_name = file.name

    _check_image_payload(image_bytes, source)
    digest = hashlib.sha256(image_bytes).hexdigest()
    if digest in cache:
        cache[source] = cache[digest]
        return cache[digest]

    mime = mimetypes.guess_type(safe_name)[0] or "application/octet-stream"
    response = session.post(
        base + "/api/v1/uploads",
        headers=api_headers(key),
        files={"file": (safe_name, io.BytesIO(image_bytes), mime)},
        timeout=120,
    )
    if response.status_code not in (200, 201):
        raise GeneratorError(f"Background upload failed (HTTP {response.status_code}): {response.text[:400]}")
    data = response.json()
    url = data.get("url")
    if not isinstance(url, str) or "/api/files/" not in url:
        raise GeneratorError(f"Beeroniza upload did not return its own media URL: {data}")

    # Confirm the media really exists and can be fetched with the same API key.
    media_url = url if url.startswith(("http://", "https://")) else base + (url if url.startswith("/") else "/" + url)
    try:
        check = session.get(media_url, headers=api_headers(key), timeout=60)
        check.raise_for_status()
        if hashlib.sha256(check.content).hexdigest() != digest:
            raise GeneratorError("Background upload verification failed: media bytes differ from original")
    except requests.RequestException as exc:
        raise GeneratorError(f"Background uploaded but cannot be read back: {exc}") from exc

    cache[source] = url
    cache[digest] = url
    print(f"    background: uploaded and verified {len(image_bytes):,} bytes (file={data.get('id', '?')})")
    return url

def build_modifications(
    row: dict[str, str], cfg: dict[str, Any], session: requests.Session,
    base: str, key: str, data_dir: Path, upload_cache: dict[str, str],
    *, dry_run: bool = False,
) -> list[dict[str, str]]:
    ph = cfg.get("placeholders") or {}
    bg_key = ph.get("background", "background-1")
    hook_key = ph.get("hook", "hook")
    cta_key = ph.get("cta", "cta")

    src = row.get("background_image", "").strip()
    if not src:
        raise GeneratorError("background_image is empty")
    if not row.get("hook_text", "").strip():
        raise GeneratorError("hook_text is empty")

    requires_sample = any(
        row.get(section + "_color", "").lower() == "auto" and not cfg.get(section + "_overlay_color")
        for section in ("hook", "cta")
    )
    img = image_for_analysis(session, src, data_dir) if requires_sample else None

    if dry_run:
        image_url = src  # no upload performed in dry run
    else:
        image_url = upload_local_image(session, base, key, src, data_dir, upload_cache)
    mods: list[dict[str, str]] = [{"name": bg_key, "image_url": image_url}]

    for section, name in (("hook", hook_key), ("cta", cta_key)):
        label = row.get(section + "_text", "").strip()
        color = choose_color(row.get(section + "_color", ""), section, cfg, img)
        if label or color:
            mod = {"name": name}
            if label:
                mod["text"] = label
            if color:
                mod["color"] = color
            mods.append(mod)
    return mods


def api_headers(api_key: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {api_key}"}


def check_templates(session: requests.Session, base: str, key: str) -> list[dict[str, Any]]:
    try:
        response = session.get(base + "/api/v1/templates", headers=api_headers(key), timeout=30)
        if response.status_code != 200:
            raise GeneratorError(f"API access failed (HTTP {response.status_code}). Check base_url and API key. {response.text[:250]}")
        return response.json().get("templates", [])
    except requests.RequestException as exc:
        raise GeneratorError(f"Cannot connect to Beeroniza: {exc}") from exc


def print_templates(templates: list[dict[str, Any]]) -> None:
    if not templates:
        print("No templates found on this Beeroniza server")
    for tpl in templates:
        print(f"\n{tpl.get('name')}  ({tpl.get('width')}x{tpl.get('height')})")
        print(f"  template_id: {tpl.get('id')}")
        print("  placeholders:")
        for item in tpl.get("placeholders", []):
            print(f"    {item.get('key')} [{item.get('type')}]")


def validate_template(tpl: dict[str, Any], cfg: dict[str, Any]) -> None:
    actual = {p.get("key"): p.get("type") for p in tpl.get("placeholders", [])}
    expected = {cfg.get("placeholders", {}).get("background", "background-1"): "image",
                cfg.get("placeholders", {}).get("hook", "hook"): "text",
                cfg.get("placeholders", {}).get("cta", "cta"): "text"}
    errors = [f"{k} [{typ}]" for k, typ in expected.items() if actual.get(k) != typ]
    if errors:
        raise GeneratorError(
            "Template placeholders don't match config.json: " + ", ".join(errors)
            + ". Check Dynamic in Beeroniza, save template, and run --list-templates."
        )


def safe_filename(name: str, fmt: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|\x00-\x1F]+", "_", (name or "pin").strip())
    name = re.sub(r"\.(png|jpg|webp)$", "", name, flags=re.I)
    return ((name or "pin")[:160] + "." + fmt)


def download_result(session: requests.Session, url: str, dest: Path, key: str) -> None:
    tmp = dest.with_name(dest.name + ".part")
    try:
        with session.get(url, headers=api_headers(key), stream=True, timeout=180, allow_redirects=True) as response:
            response.raise_for_status()
            ct = response.headers.get("content-type", "").lower()
            if "text/html" in ct or "application/json" in ct:
                snippet = response.text[:300] if hasattr(response, "text") else ""
                raise GeneratorError(f"Expected an image, got {ct}. URL={url} Response={snippet}")
            with tmp.open("wb") as stream:
                for piece in response.iter_content(chunk_size=256 * 1024):
                    if piece:
                        stream.write(piece)
        if tmp.stat().st_size == 0:
            raise GeneratorError("Downloaded image is empty")
        tmp.replace(dest)
    finally:
        if tmp.exists():
            tmp.unlink()


def warn_if_background_looks_missing(path: Path) -> bool:
    """Report suspicious output instead of treating a blank white canvas as success."""
    try:
        with Image.open(path) as img:
            image = img.convert("RGB")
            w, h = image.size
            # Sample central region, away from header and CTA. This heuristic
            # is for diagnostics; legitimate all-white designs are possible.
            total = white = 0
            for fy in (0.35, 0.42, 0.50, 0.58, 0.66, 0.74):
                for fx in (0.15, 0.25, 0.35, 0.50, 0.65, 0.75, 0.85):
                    r, g, b = image.getpixel((int(fx * (w - 1)), int(fy * (h - 1))))
                    total += 1
                    if r >= 252 and g >= 252 and b >= 252:
                        white += 1
            if white / total > 0.90:
                print("    WARNING: Most of the middle of the pin is pure white.")
                print("    Background image may be missing in Beeroniza's render.")
                print("    Check whether the 'background-1' layer is an image or a dynamic rectangle.")
                return True
    except (OSError, ValueError):
        pass
    return False


def render_one(
    session: requests.Session, base: str, key: str, template_id: str,
    mods: list[dict[str, str]], fmt: str, timeout: float, interval: float,
) -> str:
    response = session.post(
        base + "/api/v1/images",
        json={"template_id": template_id, "modifications": mods, "format": fmt},
        headers=api_headers(key), timeout=120,
    )
    if response.status_code != 202:
        raise GeneratorError(f"Image job failed (HTTP {response.status_code}): {response.text[:500]}")
    job_id = response.json().get("id")
    if not job_id:
        raise GeneratorError("Image job returned no ID")
    print(f"    job: {job_id}")
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        response = session.get(
            base + "/api/v1/images/" + str(job_id), headers=api_headers(key), timeout=40,
        )
        if response.status_code != 200:
            raise GeneratorError(f"Cannot check job status (HTTP {response.status_code}): {response.text[:300]}")
        result = response.json()
        status = str(result.get("status", "")).lower()
        if status == "completed":
            url = result.get("image_url")
            if not url:
                raise GeneratorError("Image job completed without image_url")
            return url
        if status == "failed":
            raise GeneratorError(f"Beeroniza render failed: {result.get('error') or result}")
        time.sleep(interval)
    raise GeneratorError(f"Image rendering exceeded {timeout:.0f} seconds")


def run(args: argparse.Namespace) -> int:
    config_file = Path(args.config).expanduser().resolve()
    cfg = config_from_file(config_file)
    base = str(cfg.get("base_url", "")).rstrip("/")
    key = os.environ.get("BEERONIZA_API_KEY") or str(cfg.get("api_key", ""))
    if not base.startswith(("http://", "https://")) or not key or key == "YOUR_API_KEY":
        raise GeneratorError("Fill base_url and api_key in config.json, or set BEERONIZA_API_KEY")
    session = requests.Session()
    session.headers["User-Agent"] = "Beeroniza-Pinterest-Generator/1.3"
    templates = check_templates(session, base, key)
    if args.list_templates:
        print_templates(templates)
        return 0

    template_id = args.template_id or str(cfg.get("template_id", ""))
    tpl = next((t for t in templates if t.get("id") == template_id), None)
    if not tpl:
        raise GeneratorError("Template ID not found. Run python generate_pins.py --list-templates and copy the ID into config.json")
    validate_template(tpl, cfg)

    input_name = args.input or cfg.get("input_file", "pins.xlsx")
    input_path = Path(str(input_name)).expanduser()
    if not input_path.is_absolute():
        input_path = config_file.parent / input_path
    rows = load_data(input_path)
    if args.limit > 0:
        rows = rows[:args.limit]
    output_name = args.out or cfg.get("output_dir", "generated_pins")
    out_path = Path(str(output_name)).expanduser()
    if not out_path.is_absolute():
        out_path = config_file.parent / out_path
    out_path.mkdir(parents=True, exist_ok=True)
    fmt_default = str(cfg.get("output_format", "png")).lower()
    if fmt_default not in IMAGE_FORMATS:
        raise GeneratorError(f"Unsupported output_format in config.json: {fmt_default}")

    print(f"Template: {tpl.get('name')} ({template_id})")
    print(f"Input:    {input_path} ({len(rows)} rows)")
    print(f"Output:   {out_path.resolve()}")
    if args.dry_run:
        print("DRY RUN: no images will be created")

    good = skipped = failed = 0
    cache: dict[str, str] = {}
    for index, row in enumerate(rows, 1):
        fmt = (row.get("format") or fmt_default).lower().strip()
        if fmt not in IMAGE_FORMATS:
            failed += 1
            print(f"[{index}] ERROR: unsupported format {fmt!r}")
            continue
        name = safe_filename(row.get("filename") or f"pin_{index:04d}", fmt)
        file_path = out_path / name
        print(f"\n[{index}/{len(rows)}] {name}")
        if file_path.exists() and not args.force:
            skipped += 1
            print("    already exists, skipped (use --force to regenerate)")
            continue
        try:
            mods = build_modifications(row, cfg, session, base, key, input_path.parent, cache, dry_run=args.dry_run)
            for m in mods:
                if "text" in m:
                    print(f"    {m['name']} text={m['text']!r} color={m.get('color', '(template default)')}")
            if args.dry_run:
                print("    background:", row.get("background_image"))
                good += 1
                continue
            url = render_one(session, base, key, template_id, mods, fmt,
                             float(cfg.get("job_timeout_seconds", 300)),
                             float(cfg.get("poll_interval_seconds", 2)))
            print(f"    image_url: {url}")
            download_result(session, url, file_path, key)
            print(f"    SAVED: {file_path.resolve()}")
            if warn_if_background_looks_missing(file_path):
                raise GeneratorError("Image is saved, but the background appears missing. Check the Beeroniza template. Use --force when retrying.")
            good += 1
        except (requests.RequestException, OSError, ValueError, GeneratorError) as exc:
            failed += 1
            print("    FAILED:", exc)
    print(f"\nDone. Success: {good}, skipped: {skipped}, errors: {failed}")
    return 1 if failed else 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Beeroniza Pinterest bulk image generator from Excel/CSV")
    parser.add_argument("--config", default=str(ROOT / "config.json"), help="Configuration JSON path")
    parser.add_argument("--list-templates", action="store_true", help="Show template IDs and Dynamic fields")
    parser.add_argument("--template-id", help="Override template_id from config")
    parser.add_argument("--input", help="Override input_file from config (.xlsx or .csv)")
    parser.add_argument("--out", help="Override output_dir from config")
    parser.add_argument("--limit", type=int, default=0, help="Generate first N rows (e.g., --limit 2)")
    parser.add_argument("--force", action="store_true", help="Regenerate existing files")
    parser.add_argument("--dry-run", action="store_true", help="Print planned changes; don't create images")
    args = parser.parse_args()
    try:
        return run(args)
    except (GeneratorError, requests.RequestException, OSError, ValueError) as exc:
        print("ERROR:", exc, file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("Stopped by user", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
