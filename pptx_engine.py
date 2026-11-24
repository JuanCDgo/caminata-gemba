"""Engine for reading, appending, and updating Caminata Gemba.pptx."""

from __future__ import annotations

import io
import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from PIL import Image, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

SLIDE_W = 12192000
SLIDE_H = 6858000

# Bounding boxes matching existing Caminata Gemba.pptx slides
COVER_BOX = (152400, 757937, 11887200, 5342117)
ANTES_BOX = (152400, 1603275, 5784600, 4901025)
DESPUES_BOX = (6255000, 1603275, 5784600, 4901025)

ACCENT_BLUE = RGBColor(0x05, 0x8D, 0xC7)
BORDER_NAVY = RGBColor(0x1C, 0x30, 0x52)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT = RGBColor(0x0F, 0x17, 0x2A)
NOTE_BG = RGBColor(0xF1, 0xF5, 0xF9)

MAX_PHOTO_PX = 1400
JPEG_QUALITY = 82


def backup_pptx(pptx_path: Path) -> Path:
    """Create a timestamped backup before modifying the presentation."""
    backups_dir = pptx_path.parent / "backups"
    backups_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = backups_dir / f"{pptx_path.stem}_{stamp}{pptx_path.suffix}"
    shutil.copy2(pptx_path, dest)
    # Keep last 15 backups so disk doesn't fill up
    existing = sorted(backups_dir.glob(f"{pptx_path.stem}_*{pptx_path.suffix}"))
    for old in existing[:-15]:
        old.unlink(missing_ok=True)
    return dest


def optimize_photo(raw_bytes: bytes, max_px: int = MAX_PHOTO_PX) -> tuple[bytes, int, int]:
    """Fix EXIF orientation, resize to max_px, and compress to JPEG."""
    with Image.open(io.BytesIO(raw_bytes)) as img:
        img = ImageOps.exif_transpose(img)
        if img.mode in ("RGBA", "LA", "P"):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            bg.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")

        img.thumbnail((max_px, max_px), Image.Resampling.LANCZOS)
        out = io.BytesIO()
        img.save(out, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        return out.getvalue(), img.width, img.height


def fit_box(img_w: int, img_h: int, box: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    """Fit an image of (img_w, img_h) centered inside box=(x, y, cx, cy)."""
    bx, by, bcx, bcy = box
    img_ratio = img_w / img_h
    box_ratio = bcx / bcy
    if img_ratio > box_ratio:
        new_w = bcx
        new_h = int(bcx / img_ratio)
    else:
        new_h = bcy
        new_w = int(bcy * img_ratio)
    new_x = bx + (bcx - new_w) // 2
    new_y = by + (bcy - new_h) // 2
    return new_x, new_y, new_w, new_h


def _get_blank_layout(prs: Presentation):
    """Return the slide layout used by existing slides (layout 0 in this deck)."""
    return prs.slide_layouts[0]


def _slide_texts(slide) -> list[str]:
    texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            txt = shape.text_frame.text.strip()
            if txt:
                texts.append(txt)
    return texts


def _slide_pictures(slide) -> list[Any]:
    return [s for s in slide.shapes if s.shape_type == 13]  # MSO_SHAPE_TYPE.PICTURE == 13


def _is_empty_template_slide(slide) -> bool:
    pics = _slide_pictures(slide)
    if pics:
        return False
    texts = set(_slide_texts(slide))
    return texts == {"ANTES", "DESPUÉS"}


def _remove_slide(prs: Presentation, slide_index: int) -> None:
    """Safely remove a slide by index from a python-pptx Presentation."""
    sld_id_lst = prs.slides._sldIdLst
    sld_id = sld_id_lst[slide_index]
    r_id = sld_id.rId
    prs.part.drop_rel(r_id)
    sld_id_lst.remove(sld_id)


def _add_antes_despues_chrome(slide) -> None:
    """Add the exact top bar, ANTES/DESPUÉS boxes, and bottom bar."""
    # Top blue bar
    top_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Emu(0), Emu(196952), Emu(SLIDE_W), Emu(1254000)
    )
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = ACCENT_BLUE
    top_bar.line.color.rgb = BORDER_NAVY
    top_bar.line.width = Emu(12700)

    # ANTES box
    antes_box = slide.shapes.add_textbox(Emu(152375), Emu(362250), Emu(5373300), Emu(923400))
    antes_box.line.color.rgb = WHITE
    antes_box.line.width = Emu(76200)
    tf = antes_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "ANTES"
    run.font.name = "Calibri"
    run.font.size = Pt(54)
    run.font.bold = True
    run.font.color.rgb = WHITE

    # DESPUÉS box
    despues_box = slide.shapes.add_textbox(Emu(5643241), Emu(362250), Emu(6246000), Emu(923400))
    despues_box.line.color.rgb = WHITE
    despues_box.line.width = Emu(76200)
    tf2 = despues_box.text_frame
    tf2.word_wrap = True
    tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = "DESPUÉS"
    run2.font.name = "Calibri"
    run2.font.size = Pt(54)
    run2.font.bold = True
    run2.font.color.rgb = WHITE

    # Bottom blue bar
    bot_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Emu(0), Emu(6656697), Emu(SLIDE_W), Emu(187500)
    )
    bot_bar.fill.solid()
    bot_bar.fill.fore_color.rgb = ACCENT_BLUE
    bot_bar.line.color.rgb = BORDER_NAVY
    bot_bar.line.width = Emu(12700)


def inspect_presentation(pptx_path: Path) -> dict[str, Any]:
    """Return structured list of walks (grouped by cover slides) and slide states."""
    prs = Presentation(str(pptx_path))
    slide_w = int(prs.slide_width)
    half_w = slide_w // 2

    walks: list[dict[str, Any]] = []
    current_walk: dict[str, Any] | None = None

    for idx, slide in enumerate(prs.slides):
        texts = _slide_texts(slide)
        pics = _slide_pictures(slide)
        has_headers = any(t in ("ANTES", "DESPUÉS") for t in texts)

        # Detect cover slide: 1 picture and no ANTES/DESPUÉS header
        if not has_headers and len(pics) >= 1:
            current_walk = {
                "walk_index": len(walks) + 1,
                "cover_slide_index": idx,
                "cover_slide_number": idx + 1,
                "slides": [],
            }
            walks.append(current_walk)
            continue

        if _is_empty_template_slide(slide):
            continue

        if current_walk is None:
            current_walk = {
                "walk_index": 1,
                "cover_slide_index": None,
                "cover_slide_number": None,
                "slides": [],
            }
            walks.append(current_walk)

        antes_pics = []
        despues_pics = []
        for p_idx, pic in enumerate(pics):
            center_x = int(pic.left) + int(pic.width) // 2
            if center_x < half_w:
                antes_pics.append(p_idx)
            else:
                despues_pics.append(p_idx)

        extra_notes = [t for t in texts if t not in ("ANTES", "DESPUÉS")]

        current_walk["slides"].append(
            {
                "slide_index": idx,
                "slide_number": idx + 1,
                "has_antes": len(antes_pics) > 0,
                "has_despues": len(despues_pics) > 0 or len(extra_notes) > 0,
                "antes_pic_index": antes_pics[0] if antes_pics else None,
                "despues_pic_index": despues_pics[0] if despues_pics else None,
                "notes": " ".join(extra_notes),
            }
        )

    file_size_mb = round(pptx_path.stat().st_size / (1024 * 1024), 2)
    return {
        "filename": pptx_path.name,
        "total_slides": len(prs.slides),
        "file_size_mb": file_size_mb,
        "walks": walks,
    }


def extract_slide_picture(pptx_path: Path, slide_index: int, pic_index: int, thumb_max: int = 720) -> bytes:
    """Extract a specific picture from a slide as a lightweight JPEG thumbnail."""
    prs = Presentation(str(pptx_path))
    slide = prs.slides[slide_index]
    pics = _slide_pictures(slide)
    blob = pics[pic_index].image.blob
    with Image.open(io.BytesIO(blob)) as img:
        if img.mode in ("RGBA", "LA", "P"):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            bg.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")
        img.thumbnail((thumb_max, thumb_max), Image.Resampling.LANCZOS)
        out = io.BytesIO()
        img.save(out, format="JPEG", quality=80)
        return out.getvalue()


def append_walk(
    pptx_path: Path,
    cover_png_bytes: bytes,
    antes_photos: list[bytes],
    clean_trailing_empty: bool = True,
) -> dict[str, Any]:
    """Append a cover slide + 1 ANTES/DESPUÉS slide per photo to the presentation."""
    backup_pptx(pptx_path)
    prs = Presentation(str(pptx_path))
    layout = _get_blank_layout(prs)

    removed_empty = 0
    if clean_trailing_empty:
        while len(prs.slides) > 0 and _is_empty_template_slide(prs.slides[-1]):
            _remove_slide(prs, len(prs.slides) - 1)
            removed_empty += 1

    # 1. Add cover slide
    cover_slide = prs.slides.add_slide(layout)
    with Image.open(io.BytesIO(cover_png_bytes)) as c_img:
        cw, ch = c_img.size
    cx, cy, c_width, c_height = fit_box(cw, ch, COVER_BOX)
    cover_slide.shapes.add_picture(
        io.BytesIO(cover_png_bytes), Emu(cx), Emu(cy), Emu(c_width), Emu(c_height)
    )

    # 2. Add one ANTES / DESPUÉS slide per photo
    added_slides = 0
    for raw_photo in antes_photos:
        opt_bytes, pw, ph = optimize_photo(raw_photo)
        slide = prs.slides.add_slide(layout)
        _add_antes_despues_chrome(slide)
        px, py, p_width, p_height = fit_box(pw, ph, ANTES_BOX)
        slide.shapes.add_picture(
            io.BytesIO(opt_bytes), Emu(px), Emu(py), Emu(p_width), Emu(p_height)
        )
        added_slides += 1

    prs.save(str(pptx_path))
    return {
        "added_cover": True,
        "added_photo_slides": added_slides,
        "removed_empty_slides": removed_empty,
        "total_slides": len(prs.slides),
        "file_size_mb": round(pptx_path.stat().st_size / (1024 * 1024), 2),
    }


def update_despues_slides(
    pptx_path: Path,
    updates: list[dict[str, Any]],
) -> dict[str, Any]:
    """Insert DESPUÉS photos and/or notes into existing slides by slide_index.

    Each item in updates:
      {"slide_index": int, "photo_bytes": bytes | None, "note": str | None}
    """
    backup_pptx(pptx_path)
    prs = Presentation(str(pptx_path))
    slide_w = int(prs.slide_width)
    half_w = slide_w // 2
    updated_count = 0

    for item in updates:
        idx = int(item["slide_index"])
        photo_bytes = item.get("photo_bytes")
        note = (item.get("note") or "").strip()
        if not photo_bytes and not note:
            continue

        slide = prs.slides[idx]

        # If replacing an existing DESPUÉS photo, remove old right-side picture(s)
        if photo_bytes:
            for pic in list(_slide_pictures(slide)):
                center_x = int(pic.left) + int(pic.width) // 2
                if center_x >= half_w:
                    pic._element.getparent().remove(pic._element)

            opt_bytes, pw, ph = optimize_photo(photo_bytes)
            # Leave room for note banner at bottom if a note is also provided
            box = DESPUES_BOX
            if note:
                bx, by, bcx, bcy = DESPUES_BOX
                box = (bx, by, bcx, bcy - 620000)
            px, py, p_width, p_height = fit_box(pw, ph, box)
            slide.shapes.add_picture(
                io.BytesIO(opt_bytes), Emu(px), Emu(py), Emu(p_width), Emu(p_height)
            )

        if note:
            # Remove any previous extra note textboxes on the right side
            for shp in list(slide.shapes):
                if shp.has_text_frame:
                    txt = shp.text_frame.text.strip()
                    if txt and txt not in ("ANTES", "DESPUÉS"):
                        shp._element.getparent().remove(shp._element)

            bx, by, bcx, bcy = DESPUES_BOX
            if photo_bytes:
                note_box = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    Emu(bx + 100000),
                    Emu(by + bcy - 540000),
                    Emu(bcx - 200000),
                    Emu(480000),
                )
            else:
                note_box = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    Emu(bx + 250000),
                    Emu(by + bcy // 2 - 550000),
                    Emu(bcx - 500000),
                    Emu(1100000),
                )
            note_box.fill.solid()
            note_box.fill.fore_color.rgb = NOTE_BG
            note_box.line.color.rgb = ACCENT_BLUE
            note_box.line.width = Emu(19050)
            tf = note_box.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = note
            run.font.name = "Calibri"
            run.font.size = Pt(22 if photo_bytes else 28)
            run.font.bold = True
            run.font.color.rgb = DARK_TEXT

        updated_count += 1

    prs.save(str(pptx_path))
    return {
        "updated_slides": updated_count,
        "total_slides": len(prs.slides),
        "file_size_mb": round(pptx_path.stat().st_size / (1024 * 1024), 2),
    }
