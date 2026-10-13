__author__ = "Süleyman Bozkurt and ChatGPT"
__version__ = "v1.7"
__maintainer__ = "Süleyman Bozkurt"
__email__ = "sbozkurt.mbg@gmail.com"
__date__ = '28.12.2022'
__update__ = '02.06.2023'

import os
import sys
import shutil
import subprocess
import io
import tempfile
import threading
import traceback
import re
import pytesseract
from docx import Document
from pdf2image import convert_from_path, pdfinfo_from_path
from pdf2docx import Converter
from docx.shared import Inches, Pt
from PIL import Image, ImageDraw
import webbrowser
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

# GUI imports are optional so the converter can be imported on headless servers.
# The Flask API uses the conversion helpers from this module and should not need
# tkinter/custom GUI packages during deployment.
try:
    import tkinter
    import tkinter as tk
    from tkinter import Button, Label, OptionMenu, StringVar, Tk, filedialog, ttk
    from tkinter.font import Font
    from PIL import ImageTk
except Exception:
    tkinter = None
    tk = None
    Button = Label = OptionMenu = StringVar = Tk = filedialog = Font = ImageTk = None

    class _HeadlessFrame:
        pass

    class _HeadlessTtk:
        Frame = _HeadlessFrame

    ttk = _HeadlessTtk()

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None
try:
    import cv2
    import numpy as np
except Exception:
    cv2 = None
    np = None


def open_with_default_app(path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=True)
    else:
        subprocess.run(["xdg-open", path], check=True)


def flatten_floating_shapes_in_docx(docx_path):
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }
    ET.register_namespace("w", ns["w"])
    ET.register_namespace("wp", ns["wp"])
    ET.register_namespace("a", ns["a"])

    def local_name(tag):
        return tag.split("}", 1)[1] if "}" in tag else tag

    def anchor_to_inline(anchor):
        inline = ET.Element(f'{{{ns["wp"]}}}inline')
        for attr in ("distT", "distB", "distL", "distR"):
            if attr in anchor.attrib:
                inline.set(attr, anchor.attrib[attr])

        allowed_children = {"extent", "effectExtent", "docPr", "cNvGraphicFramePr"}
        for child in list(anchor):
            lname = local_name(child.tag)
            if lname in allowed_children:
                inline.append(child)
            elif child.tag == f'{{{ns["a"]}}}graphic':
                inline.append(child)
        return inline

    converted_count = 0
    updated = False
    with zipfile.ZipFile(docx_path, "r") as zin:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            tmp_path = tmp_file.name

        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename.startswith("word/") and info.filename.endswith(".xml") and b"wp:anchor" in data:
                    try:
                        root = ET.fromstring(data)
                        parent_map = {child: parent for parent in root.iter() for child in parent}
                        anchors = list(root.iter(f'{{{ns["wp"]}}}anchor'))
                        for anchor in anchors:
                            parent = parent_map.get(anchor)
                            if parent is None:
                                continue
                            inline = anchor_to_inline(anchor)
                            idx = list(parent).index(anchor)
                            parent.remove(anchor)
                            parent.insert(idx, inline)
                            converted_count += 1
                        if anchors:
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                            updated = True
                    except Exception:
                        pass
                zout.writestr(info, data)

    if updated:
        shutil.move(tmp_path, docx_path)
    else:
        os.remove(tmp_path)
    return converted_count


def force_large_anchors_behind_text(docx_path, min_area_ratio=0.25):
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    }
    ET.register_namespace("w", ns["w"])
    ET.register_namespace("wp", ns["wp"])

    page_w_emu = 11909 * 635
    page_h_emu = 16834 * 635

    with zipfile.ZipFile(docx_path, "r") as zin:
        try:
            doc_root = ET.fromstring(zin.read("word/document.xml"))
            sect_pr = doc_root.find(f'.//{{{ns["w"]}}}sectPr')
            if sect_pr is not None:
                pg_sz = sect_pr.find(f'{{{ns["w"]}}}pgSz')
                if pg_sz is not None:
                    w_twips = int(pg_sz.get(f'{{{ns["w"]}}}w', "11909"))
                    h_twips = int(pg_sz.get(f'{{{ns["w"]}}}h', "16834"))
                    page_w_emu = max(1, w_twips * 635)
                    page_h_emu = max(1, h_twips * 635)
        except Exception:
            pass

        page_area = page_w_emu * page_h_emu
        changed = 0
        updated = False

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            tmp_path = tmp_file.name

        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename.startswith("word/") and info.filename.endswith(".xml") and b"wp:anchor" in data:
                    try:
                        root = ET.fromstring(data)
                        for anchor in root.iter(f'{{{ns["wp"]}}}anchor'):
                            extent = anchor.find(f'{{{ns["wp"]}}}extent')
                            if extent is None:
                                continue
                            try:
                                cx = int(extent.get("cx", "0"))
                                cy = int(extent.get("cy", "0"))
                            except Exception:
                                continue
                            area_ratio = (cx * cy) / page_area if page_area > 0 else 0.0
                            covers_large_area = area_ratio >= min_area_ratio
                            covers_major_width = cx >= int(page_w_emu * 0.75) and cy >= int(page_h_emu * 0.55)
                            if not (covers_large_area or covers_major_width):
                                continue

                            changed_here = False
                            if anchor.get("behindDoc") != "1":
                                anchor.set("behindDoc", "1")
                                changed_here = True
                            if anchor.get("relativeHeight") != "0":
                                anchor.set("relativeHeight", "0")
                                changed_here = True
                            if anchor.get("allowOverlap") != "1":
                                anchor.set("allowOverlap", "1")
                                changed_here = True

                            wrap_none = anchor.find(f'{{{ns["wp"]}}}wrapNone')
                            if wrap_none is None:
                                for child in list(anchor):
                                    if child.tag.startswith(f'{{{ns["wp"]}}}wrap'):
                                        anchor.remove(child)
                                anchor.insert(3, ET.Element(f'{{{ns["wp"]}}}wrapNone'))
                                changed_here = True

                            if changed_here:
                                changed += 1
                                updated = True

                        if updated:
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    except Exception:
                        pass
                zout.writestr(info, data)

    if updated:
        shutil.move(tmp_path, docx_path)
    else:
        os.remove(tmp_path)
    return changed


def remove_soft_hyphens_in_docx(docx_path):
    updated = False
    with zipfile.ZipFile(docx_path, "r") as zin:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            tmp_path = tmp_file.name
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                    try:
                        text = data.decode("utf-8")
                        cleaned = text.replace("\u00ad", "")
                        if cleaned != text:
                            data = cleaned.encode("utf-8")
                            updated = True
                    except Exception:
                        pass
                zout.writestr(info, data)
    if updated:
        shutil.move(tmp_path, docx_path)
    else:
        os.remove(tmp_path)
    return updated


def sanitize_docx_for_word_compatibility(docx_path):
    def normalize_decimal_attrs(xml_text, attr_name, allow_negative=False):
        pattern = rf'({re.escape(attr_name)}=")(-?\d+(?:\.\d+)?)(")'

        def repl(match):
            try:
                value = int(round(float(match.group(2))))
            except Exception:
                return match.group(0)
            if not allow_negative and value < 0:
                value = 0
            return f"{match.group(1)}{value}{match.group(3)}"

        return re.sub(pattern, repl, xml_text)

    updated = False
    with zipfile.ZipFile(docx_path, "r") as zin:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            tmp_path = tmp_file.name

        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                    try:
                        text = data.decode("utf-8")
                        cleaned = re.sub(r'(w:color=")#([0-9A-Fa-f]{6}")', r"\1\2", text)
                        cleaned = re.sub(r'(w:fill=")#([0-9A-Fa-f]{6}")', r"\1\2", cleaned)
                        cleaned = normalize_decimal_attrs(cleaned, "w:sz")
                        cleaned = normalize_decimal_attrs(cleaned, "w:w")
                        cleaned = normalize_decimal_attrs(cleaned, "w:val")
                        if cleaned != text:
                            data = cleaned.encode("utf-8")
                            updated = True
                    except Exception:
                        pass
                zout.writestr(info, data)

    if updated:
        shutil.move(tmp_path, docx_path)
    else:
        os.remove(tmp_path)
    return updated


def optimize_docx_images(docx_path, target_dpi=150, jpeg_quality=82, max_px=2400):
    """Downscale/recompress embedded images to match their on-page display size.

    pdf2docx extracts images at their full native resolution (often 300-800 DPI),
    which bloats the DOCX (a single photo can be 8-9 MB) and can disturb layout.
    Each image is resized so its resolution is ~target_dpi at the size it is
    actually displayed, and photographic images (no alpha) are re-encoded as JPEG.
    Returns a summary dict; never raises on a single bad image.
    """
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    EMU_PER_INCH = 914400

    with zipfile.ZipFile(docx_path, "r") as zin:
        names = set(zin.namelist())
        if "word/document.xml" not in names or "word/_rels/document.xml.rels" not in names:
            return {"updated": False, "optimized": 0, "bytes_saved": 0}

        # Map each relationship id to the largest extent it is displayed at.
        try:
            doc_root = ET.fromstring(zin.read("word/document.xml"))
        except Exception:
            return {"updated": False, "optimized": 0, "bytes_saved": 0}
        rid_max_emu = {}
        for drawing in doc_root.iter(f'{{{ns["w"]}}}drawing'):
            extent = drawing.find(f'.//{{{ns["wp"]}}}extent')
            blip = drawing.find(f'.//{{{ns["a"]}}}blip')
            if extent is None or blip is None:
                continue
            rid = blip.get(f'{{{ns["r"]}}}embed')
            if not rid:
                continue
            try:
                cx = int(extent.get("cx", "0"))
                cy = int(extent.get("cy", "0"))
            except Exception:
                continue
            prev = rid_max_emu.get(rid, (0, 0))
            if cx * cy > prev[0] * prev[1]:
                rid_max_emu[rid] = (cx, cy)

        rels_root = ET.fromstring(zin.read("word/_rels/document.xml.rels"))
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        # media target ("media/imageN.png") -> max display EMU across all refs
        media_max_emu = {}
        rid_to_target = {}
        for rel in rels_root.iter(f'{{{rel_ns}}}Relationship'):
            rid_to_target[rel.get("Id")] = rel.get("Target")
        for rid, emu in rid_max_emu.items():
            target = rid_to_target.get(rid)
            if not target:
                continue
            prev = media_max_emu.get(target, (0, 0))
            if emu[0] * emu[1] > prev[0] * prev[1]:
                media_max_emu[target] = emu

        # Decide new bytes (and possibly new extension) for each media part.
        new_media = {}          # "word/media/x.png" -> (new_arcname, new_bytes)
        rename_map = {}         # old target -> new target (relative, for rels)
        optimized = 0
        bytes_saved = 0
        for target, (cx, cy) in media_max_emu.items():
            arcname = "word/" + target
            if arcname not in names:
                continue
            raw = zin.read(arcname)
            try:
                from PIL import Image as _PILImage
                img = _PILImage.open(io.BytesIO(raw))
                img.load()
            except Exception:
                continue

            disp_w_in = max(cx / EMU_PER_INCH, 0.01)
            disp_h_in = max(cy / EMU_PER_INCH, 0.01)
            target_w = min(int(round(disp_w_in * target_dpi)), max_px)
            target_h = min(int(round(disp_h_in * target_dpi)), max_px)
            target_w = max(target_w, 1)
            target_h = max(target_h, 1)

            needs_resize = img.width > int(target_w * 1.15) or img.height > int(target_h * 1.15)
            has_alpha = img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info)

            if not needs_resize and (has_alpha or arcname.lower().endswith((".jpg", ".jpeg"))):
                continue

            work = img
            if needs_resize:
                work = img.copy()
                work.thumbnail((target_w, target_h), _PILImage.LANCZOS)

            buf = io.BytesIO()
            if has_alpha:
                work.save(buf, format="PNG", optimize=True)
                new_arc = arcname
                new_target = target
            else:
                if work.mode != "RGB":
                    work = work.convert("RGB")
                work.save(buf, format="JPEG", quality=jpeg_quality, optimize=True, progressive=True)
                new_arc = re.sub(r"\.(png|bmp|tif|tiff|gif)$", ".jpeg", arcname, flags=re.IGNORECASE)
                new_target = re.sub(r"\.(png|bmp|tif|tiff|gif)$", ".jpeg", target, flags=re.IGNORECASE)

            new_bytes = buf.getvalue()
            if len(new_bytes) >= len(raw) and new_arc == arcname:
                continue  # no benefit, keep original
            new_media[arcname] = (new_arc, new_bytes)
            if new_arc != arcname:
                rename_map[target] = new_target
            optimized += 1
            bytes_saved += max(0, len(raw) - len(new_bytes))

        if not new_media:
            return {"updated": False, "optimized": 0, "bytes_saved": 0}

        # Update rels targets and content types via string edits. These parts use
        # a default namespace; re-serializing them with ElementTree would inject
        # ns0: prefixes and corrupt the package (Word tolerates it, LibreOffice
        # does not), so patch the raw text instead.
        rels_data = zin.read("word/_rels/document.xml.rels")
        if rename_map:
            rels_text = rels_data.decode("utf-8")
            for old_target, new_target in rename_map.items():
                rels_text = rels_text.replace(
                    f'Target="{old_target}"', f'Target="{new_target}"'
                )
            rels_data = rels_text.encode("utf-8")

        ct_data = zin.read("[Content_Types].xml")
        if rename_map and b'Extension="jpeg"' not in ct_data and b'Extension="jpg"' not in ct_data:
            ct_text = ct_data.decode("utf-8")
            ct_text = ct_text.replace(
                "</Types>",
                '<Default Extension="jpeg" ContentType="image/jpeg"/></Types>',
            )
            ct_data = ct_text.encode("utf-8")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            tmp_path = tmp_file.name
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                arc = info.filename
                if arc == "word/_rels/document.xml.rels":
                    zout.writestr(info, rels_data)
                elif arc == "[Content_Types].xml":
                    zout.writestr(info, ct_data)
                elif arc in new_media:
                    new_arc, new_bytes = new_media[arc]
                    if new_arc == arc:
                        zout.writestr(info, new_bytes)
                    else:
                        zout.writestr(new_arc, new_bytes)  # renamed part
                else:
                    zout.writestr(info, zin.read(arc))

    shutil.move(tmp_path, docx_path)
    return {"updated": True, "optimized": optimized, "bytes_saved": bytes_saved}


def _normalize_line_for_compare(text):
    return re.sub(r"\s+", " ", (text or "").strip().lower())


def build_pdf_line_counter(pdf_path, min_len=20):
    counter = Counter()
    if fitz is None:
        return counter
    with fitz.open(pdf_path) as pdf_doc:
        for page in pdf_doc:
            for line in (page.get_text() or "").splitlines():
                normalized = _normalize_line_for_compare(line)
                if len(normalized) >= min_len:
                    counter[normalized] += 1
    return counter


def prune_docx_redundant_lines(docx_path, pdf_path, min_len=28, unknown_line_limit=1):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w_p = f'{{{ns["w"]}}}p'
    w_t = f'{{{ns["w"]}}}t'
    w_tc = f'{{{ns["w"]}}}tc'
    ET.register_namespace("w", ns["w"])

    pdf_counter = build_pdf_line_counter(pdf_path, min_len=min_len)
    kept_counter = Counter()
    removed = 0
    updated = False

    with zipfile.ZipFile(docx_path, "r") as zin:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            tmp_path = tmp_file.name

        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/document.xml":
                    try:
                        root = ET.fromstring(data)
                        parent_map = {child: parent for parent in root.iter() for child in parent}
                        paragraphs = [p for p in root.iter(w_p)]
                        for para in paragraphs:
                            parent = parent_map.get(para)
                            if parent is None:
                                continue

                            line_text = "".join((t.text or "") for t in para.iter(w_t))
                            normalized = _normalize_line_for_compare(line_text)
                            if len(normalized) < min_len:
                                continue

                            allowed = pdf_counter.get(normalized, 0)
                            if allowed == 0:
                                allowed = unknown_line_limit

                            if kept_counter[normalized] >= allowed:
                                if parent.tag == w_tc:
                                    p_count = sum(1 for child in list(parent) if child.tag == w_p)
                                    if p_count <= 1:
                                        # Keep table-cell structure valid, but clear duplicated text.
                                        for child in list(para):
                                            para.remove(child)
                                        removed += 1
                                        updated = True
                                        continue
                                parent.remove(para)
                                removed += 1
                                updated = True
                            else:
                                kept_counter[normalized] += 1

                        if updated:
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    except Exception:
                        pass
                zout.writestr(info, data)

    if updated:
        shutil.move(tmp_path, docx_path)
    else:
        os.remove(tmp_path)
    return removed


def _barcode_numeric_token(text, min_digits=6):
    compact = re.sub(r"\s+", "", text or "")
    if re.fullmatch(rf"\*?\d{{{min_digits},}}\*?", compact):
        return compact
    return None


def _barcode_payload_from_token(token, min_digits=6):
    if not token:
        return None
    compact = re.sub(r"\s+", "", token or "").upper()
    if compact.startswith("*") and compact.endswith("*") and len(compact) > 2:
        compact = compact[1:-1]
    if re.fullmatch(rf"\d{{{min_digits},}}", compact):
        return compact
    if re.fullmatch(r"[0-9A-Z\-\.\ $/+%]{4,}", compact):
        return compact
    return None


def _code39_patterns():
    return {
        "0": "nnnwwnwnn",
        "1": "wnnwnnnnw",
        "2": "nnwwnnnnw",
        "3": "wnwwnnnnn",
        "4": "nnnwwnnnw",
        "5": "wnnwwnnnn",
        "6": "nnwwwnnnn",
        "7": "nnnwnnwnw",
        "8": "wnnwnnwnn",
        "9": "nnwwnnwnn",
        "A": "wnnnnwnnw",
        "B": "nnwnnwnnw",
        "C": "wnwnnwnnn",
        "D": "nnnnwwnnw",
        "E": "wnnnwwnnn",
        "F": "nnwnwwnnn",
        "G": "nnnnnwwnw",
        "H": "wnnnnwwnn",
        "I": "nnwnnwwnn",
        "J": "nnnnwwwnn",
        "K": "wnnnnnnww",
        "L": "nnwnnnnww",
        "M": "wnwnnnnwn",
        "N": "nnnnwnnww",
        "O": "wnnnwnnwn",
        "P": "nnwnwnnwn",
        "Q": "nnnnnnwww",
        "R": "wnnnnnwwn",
        "S": "nnwnnnwwn",
        "T": "nnnnwnwwn",
        "U": "wwnnnnnnw",
        "V": "nwwnnnnnw",
        "W": "wwwnnnnnn",
        "X": "nwnnwnnnw",
        "Y": "wwnnwnnnn",
        "Z": "nwwnwnnnn",
        "-": "nwnnnnwnw",
        ".": "wwnnnnwnn",
        " ": "nwwnnnwnn",
        "$": "nwnwnwnnn",
        "/": "nwnwnnnwn",
        "+": "nwnnnwnwn",
        "%": "nnnwnwnwn",
        "*": "nwnnwnwnn",
    }


def _render_code39_barcode_image(payload, out_path):
    patterns = _code39_patterns()
    payload = (payload or "").upper().strip()
    if not payload:
        raise ValueError("Empty barcode payload.")

    encoded = f"*{payload}*"
    for ch in encoded:
        if ch not in patterns:
            raise ValueError(f"Unsupported character for Code39: {ch}")

    narrow = 2
    wide = 5
    inter_char_gap = 2
    quiet = 12
    height = 56

    width = quiet * 2
    for idx, ch in enumerate(encoded):
        width += sum(wide if bit == "w" else narrow for bit in patterns[ch])
        if idx < len(encoded) - 1:
            width += inter_char_gap

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    x = quiet
    for idx, ch in enumerate(encoded):
        pattern = patterns[ch]
        for bar_idx, bit in enumerate(pattern):
            bar_w = wide if bit == "w" else narrow
            if bar_idx % 2 == 0:
                draw.rectangle((x, 0, x + bar_w - 1, height - 1), fill="black")
            x += bar_w
        if idx < len(encoded) - 1:
            x += inter_char_gap

    image.save(out_path, format="PNG")
    return out_path


def _iter_doc_paragraphs(container, seen_cells=None):
    if seen_cells is None:
        seen_cells = set()
    for para in getattr(container, "paragraphs", []):
        yield para
    for table in getattr(container, "tables", []):
        for row in table.rows:
            for cell in row.cells:
                cell_key = id(cell._tc)
                if cell_key in seen_cells:
                    continue
                seen_cells.add(cell_key)
                yield from _iter_doc_paragraphs(cell, seen_cells=seen_cells)


def _run_font_values(run):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    rfonts = run._r.find("./w:rPr/w:rFonts", namespaces=ns)
    if rfonts is None:
        return ""
    attr_ascii = f'{{{ns["w"]}}}ascii'
    attr_hansi = f'{{{ns["w"]}}}hAnsi'
    attr_eastasia = f'{{{ns["w"]}}}eastAsia'
    attr_cs = f'{{{ns["w"]}}}cs'
    return " ".join(
        [
            rfonts.get(attr_ascii, ""),
            rfonts.get(attr_hansi, ""),
            rfonts.get(attr_eastasia, ""),
            rfonts.get(attr_cs, ""),
        ]
    ).lower()


def remove_barcode_text_artifacts(docx_path, min_digits=6, min_repeat=2):
    barcode_font_keys = (
        "barcode",
        "code39",
        "code 39",
        "3 of 9",
        "free 3 of 9",
        "ean",
        "upc",
        "code128",
        "code 128",
    )
    updated = False
    removed_runs = 0
    inserted_images = 0

    doc = Document(docx_path)
    paragraphs = list(_iter_doc_paragraphs(doc))
    token_counter = Counter()
    for para in paragraphs:
        token = _barcode_numeric_token(para.text, min_digits=min_digits)
        if token:
            token_counter[token] += 1

    with tempfile.TemporaryDirectory(prefix="barcode_image_") as tmp_dir:
        image_cache = {}
        for para in paragraphs:
            for run in para.runs:
                raw_text = run.text or ""
                if not raw_text.strip():
                    continue
                font_values = _run_font_values(run)
                has_barcode_font = any(key in font_values for key in barcode_font_keys)
                token = _barcode_numeric_token(raw_text, min_digits=min_digits)
                has_star_wrapped_numeric = bool(token and token.startswith("*") and token.endswith("*"))

                if not has_barcode_font and not has_star_wrapped_numeric:
                    continue

                payload = _barcode_payload_from_token(token or raw_text, min_digits=min_digits)
                if not payload:
                    if has_barcode_font:
                        run.text = ""
                        removed_runs += 1
                        updated = True
                    continue

                if token and token_counter.get(token, 0) < max(1, min_repeat) and not has_barcode_font:
                    continue

                if payload not in image_cache:
                    safe_name = re.sub(r"[^0-9A-Z]+", "_", payload) or "barcode"
                    img_path = os.path.join(tmp_dir, f"{safe_name}_{len(image_cache)}.png")
                    _render_code39_barcode_image(payload, img_path)
                    image_cache[payload] = img_path
                else:
                    img_path = image_cache[payload]

                run.text = ""
                run.add_picture(img_path, height=Pt(16))
                inserted_images += 1
                updated = True

    if updated:
        doc.save(docx_path)

    return {
        "removed_runs": removed_runs,
        "removed_paragraphs": 0,
        "images_inserted": inserted_images,
        "updated": updated,
    }


def relax_table_layout_constraints(docx_path, single_page_target=False, aggressive=False):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w_tc = f'{{{ns["w"]}}}tc'
    w_p = f'{{{ns["w"]}}}p'
    w_ppr = f'{{{ns["w"]}}}pPr'
    w_spacing = f'{{{ns["w"]}}}spacing'
    w_tr_height = f'{{{ns["w"]}}}trHeight'
    attr_h_rule = f'{{{ns["w"]}}}hRule'
    attr_line_rule = f'{{{ns["w"]}}}lineRule'
    attr_line = f'{{{ns["w"]}}}line'
    attr_before = f'{{{ns["w"]}}}before'
    attr_after = f'{{{ns["w"]}}}after'
    attr_before_lines = f'{{{ns["w"]}}}beforeLines'
    attr_after_lines = f'{{{ns["w"]}}}afterLines'
    ET.register_namespace("w", ns["w"])

    max_before_after = 18 if aggressive else (24 if single_page_target else 36)
    min_line = 240 if aggressive else 220

    rows_relaxed = 0
    spacing_relaxed = 0
    updated = False

    with zipfile.ZipFile(docx_path, "r") as zin:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            tmp_path = tmp_file.name

        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/document.xml":
                    try:
                        root = ET.fromstring(data)
                        for tr_h in root.iter(w_tr_height):
                            if tr_h.get(attr_h_rule) == "exact":
                                tr_h.set(attr_h_rule, "atLeast")
                                rows_relaxed += 1
                                updated = True

                        if aggressive:
                            for cell in root.iter(w_tc):
                                for para in cell.iter(w_p):
                                    ppr = para.find(w_ppr)
                                    if ppr is None:
                                        continue
                                    spacing = ppr.find(w_spacing)
                                    if spacing is None:
                                        continue
                                    changed = False
                                    if spacing.get(attr_line_rule) == "exact":
                                        spacing.set(attr_line_rule, "auto")
                                        changed = True
                                    if attr_line in spacing.attrib:
                                        try:
                                            if int(spacing.get(attr_line, "0")) < min_line:
                                                spacing.set(attr_line, str(min_line))
                                                changed = True
                                        except Exception:
                                            pass
                                    for attr in (attr_before, attr_after):
                                        if attr not in spacing.attrib:
                                            continue
                                        try:
                                            if int(spacing.get(attr, "0")) > max_before_after:
                                                spacing.set(attr, str(max_before_after))
                                                changed = True
                                        except Exception:
                                            pass
                                    if attr_before_lines in spacing.attrib:
                                        spacing.attrib.pop(attr_before_lines, None)
                                        changed = True
                                    if attr_after_lines in spacing.attrib:
                                        spacing.attrib.pop(attr_after_lines, None)
                                        changed = True
                                    if changed:
                                        spacing_relaxed += 1
                                        updated = True

                        if updated:
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    except Exception:
                        pass
                zout.writestr(info, data)

    if updated:
        shutil.move(tmp_path, docx_path)
    else:
        os.remove(tmp_path)
    return {
        "updated": updated,
        "rows_relaxed": rows_relaxed,
        "spacing_relaxed": spacing_relaxed,
        "cell_margins_compacted": 0,
        "blank_paragraphs_removed": 0,
        "page_breaks_removed": 0,
    }


def enforce_single_page_layout(docx_path, pdf_path):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w_br = f'{{{ns["w"]}}}br'
    w_last_break = f'{{{ns["w"]}}}lastRenderedPageBreak'
    attr_type = f'{{{ns["w"]}}}type'
    ET.register_namespace("w", ns["w"])

    if fitz is None:
        return {
            "updated": False,
            "page_size_set": False,
            "margins_compacted": False,
            "breaks_removed": 0,
        }
    try:
        with fitz.open(pdf_path) as pdf_doc:
            if len(pdf_doc) != 1:
                return {
                    "updated": False,
                    "page_size_set": False,
                    "margins_compacted": False,
                    "breaks_removed": 0,
                }
    except Exception:
        return {
            "updated": False,
            "page_size_set": False,
            "margins_compacted": False,
            "breaks_removed": 0,
        }

    updated = False
    breaks_removed = 0
    with zipfile.ZipFile(docx_path, "r") as zin:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            tmp_path = tmp_file.name
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/document.xml":
                    try:
                        root = ET.fromstring(data)
                        parent_map = {child: parent for parent in root.iter() for child in parent}
                        for br in list(root.iter(w_br)):
                            if br.get(attr_type) == "page":
                                parent = parent_map.get(br)
                                if parent is not None:
                                    parent.remove(br)
                                    breaks_removed += 1
                                    updated = True
                        for lpb in list(root.iter(w_last_break)):
                            parent = parent_map.get(lpb)
                            if parent is not None:
                                parent.remove(lpb)
                                breaks_removed += 1
                                updated = True
                        if updated:
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    except Exception:
                        pass
                zout.writestr(info, data)
    if updated:
        shutil.move(tmp_path, docx_path)
    else:
        os.remove(tmp_path)
    return {
        "updated": updated,
        "page_size_set": False,
        "margins_compacted": False,
        "breaks_removed": breaks_removed,
    }


def compute_layout_overlap_risk(docx_path):
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    }
    w_tc = f'{{{ns["w"]}}}tc'
    w_tr = f'{{{ns["w"]}}}tr'
    w_p = f'{{{ns["w"]}}}p'
    w_t = f'{{{ns["w"]}}}t'
    w_ppr = f'{{{ns["w"]}}}pPr'
    w_spacing = f'{{{ns["w"]}}}spacing'
    w_tr_height = f'{{{ns["w"]}}}trHeight'
    attr_line = f'{{{ns["w"]}}}line'
    attr_h_rule = f'{{{ns["w"]}}}hRule'
    attr_val = f'{{{ns["w"]}}}val'

    total_cells = 0
    cramped_cells = 0
    cramped_rows = 0
    front_large_anchors = 0
    page_w_emu = 11909 * 635
    page_h_emu = 16834 * 635

    with zipfile.ZipFile(docx_path, "r") as archive:
        try:
            root = ET.fromstring(archive.read("word/document.xml"))
        except Exception:
            return {
                "risk": 0.0,
                "cramped_cells": 0,
                "cramped_rows": 0,
                "front_large_anchors": 0,
                "total_cells": 0,
            }

        sect_pr = root.find(f'.//{{{ns["w"]}}}sectPr')
        if sect_pr is not None:
            pg_sz = sect_pr.find(f'{{{ns["w"]}}}pgSz')
            if pg_sz is not None:
                try:
                    page_w_emu = int(pg_sz.get(f'{{{ns["w"]}}}w', "11909")) * 635
                    page_h_emu = int(pg_sz.get(f'{{{ns["w"]}}}h', "16834")) * 635
                except Exception:
                    pass
        page_area = max(1, page_w_emu * page_h_emu)

        for anchor in root.iter(f'{{{ns["wp"]}}}anchor'):
            extent = anchor.find(f'{{{ns["wp"]}}}extent')
            if extent is None:
                continue
            try:
                cx = int(extent.get("cx", "0"))
                cy = int(extent.get("cy", "0"))
            except Exception:
                continue
            area_ratio = (cx * cy) / page_area
            if anchor.get("behindDoc") != "1" and area_ratio >= 0.20:
                front_large_anchors += 1

        for row in root.iter(w_tr):
            row_text_len = 0
            for node in row.iter(w_t):
                row_text_len += len((node.text or "").strip())
            tr_height = row.find(f'.//{w_tr_height}')
            if tr_height is not None:
                try:
                    h_val = int(tr_height.get(attr_val, "0"))
                except Exception:
                    h_val = 0
                if tr_height.get(attr_h_rule) == "exact" and h_val < 200 and row_text_len > 140:
                    cramped_rows += 1

        for cell in root.iter(w_tc):
            total_cells += 1
            text_len = 0
            min_line = None
            for para in cell.iter(w_p):
                for node in para.iter(w_t):
                    text_len += len((node.text or "").strip())
                ppr = para.find(w_ppr)
                if ppr is None:
                    continue
                spacing = ppr.find(w_spacing)
                if spacing is None:
                    continue
                try:
                    line = int(spacing.get(attr_line, "0"))
                except Exception:
                    continue
                if line > 0 and (min_line is None or line < min_line):
                    min_line = line
            if text_len > 140 and (min_line is not None and min_line < 170):
                cramped_cells += 1
            if text_len > 260 and (min_line is not None and min_line < 190):
                cramped_cells += 1

    denom = max(10, int(total_cells * 0.3))
    raw_risk = (2 * front_large_anchors + cramped_rows + cramped_cells) / denom
    risk = min(max(raw_risk, 0.0), 1.0)
    return {
        "risk": risk,
        "cramped_cells": cramped_cells,
        "cramped_rows": cramped_rows,
        "front_large_anchors": front_large_anchors,
        "total_cells": total_cells,
    }


def _tokenize_quality_text(text):
    tokens = re.findall(r"[A-Za-z0-9*#:/\.\-]+", (text or "").lower())
    return [tok for tok in tokens if len(tok) > 1]


def analyze_pdf_quality_signal(pdf_path):
    if fitz is None:
        return {
            "tokens": set(),
            "numeric_tokens": set(),
            "barcode_tokens": set(),
            "field_labels": set(),
            "image_count": 0,
            "drawing_count": 0,
            "page_count": 0,
            "total_token_count": 0,
        }

    tokens = set()
    numeric_tokens = set()
    barcode_tokens = set()
    field_labels = set()
    image_count = 0
    drawing_count = 0
    page_count = 0
    total_token_count = 0
    with fitz.open(pdf_path) as pdf_doc:
        page_count = len(pdf_doc)
        for page in pdf_doc:
            page_text = page.get_text() or ""
            page_tokens = _tokenize_quality_text(page_text)
            total_token_count += len(page_tokens)
            for token in page_tokens:
                tokens.add(token)
                if any(ch.isdigit() for ch in token):
                    numeric_tokens.add(token)
                    digits_only = "".join(ch for ch in token if ch.isdigit())
                    if len(digits_only) >= 6 or token.startswith("*") or token.endswith("*"):
                        barcode_tokens.add(token)
            for match in re.findall(r"\b([A-Za-z][A-Za-z /'\-]{1,50}:)", page_text):
                field_labels.add(re.sub(r"\s+", " ", match.strip().lower()))
            try:
                image_count += len(page.get_images(full=True))
            except Exception:
                pass
            try:
                drawing_count += len(page.get_drawings())
            except Exception:
                pass

    return {
        "tokens": tokens,
        "numeric_tokens": numeric_tokens,
        "barcode_tokens": barcode_tokens,
        "field_labels": field_labels,
        "image_count": image_count,
        "drawing_count": drawing_count,
        "page_count": page_count,
        "total_token_count": total_token_count,
    }


def analyze_docx_quality_signal(docx_path):
    doc = Document(docx_path)
    text_parts = []
    anchor_count = 0
    media_count = 0
    with zipfile.ZipFile(docx_path, "r") as archive:
        for info in archive.infolist():
            if info.filename.startswith("word/media/") and not info.is_dir():
                media_count += 1
            elif info.filename.startswith("word/") and info.filename.endswith(".xml"):
                try:
                    data = archive.read(info.filename)
                except Exception:
                    continue
                anchor_count += data.count(b"<wp:anchor")
                basename = os.path.basename(info.filename)
                if not (
                    basename == "document.xml"
                    or basename.startswith("header")
                    or basename.startswith("footer")
                    or basename in {"footnotes.xml", "endnotes.xml", "comments.xml"}
                ):
                    continue
                try:
                    root = ET.fromstring(data)
                except Exception:
                    continue
                for node in root.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
                    if node.text:
                        text_parts.append(node.text)

    merged_text = "\n".join(text_parts)
    token_list = _tokenize_quality_text(merged_text)
    tokens = set(token_list)
    numeric_tokens = {tok for tok in tokens if any(ch.isdigit() for ch in tok)}
    barcode_tokens = set()
    for tok in numeric_tokens:
        digits_only = "".join(ch for ch in tok if ch.isdigit())
        if len(digits_only) >= 6 or tok.startswith("*") or tok.endswith("*"):
            barcode_tokens.add(tok)
    field_labels = set()
    for match in re.findall(r"\b([A-Za-z][A-Za-z /'\-]{1,50}:)", merged_text or ""):
        field_labels.add(re.sub(r"\s+", " ", match.strip().lower()))


    return {
        "tokens": tokens,
        "numeric_tokens": numeric_tokens,
        "barcode_tokens": barcode_tokens,
        "field_labels": field_labels,
        "table_count": len(doc.tables),
        "section_count": len(doc.sections),
        "paragraph_count": len(doc.paragraphs),
        "anchor_count": anchor_count,
        "media_count": media_count,
        "total_token_count": len(token_list),
    }


def score_conversion_quality(pdf_signal, docx_signal, reconstruction_mode="Auto"):
    def coverage(src, dst):
        if not src:
            return 1.0
        return len(src.intersection(dst)) / len(src)

    text_cov = coverage(pdf_signal["tokens"], docx_signal["tokens"])
    numeric_cov = coverage(pdf_signal["numeric_tokens"], docx_signal["numeric_tokens"])
    barcode_cov = coverage(pdf_signal["barcode_tokens"], docx_signal["barcode_tokens"])
    field_cov = coverage(pdf_signal["field_labels"], docx_signal["field_labels"])
    if pdf_signal["image_count"] > 0:
        image_cov = min(docx_signal["media_count"] / pdf_signal["image_count"], 1.0)
    else:
        image_cov = 1.0
    table_cov = 1.0 if docx_signal["table_count"] > 0 else 0.0
    page_count = max(1, pdf_signal.get("page_count", 0))
    section_cov = min(docx_signal.get("section_count", 0) / page_count, 1.0)
    anchor_penalty = min(docx_signal["anchor_count"] / 4.0, 1.0)
    token_count = pdf_signal.get("total_token_count", 0)
    image_count = pdf_signal.get("image_count", 0)
    drawing_count = pdf_signal.get("drawing_count", 0)
    token_density = token_count / page_count
    image_density = image_count / page_count
    drawing_density = drawing_count / page_count

    # Keep floating objects for visually-rich docs (slide decks, brochures, image-heavy PDFs).
    if pdf_signal.get("page_count", 0) >= 2 and (
        token_count <= 200 or token_density <= 80 or image_density >= 0.7 or drawing_density >= 8
    ):
        anchor_weight = 0.0
    elif image_density >= 0.3:
        anchor_weight = 0.05
    else:
        anchor_weight = 0.15
    if pdf_signal["total_token_count"] > 0:
        inflation_ratio = docx_signal["total_token_count"] / pdf_signal["total_token_count"]
    else:
        inflation_ratio = 1.0
    inflation_penalty = min(max(inflation_ratio - 1.8, 0.0) / 2.0, 1.0)

    score = (
        (0.35 * text_cov)
        + (0.15 * numeric_cov)
        + (0.20 * barcode_cov)
        + (0.15 * field_cov)
        + (0.10 * image_cov)
        + (0.05 * table_cov)
        - (anchor_weight * anchor_penalty)
        - (0.10 * inflation_penalty)
    )
    if reconstruction_mode == "Continuous Reconstruction":
        score += 0.08 * section_cov
    elif reconstruction_mode == "Flowing Reconstruction":
        score += 0.03 * (1.0 - min(abs(section_cov - 0.35), 1.0))

    metrics = {
        "text_coverage": text_cov,
        "numeric_coverage": numeric_cov,
        "barcode_coverage": barcode_cov,
        "field_label_coverage": field_cov,
        "image_coverage": image_cov,
        "table_coverage": table_cov,
        "section_coverage": section_cov,
        "anchor_penalty": anchor_penalty,
        "anchor_weight": anchor_weight,
        "inflation_ratio": inflation_ratio,
        "inflation_penalty": inflation_penalty,
    }
    return max(score, 0.0), metrics


def is_visual_layout_pdf(pdf_signal):
    page_count = max(1, pdf_signal.get("page_count", 0))
    token_count = pdf_signal.get("total_token_count", 0)
    image_count = pdf_signal.get("image_count", 0)
    drawing_count = pdf_signal.get("drawing_count", 0)
    token_density = token_count / page_count
    image_density = image_count / page_count
    drawing_density = drawing_count / page_count
    return pdf_signal.get("page_count", 0) >= 2 and (
        token_count <= 200 or token_density <= 80 or image_density >= 0.7 or drawing_density >= 8
    )


# --- Per-page classification & hybrid routing ---------------------------------
# Thresholds are intentionally named module constants so they can be tuned
# against real-world PDFs without touching the routing logic below.
PAGE_MIN_TEXT_CHARS = 100          # at/above this a page is considered to carry real text
PAGE_SCANNED_MAX_CHARS = 20        # at/below this the page has effectively no text layer
PAGE_SCANNED_MIN_IMAGE_RATIO = 0.55  # raster coverage that marks an image-only (scanned) page
PAGE_HYBRID_MIN_IMAGE_RATIO = 0.20   # raster coverage that makes a text page also "hybrid"
PAGE_GRAPHICS_MIN_DRAWINGS = 800   # vector-path count typical of charts/plots/diagrams
PAGE_GRAPHICS_MAX_CHARS = 500      # graphics pages carry only labels, not paragraphs
# A "form" page is a dense fixed-layout page (challan, invoice, certificate):
# many border/box drawings AND several embedded images (logos/photos/barcodes).
# These reconstruct poorly as editable Word tables, so they are rendered to an
# image for pixel-faithful layout.
PAGE_FORM_MIN_DRAWINGS = 30
PAGE_FORM_MIN_IMAGES = 3


def classify_page_kind(char_count, image_ratio, drawing_count, image_count=0):
    """Classify a single page from its extracted-text/image/vector signals.

    Returns one of: "text", "scanned", "hybrid", "graphics", "form".
    """
    near_no_text = char_count < PAGE_SCANNED_MAX_CHARS
    has_text = char_count >= PAGE_MIN_TEXT_CHARS

    # Image-only page with no usable text layer -> needs OCR.
    if near_no_text and image_ratio >= PAGE_SCANNED_MIN_IMAGE_RATIO:
        return "scanned"
    # Vector-heavy page with little text (charts, plots, schematics) -> image fidelity.
    if drawing_count >= PAGE_GRAPHICS_MIN_DRAWINGS and char_count < PAGE_GRAPHICS_MAX_CHARS:
        return "graphics"
    # Dense fixed-layout form (boxes + multiple images) -> image fidelity, since
    # editable table reconstruction inflates cell heights and breaks alignment.
    if drawing_count >= PAGE_FORM_MIN_DRAWINGS and image_count >= PAGE_FORM_MIN_IMAGES:
        return "form"
    # Real text plus meaningful raster imagery -> OCR fills the image regions.
    if has_text and image_ratio >= PAGE_HYBRID_MIN_IMAGE_RATIO:
        return "hybrid"
    return "text"


def classify_pdf_pages(pdf_path):
    """Inspect every page with PyMuPDF and classify it before conversion.

    Returns a list of per-page dicts: {index, kind, char_count,
    image_area_ratio, drawing_count}. Empty list if PyMuPDF is unavailable.
    """
    results = []
    if fitz is None:
        return results
    with fitz.open(pdf_path) as doc:
        for index, page in enumerate(doc):
            text = (page.get_text() or "").strip()
            char_count = len(text)
            page_area = abs(page.rect.width * page.rect.height) or 1.0
            image_area = 0.0
            try:
                for info in page.get_image_info():
                    bbox = info.get("bbox")
                    if bbox:
                        image_area += abs((bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
            except Exception:
                pass
            image_ratio = min(image_area / page_area, 1.0)
            try:
                image_count = len(page.get_images())
            except Exception:
                image_count = 0
            try:
                drawing_count = len(page.get_drawings())
            except Exception:
                drawing_count = 0
            results.append(
                {
                    "index": index,
                    "kind": classify_page_kind(char_count, image_ratio, drawing_count, image_count),
                    "char_count": char_count,
                    "image_area_ratio": round(image_ratio, 3),
                    "image_count": image_count,
                    "drawing_count": drawing_count,
                }
            )
    return results


def summarize_page_kinds(page_infos):
    """Reduce per-page classification into a compact summary for routing/diagnostics."""
    counts = Counter(info["kind"] for info in page_infos)
    kinds = set(counts)
    return {
        "counts": dict(counts),
        "page_count": len(page_infos),
        "needs_ocr": bool(kinds & {"scanned", "hybrid"}),
        "needs_image_fallback": "graphics" in kinds,
        "all_text": kinds == {"text"} or not kinds,
        "all_scanned": kinds == {"scanned"},
        "kinds_sequence": [info["kind"] for info in page_infos],
    }


def _ocrmypdf_invocation():
    """Return the argv prefix to invoke OCRmyPDF, or None if unavailable.

    Prefers `python -m ocrmypdf` with the current interpreter so it runs in the
    same environment that has the package installed (avoids picking up a system
    binary tied to a different, possibly broken, Python). Falls back to a
    standalone `ocrmypdf` on PATH.
    """
    import importlib.util

    if importlib.util.find_spec("ocrmypdf") is not None:
        return [sys.executable, "-m", "ocrmypdf"]
    standalone = shutil.which("ocrmypdf")
    if standalone:
        return [standalone]
    return None


def enhance_pdf_with_text_layer(pdf_path, page_infos, lang="eng", out_path=None):
    """Inject an OCR text layer onto scanned/hybrid pages via OCRmyPDF.

    Uses --skip-text so pages that already have a text layer are left untouched;
    only image pages get OCR'd. Returns a dict describing the result. On any
    failure (missing binary, OCRmyPDF error) it degrades gracefully and reports
    applied=False so callers fall back to the original PDF.
    """
    ocr_pages = [info["index"] for info in page_infos if info["kind"] in ("scanned", "hybrid")]
    if not ocr_pages:
        return {"applied": False, "reason": "no_ocr_pages", "pdf_path": pdf_path}

    invocation = _ocrmypdf_invocation()
    if not invocation:
        return {"applied": False, "reason": "ocrmypdf_unavailable", "pdf_path": pdf_path}

    if out_path is None:
        fd, out_path = tempfile.mkstemp(suffix="_ocr.pdf")
        os.close(fd)

    cmd = invocation + [
        "--skip-text",          # keep existing text layers; OCR only image pages
        "--optimize", "0",
        "--language", lang,
        "--quiet",
        pdf_path,
        out_path,
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    except Exception as exc:
        return {"applied": False, "reason": f"ocrmypdf_exception: {exc}", "pdf_path": pdf_path}

    # Exit code 0 == success. OCRmyPDF returns specific non-zero codes; treat any
    # non-zero as a soft failure and fall back to the original PDF.
    if proc.returncode != 0 or not (os.path.exists(out_path) and os.path.getsize(out_path) > 0):
        return {
            "applied": False,
            "reason": f"ocrmypdf_failed(code={proc.returncode}): {proc.stderr.strip()[:300]}",
            "pdf_path": pdf_path,
        }

    return {
        "applied": True,
        "reason": "ok",
        "pdf_path": out_path,
        "ocr_pages": ocr_pages,
        "lang": lang,
    }


def rasterize_pdf_pages(pdf_path, page_indices, dpi=200, jpg_quality=85, out_path=None):
    """Replace the given pages with a flattened raster image of themselves.

    Vector-heavy pages (charts/plots with thousands of paths) are extremely slow
    for pdf2docx to parse and produce poor editable output. Rasterizing such a
    page to a single JPEG-backed page keeps it visually identical, makes the
    conversion ~100x faster, and lets the rest of the document convert normally.
    Other pages are copied through untouched. Returns a dict with the output
    path and which pages were rasterized.
    """
    page_set = set(page_indices or [])
    if not page_set or fitz is None:
        return {"applied": False, "reason": "nothing_to_rasterize", "pdf_path": pdf_path}

    if out_path is None:
        fd, out_path = tempfile.mkstemp(suffix="_raster.pdf")
        os.close(fd)

    rasterized = []
    src = fitz.open(pdf_path)
    out = fitz.open()
    try:
        for index, page in enumerate(src):
            rect = page.rect
            if index in page_set:
                pix = page.get_pixmap(dpi=dpi, alpha=False)
                jpeg_bytes = pix.tobytes("jpeg", jpg_quality=jpg_quality)
                new_page = out.new_page(width=rect.width, height=rect.height)
                new_page.insert_image(rect, stream=jpeg_bytes)
                rasterized.append(index)
            else:
                out.insert_pdf(src, from_page=index, to_page=index)
        out.save(out_path, garbage=3, deflate=True)
    finally:
        out.close()
        src.close()

    return {"applied": True, "reason": "ok", "pdf_path": out_path, "rasterized_pages": rasterized}


def resolve_ocr_language(job):
    """Resolve the Tesseract/OCRmyPDF language code for a conversion job."""
    lang_attr = getattr(job, "language", None)
    if lang_attr is not None and hasattr(lang_attr, "get"):
        try:
            return PDF2Doc.LANG_MAP.get(lang_attr.get(), "eng")
        except Exception:
            return "eng"
    if isinstance(lang_attr, str) and lang_attr:
        return PDF2Doc.LANG_MAP.get(lang_attr, lang_attr)
    return os.environ.get("PDF2DOCX_OCR_LANG", "eng")


# GUI-only code. It stays import-safe for server/API deployment because tkinter
# is optional above and these classes are only instantiated by launch_gui().
class about(ttk.Frame):
    def __init__(self, parent):

        textAboutMe = """
        Hello! My name is Süleyman Bozkurt and I am a senior PhD student in the field of biochemistry. 
        My research focuses on mitochondrial protein import and proteomics, 
        and I am particularly interested in understanding how it's impact on cellular function.
        In addition to my research, I am also interested in building my own apps and sharing them with others. 
        I am a huge fan of Python and enjoy using R for visualizations. I believe that these tools 
        can help make complex scientific data more accessible and easier to understand.
        I am passionate about using my skills and knowledge to make a positive impact in the scientific community. 
        I am always looking for new ways to learn and grow, and I enjoy collaborating with others 
        to find innovative solutions to complex problems.
        I hope this gives you a sense of who I am and what I am interested in. 
        If you have any questions or would like to learn more about my work, please don't hesitate to contact me.
        """

        ttk.Frame.__init__(self, parent)
        self.font = Font(family="Times New Roman", size=20)
        label2 = Label(self, text="")
        label2.grid(padx=1000, pady=800)
        label2 = Label(self, text="ABOUT ME!",font=self.font)
        label2.place(x=420, y=35)

        label = tk.Label(self, font = Font(family="Times New Roman", size=14),
                         text = textAboutMe,
                         )

        label.place(x=40, y=75)

        # Open the image using PIL
        self.github_logo = Image.open("./files/github.png")
        self.linkedin_logo = Image.open("./files/linkedin.png")
        self.twitter_logo = Image.open("./files/twitter.png")
        self.ibc2_logo = Image.open("./files/ibc2logo.png")

        # Resize the image
        self.github_logo = self.github_logo.resize((100, 100), resample=Image.Resampling.LANCZOS)
        self.linkedin_logo = self.linkedin_logo.resize((100, 100), resample=Image.Resampling.LANCZOS)
        self.twitter_logo = self.twitter_logo.resize((100, 100), resample=Image.Resampling.LANCZOS)
        self.ibc2_logo = self.ibc2_logo.resize((100, 100), resample=Image.Resampling.LANCZOS)

        # Create a PhotoImage object from the resized image
        self.github_logo = ImageTk.PhotoImage(self.github_logo)
        self.linkedin_logo = ImageTk.PhotoImage(self.linkedin_logo)
        self.twitter_logo = ImageTk.PhotoImage(self.twitter_logo)
        self.ibc2_logo = ImageTk.PhotoImage(self.ibc2_logo)

        # Create clickable link to your website
        website_link = tk.Label(self, image=self.ibc2_logo, text="Website", fg="blue", cursor="hand2")
        website_link.place(x=220, y=370)
        website_link.bind("<Button-1>", self.open_website)

        # Create clickable link to your GitHub page with logo
        github_link = tk.Button(self, image=self.github_logo, cursor="hand2", bd=0, highlightthickness=0,
                                activebackground="white")
        github_link.place(x=370, y=370)
        github_link.bind("<Button-1>", self.open_github)

        # Create clickable link to your LinkedIn profile with logo
        linkedin_link = tk.Button(self, image=self.linkedin_logo, cursor="hand2", bd=0, highlightthickness=0,
                                  activebackground="white")
        linkedin_link.place(x=520, y=370)
        linkedin_link.bind("<Button-1>", self.open_linkedin)

        # Create clickable link to your Twitter profile with logo
        twitter_link = tk.Button(self, image=self.twitter_logo, cursor="hand2", bd=0, highlightthickness=0,
                                 activebackground="white")
        twitter_link.place(x=670, y=370)
        twitter_link.bind("<Button-1>", self.open_twitter)

        label2 = Label(self, text="You can reach me one of these profiles!",font=self.font)
        label2.place(x=270, y=500)

    def open_github(self, event):
        # Open your GitHub page in the default web browser
        import webbrowser
        webbrowser.open("https://github.com/science64")

    def open_website(self, event):
        # Open your website in the default web browser
        import webbrowser
        webbrowser.open("https://biochem2.com/people/bozkurt-sueleyman")

    def open_linkedin(self, event):
        # Open your LinkedIn profile in the default web browser
        import webbrowser
        webbrowser.open("https://www.linkedin.com/in/s%C3%BCleyman-bozkurt-51080b60/")

    def open_twitter(self, event):
        # Open your LinkedIn profile in the default web browser
        import webbrowser
        webbrowser.open("https://twitter.com/science_mbg")

    def callback(self, url):
        webbrowser.open_new(url)

class support(ttk.Frame):
    def __init__(self, parent):
        ttk.Frame.__init__(self, parent)
        self.font = Font(family="Times New Roman", size=16)

        img = Image.open("files/buymecofee.jpeg")
        img = img.resize((800, 400), Image.Resampling.LANCZOS)
        img = ImageTk.PhotoImage(img)
        panel = Label(self, image=img)
        panel.image = img
        panel.place(x=80, y=20, width=800, height=400)

        label2 = Label(self, text="You can support by buying a coffee ☕️ here ", font=self.font)
        label2.grid(padx=1000, pady=800)
        label2.place(x=300, y=450)

        self.clickme = Button(self, text='https://www.buymeacoffee.com/science64', fg='black', bg='#FFC19E',
                                font=Font(family="Times New Roman", size=18, weight='bold'), command=self.callback)
        self.clickme.place(x=270, y=500)

    def callback(self):
        webbrowser.open_new('https://www.buymeacoffee.com/science64')

class PDF2Doc(ttk.Frame):
    FORCE_STRONG_CONTINUOUS = True
    LANG_MAP = {
        "English": "eng",
        "French": "fra",
        "German": "deu",
        "Turkish": "tur",
        "Polish": "pol",
        "Spanish": "spa",
        "Italian": "ita",
        "Arabic": "ara",
        "Chinese": "chi_sim",
        "Hindi": "hin",
        "Croatian": "hrv",
        "Russian": "rus",
        "Ukrainian": "ukr",
        "Persian": "fas",
    }

    def __init__(self, parent):
        ttk.Frame.__init__(self, parent)

        self.pdf_label = tk.Label(self, text="Select a PDF file:", font=("Helvetica", 14), pady=10)
        self.pdf_label.pack()

        self.pdf_button = tk.Button(self, text="Browse", font=("Helvetica", 14), pady=10, command=self.select_pdf)
        self.pdf_button.pack()

        self.conversion_method_label = tk.Label(self, text="Select a conversion method:", font=("Helvetica", 14), pady=10)
        self.conversion_method_label.pack()

        self.conversion_method_var = tk.IntVar()

        # Set the default value of the IntVar to 1
        self.conversion_method_var.set(1)

        self.pdf_to_word_radio = tk.Radiobutton(self, text="Preserve Format (Editable)", value=1, font=("Helvetica", 14), pady=10,
                                                variable=self.conversion_method_var)
        self.pdf_to_word_radio.pack()

        self.ocr_radio = tk.Radiobutton(self, text="OCR (Optical Character Recognition)", value=2, font=("Helvetica", 14), pady=10,
                                        variable=self.conversion_method_var)
        self.ocr_radio.pack()

        self.layout_radio = tk.Radiobutton(
            self,
            text="Exact Layout (No Overlap, image-based)",
            value=3,
            font=("Helvetica", 14),
            pady=10,
            variable=self.conversion_method_var,
        )
        self.layout_radio.pack()

        self.reconstruction_mode_label = tk.Label(
            self,
            text="Reconstruction Mode (Editable):",
            font=("Helvetica", 12),
            pady=6,
        )
        self.reconstruction_mode_label.pack()

        self.reconstruction_mode = StringVar()
        self.reconstruction_mode.set("Continuous Reconstruction")
        reconstruction_modes = ["Auto", "Flowing Reconstruction", "Continuous Reconstruction"]
        self.reconstruction_options = OptionMenu(self, self.reconstruction_mode, *reconstruction_modes)
        self.reconstruction_options.pack()

        self.language = StringVar()
        self.language.set("English")  # default value

        language_formats = ["English", "German", "Turkish",
                            "French", "Spanish", "Italian",
                            "Polish", "Russian", "Ukrainian", "Croatian",
                            "Arabic", "Chinese", "Hindi", "Persian"] # 13 different language format so far supported!

        self.options = OptionMenu(self, self.language, *language_formats)
        self.options.pack()

        # self.output_text = tk.Text(root, height=40, width=30, font=("Helvetica", 14))
        # self.output_text.pack(padx=10, pady=10)
        self.output_frame = tk.Frame(self)
        self.output_frame.pack(padx=10, pady=10)

        # Create a Scrollbar widget and associate it with the Text widget
        self.scrollbar = tk.Scrollbar(self.output_frame, orient="vertical")
        self.scrollbar.pack(side="right", fill="y")

        # Create the Text widget and associate it with the Scrollbar widget
        self.output_text = tk.Text(self.output_frame, height=10, width=60, font=("Helvetica", 14),
                                   yscrollcommand=self.scrollbar.set)
        self.output_text.pack(side="left", fill="both", expand=True)

        # Set the command of the Scrollbar widget to the yview method of the Text widget
        self.scrollbar.config(command=self.output_text.yview)

        self.button_frame = tk.Frame(self)
        self.button_frame.pack()

        self.convert_button = tk.Button(self.button_frame, text="Convert", font=("Helvetica", 14), pady=10,
                                        command=self.convert,
                                        state=tk.DISABLED)
        self.convert_button.pack(side=tk.LEFT)

        self.open_button = tk.Button(self.button_frame, text="Open", font=("Helvetica", 14), pady=10,
                                        command=self.open_docx,)
        self.open_button.pack(side=tk.RIGHT)

    def select_pdf(self):
        try:
            self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if self.pdf_path:
                self.output_text.insert(tk.END, "- Selected PDF: " + self.pdf_path + "\n")
                self.convert_button.config(state=tk.NORMAL)
        except Exception:
            # Display an error message in a pop-up window
            tkinter.messagebox.showerror("Error",
                                         "An error occurred while selecting the PDF file.\n\n" + traceback.format_exc())

    def convert(self):
        # Create a new thread to perform the conversion
        conversion_thread = threading.Thread(target=self.conversion_worker)
        conversion_thread.start()

    def open_docx(self):
        try:
            # Open the docx file with the default program
            open_with_default_app(self.docx_path)
        except Exception:
            import tkinter.messagebox

            tkinter.messagebox.showerror("Error",
                                         "An error occurred while opening the docx file.\n\n" + traceback.format_exc())

    def _resolve_poppler_path(self):
        bundled = os.path.join(os.path.dirname(__file__), "files", "poppler-22.04.0", "Library", "bin")
        if sys.platform.startswith("win") and os.path.isfile(os.path.join(bundled, "pdftoppm.exe")):
            return bundled
        system_pdftoppm = shutil.which("pdftoppm")
        if system_pdftoppm:
            return os.path.dirname(system_pdftoppm)
        return None

    def _resolve_tesseract_cmd(self):
        bundled = os.path.join(os.path.dirname(__file__), "files", "Tesseract-OCR", "tesseract.exe")
        if sys.platform.startswith("win") and os.path.isfile(bundled):
            return bundled
        system_cmd = shutil.which("tesseract")
        return system_cmd

    def _iter_rendered_pages(self, dpi=300):
        # Prefer PyMuPDF (fitz) for rendering EVERYWHERE so the output is
        # identical between local dev (often no Poppler installed) and the
        # deployed Linux container (Poppler present). Poppler and fitz produce
        # different pixmaps/dimensions, which previously made the same PDF look
        # different after deploy. rasterize_pdf_pages() already standardises on
        # fitz, so this keeps the whole pipeline on one renderer.
        if fitz is not None:
            with fitz.open(self.pdf_path) as doc:
                for page in doc:
                    pix = page.get_pixmap(dpi=dpi, alpha=False)
                    image_bytes = pix.tobytes("png")
                    image = Image.open(io.BytesIO(image_bytes))
                    yield image.convert("RGB")
            return

        # Fallback only when PyMuPDF is unavailable (e.g. minimal Windows GUI
        # build that ships bundled Poppler instead).
        poppler_path = self._resolve_poppler_path()
        if poppler_path:
            kwargs = {"dpi": dpi, "poppler_path": poppler_path}
            info = pdfinfo_from_path(self.pdf_path, poppler_path=poppler_path)
            page_count = int(info.get("Pages", 0))
            for page_no in range(1, page_count + 1):
                images = convert_from_path(
                    self.pdf_path,
                    first_page=page_no,
                    last_page=page_no,
                    **kwargs,
                )
                if images:
                    yield images[0]
            return

        raise RuntimeError(
            "No PDF renderer found. Install PyMuPDF, or Poppler (pdftoppm in PATH)."
        )

    def _preprocess_ocr_image(self, image):
        if cv2 is None or np is None:
            return image

        rgb = np.array(image.convert("RGB"))
        gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
        gray = cv2.fastNlMeansDenoising(gray, None, 12, 7, 21)
        binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

        coords = np.column_stack(np.where(binary < 255))
        if len(coords) > 100:
            angle = cv2.minAreaRect(coords)[-1]
            if angle < -45:
                angle = -(90 + angle)
            else:
                angle = -angle
            if abs(angle) > 0.15:
                h, w = binary.shape[:2]
                center = (w // 2, h // 2)
                rot = cv2.getRotationMatrix2D(center, angle, 1.0)
                binary = cv2.warpAffine(
                    binary,
                    rot,
                    (w, h),
                    flags=cv2.INTER_CUBIC,
                    borderMode=cv2.BORDER_REPLICATE,
                )
        return Image.fromarray(binary)

    def _run_pdf2docx(self):
        pdf_signal = analyze_pdf_quality_signal(self.pdf_path)
        reconstruction_mode = self.reconstruction_mode.get() if hasattr(self, "reconstruction_mode") else "Auto"
        if self.FORCE_STRONG_CONTINUOUS:
            reconstruction_mode = "Continuous Reconstruction"
        cpu_cap = os.cpu_count() or 1
        multiprocessing_settings = {}
        if pdf_signal["page_count"] >= 4 and __name__ == "__main__":
            multiprocessing_settings = {
                "multi_processing": True,
                "cpu_count": min(4, cpu_cap),
            }

        base_settings = {
            "ignore_page_error": False,
            "raw_exceptions": True,
            "delete_end_line_hyphen": True,
            # 0 = ignore hidden text (default), 2 = extract ONLY hidden text.
            # The auto pipeline sets 2 when converting an OCR sandwich PDF whose
            # text layer is invisible (added by OCRmyPDF on scanned pages).
            "ocr": getattr(self, "_pdf2docx_ocr_mode", 0),
        }
        profiles = [
            (
                "balanced_layout",
                {
                    **base_settings,
                    **multiprocessing_settings,
                    "parse_stream_table": False,
                    "list_not_table": True,
                    "line_break_width_ratio": 0.35,
                    "line_break_free_space_ratio": 0.20,
                    "line_separate_threshold": 3.0,
                    "new_paragraph_free_space_ratio": 0.70,
                    "lines_center_aligned_threshold": 1.2,
                    "float_image_ignorable_gap": 2.0,
                },
            ),
            (
                "continuous_strict",
                {
                    **base_settings,
                    **multiprocessing_settings,
                    "parse_stream_table": True,
                    "extract_stream_table": True,
                    "parse_lattice_table": True,
                    "list_not_table": True,
                    "page_margin_factor_top": 0.25,
                    "page_margin_factor_bottom": 0.25,
                    "line_break_width_ratio": 0.78,
                    "line_break_free_space_ratio": 0.03,
                    "line_separate_threshold": 1.4,
                    "new_paragraph_free_space_ratio": 0.45,
                    "lines_left_aligned_threshold": 0.8,
                    "lines_right_aligned_threshold": 0.8,
                    "lines_center_aligned_threshold": 0.7,
                    "max_line_spacing_ratio": 1.3,
                    "float_image_ignorable_gap": 0.6,
                    "min_border_clearance": 1.0,
                },
            ),
            (
                "continuous_layout",
                {
                    **base_settings,
                    **multiprocessing_settings,
                    "parse_stream_table": True,
                    "extract_stream_table": True,
                    "parse_lattice_table": True,
                    "list_not_table": True,
                    "page_margin_factor_top": 1.0,
                    "page_margin_factor_bottom": 1.0,
                    "line_break_width_ratio": 0.65,
                    "line_break_free_space_ratio": 0.05,
                    "line_separate_threshold": 2.0,
                    "new_paragraph_free_space_ratio": 0.55,
                    "lines_center_aligned_threshold": 0.8,
                    "float_image_ignorable_gap": 0.8,
                },
            ),
            (
                "table_recovery",
                {
                    **base_settings,
                    **multiprocessing_settings,
                    "parse_stream_table": True,
                    "extract_stream_table": True,
                    "parse_lattice_table": True,
                    "list_not_table": False,
                    "line_break_width_ratio": 0.50,
                    "line_break_free_space_ratio": 0.10,
                    "line_separate_threshold": 5.0,
                    "new_paragraph_free_space_ratio": 0.85,
                    "float_image_ignorable_gap": 1.0,
                    "min_border_clearance": 1.0,
                },
            ),
            (
                "continuous_table_strict",
                {
                    **base_settings,
                    **multiprocessing_settings,
                    "parse_stream_table": True,
                    "extract_stream_table": True,
                    "parse_lattice_table": True,
                    "list_not_table": False,
                    "page_margin_factor_top": 0.25,
                    "page_margin_factor_bottom": 0.25,
                    "line_break_width_ratio": 0.70,
                    "line_break_free_space_ratio": 0.04,
                    "line_separate_threshold": 1.8,
                    "new_paragraph_free_space_ratio": 0.50,
                    "lines_left_aligned_threshold": 0.8,
                    "lines_right_aligned_threshold": 0.8,
                    "lines_center_aligned_threshold": 0.7,
                    "max_line_spacing_ratio": 1.25,
                    "float_image_ignorable_gap": 0.6,
                    "min_border_clearance": 0.8,
                },
            ),
            (
                "barcode_image_focus",
                {
                    **base_settings,
                    **multiprocessing_settings,
                    "parse_stream_table": True,
                    "extract_stream_table": True,
                    "parse_lattice_table": True,
                    "list_not_table": True,
                    "float_image_ignorable_gap": 0.1,
                    "shape_min_dimension": 0.5,
                    "min_svg_w": 0.5,
                    "min_svg_h": 0.5,
                    "clip_image_res_ratio": 6.0,
                    "line_overlap_threshold": 0.97,
                },
            ),
        ]
        profile_map = {name: settings for name, settings in profiles}
        if reconstruction_mode == "Flowing Reconstruction":
            profile_sequence = [
                ("balanced_layout", profile_map["balanced_layout"]),
                ("table_recovery", profile_map["table_recovery"]),
                ("continuous_layout", profile_map["continuous_layout"]),
                ("continuous_strict", profile_map["continuous_strict"]),
            ]
        elif reconstruction_mode == "Continuous Reconstruction":
            profile_sequence = [
                ("continuous_strict", profile_map["continuous_strict"]),
                ("continuous_layout", profile_map["continuous_layout"]),
                ("continuous_table_strict", profile_map["continuous_table_strict"]),
                ("barcode_image_focus", profile_map["barcode_image_focus"]),
            ]
        else:
            profile_sequence = [
                ("balanced_layout", profile_map["balanced_layout"]),
                ("continuous_strict", profile_map["continuous_strict"]),
                ("continuous_layout", profile_map["continuous_layout"]),
                ("table_recovery", profile_map["table_recovery"]),
                ("barcode_image_focus", profile_map["barcode_image_focus"]),
            ]
        max_profiles = getattr(self, "max_profiles", None)
        if max_profiles:
            try:
                profile_sequence = profile_sequence[:max(1, int(max_profiles))]
            except Exception:
                pass

        best_result = None
        attempts = []
        with tempfile.TemporaryDirectory(prefix="pdf2docx_profiles_") as temp_dir:
            for idx, (profile_name, profile_settings) in enumerate(profile_sequence, start=1):
                candidate_docx = os.path.join(temp_dir, f"candidate_{idx}.docx")
                try:
                    cv = Converter(self.pdf_path)
                    try:
                        cv.convert(candidate_docx, **profile_settings)
                    finally:
                        cv.close()

                    pre_signal = analyze_docx_quality_signal(candidate_docx)
                    pre_score, pre_metrics = score_conversion_quality(
                        pdf_signal,
                        pre_signal,
                        reconstruction_mode=reconstruction_mode,
                    )
                    pre_metrics["flattened_shapes"] = 0
                    pre_metrics["flatten_applied"] = False

                    flat_docx = os.path.join(temp_dir, f"candidate_{idx}_flat.docx")
                    shutil.copy2(candidate_docx, flat_docx)
                    flattened = flatten_floating_shapes_in_docx(flat_docx)
                    flat_signal = analyze_docx_quality_signal(flat_docx)
                    flat_score, flat_metrics = score_conversion_quality(
                        pdf_signal,
                        flat_signal,
                        reconstruction_mode=reconstruction_mode,
                    )
                    flat_metrics["flattened_shapes"] = flattened
                    flat_metrics["flatten_applied"] = True

                    prefers_floating_layout = is_visual_layout_pdf(pdf_signal)
                    keep_floating_margin = 0.02
                    if (
                        prefers_floating_layout
                        and pre_signal["anchor_count"] > 0
                        and pre_score >= (flat_score - keep_floating_margin)
                    ):
                        chosen_docx = candidate_docx
                        score = pre_score
                        metrics = pre_metrics
                    elif flat_score > pre_score:
                        chosen_docx = flat_docx
                        score = flat_score
                        metrics = flat_metrics
                    else:
                        chosen_docx = candidate_docx
                        score = pre_score
                        metrics = pre_metrics

                    attempts.append((profile_name, score, metrics, None))

                    if best_result is None or score > best_result["score"]:
                        best_result = {
                            "profile_name": profile_name,
                            "score": score,
                            "metrics": metrics,
                            "candidate_docx": chosen_docx,
                        }

                    if (
                        metrics["text_coverage"] >= 0.97
                        and metrics["numeric_coverage"] >= 0.90
                        and metrics["barcode_coverage"] >= 0.85
                        and metrics["field_label_coverage"] >= 0.90
                        and metrics["image_coverage"] >= 0.70
                        and (
                            metrics["anchor_penalty"] <= 0.0
                            or metrics.get("anchor_weight", 0.15) == 0.0
                        )
                    ):
                        break
                except Exception as profile_err:
                    attempts.append((profile_name, 0.0, None, str(profile_err)))

            if best_result is None:
                detail = "; ".join(
                    f"{name}: {err}" for name, _, _, err in attempts if err
                )
                raise RuntimeError(f"All conversion profiles failed. {detail}")

            single_page_target = pdf_signal.get("page_count", 0) == 1
            post_variants = []
            base_overlap = compute_layout_overlap_risk(best_result["candidate_docx"])
            should_try_deoverlap = (
                base_overlap.get("risk", 0.0) > 0.08
                or base_overlap.get("front_large_anchors", 0) > 0
            )

            with tempfile.TemporaryDirectory(prefix="pdf2docx_post_variants_") as post_dir:
                variant_specs = [
                    ("balanced", False, False),
                ]
                if should_try_deoverlap:
                    force_flatten_deoverlap = (
                        base_overlap.get("front_large_anchors", 0) > 0
                        and not is_visual_layout_pdf(pdf_signal)
                    )
                    aggressive_deoverlap = False
                    variant_specs.append(("deoverlap", force_flatten_deoverlap, aggressive_deoverlap))
                for variant_name, force_flatten, aggressive_spacing in variant_specs:
                    variant_docx = os.path.join(post_dir, f"{variant_name}.docx")
                    shutil.copy2(best_result["candidate_docx"], variant_docx)

                    extra_flattened = 0
                    if force_flatten:
                        extra_flattened = flatten_floating_shapes_in_docx(variant_docx)

                    anchors_moved_behind = force_large_anchors_behind_text(variant_docx)
                    dedup_removed = 0
                    if best_result["metrics"].get("inflation_ratio", 1.0) > 2.2:
                        dedup_removed = prune_docx_redundant_lines(
                            variant_docx,
                            self.pdf_path,
                            min_len=28,
                            unknown_line_limit=1,
                        )
                    hyphen_cleaned = remove_soft_hyphens_in_docx(variant_docx)
                    layout_relax = relax_table_layout_constraints(
                        variant_docx,
                        single_page_target=single_page_target,
                        aggressive=aggressive_spacing,
                    )
                    single_page_fit = enforce_single_page_layout(variant_docx, self.pdf_path)
                    barcode_cleanup = remove_barcode_text_artifacts(
                        variant_docx,
                        min_digits=6,
                        min_repeat=2,
                    )

                    post_signal = analyze_docx_quality_signal(variant_docx)
                    post_score, post_metrics = score_conversion_quality(
                        pdf_signal,
                        post_signal,
                        reconstruction_mode=reconstruction_mode,
                    )
                    overlap_info = compute_layout_overlap_risk(variant_docx)
                    overlap_risk = overlap_info.get("risk", 0.0)
                    combined_score = post_score - (0.35 * overlap_risk)

                    flatten_base = best_result["metrics"].get("flattened_shapes", 0)
                    flatten_applied = best_result["metrics"].get("flatten_applied", False) or (extra_flattened > 0)
                    preserve_keys = {
                        "flatten_applied": flatten_applied,
                        "flattened_shapes": flatten_base + extra_flattened,
                        "anchors_forced_behind_text": anchors_moved_behind,
                        "soft_hyphen_cleaned": bool(hyphen_cleaned),
                        "dedup_removed_paragraphs": dedup_removed,
                        "layout_rows_relaxed": layout_relax.get("rows_relaxed", 0),
                        "layout_spacing_relaxed": layout_relax.get("spacing_relaxed", 0),
                        "layout_cell_margins_compacted": layout_relax.get("cell_margins_compacted", 0),
                        "layout_blank_paragraphs_removed": layout_relax.get("blank_paragraphs_removed", 0),
                        "layout_page_breaks_removed": layout_relax.get("page_breaks_removed", 0),
                        "single_page_target": single_page_target,
                        "single_page_size_set": single_page_fit.get("page_size_set", False),
                        "single_page_margins_compacted": single_page_fit.get("margins_compacted", False),
                        "single_page_breaks_removed": single_page_fit.get("breaks_removed", 0),
                        "barcode_text_runs_removed": barcode_cleanup.get("removed_runs", 0),
                        "barcode_text_paragraphs_removed": barcode_cleanup.get("removed_paragraphs", 0),
                        "barcode_images_inserted": barcode_cleanup.get("images_inserted", 0),
                        "postprocess_variant": variant_name,
                        "overlap_risk": overlap_risk,
                        "overlap_cramped_cells": overlap_info.get("cramped_cells", 0),
                        "overlap_cramped_rows": overlap_info.get("cramped_rows", 0),
                        "overlap_front_large_anchors": overlap_info.get("front_large_anchors", 0),
                    }
                    post_metrics.update(preserve_keys)
                    post_variants.append(
                        {
                            "name": variant_name,
                            "docx": variant_docx,
                            "quality_score": post_score,
                            "combined_score": combined_score,
                            "metrics": post_metrics,
                        }
                    )

                balanced_post = next((v for v in post_variants if v["name"] == "balanced"), post_variants[0])
                deoverlap_post = next((v for v in post_variants if v["name"] == "deoverlap"), None)
                best_post = balanced_post
                if deoverlap_post is not None:
                    balanced_metrics = balanced_post["metrics"]
                    deoverlap_metrics = deoverlap_post["metrics"]
                    balanced_has_overlap_issue = (
                        balanced_metrics.get("overlap_risk", 0.0) >= 0.12
                        or balanced_metrics.get("overlap_front_large_anchors", 0) > 0
                        or balanced_metrics.get("overlap_cramped_rows", 0) > 0
                        or balanced_metrics.get("overlap_cramped_cells", 0) >= 2
                    )
                    deoverlap_better = (
                        balanced_has_overlap_issue
                        and deoverlap_metrics.get("overlap_risk", 1.0) <= (balanced_metrics.get("overlap_risk", 1.0) - 0.03)
                        and deoverlap_post["combined_score"] >= (balanced_post["combined_score"] + 0.03)
                        and deoverlap_metrics.get("text_coverage", 0.0) >= (balanced_metrics.get("text_coverage", 0.0) - 0.01)
                        and deoverlap_metrics.get("field_label_coverage", 0.0) >= (balanced_metrics.get("field_label_coverage", 0.0) - 0.01)
                    )
                    if deoverlap_better:
                        best_post = deoverlap_post
                shutil.copy2(best_post["docx"], self.docx_path)
                compatibility_sanitized = sanitize_docx_for_word_compatibility(self.docx_path)
                best_post["metrics"]["word_compatibility_sanitized"] = bool(compatibility_sanitized)
                best_result["metrics"] = best_post["metrics"]
                best_result["score"] = best_post["quality_score"]

        return {
            "profile_name": best_result["profile_name"],
            "score": best_result["score"],
            "metrics": best_result["metrics"],
            "attempts": attempts,
            "reconstruction_mode": reconstruction_mode,
        }

    def _run_auto_pipeline(self, pdf_name):
        """Smart per-document conversion driven by per-page classification.

        1. Classify every page (text / scanned / hybrid / graphics).
        2. Normalize the PDF so a single pdf2docx pass produces good output:
             - rasterize graphics pages (fast, faithful, avoids vector blowup)
             - OCR-inject a text layer onto scanned/hybrid pages
        3. Run the existing multi-profile pdf2docx sweep on the normalized PDF.
        4. If the whole document is image-only and no text could be recovered,
           fall back to exact-layout image rendering so the user still gets a
           faithful document.

        Returns the pdf2docx report augmented with pipeline diagnostics.
        """
        page_infos = classify_pdf_pages(self.pdf_path)
        summary = summarize_page_kinds(page_infos)
        counts = summary["counts"]
        pipeline = {
            "page_kinds": counts,
            "kinds_sequence": summary["kinds_sequence"],
            "rasterized_pages": [],
            "ocr_applied": False,
            "ocr_pages": [],
            "ocr_reason": None,
            "image_fallback": False,
            "segmented": False,
        }

        # Fast path: a pure text document needs no normalization.
        if summary["all_text"] or not page_infos:
            report = self._run_pdf2docx()
            report.setdefault("pipeline", {}).update(pipeline)
            return report

        # Whole document is image-fidelity (dense forms and/or vector graphics
        # with no editable-text pages): editable reconstruction inflates cell
        # heights and breaks alignment, so render the pages as images for a
        # pixel-faithful result.
        if page_infos and all(p["kind"] in ("form", "graphics") for p in page_infos):
            self._run_exact_layout_conversion(pdf_name)
            pipeline["image_fallback"] = True
            report = {
                "profile_name": "exact_layout_form",
                "score": 1.0,
                "metrics": {"mode": "auto"},
                "reconstruction_mode": "Auto",
            }
            report.setdefault("pipeline", {}).update(pipeline)
            return report

        scanned_pages = [p["index"] for p in page_infos if p["kind"] == "scanned"]
        # Pages whose text is read normally (displayed) by pdf2docx.
        native_pages = [p["index"] for p in page_infos if p["kind"] in ("text", "hybrid")]

        working_pdf = self.pdf_path
        temp_files = []
        original_pdf = self.pdf_path
        try:
            # Step 1: rasterize image-fidelity pages (charts and dense forms) so
            # text pages can still convert to editable Word content around them.
            graphics_pages = [p["index"] for p in page_infos if p["kind"] in ("graphics", "form")]
            if graphics_pages:
                raster = rasterize_pdf_pages(working_pdf, graphics_pages)
                if raster.get("applied"):
                    working_pdf = raster["pdf_path"]
                    temp_files.append(working_pdf)
                    pipeline["rasterized_pages"] = raster.get("rasterized_pages", [])

            # Step 2: OCR-inject a hidden text layer onto scanned pages.
            if scanned_pages:
                lang = resolve_ocr_language(self)
                ocr = enhance_pdf_with_text_layer(working_pdf, page_infos, lang=lang)
                pipeline["ocr_reason"] = ocr.get("reason")
                if ocr.get("applied"):
                    working_pdf = ocr["pdf_path"]
                    temp_files.append(working_pdf)
                    pipeline["ocr_applied"] = True
                    pipeline["ocr_pages"] = ocr.get("ocr_pages", [])

            ocr_succeeded = pipeline["ocr_applied"] or not scanned_pages

            if scanned_pages and native_pages and ocr_succeeded:
                # Mixed: native text needs ocr=0, scanned (hidden) text needs
                # ocr=2. pdf2docx applies the flag per-document, so convert each
                # page in its correct mode and merge in page order.
                report = self._convert_mixed_document(working_pdf, page_infos)
                pipeline["segmented"] = True
            elif scanned_pages and ocr_succeeded and not native_pages:
                # Only scanned (+ rasterized graphics): single hidden-text pass.
                self.pdf_path = working_pdf
                self._pdf2docx_ocr_mode = 2
                report = self._run_pdf2docx()
            else:
                # No scanned pages, or OCR unavailable: normal displayed-text pass.
                self.pdf_path = working_pdf
                self._pdf2docx_ocr_mode = 0
                report = self._run_pdf2docx()
        finally:
            self.pdf_path = original_pdf
            self._pdf2docx_ocr_mode = 0
            for path in temp_files:
                try:
                    os.remove(path)
                except OSError:
                    pass

        # Step 4: if the document is fundamentally image-only and text recovery
        # was impossible (e.g. OCR unavailable on a scanned doc), guarantee
        # fidelity with an exact-layout render rather than shipping empty text.
        image_only = set(counts) <= {"scanned", "graphics", "form"}
        if scanned_pages and not pipeline["ocr_applied"] and image_only:
            self._run_exact_layout_conversion(pdf_name)
            pipeline["image_fallback"] = True
            report = {
                "profile_name": "exact_layout_fallback",
                "score": report.get("score", 0.0),
                "metrics": report.get("metrics", {}),
                "reconstruction_mode": report.get("reconstruction_mode", "Auto"),
            }

        report.setdefault("pipeline", {}).update(pipeline)
        return report

    def _convert_pdf_subset(self, pdf_path, out_docx, pages, ocr_mode):
        """Convert a subset of pages to DOCX with a fixed pdf2docx OCR mode."""
        from pdf2docx import Converter

        cv = Converter(pdf_path)
        try:
            cv.convert(out_docx, pages=pages, ocr=ocr_mode)
        finally:
            cv.close()

    def _convert_mixed_document(self, pdf_path, page_infos):
        """Convert a doc that mixes native-text and scanned pages.

        Contiguous runs of the same conversion mode are converted separately
        (native pages with ocr=0, scanned pages with ocr=2) and merged in page
        order with docxcompose, which preserves styles, numbering and images.
        """
        from docxcompose.composer import Composer

        # Build (ocr_mode, [page_indices]) runs preserving page order.
        runs = []
        for info in page_infos:
            mode = 2 if info["kind"] == "scanned" else 0
            if runs and runs[-1][0] == mode:
                runs[-1][1].append(info["index"])
            else:
                runs.append((mode, [info["index"]]))

        with tempfile.TemporaryDirectory(prefix="mixed_segments_") as seg_dir:
            segment_paths = []
            for seg_idx, (mode, pages) in enumerate(runs):
                seg_path = os.path.join(seg_dir, f"segment_{seg_idx}.docx")
                self._convert_pdf_subset(pdf_path, seg_path, pages, mode)
                segment_paths.append(seg_path)

            base = Document(segment_paths[0])
            composer = Composer(base)
            for seg_path in segment_paths[1:]:
                composer.append(Document(seg_path))
            composer.save(self.docx_path)

        # Score the merged result against the (OCR-enhanced) source.
        pdf_signal = analyze_pdf_quality_signal(pdf_path)
        docx_signal = analyze_docx_quality_signal(self.docx_path)
        score, metrics = score_conversion_quality(pdf_signal, docx_signal)
        sanitize_docx_for_word_compatibility(self.docx_path)
        return {
            "profile_name": "mixed_segmented",
            "score": score,
            "metrics": metrics,
            "reconstruction_mode": "Auto",
        }

    def _run_ocr_conversion(self, pdf_name):
        lang_pref = self.LANG_MAP.get(self.language.get(), "eng")
        tesseract_cmd = self._resolve_tesseract_cmd()
        if not tesseract_cmd:
            raise RuntimeError(
                "Tesseract is required for OCR but was not found. Install Tesseract and add it to PATH."
            )
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

        document = Document()
        rendered_pages = 0
        for image in self._iter_rendered_pages(dpi=300):
            rendered_pages += 1
            processed_image = self._preprocess_ocr_image(image)
            text = pytesseract.image_to_string(processed_image, lang=lang_pref)
            document.add_paragraph(text.strip() if text else "")
        if rendered_pages < 1:
            raise RuntimeError("No pages found in PDF.")

        core_properties = document.core_properties
        core_properties.author = 'PDF Converter v1.6 by Suleyman Bozkurt'
        core_properties.comments = 'PDF converted into DOCX with PDF Converter v1.6'
        core_properties.title = pdf_name
        document.save(self.docx_path)

    def _run_exact_layout_conversion(self, pdf_name):
        # Pages are rendered and embedded one at a time (holding at most the
        # current page plus one look-ahead page, which decides whether a page
        # break follows). Materialising every 300 DPI render first could need
        # gigabytes for a long scanned PDF. The output is unchanged.
        rendered = self._iter_rendered_pages(dpi=300)
        image = next(rendered, None)
        if image is None:
            raise RuntimeError("No pages found in PDF.")

        document = Document()
        section = document.sections[0]
        section.top_margin = Inches(0.3)
        section.bottom_margin = Inches(0.3)
        section.left_margin = Inches(0.3)
        section.right_margin = Inches(0.3)
        available_width = section.page_width - section.left_margin - section.right_margin
        available_height = section.page_height - section.top_margin - section.bottom_margin

        while image is not None:
            image_buffer = io.BytesIO()
            image.save(image_buffer, format="PNG")
            image_buffer.seek(0)
            # Fit the full page image inside BOTH the available width and height
            # so a tall (e.g. A4) page rendered onto a Letter section is never
            # clipped at the bottom. Scale by whichever dimension binds first
            # while preserving the page aspect ratio.
            img_w, img_h = image.width, image.height
            if img_w > 0 and img_h > 0 and (available_width / img_w) * img_h > available_height:
                document.add_picture(image_buffer, height=available_height)
            else:
                document.add_picture(image_buffer, width=available_width)
            image = None  # free this render before producing the next one
            next_image = next(rendered, None)
            if next_image is not None:
                document.add_page_break()
            image = next_image

        core_properties = document.core_properties
        core_properties.author = 'PDF Converter v1.6 by Suleyman Bozkurt'
        core_properties.comments = 'PDF converted into DOCX with PDF Converter v1.6 (exact layout mode)'
        core_properties.title = pdf_name
        document.save(self.docx_path)

    def conversion_worker(self):
        pdf_name, pdf_ext = os.path.splitext(os.path.basename(self.pdf_path))

        # Create the Word file path with the same file name and a .docx extension
        self.docx_path = os.path.join(os.path.dirname(self.pdf_path), pdf_name + ".docx")
        self.conversion_methodResult = self.conversion_method_var.get()

        conversion_succeeded = False
        try:
            if self.conversion_methodResult == 1:
                self.output_text.insert(tk.END, "- Convertion method: Preserve Format (Editable)!\n")
                self.output_text.insert(tk.END, "- Converting PDF...\n")
                report = self._run_pdf2docx()
                conversion_succeeded = True
                self.output_text.insert(
                    tk.END,
                    f"- Reconstruction mode: {report.get('reconstruction_mode', 'Auto')}.\n",
                )
                if self.FORCE_STRONG_CONTINUOUS:
                    self.output_text.insert(
                        tk.END,
                        "- Strong Continuous Reconstruction is enforced for all PDFs.\n",
                    )
                self.output_text.insert(
                    tk.END,
                    f"- Best profile: {report['profile_name']} (score {report['score']:.2f}).\n",
                )
                metrics = report["metrics"]
                self.output_text.insert(
                    tk.END,
                    "- Quality: "
                    f"text {metrics['text_coverage']:.0%}, "
                    f"numbers {metrics['numeric_coverage']:.0%}, "
                    f"barcode-like {metrics['barcode_coverage']:.0%}, "
                    f"labels {metrics['field_label_coverage']:.0%}, "
                    f"images {metrics['image_coverage']:.0%}, "
                    f"sections {metrics.get('section_coverage', 0):.0%}.\n",
                )
                if metrics.get("postprocess_variant"):
                    self.output_text.insert(
                        tk.END,
                        f"- Postprocess variant: {metrics['postprocess_variant']} (overlap risk {metrics.get('overlap_risk', 0.0):.2f}).\n",
                    )
                if metrics.get("flatten_applied"):
                    self.output_text.insert(
                        tk.END,
                        f"- Fixed {metrics.get('flattened_shapes', 0)} floating layout objects to prevent overlap.\n",
                    )
                else:
                    self.output_text.insert(
                        tk.END,
                        "- Kept original floating layout objects for better visual fidelity.\n",
                    )
                if metrics.get("anchors_forced_behind_text", 0) > 0:
                    self.output_text.insert(
                        tk.END,
                        f"- Forced {metrics['anchors_forced_behind_text']} large background image layer(s) behind text.\n",
                    )
                if metrics.get("soft_hyphen_cleaned"):
                    self.output_text.insert(
                        tk.END,
                        "- Removed soft hyphens from output text.\n",
                    )
                if metrics.get("dedup_removed_paragraphs", 0) > 0:
                    self.output_text.insert(
                        tk.END,
                        f"- Removed {metrics['dedup_removed_paragraphs']} redundant repeated lines.\n",
                    )
                if (
                    metrics.get("layout_rows_relaxed", 0) > 0
                    or metrics.get("layout_spacing_relaxed", 0) > 0
                    or metrics.get("layout_cell_margins_compacted", 0) > 0
                    or metrics.get("layout_blank_paragraphs_removed", 0) > 0
                ):
                    self.output_text.insert(
                        tk.END,
                        "- Stabilized table/frame row heights and text spacing to reduce overlap.\n",
                    )
                if metrics.get("single_page_target"):
                    self.output_text.insert(
                        tk.END,
                        "- Applied single-page fit for one-page source PDF.\n",
                    )
                    if metrics.get("single_page_breaks_removed", 0) > 0:
                        self.output_text.insert(
                            tk.END,
                            f"- Removed {metrics['single_page_breaks_removed']} forced page-break marker(s).\n",
                        )
                if metrics.get("barcode_images_inserted", 0) > 0:
                    self.output_text.insert(
                        tk.END,
                        f"- Replaced barcode font text with {metrics['barcode_images_inserted']} barcode image(s) in-place.\n",
                    )
                elif metrics.get("barcode_text_runs_removed", 0) > 0:
                    self.output_text.insert(
                        tk.END,
                        "- Removed unreadable barcode text artifacts.\n",
                    )
            elif self.conversion_methodResult == 2:
                # Use OCR to convert the PDF
                self.output_text.insert(tk.END, "- Convertion method: OCR Converter!\n")
                self.output_text.insert(tk.END, "- Converting PDF...\n")
                self._run_ocr_conversion(pdf_name)
                conversion_succeeded = True
            elif self.conversion_methodResult == 3:
                self.output_text.insert(tk.END, "- Convertion method: Exact Layout Converter!\n")
                self.output_text.insert(tk.END, "- Rendering pages to preserve original arrangement...\n")
                self._run_exact_layout_conversion(pdf_name)
                conversion_succeeded = True
        except Exception as e:
            # Print an error message if the conversion fails
            print(f"An error occurred while converting the PDF to Word: {e}")
            tkinter.messagebox.showerror("Error",
                                         f"An error occurred, contact to the programmer. Error is {e}.\n\n" + traceback.format_exc())

        if conversion_succeeded:
            self.output_text.insert(tk.END, f"- Saved here: {self.docx_path}\n")
            self.output_text.insert(tk.END, "- Finished!\n")
            tkinter.messagebox.showinfo("Task complete", "The task is complete!")

class PDF2Image(ttk.Frame):
    def __init__(self, parent):
        ttk.Frame.__init__(self, parent)

        self.pdf_label = tk.Label(self, text="Select a PDF file:", font=("Helvetica", 14), pady=10)
        self.pdf_label.pack()

        self.pdf_button = tk.Button(self, text="Browse", font=("Helvetica", 14), pady=10, command=self.select_pdf)
        self.pdf_button.pack()

        self.conversion_method_label = tk.Label(self, text="Select a image format:", font=("Helvetica", 14), pady=10)
        self.conversion_method_label.pack()

        self.variable = StringVar()
        self.variable.set("tiff")  # default value
        image_formats = ["tiff", "jpg", "png", "eps", "bmp", "ico"] # 6 different file format so far

        self.options = OptionMenu(self, self.variable, *image_formats)
        self.options.pack()

        # self.output_text = tk.Text(root, height=40, width=30, font=("Helvetica", 14))
        # self.output_text.pack(padx=10, pady=10)
        self.output_frame = tk.Frame(self)
        self.output_frame.pack(padx=10, pady=10)

        # Create a Scrollbar widget and associate it with the Text widget
        self.scrollbar = tk.Scrollbar(self.output_frame, orient="vertical")
        self.scrollbar.pack(side="right", fill="y")

        # Create the Text widget and associate it with the Scrollbar widget
        self.output_text = tk.Text(self.output_frame, height=13, width=60, font=("Helvetica", 14),
                                   yscrollcommand=self.scrollbar.set)
        self.output_text.pack(side="left", fill="both", expand=True)

        # Set the command of the Scrollbar widget to the yview method of the Text widget
        self.scrollbar.config(command=self.output_text.yview)

        self.button_frame = tk.Frame(self)
        self.button_frame.pack()

        self.convert_button = tk.Button(self.button_frame, text="Convert", font=("Helvetica", 14), pady=10,
                                        command=self.convert,
                                        state=tk.DISABLED)
        self.convert_button.pack(side=tk.LEFT)

        self.open_button = tk.Button(self.button_frame, text="Open", font=("Helvetica", 14), pady=10,
                                        command=self.open_file,)
        self.open_button.pack(side=tk.RIGHT)

    def pdf_to_image(self, path, out_format):
        pdf_name, pdf_ext = os.path.splitext(os.path.basename(path))
        self.image_path = os.path.join(os.path.dirname(self.pdf_path), pdf_name)
        try:
            kwargs = {"dpi": 300}
            bundled_poppler = os.path.join(os.path.dirname(__file__), "files", "poppler-22.04.0", "Library", "bin")
            if sys.platform.startswith("win") and os.path.isfile(os.path.join(bundled_poppler, "pdftoppm.exe")):
                kwargs["poppler_path"] = bundled_poppler
            else:
                system_pdftoppm = shutil.which("pdftoppm")
                if system_pdftoppm:
                    kwargs["poppler_path"] = os.path.dirname(system_pdftoppm)
            if "poppler_path" not in kwargs:
                raise RuntimeError(
                    "Poppler is required for PDF to image conversion but was not found. Install Poppler and add 'pdftoppm' to PATH."
                )
            img = convert_from_path(path, **kwargs)
            num = len(img)

            for i, page in enumerate(img):
                if num == 1:
                    self.output_file = f'{self.image_path}.{out_format}'
                else:
                    self.output_file = f'{self.image_path}_page_{i+1}.{out_format}'

                page.save(self.output_file, out_format)
                self.output_text.insert(tk.END, f"- Saved here: {self.output_file}\n")
        except Exception as e:
            tkinter.messagebox.showerror("Error",
                                     "An error occurred while converting the PDF file.\n\n" + traceback.format_exc())

    def select_pdf(self):
        try:
            self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if self.pdf_path:
                self.output_text.insert(tk.END, "- Selected PDF: " + self.pdf_path + "\n")
                self.convert_button.config(state=tk.NORMAL)
        except Exception:
            # Display an error message in a pop-up window
            tkinter.messagebox.showerror("Error",
                                         "An error occurred while selecting the PDF file.\n\n" + traceback.format_exc())

    def convert(self):
        # Create a new thread to perform the conversion
        conversion_thread = threading.Thread(target=self.conversion_worker)
        conversion_thread.start()

    def open_file(self):
        try:
            # Open the docx file with the default program
            open_with_default_app(self.output_file)
        except Exception:
            import tkinter.messagebox

            tkinter.messagebox.showerror("Error",
                                         "An error occurred while opening the docx file.\n\n" + traceback.format_exc())

    def conversion_worker(self):

        self.options_method = self.variable.get()
        self.output_text.insert(tk.END, f"- Convertion format: PDF to {self.options_method}!\n")

        try:
            self.output_text.insert(tk.END, "- Converting PDF...\n")
            if self.options_method == 'tiff':
                self.pdf_to_image(self.pdf_path, 'tiff')  # Convert PDF to TIFF
            elif self.options_method == 'jpg':
                self.pdf_to_image(self.pdf_path, 'JPEG') # Convert PDF to JPEG
            elif self.options_method == 'png':
                self.pdf_to_image(self.pdf_path, 'png')  # Convert PDF to PNG
            elif self.options_method == 'ico':
                self.pdf_to_image(self.pdf_path, 'ICO')  # Convert PDF to ICO
            elif self.options_method == 'bmp':
                self.pdf_to_image(self.pdf_path, 'BMP')  # Convert PDF to SVG
            elif self.options_method == 'eps':
                self.pdf_to_image(self.pdf_path, 'eps') # Convert PDF to EPS

        except Exception as e:
            # Print an error message if the conversion fails
            print(f"An error occurred while converting the PDF to Word: {e}")
            tkinter.messagebox.showerror("Error",
                                         f"An error occurred, contact to the programmer. Error is {e}.\n\n" + traceback.format_exc())

        self.output_text.insert(tk.END, "- Finished!\n")
        tkinter.messagebox.showinfo("Task complete", "The task is complete!")

class MyWindow():

    def __init__(self, root):
        self.root = root
        self.notebook = ttk.Notebook(self.root)

        self.notebook.pack(expand=1, fill="both")
        self.PDFConverterFrame = PDF2Doc(self.notebook)
        self.PDFConverterFrame.bind("<<NotebookTabChanged>>", self.on_tab_selected)
        self.notebook.add(self.PDFConverterFrame, text="PDF to Docx")

        self.PDFimageFrame = PDF2Image(self.notebook)
        self.notebook.add(self.PDFimageFrame, text="PDF to Image")

        self.aboutFrame = about(self.notebook)
        self.notebook.add(self.aboutFrame, text="About")

        self.supportFrame = support(self.notebook)
        self.notebook.add(self.supportFrame, text="Support")

        #self.notebook.grid()
    def on_tab_selected(self):
        selected_tab = self.widget.select()
        tab_text = self.widget.tab(selected_tab, "text")

def launch_gui():
    if Tk is None:
        raise RuntimeError(
            "GUI is disabled because tkinter is not available in this environment."
        )
    root = Tk()
    root.title("PDF Converter v1.6 @2023", )
    root.geometry("960x600+480+250")
    root.resizable(0, 0)
    root.wm_iconbitmap('./files/icon.ico')
    MyWindow(root)
    root.mainloop()


if __name__ == '__main__':
    launch_gui()
