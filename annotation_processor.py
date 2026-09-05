"""
AnotAI - Core Engine
====================
Extracts annotations/comments from PDF and Word (.docx) documents and builds a
branded Excel report.

Each row carries:
  * a page-scoped reference  (1.1, 1.2, 2.1 ...)
  * a high-resolution DETAIL crop, ringed in red
  * a full-page OVERVIEW thumbnail with the annotation marked

100% offline. No Azure, no SharePoint, no network calls.

Powered by Maehan Solutions.
"""

import io
import os
import re
import shutil
import logging
import subprocess
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import List, Dict, Optional, Tuple

import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU

from docx import Document

import brand

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


class Cancelled(Exception):
    """Raised internally when the user cancels a run."""


# ---------------------------------------------------------------------------
# Excel layout
# ---------------------------------------------------------------------------
HEADERS = [
    "Ref", "Source", "Page", "Annotation Type", "Author", "Date",
    "Annotation Text", "Location Snapshot", "Page Overview",
    "Status", "Reviewer Remarks",
]

COL_WIDTHS = {
    1: 8, 2: 9, 3: 6, 4: 17, 5: 21, 6: 17, 7: 46,
    8: 96, 9: 32, 10: 11, 11: 24,
}

DETAIL_COL = 8
OVERVIEW_COL = 9

# Detail crop: rendered wide so fine drawing text stays legible
DETAIL_PX_WIDTH = 900
DETAIL_MAX_PX_HEIGHT = 620

# Overview thumbnail
OVERVIEW_PX_WIDTH = 300
OVERVIEW_MAX_PX_HEIGHT = 420

TITLE_FILL  = PatternFill("solid", fgColor=brand.XL_DARK_BLUE)
HEADER_FILL = PatternFill("solid", fgColor=brand.XL_DARK_BLUE)
HEADER_FONT = Font(bold=True, size=11, color=brand.XL_WHITE, name="Segoe UI")
TITLE_FONT  = Font(bold=True, size=16, color=brand.XL_WHITE, name="Segoe UI")
BODY_FONT   = Font(size=10, color=brand.XL_BLACK, name="Segoe UI")
REF_FONT    = Font(size=11, bold=True, color=brand.XL_DARK_BLUE, name="Segoe UI")
ALT_FILL    = PatternFill("solid", fgColor=brand.XL_CREAM)

THIN = Side(style="thin", color=brand.XL_GREY_LINE)
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

BOX_COLOR = brand.BOX_COLOR

# Viewer scaffolding - never review content
NOISE_TYPES = {"Popup", "Link"}

TYPE_LABELS = {
    "Text": "Sticky Note", "FreeText": "Text Box", "Line": "Line",
    "Square": "Rectangle", "Circle": "Ellipse", "Polygon": "Polygon",
    "PolyLine": "Polyline", "Highlight": "Highlight", "Underline": "Underline",
    "Squiggly": "Squiggly Underline", "StrikeOut": "Strikeout",
    "Stamp": "Stamp", "Caret": "Caret", "Ink": "Freehand (Ink)",
    "Popup": "Popup", "FileAttachment": "File Attachment", "Sound": "Sound",
    "Movie": "Movie", "Widget": "Form Field", "Screen": "Screen",
    "PrinterMark": "Printer Mark", "TrapNet": "Trap Network",
    "Watermark": "Watermark", "3D": "3D Object", "Redact": "Redaction",
    "Link": "Link",
}

# Text that markup tools auto-write into 'subject'. Not typed by a reviewer,
# so an annotation carrying only one of these has no review content.
TOOL_NAME_WORDS = {
    "rectangle", "square", "box", "circle", "oval", "ellipse", "triangle",
    "polygon", "polyline", "line", "arrow", "arrows", "cloud", "cloud+",
    "callout", "stamp", "ink", "pen", "pencil", "freehand", "highlight",
    "underline", "strikeout", "squiggly", "text box", "textbox", "note",
    "sticky note", "shape", "markup", "measurement", "length", "area",
    "perimeter", "dimension", "leader", "polygon cloud", "revision cloud",
}


class AnnotationExtractor:
    """Extract annotations from PDF / DOCX and emit a branded Excel report."""

    def __init__(self, logger=None, include_noise: bool = False,
                 keep_markup_only: bool = False,
                 progress_cb=None, cancel_cb=None):
        self.logger = logger or logging.getLogger(__name__)
        self.include_noise = include_noise
        # False  -> drop shapes that carry no reviewer-typed text
        self.keep_markup_only = keep_markup_only
        self.progress_cb = progress_cb
        self.cancel_cb = cancel_cb
        Image.MAX_IMAGE_PIXELS = None
        self._temp_dirs: List[str] = []
        self._page_cache: Dict[int, Tuple[Image.Image, float]] = {}
        self.skipped_markup = 0

    # ------------------------------------------------------------------
    def _say(self, msg: str, pct: int = -1):
        if self.progress_cb:
            try:
                self.progress_cb(msg if pct < 0 else f"{msg}||{pct}")
            except Exception:
                pass

    def _check_cancel(self):
        if self.cancel_cb:
            try:
                if self.cancel_cb():
                    raise Cancelled()
            except Cancelled:
                raise
            except Exception:
                pass

    @staticmethod
    def _clean_text(text) -> str:
        return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", str(text or "")).strip()

    @staticmethod
    def _squash(text: str) -> str:
        return re.sub(r"\s+", " ", text or "").strip()

    @staticmethod
    def _is_tool_name(text: str) -> bool:
        """True when the text is only an auto-generated tool name."""
        if not text:
            return True
        t = re.sub(r"[\s\-_.#0-9]+", " ", text.strip().lower()).strip()
        if not t:
            return True
        return t in TOOL_NAME_WORDS

    def _parse_pdf_date(self, date_str: str) -> Optional[str]:
        if not date_str:
            return None
        s = str(date_str).strip()
        if not s.startswith("D:"):
            return None
        try:
            core = s[2:16]
            if len(core) < 8:
                return s
            core = core.ljust(14, "0")
            dt = datetime.strptime(core, "%Y%m%d%H%M%S")
            m = re.search(r"([+-])(\d{2})'?(\d{2})?'?", s[16:])
            if m:
                sign, hh, mm = m.group(1), int(m.group(2)), int(m.group(3) or 0)
                off = timedelta(hours=hh, minutes=mm)
                if sign == "-":
                    off = -off
                dt = dt.replace(tzinfo=timezone(off)).astimezone()
            return dt.strftime("%d-%m-%Y %H:%M")
        except Exception:
            return s

    @staticmethod
    def _parse_docx_date(date_attr: str) -> str:
        if not date_attr:
            return "Not recorded"
        try:
            dt = datetime.fromisoformat(str(date_attr).replace("Z", "+00:00"))
            if dt.tzinfo:
                dt = dt.astimezone()
            return dt.strftime("%d-%m-%Y %H:%M")
        except Exception:
            return str(date_attr)

    def _mktemp(self) -> str:
        d = tempfile.mkdtemp(prefix="anotai_")
        self._temp_dirs.append(d)
        return d

    def cleanup(self):
        for d in self._temp_dirs:
            shutil.rmtree(d, ignore_errors=True)
        self._temp_dirs = []
        self._page_cache.clear()

    # ==================================================================
    # ORDERING  -  column-major, top-to-bottom then left-to-right
    # ==================================================================
    @staticmethod
    def _column_major_order(items: List[Tuple[object, fitz.Rect]],
                            page_width: float) -> List[object]:
        """
        Sort annotations the way an engineer reads a drawing sheet:
        down the leftmost column first, then the next column to the right.

        Annotations whose horizontal centres fall within one band are
        treated as the same column, so a slightly ragged column still
        reads top-to-bottom rather than zig-zagging.
        """
        if not items:
            return []

        # Band width: wide enough to absorb ragged alignment, narrow enough
        # to separate genuine columns on a large sheet.
        band = max(page_width * 0.16, 120.0)

        # Order candidates by horizontal position first.
        by_x = sorted(items, key=lambda it: (it[1].x0, it[1].y0))

        columns: List[List[Tuple[object, fitz.Rect]]] = []
        col_left: List[float] = []

        for obj, rect in by_x:
            placed = False
            for i, left in enumerate(col_left):
                if rect.x0 - left <= band:
                    columns[i].append((obj, rect))
                    placed = True
                    break
            if not placed:
                columns.append([(obj, rect)])
                col_left.append(rect.x0)

        ordered: List[object] = []
        for col in columns:                      # already left-to-right
            col.sort(key=lambda it: (round(it[1].y0, 1), round(it[1].x0, 1)))
            ordered.extend(obj for obj, _ in col)
        return ordered

    # ==================================================================
    # SNAPSHOT ENGINE
    # ==================================================================
    @staticmethod
    def _detail_clip(page: fitz.Page, rects: List[fitz.Rect],
                     full_width: bool = False) -> fitz.Rect:
        """
        Crop window sized from the ANNOTATION, not from the page.

        Sizing from the page meant a small callout on an E-size drawing
        got a crop 24 inches wide - unreadable once scaled into a cell.
        """
        pr = page.rect
        target = fitz.Rect(rects[0])
        for r in rects[1:]:
            target |= r

        if target.width < 1:
            cx = (target.x0 + target.x1) / 2
            target.x0, target.x1 = cx - 8, cx + 8
        if target.height < 1:
            cy = (target.y0 + target.y1) / 2
            target.y0, target.y1 = cy - 8, cy + 8

        if full_width:
            want_w = pr.width
            want_h = max(target.height * 3.0, 190.0)
        else:
            # ~2x the annotation, with sane floor and ceiling
            pad_x = max(target.width * 0.55, 80.0)
            pad_y = max(target.height * 0.55, 70.0)
            want_w = min(target.width + 2 * pad_x, pr.width, 1500.0)
            want_h = min(target.height + 2 * pad_y, pr.height, 1200.0)
            want_w = max(want_w, 240.0)
            want_h = max(want_h, 170.0)

        want_w = min(want_w, pr.width)
        want_h = min(want_h, pr.height)

        cx = (target.x0 + target.x1) / 2
        cy = (target.y0 + target.y1) / 2
        clip = fitz.Rect(cx - want_w / 2, cy - want_h / 2,
                         cx + want_w / 2, cy + want_h / 2)

        # Slide fully inside the page rather than clipping to a sliver
        if clip.x0 < pr.x0:
            clip.x1 += (pr.x0 - clip.x0); clip.x0 = pr.x0
        if clip.x1 > pr.x1:
            clip.x0 -= (clip.x1 - pr.x1); clip.x1 = pr.x1
        if clip.y0 < pr.y0:
            clip.y1 += (pr.y0 - clip.y0); clip.y0 = pr.y0
        if clip.y1 > pr.y1:
            clip.y0 -= (clip.y1 - pr.y1); clip.y1 = pr.y1

        clip = clip & pr
        if clip.is_empty or clip.width <= 1 or clip.height <= 1:
            clip = fitz.Rect(pr)
        return clip

    def _render_detail(self, page: fitz.Page, clip: fitz.Rect,
                       boxes: List[fitz.Rect]) -> Optional[Image.Image]:
        """
        Render `clip` so the OUTPUT is ~DETAIL_PX_WIDTH wide regardless of
        how large the source area is. Small callouts therefore render at a
        very high effective DPI, huge areas at a modest one - consistent
        legibility, bounded memory, no oversized-pixmap failures.
        """
        target_w = DETAIL_PX_WIDTH
        zoom = target_w / max(clip.width, 1.0)
        zoom = max(0.5, min(zoom, 9.0))

        # Guard total pixels so a giant clip can never blow up
        while clip.width * zoom * clip.height * zoom > 14_000_000 and zoom > 0.5:
            zoom *= 0.8

        img = None
        used = zoom
        for attempt in (zoom, zoom * 0.6, zoom * 0.35, 1.0):
            try:
                pix = page.get_pixmap(matrix=fitz.Matrix(attempt, attempt),
                                      clip=clip, annots=True, alpha=False)
                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                used = attempt
                break
            except Exception as e:
                self.logger.warning(f"Detail render retry (zoom {attempt:.2f}): {e}")
        if img is None:
            return None

        draw = ImageDraw.Draw(img)
        lw = max(3, int(img.width / 260))
        for b in boxes:
            x0 = (b.x0 - clip.x0) * used
            y0 = (b.y0 - clip.y0) * used
            x1 = (b.x1 - clip.x0) * used
            y1 = (b.y1 - clip.y0) * used
            pad = lw + 2
            x0, y0 = max(0, x0 - pad), max(0, y0 - pad)
            x1 = min(img.width - 1, x1 + pad)
            y1 = min(img.height - 1, y1 + pad)
            if x1 <= x0:
                x1 = min(img.width - 1, x0 + 14)
            if y1 <= y0:
                y1 = min(img.height - 1, y0 + 14)
            draw.rectangle([x0, y0, x1, y1], outline=BOX_COLOR, width=lw)

        draw.rectangle([0, 0, img.width - 1, img.height - 1],
                       outline=(140, 140, 140), width=2)

        if img.width > DETAIL_PX_WIDTH:
            r = DETAIL_PX_WIDTH / float(img.width)
            img = img.resize((DETAIL_PX_WIDTH, max(1, int(img.height * r))),
                             Image.Resampling.LANCZOS)
        if img.height > DETAIL_MAX_PX_HEIGHT:
            r = DETAIL_MAX_PX_HEIGHT / float(img.height)
            img = img.resize((max(1, int(img.width * r)), DETAIL_MAX_PX_HEIGHT),
                             Image.Resampling.LANCZOS)
        return img

    def _detail_snapshot(self, page, rects, full_width=False):
        try:
            clip = self._detail_clip(page, rects, full_width)
            return self._render_detail(page, clip, rects)
        except Exception as e:
            self.logger.error(f"Detail snapshot failed: {e}")
            return None

    # ------------------------------------------------------------------
    def _page_base(self, page: fitz.Page) -> Optional[Tuple[Image.Image, float]]:
        """Full-page render, cached - a busy page is reused many times."""
        key = page.number
        if key in self._page_cache:
            return self._page_cache[key]
        try:
            zoom = OVERVIEW_PX_WIDTH / max(page.rect.width, 1.0)
            zoom = max(0.08, min(zoom, 2.0))
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom),
                                  annots=True, alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            if img.height > OVERVIEW_MAX_PX_HEIGHT:
                r = OVERVIEW_MAX_PX_HEIGHT / float(img.height)
                img = img.resize((max(1, int(img.width * r)), OVERVIEW_MAX_PX_HEIGHT),
                                 Image.Resampling.LANCZOS)
                zoom *= r
            self._page_cache[key] = (img, zoom)
            return self._page_cache[key]
        except Exception as e:
            self.logger.warning(f"Page overview render failed: {e}")
            return None

    def _overview_snapshot(self, page: fitz.Page,
                           rects: List[fitz.Rect]) -> Optional[Image.Image]:
        """Whole page, with this annotation marked - answers 'where on the sheet?'."""
        base = self._page_base(page)
        if base is None:
            return None
        src, zoom = base
        try:
            img = src.copy()
            d = ImageDraw.Draw(img)

            target = fitz.Rect(rects[0])
            for r in rects[1:]:
                target |= r

            x0 = target.x0 * zoom
            y0 = target.y0 * zoom
            x1 = target.x1 * zoom
            y1 = target.y1 * zoom

            # Guarantee a visible marker even for a hairline annotation
            min_side = 11
            if x1 - x0 < min_side:
                cx = (x0 + x1) / 2
                x0, x1 = cx - min_side / 2, cx + min_side / 2
            if y1 - y0 < min_side:
                cy = (y0 + y1) / 2
                y0, y1 = cy - min_side / 2, cy + min_side / 2

            x0 = max(0, x0); y0 = max(0, y0)
            x1 = min(img.width - 1, x1); y1 = min(img.height - 1, y1)

            d.rectangle([x0, y0, x1, y1], outline=BOX_COLOR, width=3)

            # Crosshair so the eye lands on it instantly
            cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
            arm = 9
            d.line([cx - arm, cy, cx + arm, cy], fill=BOX_COLOR, width=2)
            d.line([cx, cy - arm, cx, cy + arm], fill=BOX_COLOR, width=2)

            d.rectangle([0, 0, img.width - 1, img.height - 1],
                        outline=(140, 140, 140), width=2)
            return img
        except Exception as e:
            self.logger.warning(f"Overview marker failed: {e}")
            return None

    # ------------------------------------------------------------------
    @staticmethod
    def _font(bold: bool, size: int):
        names = (["arialbd.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
                 if bold else
                 ["arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"])
        for n in names:
            try:
                return ImageFont.truetype(n, size)
            except Exception:
                continue
        return ImageFont.load_default()

    def _placeholder(self, title, subtitle="", width=520, height=200):
        img = Image.new("RGB", (width, height), brand.CREAM)
        d = ImageDraw.Draw(img)
        d.rectangle([0, 0, width - 1, height - 1], outline=brand.GREY_LINE, width=2)
        f_big, f_sm = self._font(True, 17), self._font(False, 13)
        lines = [(title, f_big, brand.DARK_BLUE)]
        for chunk in self._wrap(subtitle, 52):
            lines.append((chunk, f_sm, brand.MUTED))
        total = sum((f.getbbox(t)[3] - f.getbbox(t)[1]) + 11 for t, f, _ in lines)
        y = max(12, (height - total) // 2)
        for text, font, colour in lines:
            bb = d.textbbox((0, 0), text, font=font)
            d.text(((width - (bb[2] - bb[0])) // 2, y), text, fill=colour, font=font)
            y += (bb[3] - bb[1]) + 11
        return img

    @staticmethod
    def _wrap(text, width):
        if not text:
            return []
        words, lines, cur = text.split(), [], ""
        for w in words:
            if len(cur) + len(w) + 1 <= width:
                cur = f"{cur} {w}".strip()
            else:
                lines.append(cur); cur = w
            if len(lines) >= 4:
                break
        if cur and len(lines) < 5:
            lines.append(cur)
        return lines

    # ==================================================================
    # PDF EXTRACTION
    # ==================================================================
    def _annot_rects(self, annot) -> List[fitz.Rect]:
        rects = []
        try:
            vs = annot.vertices
            if vs and len(vs) >= 4 and annot.type[1] in (
                    "Highlight", "Underline", "Squiggly", "StrikeOut"):
                for i in range(0, len(vs) - 3, 4):
                    r = fitz.Quad(vs[i], vs[i+1], vs[i+2], vs[i+3]).rect
                    if r.width > 0 and r.height > 0:
                        rects.append(r)
        except Exception:
            pass
        if not rects:
            rects = [fitz.Rect(annot.rect)]
        return rects

    def _annot_info(self, annot) -> Dict:
        info = annot.info or {}
        subtype = annot.type[1] if annot.type else "Unknown"
        label = TYPE_LABELS.get(subtype, subtype)

        # --- reviewer-typed content -----------------------------------
        typed = ""
        try:
            popup = annot.get_popup()
            if popup is not None and popup.info:
                c = self._clean_text(popup.info.get("content", ""))
                if c and c.lower() not in ("none", "null"):
                    typed = c
        except Exception:
            pass
        if not typed:
            c = self._clean_text(info.get("content", ""))
            if c and c.lower() not in ("none", "null"):
                typed = c

        # 'subject' is auto-filled with the tool name by most markup tools
        subject = self._clean_text(info.get("subject", ""))
        if not typed and subject and not self._is_tool_name(subject):
            typed = subject

        # --- text the markup sits on ----------------------------------
        underlying = ""
        if subtype in ("Highlight", "Underline", "Squiggly", "StrikeOut"):
            try:
                page = annot.parent
                parts = []
                for r in self._annot_rects(annot):
                    t = page.get_textbox(r)
                    if t:
                        parts.append(self._squash(t))
                underlying = self._squash(" ".join(parts))
            except Exception:
                pass

        has_content = bool(typed) or bool(underlying)

        if typed and underlying:
            content = f'{typed}\n\nMarked text: "{underlying}"'
        elif typed:
            content = typed
        elif underlying:
            content = f'[{label} on: "{underlying}"]'
        elif subtype in ("Text", "FreeText"):
            content = "[Empty sticky note - no text entered]"
        else:
            content = f"[{label} - no text entered]"

        author = "Unknown"
        for key in ("title", "author", "T"):
            v = self._clean_text(info.get(key, ""))
            if v and v.lower() not in ("unknown", "none", "null"):
                author = v
                break

        date_out = "Not recorded"
        for key in ("modDate", "M", "creationDate", "CreationDate"):
            parsed = self._parse_pdf_date(info.get(key, ""))
            if parsed:
                date_out = parsed
                break

        return {"content": content, "author": author, "type": label,
                "raw_type": subtype, "date": date_out,
                "has_content": has_content}

    def extract_pdf_annotations(self, pdf_path, source_label="PDF"):
        self._say(f"Opening {Path(pdf_path).name}", 3)
        doc = fitz.open(pdf_path)
        out: List[Dict] = []
        total_pages = len(doc)
        self.skipped_markup = 0

        try:
            for pno in range(total_pages):
                self._check_cancel()
                self._page_cache.pop(pno - 1, None)     # keep memory flat
                page = doc[pno]
                try:
                    annots = list(page.annots() or [])
                except Exception as e:
                    self.logger.error(f"Page {pno+1} annots unreadable: {e}")
                    continue

                pct = 5 + int(72 * (pno + 1) / max(1, total_pages))
                if not annots:
                    self._say(f"Page {pno+1} of {total_pages}", pct)
                    continue

                # ---- keep only annotations with real review content ----
                kept = []
                for a in annots:
                    st = a.type[1] if a.type else "Unknown"
                    if not self.include_noise and st in NOISE_TYPES:
                        continue
                    meta = self._annot_info(a)
                    if not meta["has_content"] and not self.keep_markup_only:
                        self.skipped_markup += 1
                        continue
                    kept.append((a, meta))

                if not kept:
                    self._say(f"Page {pno+1} of {total_pages} - markup only", pct)
                    continue

                self._say(f"Page {pno+1} of {total_pages} - {len(kept)} annotation(s)",
                          pct)

                # ---- column-major ordering -----------------------------
                pairs = [(item, fitz.Rect(item[0].rect)) for item in kept]
                ordered = self._column_major_order(pairs, page.rect.width)

                for seq, (annot, meta) in enumerate(ordered, 1):
                    self._check_cancel()
                    try:
                        rects = self._annot_rects(annot)
                        detail = self._detail_snapshot(page, rects)
                        if detail is None:
                            detail = self._placeholder(
                                meta["type"],
                                f"Page {pno+1} - detail could not be rendered")
                        overview = self._overview_snapshot(page, rects)

                        out.append({
                            "ref": f"{pno+1}.{seq}",
                            "page": pno + 1,
                            "type": meta["type"],
                            "content": meta["content"],
                            "author": meta["author"],
                            "date": meta["date"],
                            "image": detail,
                            "overview": overview,
                            "source": source_label,
                        })
                    except Cancelled:
                        raise
                    except Exception as e:
                        self.logger.error(f"Annotation error, page {pno+1}: {e}")
        finally:
            doc.close()
            self._page_cache.clear()

        if self.skipped_markup:
            self.logger.info(
                f"Skipped {self.skipped_markup} markup object(s) with no text.")
        self._say(f"Found {len(out)} annotation(s) with content", 80)
        return out

    # ==================================================================
    # WORD EXTRACTION
    # ==================================================================
    def _find_word_converter(self):
        for name in ("soffice", "soffice.exe", "libreoffice"):
            p = shutil.which(name)
            if p:
                return p
        for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
            if os.path.exists(c):
                return c
        return None

    def _docx_to_pdf(self, docx_path):
        outdir = self._mktemp()
        target = os.path.join(outdir, Path(docx_path).stem + ".pdf")

        try:
            import win32com.client  # type: ignore
            import pythoncom        # type: ignore
            self._say("Rendering document via Microsoft Word...", 25)
            pythoncom.CoInitialize()
            word = docm = None
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                docm = word.Documents.Open(os.path.abspath(docx_path),
                                           ReadOnly=True, Visible=False)
                try:
                    docm.ActiveWindow.View.MarkupMode = 0
                    docm.ActiveWindow.View.ShowComments = True
                    docm.ActiveWindow.View.RevisionsView = 0
                except Exception:
                    pass
                docm.ExportAsFixedFormat(
                    OutputFileName=target, ExportFormat=17,
                    OpenAfterExport=False, OptimizeFor=0, Item=7,
                    IncludeDocProps=True, CreateBookmarks=0)
            finally:
                for fn in (lambda: docm.Close(False) if docm else None,
                           lambda: word.Quit() if word else None,
                           pythoncom.CoUninitialize):
                    try:
                        fn()
                    except Exception:
                        pass
            if os.path.exists(target):
                self._say("Word rendering complete", 40)
                return target
        except ImportError:
            pass
        except Exception as e:
            self.logger.warning(f"Microsoft Word unavailable: {e}")

        soffice = self._find_word_converter()
        if soffice:
            try:
                self._say("Rendering document via LibreOffice...", 25)
                subprocess.run(
                    [soffice, "--headless", "--norestore", "--convert-to", "pdf",
                     "--outdir", outdir, os.path.abspath(docx_path)],
                    check=True, timeout=240,
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                if os.path.exists(target):
                    self._say("Rendering complete", 40)
                    return target
                pdfs = list(Path(outdir).glob("*.pdf"))
                if pdfs:
                    return str(pdfs[0])
            except Exception as e:
                self.logger.warning(f"LibreOffice failed: {e}")

        self.logger.warning("No renderer available - using text tiles.")
        return None

    def _read_docx_comments(self, docx_path):
        doc = Document(docx_path)
        comments_part = None
        for rel in doc.part.rels.values():
            if rel.reltype.endswith("/comments"):
                comments_part = rel.target_part
                break
        if comments_part is None:
            return []

        meta, order = {}, []
        for c in comments_part.element.findall(f".//{W_NS}comment"):
            cid = c.get(f"{W_NS}id") or c.get("id") or str(len(order))
            texts = [t.text for t in c.findall(f".//{W_NS}t") if t.text]
            meta[cid] = {
                "id": cid,
                "author": self._clean_text(
                    c.get(f"{W_NS}author") or c.get("author") or "") or "Unknown",
                "date": self._parse_docx_date(
                    c.get(f"{W_NS}date") or c.get("date") or ""),
                "content": self._squash(" ".join(texts)) or "[Empty comment]",
                "anchor": "", "context": "",
            }
            order.append(cid)

        body = doc.element.body
        active, anchors, contexts = set(), {cid: [] for cid in meta}, {}
        para_text, para_open = [], set()

        for el in body.iter():
            tag = el.tag
            if tag == f"{W_NS}p":
                joined = self._squash("".join(para_text))
                for cid in para_open:
                    if joined and not contexts.get(cid):
                        contexts[cid] = joined
                para_text, para_open = [], set(active)
            elif tag == f"{W_NS}commentRangeStart":
                cid = el.get(f"{W_NS}id") or el.get("id")
                if cid is not None:
                    active.add(cid); para_open.add(cid)
            elif tag == f"{W_NS}commentRangeEnd":
                active.discard(el.get(f"{W_NS}id") or el.get("id"))
            elif tag == f"{W_NS}t":
                txt = el.text or ""
                para_text.append(txt)
                for cid in active:
                    anchors.setdefault(cid, []).append(txt)
            elif tag in (f"{W_NS}tab", f"{W_NS}br", f"{W_NS}cr"):
                para_text.append(" ")

        joined = self._squash("".join(para_text))
        for cid in para_open:
            if joined and not contexts.get(cid):
                contexts[cid] = joined
        for cid, info in meta.items():
            info["anchor"] = self._squash("".join(anchors.get(cid, [])))
            info["context"] = contexts.get(cid, "")
        return [meta[cid] for cid in order]

    def _locate_in_pdf(self, doc, needle, used):
        if not needle:
            return None
        n = self._squash(needle)
        probes = []
        for length in (len(n), 90, 60, 40, 25, 15):
            p = n[:length].strip()
            if len(p) >= 4 and p not in probes:
                probes.append(p)
        for probe in probes:
            for pno in range(len(doc)):
                try:
                    hits = doc[pno].search_for(probe, quads=False)
                except Exception:
                    continue
                if not hits:
                    continue
                for h in hits:
                    key = (pno, round(h.x0), round(h.y0), round(h.x1))
                    if key in used:
                        continue
                    used.add(key)
                    same = [r for r in hits if abs(r.y0 - h.y0) < 2 and r.x0 >= h.x0 - 1]
                    return pno, (same[:1] or [h])
        return None

    def _context_tile(self, cm, width=520, height=210):
        img = Image.new("RGB", (width, height), brand.WHITE)
        d = ImageDraw.Draw(img)
        f_lbl, f_body = self._font(True, 13), self._font(False, 14)
        d.rectangle([0, 0, width - 1, height - 1], outline=brand.GREY_LINE, width=2)
        d.rectangle([0, 0, width - 1, 28], fill=brand.DARK_BLUE)
        d.text((11, 7), "Commented text (page view unavailable)",
               fill=brand.WHITE, font=f_lbl)
        context = cm.get("context") or cm.get("anchor") or "(anchor text not found)"
        anchor = cm.get("anchor") or ""
        y = 44
        for line in self._wrap(context, 58)[:5]:
            if anchor and anchor[:20] and anchor[:20].lower() in line.lower():
                bb = d.textbbox((15, y), line, font=f_body)
                d.rectangle([bb[0]-2, bb[1]-2, bb[2]+2, bb[3]+2], fill=brand.SKY_PALE)
            d.text((15, y), line, fill=brand.INK, font=f_body)
            y += 23
        d.rectangle([11, 36, width - 11, min(height - 34, y + 4)],
                    outline=BOX_COLOR, width=2)
        d.text((15, height - 25),
               f"{cm.get('author','Unknown')}  |  {cm.get('date','')}",
               fill=brand.MUTED, font=f_lbl)
        return img

    def extract_word_annotations(self, docx_path):
        self._say(f"Reading {Path(docx_path).name}", 5)
        try:
            comments = self._read_docx_comments(docx_path)
        except Exception as e:
            self.logger.error(f"Word parse failed: {e}")
            comments = []
        if not comments:
            self._say("No comments found", 100)
            return []

        self._say(f"Found {len(comments)} comment(s)", 15)
        self._check_cancel()

        pdf_path = self._docx_to_pdf(docx_path)
        doc = None
        if pdf_path:
            try:
                doc = fitz.open(pdf_path)
            except Exception as e:
                self.logger.warning(f"Rendered PDF unreadable: {e}")

        staged, used = [], set()
        total = len(comments)
        try:
            for idx, cm in enumerate(comments, 1):
                self._check_cancel()
                page_no, detail, overview = "n/a", None, None
                if doc is not None:
                    target = cm["anchor"] or cm["context"] or cm["content"]
                    found = self._locate_in_pdf(doc, target, used)
                    if not found and cm["context"]:
                        found = self._locate_in_pdf(doc, cm["context"], used)
                    if found:
                        pno, rects = found
                        page_no = pno + 1
                        detail = self._detail_snapshot(doc[pno], rects,
                                                       full_width=True)
                        overview = self._overview_snapshot(doc[pno], rects)
                if detail is None:
                    detail = self._context_tile(cm)
                staged.append({
                    "page": page_no, "type": "Comment", "content": cm["content"],
                    "author": cm["author"], "date": cm["date"],
                    "image": detail, "overview": overview, "source": "Word",
                })
                self._say(f"Comment {idx} of {total}",
                          40 + int(38 * idx / max(1, total)))
        finally:
            if doc is not None:
                doc.close()
            self._page_cache.clear()

        # Page-scoped references, preserving document order
        counters: Dict[object, int] = {}
        for item in staged:
            key = item["page"]
            counters[key] = counters.get(key, 0) + 1
            item["ref"] = (f"{key}.{counters[key]}"
                           if key != "n/a" else f"-.{counters[key]}")
        return staged

    # ==================================================================
    def extract(self, file_path):
        ext = Path(file_path).suffix.lower()
        if ext == ".pdf":
            return self.extract_pdf_annotations(file_path)
        if ext == ".docx":
            return self.extract_word_annotations(file_path)
        if ext == ".doc":
            raise ValueError(
                "Legacy .doc is not supported. Please save the file as .docx and retry.")
        raise ValueError(f"Unsupported file format: {ext}")

    # ==================================================================
    # BRANDED EXCEL REPORT
    # ==================================================================
    def create_excel_report(self, annotations, output_path,
                            source_name="", template_path=None):
        self._say(f"Writing report ({len(annotations)} rows)...", 84)

        if template_path and os.path.exists(template_path):
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            start_row = self._next_free_row(ws)
            branded = False
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Annotations"
            self._write_branded_header(ws, source_name, len(annotations))
            start_row = 5
            branded = True

        row = start_row
        total = len(annotations)
        for idx, a in enumerate(annotations, 1):
            self._check_cancel()
            try:
                h1 = self._place_image(ws, a.get("image"), row, DETAIL_COL)
                h2 = self._place_image(ws, a.get("overview"), row, OVERVIEW_COL)

                ws.cell(row=row, column=1, value=a.get("ref", str(idx)))
                ws.cell(row=row, column=2, value=a.get("source", ""))
                ws.cell(row=row, column=3, value=a.get("page", ""))
                ws.cell(row=row, column=4, value=a.get("type", ""))
                ws.cell(row=row, column=5, value=a.get("author", ""))
                ws.cell(row=row, column=6, value=a.get("date", ""))
                ws.cell(row=row, column=7, value=a.get("content", ""))
                ws.cell(row=row, column=10, value="Open")
                ws.cell(row=row, column=11, value="")

                for col in range(1, len(HEADERS) + 1):
                    c = ws.cell(row=row, column=col)
                    c.font = REF_FONT if col == 1 else BODY_FONT
                    c.alignment = Alignment(
                        wrap_text=True, vertical="top",
                        horizontal="left" if col in (7, 11) else "center")
                    c.border = CELL_BORDER
                    if idx % 2 == 0:
                        c.fill = ALT_FILL

                ws.row_dimensions[row].height = max(
                    100, min(500, max(h1, h2) * 0.76 + 12))
                row += 1
                if idx % 10 == 0 or idx == total:
                    self._say(f"Writing row {idx} of {total}",
                              84 + int(13 * idx / max(1, total)))
            except Cancelled:
                raise
            except Exception as e:
                self.logger.error(f"Row {idx} failed: {e}")

        if branded:
            ws.freeze_panes = "A5"
            try:
                ws.auto_filter.ref = f"A4:{get_column_letter(len(HEADERS))}{row-1}"
            except Exception:
                pass

        self._say("Saving workbook...", 98)
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        return output_path

    def _write_branded_header(self, ws, source_name, count):
        ncols = len(HEADERS)
        ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=ncols)
        t = ws.cell(row=1, column=1)
        t.value = f"        {brand.APP_NAME}   |   {brand.APP_STRAP}"
        t.font = TITLE_FONT
        t.fill = TITLE_FILL
        t.alignment = Alignment(horizontal="left", vertical="center")
        for col in range(1, ncols + 1):
            for r in (1, 2):
                ws.cell(row=r, column=col).fill = TITLE_FILL
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 20

        try:
            mark = brand.logo_pil()
            canvas = Image.new("RGB", mark.size, brand.DARK_BLUE)
            canvas.paste(mark, (0, 0), mark)
            canvas = canvas.resize((58, 58), Image.Resampling.LANCZOS)
            buf = io.BytesIO(); canvas.save(buf, format="PNG"); buf.seek(0)
            xl = XLImage(buf)
            xl.anchor = OneCellAnchor(
                _from=AnchorMarker(col=0, row=0, colOff=pixels_to_EMU(5),
                                   rowOff=pixels_to_EMU(4)),
                ext=XDRPositiveSize2D(pixels_to_EMU(58), pixels_to_EMU(58)))
            ws.add_image(xl)
        except Exception as e:
            self.logger.warning(f"Brand mark not embedded: {e}")

        extra = ""
        if self.skipped_markup:
            extra = (f"     |     Markup without text skipped: "
                     f"{self.skipped_markup}")
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
        m = ws.cell(row=3, column=1)
        m.value = (f"Source document: {source_name}     |     "
                   f"Annotations reported: {count}{extra}     |     "
                   f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')}     |     "
                   f"{brand.APP_TAGLINE}")
        m.font = Font(size=9.5, color=brand.XL_DARK_BLUE, name="Segoe UI", bold=True)
        m.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in range(1, ncols + 1):
            ws.cell(row=3, column=col).fill = PatternFill(
                "solid", fgColor=brand.XL_SKY_PALE)
        ws.row_dimensions[3].height = 20

        for col, name in enumerate(HEADERS, 1):
            c = ws.cell(row=4, column=col, value=name)
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
            c.border = CELL_BORDER
        ws.row_dimensions[4].height = 30

        for col, width in COL_WIDTHS.items():
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.sheet_view.showGridLines = False

    @staticmethod
    def _next_free_row(ws):
        r = 2
        while any(ws.cell(row=r, column=c).value is not None for c in (1, 7)):
            r += 1
        return max(r, 2)

    def _place_image(self, ws, img, row, col):
        if img is None:
            return 0
        try:
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            xl = XLImage(buf)
            xl.anchor = OneCellAnchor(
                _from=AnchorMarker(col=col - 1, row=row - 1,
                                   colOff=pixels_to_EMU(4), rowOff=pixels_to_EMU(4)),
                ext=XDRPositiveSize2D(pixels_to_EMU(img.width),
                                      pixels_to_EMU(img.height)))
            ws.add_image(xl)
            return img.height
        except Exception as e:
            self.logger.error(f"Image not embedded: {e}")
            return 0

    # ==================================================================
    def process(self, file_path, output_dir=None, template_path=None):
        """End-to-end: extract + report. Returns (output_path, count, skipped)."""
        try:
            annotations = self.extract(file_path)
            if not annotations:
                if self.skipped_markup:
                    raise ValueError(
                        f"No annotations with text were found.\n\n"
                        f"{self.skipped_markup} shape/markup object(s) were present "
                        f"but none contained reviewer text.")
                raise ValueError(
                    "No annotations or comments were found in this document.")
            self._check_cancel()

            out_dir = Path(output_dir) if output_dir else (
                Path.home() / f"{brand.APP_NAME}_Reports")
            out_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = out_dir / f"{Path(file_path).stem}_{brand.APP_NAME}_{stamp}.xlsx"

            self.create_excel_report(annotations, str(out_path),
                                     source_name=Path(file_path).name,
                                     template_path=template_path)
            self._say("Complete", 100)
            return str(out_path), len(annotations), self.skipped_markup
        finally:
            self.cleanup()
