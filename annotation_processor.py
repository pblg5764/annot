"""
AnotAI - Core Engine
====================
Extracts annotations/comments from PDF and Word (.docx) documents and builds a
branded Excel report where every row carries a real cropped snapshot showing
exactly where that annotation sits on the page.

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

# ---------------------------------------------------------------------------
# Excel layout
# ---------------------------------------------------------------------------
HEADERS = [
    "#",
    "Source",
    "Page",
    "Annotation Type",
    "Author",
    "Date",
    "Annotation Text",
    "Location Snapshot",
    "Status",
    "Reviewer Remarks",
]

COL_WIDTHS = {
    1: 6, 2: 10, 3: 7, 4: 18, 5: 22, 6: 18, 7: 52, 8: 62, 9: 12, 10: 26,
}

SNAPSHOT_COL = 8
SNAPSHOT_PX_WIDTH = 430
SNAPSHOT_MAX_PX_HEIGHT = 360

# Branded styles
TITLE_FILL  = PatternFill("solid", fgColor=brand.XL_DARK_BLUE)
HEADER_FILL = PatternFill("solid", fgColor=brand.XL_DARK_BLUE)
HEADER_FONT = Font(bold=True, size=11, color=brand.XL_WHITE, name="Segoe UI")
TITLE_FONT  = Font(bold=True, size=16, color=brand.XL_WHITE, name="Segoe UI")
BODY_FONT   = Font(size=10, color=brand.XL_BLACK, name="Segoe UI")
ALT_FILL    = PatternFill("solid", fgColor=brand.XL_CREAM)

THIN = Side(style="thin", color=brand.XL_GREY_LINE)
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

BOX_COLOR = brand.BOX_COLOR
BOX_WIDTH = 4

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


class AnnotationExtractor:
    """Extract annotations from PDF / DOCX and emit a branded Excel report."""

    def __init__(self, logger=None, include_noise: bool = False,
                 progress_cb=None):
        self.logger = logger or logging.getLogger(__name__)
        self.include_noise = include_noise
        self.progress_cb = progress_cb
        Image.MAX_IMAGE_PIXELS = None
        self._temp_dirs: List[str] = []

    # ------------------------------------------------------------------
    def _say(self, msg: str):
        self.logger.info(msg)
        if self.progress_cb:
            try:
                self.progress_cb(msg)
            except Exception:
                pass

    @staticmethod
    def _clean_text(text) -> str:
        return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", str(text or "")).strip()

    @staticmethod
    def _squash(text: str) -> str:
        return re.sub(r"\s+", " ", text or "").strip()

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

    # ==================================================================
    # SNAPSHOT ENGINE
    # ==================================================================
    def _context_clip(self, page: fitz.Page, rects: List[fitz.Rect],
                      full_width: bool = False) -> fitz.Rect:
        pr = page.rect
        target = fitz.Rect(rects[0])
        for r in rects[1:]:
            target |= r

        if target.width < 1:
            cx = (target.x0 + target.x1) / 2
            target.x0, target.x1 = cx - 6, cx + 6
        if target.height < 1:
            cy = (target.y0 + target.y1) / 2
            target.y0, target.y1 = cy - 6, cy + 6

        want_w = pr.width if full_width else max(
            target.width * 2.6, pr.width * 0.55, 260.0)
        want_h = max(target.height * 2.6, 150.0)
        want_w = min(want_w, pr.width)
        want_h = min(want_h, pr.height)

        cx = (target.x0 + target.x1) / 2
        cy = (target.y0 + target.y1) / 2
        clip = fitz.Rect(cx - want_w / 2, cy - want_h / 2,
                         cx + want_w / 2, cy + want_h / 2)

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

    def _render_clip(self, page: fitz.Page, clip: fitz.Rect,
                     boxes: List[fitz.Rect]) -> Optional[Image.Image]:
        longest = max(clip.width, clip.height)
        if longest < 160:
            dpi = 400
        elif longest < 340:
            dpi = 300
        elif longest < 700:
            dpi = 220
        else:
            dpi = 170

        img = None
        used_zoom = 1.0
        for attempt_dpi in (dpi, 150, 100):
            try:
                zoom = attempt_dpi / 72.0
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, clip=clip, annots=True, alpha=False)
                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                used_zoom = zoom
                break
            except Exception as e:
                self.logger.warning(f"Pixmap render failed at {attempt_dpi} DPI: {e}")
        if img is None:
            return None

        draw = ImageDraw.Draw(img)
        for b in boxes:
            x0 = (b.x0 - clip.x0) * used_zoom
            y0 = (b.y0 - clip.y0) * used_zoom
            x1 = (b.x1 - clip.x0) * used_zoom
            y1 = (b.y1 - clip.y0) * used_zoom
            pad = 3
            x0, y0 = max(0, x0 - pad), max(0, y0 - pad)
            x1 = min(img.width - 1, x1 + pad)
            y1 = min(img.height - 1, y1 + pad)
            if x1 <= x0:
                x1 = min(img.width - 1, x0 + 10)
            if y1 <= y0:
                y1 = min(img.height - 1, y0 + 10)
            draw.rectangle([x0, y0, x1, y1], outline=BOX_COLOR, width=BOX_WIDTH)

        draw.rectangle([0, 0, img.width - 1, img.height - 1],
                       outline=(150, 150, 150), width=2)
        return img

    def _snapshot_for_rects(self, page: fitz.Page, rects: List[fitz.Rect],
                            full_width: bool = False) -> Optional[Image.Image]:
        try:
            clip = self._context_clip(page, rects, full_width=full_width)
            return self._render_clip(page, clip, rects)
        except Exception as e:
            self.logger.error(f"Snapshot failed: {e}")
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

    def _placeholder(self, title: str, subtitle: str = "",
                     width: int = 420, height: int = 190) -> Image.Image:
        img = Image.new("RGB", (width, height), brand.CREAM)
        d = ImageDraw.Draw(img)
        d.rectangle([0, 0, width - 1, height - 1], outline=brand.GREY_LINE, width=2)
        f_big, f_sm = self._font(True, 16), self._font(False, 12)

        lines = [(title, f_big, brand.DARK_BLUE)]
        for chunk in self._wrap(subtitle, 46):
            lines.append((chunk, f_sm, brand.MUTED))

        total = sum((f.getbbox(t)[3] - f.getbbox(t)[1]) + 10 for t, f, _ in lines)
        y = max(12, (height - total) // 2)
        for text, font, colour in lines:
            bb = d.textbbox((0, 0), text, font=font)
            d.text(((width - (bb[2] - bb[0])) // 2, y), text, fill=colour, font=font)
            y += (bb[3] - bb[1]) + 10
        return img

    @staticmethod
    def _wrap(text: str, width: int) -> List[str]:
        if not text:
            return []
        words, lines, cur = text.split(), [], ""
        for w in words:
            if len(cur) + len(w) + 1 <= width:
                cur = f"{cur} {w}".strip()
            else:
                lines.append(cur)
                cur = w
            if len(lines) >= 4:
                break
        if cur and len(lines) < 5:
            lines.append(cur)
        return lines

    # ==================================================================
    # PDF EXTRACTION
    # ==================================================================
    def _annot_rects(self, annot) -> List[fitz.Rect]:
        rects: List[fitz.Rect] = []
        try:
            vs = annot.vertices
            if vs and len(vs) >= 4 and annot.type[1] in (
                    "Highlight", "Underline", "Squiggly", "StrikeOut"):
                for i in range(0, len(vs) - 3, 4):
                    r = fitz.Quad(vs[i], vs[i + 1], vs[i + 2], vs[i + 3]).rect
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

        candidates = []
        try:
            popup = annot.get_popup()
            if popup is not None and popup.info:
                candidates.append(popup.info.get("content", ""))
        except Exception:
            pass
        candidates.append(info.get("content", ""))
        candidates.append(info.get("subject", ""))

        content = ""
        for c in candidates:
            c = self._clean_text(c)
            if c and c.lower() not in ("none", "null"):
                content = c
                break

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

        label = TYPE_LABELS.get(subtype, subtype)
        if not content:
            if underlying:
                content = f'[{label} on: "{underlying}"]'
            elif subtype in ("Text", "FreeText"):
                content = "[Empty sticky note - no text entered]"
            else:
                content = f"[{label} - no text entered]"
        elif underlying:
            content = f'{content}\n\nMarked text: "{underlying}"'

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
                "raw_type": subtype, "date": date_out}

    def extract_pdf_annotations(self, pdf_path: str,
                                source_label: str = "PDF") -> List[Dict]:
        self._say(f"Opening PDF: {Path(pdf_path).name}")
        doc = fitz.open(pdf_path)
        out: List[Dict] = []
        total_pages = len(doc)

        for pno in range(total_pages):
            page = doc[pno]
            try:
                annots = list(page.annots() or [])
            except Exception as e:
                self.logger.error(f"Cannot read annotations on page {pno+1}: {e}")
                continue
            if not annots:
                continue

            annots.sort(key=lambda a: (round(a.rect.y0, 1), round(a.rect.x0, 1)))
            self._say(f"Page {pno+1}/{total_pages}: {len(annots)} annotation(s)")

            for annot in annots:
                try:
                    subtype = annot.type[1] if annot.type else "Unknown"
                    if not self.include_noise and subtype in NOISE_TYPES:
                        continue
                    meta = self._annot_info(annot)
                    rects = self._annot_rects(annot)
                    img = self._snapshot_for_rects(page, rects)
                    if img is None:
                        img = self._placeholder(
                            meta["type"], f"Page {pno+1} - snapshot unavailable")
                    out.append({
                        "page": pno + 1, "type": meta["type"],
                        "content": meta["content"], "author": meta["author"],
                        "date": meta["date"], "image": img, "source": source_label,
                    })
                except Exception as e:
                    self.logger.error(f"Error on annotation, page {pno+1}: {e}")

        doc.close()
        self._say(f"Extracted {len(out)} annotation(s) from PDF")
        return out

    # ==================================================================
    # WORD EXTRACTION
    # ==================================================================
    def _find_word_converter(self) -> Optional[str]:
        for name in ("soffice", "soffice.exe", "libreoffice"):
            p = shutil.which(name)
            if p:
                return p
        for candidate in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if os.path.exists(candidate):
                return candidate
        return None

    def _docx_to_pdf(self, docx_path: str) -> Optional[str]:
        outdir = self._mktemp()
        target = os.path.join(outdir, Path(docx_path).stem + ".pdf")

        try:
            import win32com.client  # type: ignore
            self._say("Rendering Word document via Microsoft Word...")
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            docm = None
            try:
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
                if docm is not None:
                    docm.Close(False)
                word.Quit()
            if os.path.exists(target):
                self._say("Word rendering complete (comments shown in margin).")
                return target
        except ImportError:
            pass
        except Exception as e:
            self.logger.warning(f"Microsoft Word conversion unavailable: {e}")

        soffice = self._find_word_converter()
        if soffice:
            try:
                self._say("Rendering Word document via LibreOffice...")
                subprocess.run(
                    [soffice, "--headless", "--norestore", "--convert-to", "pdf",
                     "--outdir", outdir, os.path.abspath(docx_path)],
                    check=True, timeout=240,
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                if os.path.exists(target):
                    self._say("LibreOffice rendering complete.")
                    return target
                pdfs = list(Path(outdir).glob("*.pdf"))
                if pdfs:
                    return str(pdfs[0])
            except Exception as e:
                self.logger.warning(f"LibreOffice conversion failed: {e}")

        self.logger.warning("No DOCX renderer available - using text-context tiles.")
        return None

    # ------------------------------------------------------------------
    def _read_docx_comments(self, docx_path: str) -> List[Dict]:
        doc = Document(docx_path)

        comments_part = None
        for rel in doc.part.rels.values():
            if rel.reltype.endswith("/comments"):
                comments_part = rel.target_part
                break
        if comments_part is None:
            return []

        meta: Dict[str, Dict] = {}
        order: List[str] = []
        for c in comments_part.element.findall(f".//{W_NS}comment"):
            cid = c.get(f"{W_NS}id") or c.get("id") or str(len(order))
            texts = [t.text for t in c.findall(f".//{W_NS}t") if t.text]
            meta[cid] = {
                "id": cid,
                "author": self._clean_text(
                    c.get(f"{W_NS}author") or c.get("author") or "") or "Unknown",
                "initials": c.get(f"{W_NS}initials") or "",
                "date": self._parse_docx_date(
                    c.get(f"{W_NS}date") or c.get("date") or ""),
                "content": self._squash(" ".join(texts)) or "[Empty comment]",
                "anchor": "", "context": "",
            }
            order.append(cid)

        body = doc.element.body
        active: set = set()
        anchors: Dict[str, List[str]] = {cid: [] for cid in meta}
        contexts: Dict[str, str] = {}
        para_text: List[str] = []
        para_open: set = set()

        for el in body.iter():
            tag = el.tag
            if tag == f"{W_NS}p":
                joined = self._squash("".join(para_text))
                for cid in para_open:
                    if joined and not contexts.get(cid):
                        contexts[cid] = joined
                para_text = []
                para_open = set(active)
            elif tag == f"{W_NS}commentRangeStart":
                cid = el.get(f"{W_NS}id") or el.get("id")
                if cid is not None:
                    active.add(cid)
                    para_open.add(cid)
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

    # ------------------------------------------------------------------
    def _locate_in_pdf(self, doc: fitz.Document, needle: str,
                       used: set) -> Optional[Tuple[int, List[fitz.Rect]]]:
        if not needle:
            return None
        n = self._squash(needle)
        probes = []
        for length in (len(n), 90, 60, 40, 25, 15):
            probe = n[:length].strip()
            if len(probe) >= 4 and probe not in probes:
                probes.append(probe)

        for probe in probes:
            for pno in range(len(doc)):
                page = doc[pno]
                try:
                    hits = page.search_for(probe, quads=False)
                except Exception:
                    continue
                if not hits:
                    continue
                for h in hits:
                    key = (pno, round(h.x0), round(h.y0), round(h.x1))
                    if key in used:
                        continue
                    used.add(key)
                    same_line = [r for r in hits
                                 if abs(r.y0 - h.y0) < 2 and r.x0 >= h.x0 - 1]
                    return pno, (same_line[:1] or [h])
        return None

    def extract_word_annotations(self, docx_path: str) -> List[Dict]:
        self._say(f"Reading Word comments: {Path(docx_path).name}")
        try:
            comments = self._read_docx_comments(docx_path)
        except Exception as e:
            self.logger.error(f"Failed to parse Word comments: {e}")
            comments = []

        if not comments:
            self._say("No comments found in the Word document.")
            return []

        self._say(f"Found {len(comments)} comment(s). Preparing snapshots...")

        pdf_path = self._docx_to_pdf(docx_path)
        doc = None
        if pdf_path:
            try:
                doc = fitz.open(pdf_path)
            except Exception as e:
                self.logger.warning(f"Could not open rendered PDF: {e}")

        results: List[Dict] = []
        used: set = set()

        for idx, cm in enumerate(comments, 1):
            page_no = "n/a"
            img = None
            if doc is not None:
                target = cm["anchor"] or cm["context"] or cm["content"]
                found = self._locate_in_pdf(doc, target, used)
                if not found and cm["context"]:
                    found = self._locate_in_pdf(doc, cm["context"], used)
                if found:
                    pno, rects = found
                    page_no = pno + 1
                    img = self._snapshot_for_rects(doc[pno], rects, full_width=True)
            if img is None:
                img = self._context_tile(cm)

            results.append({
                "page": page_no, "type": "Comment", "content": cm["content"],
                "author": cm["author"], "date": cm["date"],
                "image": img, "source": "Word",
            })
            if idx % 5 == 0:
                self._say(f"Processed {idx}/{len(comments)} comments...")

        if doc is not None:
            doc.close()
        self._say(f"Extracted {len(results)} comment(s) from Word")
        return results

    # ------------------------------------------------------------------
    def _context_tile(self, cm: Dict, width: int = 430,
                      height: int = 200) -> Image.Image:
        """Branded fallback tile with the anchored span highlighted."""
        img = Image.new("RGB", (width, height), brand.WHITE)
        d = ImageDraw.Draw(img)
        f_lbl, f_body = self._font(True, 12), self._font(False, 13)

        d.rectangle([0, 0, width - 1, height - 1], outline=brand.GREY_LINE, width=2)
        d.rectangle([0, 0, width - 1, 26], fill=brand.DARK_BLUE)
        d.text((10, 6), "Commented text (document view unavailable)",
               fill=brand.WHITE, font=f_lbl)

        context = cm.get("context") or cm.get("anchor") or "(anchor text not found)"
        anchor = cm.get("anchor") or ""

        y = 40
        for line in self._wrap(context, 52)[:5]:
            if anchor and anchor[:20] and anchor[:20].lower() in line.lower():
                bb = d.textbbox((14, y), line, font=f_body)
                d.rectangle([bb[0] - 2, bb[1] - 2, bb[2] + 2, bb[3] + 2],
                            fill=brand.SKY_PALE)
            d.text((14, y), line, fill=brand.INK, font=f_body)
            y += 22

        d.rectangle([10, 34, width - 10, min(height - 34, y + 4)],
                    outline=BOX_COLOR, width=2)
        d.text((14, height - 24),
               f"{cm.get('author','Unknown')}  |  {cm.get('date','')}",
               fill=brand.MUTED, font=f_lbl)
        return img

    # ==================================================================
    def extract(self, file_path: str) -> List[Dict]:
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
    def create_excel_report(self, annotations: List[Dict], output_path: str,
                            source_name: str = "", template_path: str = None) -> str:
        self._say(f"Building {brand.APP_NAME} report ({len(annotations)} rows)...")

        if template_path and os.path.exists(template_path):
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            start_row = self._next_free_row(ws)
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Annotations"
            self._write_branded_header(ws, source_name, len(annotations))
            start_row = 5

        row = start_row
        for idx, a in enumerate(annotations, 1):
            try:
                img_h_px = self._place_image(ws, a.get("image"), row)

                ws.cell(row=row, column=1, value=idx)
                ws.cell(row=row, column=2, value=a.get("source", ""))
                ws.cell(row=row, column=3, value=a.get("page", ""))
                ws.cell(row=row, column=4, value=a.get("type", ""))
                ws.cell(row=row, column=5, value=a.get("author", ""))
                ws.cell(row=row, column=6, value=a.get("date", ""))
                ws.cell(row=row, column=7, value=a.get("content", ""))
                ws.cell(row=row, column=9, value="Open")
                ws.cell(row=row, column=10, value="")

                for col in range(1, len(HEADERS) + 1):
                    c = ws.cell(row=row, column=col)
                    c.font = BODY_FONT
                    c.alignment = Alignment(
                        wrap_text=True, vertical="top",
                        horizontal="left" if col in (7, 10) else "center")
                    c.border = CELL_BORDER
                    if idx % 2 == 0:
                        c.fill = ALT_FILL

                ws.row_dimensions[row].height = max(
                    90, min(SNAPSHOT_MAX_PX_HEIGHT * 0.78, img_h_px * 0.78 + 10))
                row += 1
            except Exception as e:
                self.logger.error(f"Failed writing row {idx}: {e}")

        if start_row == 5:
            ws.freeze_panes = "A5"
            try:
                ws.auto_filter.ref = f"A4:{get_column_letter(len(HEADERS))}{row-1}"
            except Exception:
                pass

        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        self._say(f"Report saved: {output_path}")
        return output_path

    # ------------------------------------------------------------------
    def _write_branded_header(self, ws, source_name: str, count: int):
        ncols = len(HEADERS)

        # Rows 1-2: navy brand banner with logo
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
            buf = io.BytesIO()
            canvas.save(buf, format="PNG")
            buf.seek(0)
            xl = XLImage(buf)
            xl.anchor = OneCellAnchor(
                _from=AnchorMarker(col=0, row=0,
                                   colOff=pixels_to_EMU(5), rowOff=pixels_to_EMU(4)),
                ext=XDRPositiveSize2D(pixels_to_EMU(58), pixels_to_EMU(58)))
            ws.add_image(xl)
        except Exception as e:
            self.logger.warning(f"Could not embed brand mark: {e}")

        # Row 3: run metadata band
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
        m = ws.cell(row=3, column=1)
        m.value = (f"Source document: {source_name}     |     "
                   f"Annotations found: {count}     |     "
                   f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')}     |     "
                   f"{brand.APP_TAGLINE}")
        m.font = Font(size=9.5, color=brand.XL_DARK_BLUE, name="Segoe UI", bold=True)
        m.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in range(1, ncols + 1):
            ws.cell(row=3, column=col).fill = PatternFill(
                "solid", fgColor=brand.XL_SKY_PALE)
        ws.row_dimensions[3].height = 20

        # Row 4: column headers
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
    def _next_free_row(ws) -> int:
        r = 2
        while any(ws.cell(row=r, column=c).value is not None for c in (1, 7)):
            r += 1
        return max(r, 2)

    def _place_image(self, ws, img: Optional[Image.Image], row: int) -> int:
        if img is None:
            return 90
        try:
            work = img.copy()
            if work.width != SNAPSHOT_PX_WIDTH:
                ratio = SNAPSHOT_PX_WIDTH / float(work.width)
                work = work.resize((SNAPSHOT_PX_WIDTH, max(1, int(work.height * ratio))),
                                   Image.Resampling.LANCZOS)
            if work.height > SNAPSHOT_MAX_PX_HEIGHT:
                ratio = SNAPSHOT_MAX_PX_HEIGHT / float(work.height)
                work = work.resize((max(1, int(work.width * ratio)), SNAPSHOT_MAX_PX_HEIGHT),
                                   Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            work.save(buf, format="PNG")
            buf.seek(0)
            xl = XLImage(buf)
            xl.anchor = OneCellAnchor(
                _from=AnchorMarker(col=SNAPSHOT_COL - 1, row=row - 1,
                                   colOff=pixels_to_EMU(4), rowOff=pixels_to_EMU(4)),
                ext=XDRPositiveSize2D(pixels_to_EMU(work.width),
                                      pixels_to_EMU(work.height)))
            ws.add_image(xl)
            return work.height
        except Exception as e:
            self.logger.error(f"Could not embed snapshot: {e}")
            return 90

    # ==================================================================
    def process(self, file_path: str, output_dir: str = None,
                template_path: str = None) -> Tuple[str, int]:
        """End-to-end: extract + report. Returns (output_path, count)."""
        try:
            annotations = self.extract(file_path)
            if not annotations:
                raise ValueError(
                    "No annotations or comments were found in this document.")

            out_dir = Path(output_dir) if output_dir else (
                Path.home() / f"{brand.APP_NAME}_Reports")
            out_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = out_dir / f"{Path(file_path).stem}_{brand.APP_NAME}_{stamp}.xlsx"

            self.create_excel_report(annotations, str(out_path),
                                     source_name=Path(file_path).name,
                                     template_path=template_path)
            return str(out_path), len(annotations)
        finally:
            self.cleanup()
