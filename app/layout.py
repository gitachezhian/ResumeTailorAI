"""Approximate Word page-count estimation for a .docx.

There is no Word / LibreOffice on this machine, so we simulate line wrapping
with real font metrics (fontTools) and the document's own page size, margins,
indents, spacing and font sizes. Accurate to within a few lines for simple,
single-column documents like resumes.
"""
from __future__ import annotations

import io
import os
import zipfile
from functools import lru_cache

from docx.oxml.ns import qn
from fontTools.ttLib import TTFont

WIN_FONTS = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")

# family -> (regular, bold) file names in C:\Windows\Fonts
FONT_FILES = {
    "garamond": ("GARA.TTF", "GARABD.TTF"),
    "eb garamond": ("GARA.TTF", "GARABD.TTF"),
    "calibri": ("calibri.ttf", "calibrib.ttf"),
    "arial": ("arial.ttf", "arialbd.ttf"),
    "times new roman": ("times.ttf", "timesbd.ttf"),
    "georgia": ("georgia.ttf", "georgiab.ttf"),
    "cambria": ("cambria.ttc", "cambriab.ttf"),
    "verdana": ("verdana.ttf", "verdanab.ttf"),
    "tahoma": ("tahoma.ttf", "tahomabd.ttf"),
    "segoe ui": ("segoeui.ttf", "segoeuib.ttf"),
    "helvetica": ("arial.ttf", "arialbd.ttf"),
    "aptos": ("aptos.ttf", "aptos-bold.ttf"),
    "trebuchet ms": ("trebuc.ttf", "trebucbd.ttf"),
    "book antiqua": ("BKANT.TTF", "ANTQUAB.TTF"),
    "palatino linotype": ("pala.ttf", "palab.ttf"),
}


class FontMetrics:
    def __init__(self, data: bytes | str):
        f = TTFont(io.BytesIO(data) if isinstance(data, bytes) else data, fontNumber=0, lazy=True)
        self.upm = f["head"].unitsPerEm
        os2 = f["OS/2"]
        hhea = f["hhea"]
        asc = max(os2.usWinAscent, hhea.ascent)
        desc = max(os2.usWinDescent, -hhea.descent)
        # Word "single" spacing ~= win ascent + win descent (+ hhea line gap)
        self.line_factor = (asc + desc + max(0, hhea.lineGap)) / self.upm
        self.cmap = f.getBestCmap() or {}
        self.hmtx = f["hmtx"].metrics
        self._cache: dict[str, int] = {}
        self.default_w = self.hmtx.get(self.cmap.get(ord("n"), ".notdef"), (self.upm // 2, 0))[0]

    def char_w(self, ch: str) -> int:
        w = self._cache.get(ch)
        if w is None:
            g = self.cmap.get(ord(ch))
            w = self.hmtx[g][0] if g in self.hmtx else self.default_w
            self._cache[ch] = w
        return w

    def width_tw(self, text: str, half_pts: int) -> float:
        units = sum(self.char_w(c) for c in text)
        return units / self.upm * half_pts * 10  # half-points * 10 = twips per em


@lru_cache(maxsize=64)
def _load(path: str) -> FontMetrics | None:
    try:
        return FontMetrics(path)
    except Exception:
        return None


class FontBook:
    def __init__(self, docx_path: str):
        self.embedded: dict[tuple[str, bool], FontMetrics] = {}
        try:
            with zipfile.ZipFile(docx_path) as z:
                for n in z.namelist():
                    low = n.lower()
                    if low.startswith("word/fonts/") and low.endswith(".ttf"):
                        base = os.path.basename(low)[:-4]
                        fam, _, style = base.partition("-")
                        bold = "bold" in style and "italic" not in style
                        if style in ("regular", "bold") or (not style):
                            try:
                                self.embedded[(fam, bold)] = FontMetrics(z.read(n))
                            except Exception:
                                pass
        except Exception:
            pass

    def get(self, family: str, bold: bool) -> FontMetrics:
        fam = (family or "times new roman").lower()
        m = self.embedded.get((fam, bold)) or self.embedded.get((fam, False))
        if m:
            return m
        files = FONT_FILES.get(fam)
        if files:
            m = _load(os.path.join(WIN_FONTS, files[1 if bold else 0])) or _load(os.path.join(WIN_FONTS, files[0]))
            if m:
                return m
        return _load(os.path.join(WIN_FONTS, "timesbd.ttf" if bold else "times.ttf"))


def _int(el, attr, default=0):
    if el is None:
        return default
    v = el.get(qn(attr))
    try:
        return int(v) if v is not None else default
    except ValueError:
        return default


def estimate_pages(document, docx_path: str) -> float:
    """Return an estimated page count (e.g. 1.84) for a python-docx Document."""
    book = FontBook(docx_path)

    # defaults
    styles = document.styles.element
    dflt_rpr = styles.find(f"{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}")
    dflt_font = "Times New Roman"
    dflt_sz = 20
    if dflt_rpr is not None:
        rf = dflt_rpr.find(qn("w:rFonts"))
        if rf is not None and rf.get(qn("w:ascii")):
            dflt_font = rf.get(qn("w:ascii"))
        dflt_sz = _int(dflt_rpr.find(qn("w:sz")), "w:val", dflt_sz)

    sect = document.sections[-1]
    tw = lambda emu: (emu or 0) / 635  # EMU -> twips
    page_h = tw(sect.page_height)
    top = tw(sect.top_margin)
    bottom = tw(sect.bottom_margin)
    text_w = tw(sect.page_width) - tw(sect.left_margin) - tw(sect.right_margin)
    header_dist = tw(sect.header_distance) if sect.header_distance is not None else 720
    # an (empty) header paragraph pushes the body down if it is taller than the margin
    try:
        hdr_has = sect.header is not None and not sect.is_linked_to_previous
        if hdr_has and len(sect.header.paragraphs):
            hm = book.get(dflt_font, False)
            hdr_h = sum(dflt_sz * 10 * hm.line_factor for _ in sect.header.paragraphs)
            top = max(top, header_dist + hdr_h)
    except Exception:
        pass
    usable = page_h - top - bottom

    pages = 1
    y = 0.0

    for p in document.element.body.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        spacing = ppr.find(qn("w:spacing")) if ppr is not None else None
        before = _int(spacing, "w:before")
        after = _int(spacing, "w:after")
        line = _int(spacing, "w:line", 240)
        rule = spacing.get(qn("w:lineRule")) if spacing is not None else None
        ind = ppr.find(qn("w:ind")) if ppr is not None else None
        left = _int(ind, "w:left", _int(ind, "w:start"))
        right = _int(ind, "w:right", _int(ind, "w:end"))
        avail = max(1000.0, text_w - left - right)

        # paragraph mark size
        mark_sz = dflt_sz
        mark_font = dflt_font
        if ppr is not None:
            prpr = ppr.find(qn("w:rPr"))
            if prpr is not None:
                mark_sz = _int(prpr.find(qn("w:sz")), "w:val", mark_sz)
                rf = prpr.find(qn("w:rFonts"))
                if rf is not None and rf.get(qn("w:ascii")):
                    mark_font = rf.get(qn("w:ascii"))

        # gather text pieces
        pieces = []  # (text, metrics, sz)
        has_tab = False
        max_sz = 0
        factor = book.get(mark_font, False).line_factor
        for r in p.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            sz = _int(rpr.find(qn("w:sz")) if rpr is not None else None, "w:val", dflt_sz)
            fam = dflt_font
            bold = False
            if rpr is not None:
                rf = rpr.find(qn("w:rFonts"))
                if rf is not None and rf.get(qn("w:ascii")):
                    fam = rf.get(qn("w:ascii"))
                b = rpr.find(qn("w:b"))
                bold = b is not None and b.get(qn("w:val")) not in ("0", "false")
            m = book.get(fam, bold)
            for child in r:
                if child.tag == qn("w:t") and child.text:
                    pieces.append((child.text, m, sz))
                    max_sz = max(max_sz, sz)
                    factor = max(factor, m.line_factor) if max_sz == sz else factor
                elif child.tag == qn("w:tab"):
                    has_tab = True
                    pieces.append(("    ", m, sz))
                elif child.tag in (qn("w:br"), qn("w:cr")):
                    pieces.append(("\n", m, sz))

        sz_for_height = max(max_sz, mark_sz) if max_sz else mark_sz
        if rule in (None, "auto"):
            line_h = sz_for_height * 10 * factor * (line / 240)
        elif rule == "exact":
            line_h = line
        else:  # atLeast
            line_h = max(line, sz_for_height * 10 * factor)

        # count lines
        text_total = "".join(t for t, _, _ in pieces)
        if not text_total.strip():
            n_lines = 1
        elif has_tab and p.find(f".//{qn('w:numPr')}") is None:
            n_lines = 1 + text_total.count("\n")
        else:
            n_lines = 1
            x = 0.0
            for text, m, sz in pieces:
                for seg_i, seg in enumerate(text.split("\n")):
                    if seg_i > 0:
                        n_lines += 1
                        x = 0.0
                    for word in seg.split(" "):
                        w = m.width_tw(word + " ", sz)
                        if x + m.width_tw(word, sz) > avail and x > 0:
                            n_lines += 1
                            x = w
                        else:
                            x += w

        border = 0.0
        if ppr is not None:
            bdr = ppr.find(qn("w:pBdr"))
            if bdr is not None:
                for side in ("w:top", "w:bottom"):
                    b = bdr.find(qn(side))
                    if b is not None:
                        border += _int(b, "w:sz") * 2.5 + _int(b, "w:space") * 20

        block = before + n_lines * line_h + after + border
        # simple pagination: lines that don't fit flow to the next page
        while y + block > usable:
            fit_lines = max(0, int((usable - y - before) // line_h))
            if fit_lines >= 2 and fit_lines < n_lines:
                n_lines -= fit_lines
                block = n_lines * line_h + after + border
            pages += 1
            y = 0.0
            if block <= usable:
                break
        y += block

    return pages - 1 + min(1.0, y / usable)
