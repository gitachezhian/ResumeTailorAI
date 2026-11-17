"""Template-preserving resume engine.

parse_template()  -> list of Para (what the AI sees)
build_resume()    -> new Document with rewritten text dropped into the
                     template's own paragraphs/runs (styling untouched)
build_cover_letter() -> Document that reuses the template's header + fonts
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field

import docx
from docx.oxml.ns import qn

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass
class Para:
    id: int
    kind: str          # contact | heading | entry | bullet | text | empty
    text: str          # markdown-ish: **bold**, \t for tabs
    section: str
    locked: bool
    el: object = field(repr=False, default=None)

    def as_dict(self):
        return {"id": self.id, "kind": self.kind, "text": self.text,
                "section": self.section, "locked": self.locked}


# ---------------------------------------------------------------- helpers
def _is_bold(rpr) -> bool:
    if rpr is None:
        return False
    b = rpr.find(qn("w:b"))
    return b is not None and b.get(qn("w:val")) not in ("0", "false")


def _runs(p_el):
    return list(p_el.iter(qn("w:r")))


def para_markdown(p_el) -> str:
    out = []
    for r in _runs(p_el):
        bold = _is_bold(r.find(qn("w:rPr")))
        buf = ""
        for ch in r:
            if ch.tag == qn("w:t") and ch.text:
                buf += ch.text
            elif ch.tag == qn("w:tab"):
                buf += "\t"
            elif ch.tag in (qn("w:br"), qn("w:cr")):
                buf += "\n"
        if not buf:
            continue
        if bold and buf.strip():
            lead = buf[: len(buf) - len(buf.lstrip())]
            trail = buf[len(buf.rstrip()):]
            out.append(f"{lead}**{buf.strip()}**{trail}")
        else:
            out.append(buf)
    text = "".join(out)
    return text.replace("****", "")


def plain(md: str) -> str:
    return md.replace("**", "")


def _is_heading(p_el, text: str) -> bool:
    ppr = p_el.find(qn("w:pPr"))
    has_border = ppr is not None and ppr.find(qn("w:pBdr")) is not None
    t = plain(text).strip()
    if not t:
        return False
    if has_border:
        return True
    all_bold = all(_is_bold(r.find(qn("w:rPr"))) for r in _runs(p_el)
                   if "".join(x.text or "" for x in r.iter(qn("w:t"))).strip())
    return all_bold and t.isupper() and len(t) < 40 and "\t" not in text.strip()


# ---------------------------------------------------------------- parsing
def parse_template(path: str):
    doc = docx.Document(path)
    paras: list[Para] = []
    section = ""
    seen_heading = False
    for i, p_el in enumerate(doc.element.body.iterchildren(qn("w:p"))):
        text = para_markdown(p_el)
        stripped = plain(text).strip()
        if not stripped:
            kind = "empty"
        elif _is_heading(p_el, text):
            kind = "heading"
            seen_heading = True
            section = stripped.upper()
        elif not seen_heading:
            kind = "contact"
        elif p_el.find(f".//{qn('w:numPr')}") is not None:
            kind = "bullet"
        elif "\t" in text.strip("\t ") or text.rstrip().endswith("\t") or "\t" in text:
            kind = "entry"
        else:
            kind = "text"
        locked = kind in ("contact", "heading", "entry", "empty")
        if kind == "text":
            low = stripped.lower()
            if "EDUCATION" in section or "INTERNSHIP" in section or low.startswith("clients"):
                locked = True  # facts: GPA, coursework, client lists
        paras.append(Para(i, kind, text, section, locked, p_el))
    return doc, paras


def template_outline(paras) -> str:
    """Human-readable dump given to the model."""
    lines = []
    for p in paras:
        if p.kind == "empty":
            continue
        tag = "LOCKED" if p.locked else "EDIT"
        t = p.text.replace("\t", " | ").strip()
        lines.append(f"[{p.id}] ({p.kind}, {tag}) {t}")
    return "\n".join(lines)


def resume_plain_text(paras) -> str:
    return "\n".join(plain(p.text).replace("\t", "  ").strip() for p in paras if p.kind != "empty")


# ---------------------------------------------------------------- writing
_TOKEN = re.compile(r"(\*\*|\t)")


def _segments(md: str):
    """Yield (text, bold, is_tab)."""
    bold = False
    for tok in _TOKEN.split(md):
        if tok == "**":
            bold = not bold
        elif tok == "\t":
            yield ("", bold, True)
        elif tok:
            yield (tok, bold, False)


def _ensure_bold(rpr, on: bool):
    rpr = copy.deepcopy(rpr) if rpr is not None else docx.oxml.parse_xml(
        '<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    for tag in ("w:b", "w:bCs"):
        for e in rpr.findall(qn(tag)):
            rpr.remove(e)
    if on:
        # w:b must come after rFonts/rStyle per schema order
        b = docx.oxml.OxmlElement("w:b")
        bcs = docx.oxml.OxmlElement("w:bCs")
        anchor = rpr.find(qn("w:rFonts"))
        if anchor is None:
            anchor = rpr.find(qn("w:rStyle"))
        if anchor is not None:
            anchor.addnext(b)
        else:
            rpr.insert(0, b)
        b.addnext(bcs)
    return rpr


def set_paragraph_text(p_el, md: str):
    """Replace the runs of p_el with md text, cloning the paragraph's own
    run formatting (font, size, colour...). Bold segments reuse the template's
    bold run formatting when present."""
    runs = [r for r in _runs(p_el) if "".join(x.text or "" for x in r.iter(qn("w:t"))).strip()]
    normal_rpr = bold_rpr = None
    for r in runs:
        rpr = r.find(qn("w:rPr"))
        if _is_bold(rpr):
            bold_rpr = bold_rpr if bold_rpr is not None else rpr
        else:
            normal_rpr = normal_rpr if normal_rpr is not None else rpr
    if normal_rpr is None:
        base = bold_rpr if bold_rpr is not None else (runs[0].find(qn("w:rPr")) if runs else None)
        if base is None:
            ppr = p_el.find(qn("w:pPr"))
            base = ppr.find(qn("w:rPr")) if ppr is not None else None
        normal_rpr = _ensure_bold(base, False)
    if bold_rpr is None:
        bold_rpr = _ensure_bold(normal_rpr, True)
    else:
        # keep the template's bold look but drop label-only colour quirks? keep as-is (strict)
        bold_rpr = copy.deepcopy(bold_rpr)

    # remove everything except paragraph properties
    for child in list(p_el):
        if child.tag != qn("w:pPr"):
            p_el.remove(child)

    for text, bold, is_tab in _segments(md):
        r = docx.oxml.OxmlElement("w:r")
        r.append(copy.deepcopy(bold_rpr if bold else normal_rpr))
        if is_tab:
            r.append(docx.oxml.OxmlElement("w:tab"))
        else:
            t = docx.oxml.OxmlElement("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = text
            r.append(t)
        p_el.append(r)


def _strip_ids(el):
    for e in el.iter():
        for a in (f"{{{W14}}}paraId", f"{{{W14}}}textId"):
            if a in e.attrib:
                del e.attrib[a]


def clean_md(text: str) -> str:
    text = text.replace("\r", "").replace("\n", " ").strip()
    text = re.sub(r"^\s*[-•●*]\s+(?!\*)", "", text)  # stray bullet glyphs
    text = re.sub(r"\s{2,}", " ", text)
    if text.count("**") % 2:
        text = text.replace("**", "")
    return text


def build_resume(template_path: str, items: list[dict]):
    """items: [{"src": id, "text": md}] for EDIT paragraphs, in order.
    Locked paragraphs are kept verbatim. An editable paragraph with no item is
    removed; extra items with the same src are inserted as clones."""
    doc, paras = parse_template(template_path)
    by_src: dict[int, list[tuple[int, str]]] = {}
    for idx, it in enumerate(items):
        try:
            src = int(it["src"])
        except (KeyError, TypeError, ValueError):
            continue
        txt = clean_md(str(it.get("text", "")))
        if txt:
            by_src.setdefault(src, []).append((idx, txt))
    item_map = {}  # paragraph element -> index into items (for comments/fixes)

    for p in paras:
        if p.locked:
            continue
        texts = by_src.get(p.id)
        if not texts:
            p.el.getparent().remove(p.el)
            continue
        template_el = copy.deepcopy(p.el)
        set_paragraph_text(p.el, texts[0][1])
        item_map[p.el] = texts[0][0]
        anchor = p.el
        for idx, extra in texts[1:]:
            clone = copy.deepcopy(template_el)
            _strip_ids(clone)
            set_paragraph_text(clone, extra)
            anchor.addnext(clone)
            item_map[clone] = idx
            anchor = clone
    doc.item_map = item_map
    return doc


def live_items(doc, items: list[dict]) -> list[dict]:
    """The items still present in doc (after any emergency trimming), in order."""
    m = getattr(doc, "item_map", {})
    return [items[m[el]] for el in doc.element.body.iterchildren(qn("w:p")) if el in m]



def doc_to_blocks(doc):
    """Ordered list of {kind, text} for the on-screen preview."""
    blocks = []
    seen_heading = False
    item_map = getattr(doc, "item_map", {})
    for p_el in doc.element.body.iterchildren(qn("w:p")):
        n_before = len(blocks)
        md = para_markdown(p_el)
        t = plain(md).strip()
        if not t:
            blocks.append({"kind": "empty", "text": ""})
        elif _is_heading(p_el, md):
            seen_heading = True
            blocks.append({"kind": "heading", "text": t})
        elif not seen_heading:
            blocks.append({"kind": "contact", "text": md.strip()})
        elif p_el.find(f".//{qn('w:numPr')}") is not None:
            blocks.append({"kind": "bullet", "text": md.strip()})
        elif "\t" in md.strip():
            left, _, right = md.strip().partition("\t")
            blocks.append({"kind": "entry", "text": left.strip(), "right": right.replace("\t", " ").strip()})
        else:
            blocks.append({"kind": "text", "text": md.strip()})
        if len(blocks) > n_before and p_el in item_map:
            blocks[-1]["item"] = item_map[p_el]
    # collapse consecutive empties
    out = []
    for b in blocks:
        if b["kind"] == "empty" and (not out or out[-1]["kind"] == "empty"):
            continue
        out.append(b)
    return out


def drop_last_bullet(doc) -> bool:
    """Emergency trim: remove the last bullet of the longest bullet run."""
    body = doc.element.body
    groups, cur = [], []
    for p_el in body.iterchildren(qn("w:p")):
        if p_el.find(f".//{qn('w:numPr')}") is not None:
            cur.append(p_el)
        else:
            if cur:
                groups.append(cur)
            cur = []
    if cur:
        groups.append(cur)
    groups = [g for g in groups if len(g) > 2]
    if not groups:
        return False
    # prefer trimming older roles (later in the doc) when equally long
    longest = max(groups, key=lambda g: (len(g), groups.index(g)))
    longest[-1].getparent().remove(longest[-1])
    return True


# ---------------------------------------------------------------- cover letter
def build_cover_letter(template_path: str, letter_md: str):
    """Reuse template: keep the contact header + first heading's look,
    remove the rest, then write the letter in the profile paragraph's style."""
    doc, paras = parse_template(template_path)
    body_style = next((p for p in paras if p.kind == "text" and not p.locked), None) \
        or next((p for p in paras if p.kind == "text"), None)
    bullet_style = next((p for p in paras if p.kind == "bullet"), None)
    blank_style = next((p for p in paras if p.kind == "empty" and p.id > 0), None)
    proto_body = copy.deepcopy(body_style.el) if body_style else None
    proto_bullet = copy.deepcopy(bullet_style.el) if bullet_style else None
    proto_blank = copy.deepcopy(blank_style.el) if blank_style else None

    # keep contact block only
    last_contact = None
    for p in paras:
        if p.kind == "contact":
            last_contact = p.el
        else:
            p.el.getparent().remove(p.el)

    body = doc.element.body
    sect = body.find(qn("w:sectPr"))

    def add(el):
        _strip_ids(el)
        if sect is not None:
            sect.addprevious(el)
        else:
            body.append(el)

    def blank():
        if proto_blank is not None:
            add(copy.deepcopy(proto_blank))
        elif proto_body is not None:
            e = copy.deepcopy(proto_body)
            set_paragraph_text(e, "")
            add(e)

    blank()
    lines = [l.rstrip() for l in letter_md.replace("\r", "").split("\n")]
    prev_blank = True
    for line in lines:
        s = line.strip()
        if not s:
            if not prev_blank:
                blank()
            prev_blank = True
            continue
        prev_blank = False
        m = re.match(r"^[-•*●]\s+(.*)$", s)
        if m and proto_bullet is not None and not s.startswith("**"):
            e = copy.deepcopy(proto_bullet)
            set_paragraph_text(e, clean_md(m.group(1)))
        else:
            e = copy.deepcopy(proto_body)
            set_paragraph_text(e, s.replace("\n", " "))
        add(e)
    _ = last_contact
    return doc
