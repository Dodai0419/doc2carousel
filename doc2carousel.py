#!/usr/bin/env python3
"""doc2carousel: turn a Word file, a slide deck, a web page or Markdown into a ready-to-post 3:4 carousel.

    python doc2carousel.py report.docx -o out          # out/slide-01.png ... + overview.png + carousel.html
    python doc2carousel.py deck.pptx --theme dark -o out
    python doc2carousel.py notes.md --md                # just print the Markdown it read

Reading follows microsoft/markitdown's docx and pptx converters, in the standard library only. The slides use the
design system of helloianneo/ian-xhs-carousel (assets/carousel.css, MIT). PNGs are screenshots from the Chrome or
Edge already installed on the machine; without one you still get carousel.html.
"""
import argparse
import base64
import html
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request
import zipfile
from html.parser import HTMLParser
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
W, A, P, R = (f"{{{NS[k]}}}" for k in ("w", "a", "p", "r"))


def md_table(rows):
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [[" ".join(c.split()) for c in r] + [""] * (width - len(r)) for r in rows]
    lines = ["| " + " | ".join(r) + " |" for r in rows]
    return "\n".join([lines[0], "| " + " | ".join(["---"] * width) + " |"] + lines[1:])


# ---------------------------------------------------------------- docx
def _rels(z, path):
    base = path.rsplit("/", 1)[0]
    rels_path = f"{base}/_rels/{path.rsplit('/', 1)[1]}.rels"
    if rels_path not in z.namelist():
        return {}
    out = {}
    for rel in ET.fromstring(z.read(rels_path)).findall("rel:Relationship", NS):
        t = rel.get("Target")
        out[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1],
                              t.lstrip("/") if t.startswith("/") else os.path.normpath(f"{base}/{t}").replace("\\", "/"))
    return out


M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
BIG_OPS = {"∑": r"\sum", "∏": r"\prod", "∐": r"\coprod", "∫": r"\int", "∬": r"\iint", "∭": r"\iiint",
           "∮": r"\oint", "⋃": r"\bigcup", "⋂": r"\bigcap"}
FUNCS = ("sin", "cos", "tan", "cot", "sec", "csc", "log", "ln", "exp", "max", "min", "lim", "det")


def omml(el):
    """Word equation (OMML) -> LaTeX; the same output as markitdown's omml.py for the common objects."""
    tag = el.tag.replace(M, "")
    kid = {c.tag.replace(M, ""): c for c in el if c.tag.startswith(M)}
    g = lambda name: omml(kid[name]) if name in kid else ""
    val = lambda path, default: e.get(M + "val") if (e := el.find(path)) is not None else default
    if tag == "r":
        return re.sub(r"(?<!\\)([{}_^#&$%~])", r"\\\1", "".join(t.text or "" for t in el.iter(M + "t")))
    if tag == "f":
        return rf"\frac{{{g('num')}}}{{{g('den')}}}"
    if tag in ("sSub", "sSup", "sSubSup", "nary"):
        scripts = (f"_{{{g('sub')}}}" if "sub" in kid else "") + (f"^{{{g('sup')}}}" if "sup" in kid else "")
        if tag == "nary":
            return BIG_OPS.get(val(f"{M}naryPr/{M}chr", "∫"), "") + scripts + g("e")
        return g("e") + scripts
    if tag == "rad":
        return rf"\sqrt[{g('deg')}]{{{g('e')}}}" if g("deg") else rf"\sqrt{{{g('e')}}}"
    if tag == "d":
        beg, end = val(f"{M}dPr/{M}begChr", "(") or ".", val(f"{M}dPr/{M}endChr", ")") or "."
        return rf"\left{beg}" + "".join(omml(e) for e in el.findall(M + "e")) + rf"\right{end}"
    if tag == "func":
        name = g("fName")
        return (f"\\{name}" if name in FUNCS else rf"\operatorname{{{name}}}") + f"({g('e')})"
    if tag.endswith("Pr") or not el.tag.startswith(M):
        return ""
    return "".join(omml(c) for c in el)  # e, num, den, sub, sup, fName, box, acc, bar, ...


def docx_to_md(path):
    z = zipfile.ZipFile(path)
    styles, ordered = {}, set()
    if "word/styles.xml" in z.namelist():
        for s in ET.fromstring(z.read("word/styles.xml")).findall("w:style", NS):
            name = s.find("w:name", NS)
            styles[s.get(W + "styleId")] = (name.get(W + "val") if name is not None else "").lower()
    if "word/numbering.xml" in z.namelist():  # numIds whose first level counts (1. 2. 3.) instead of bullets
        num = ET.fromstring(z.read("word/numbering.xml"))
        dec = {a.get(W + "abstractNumId") for a in num.findall("w:abstractNum", NS)
               if (f := a.find("w:lvl/w:numFmt", NS)) is not None and f.get(W + "val") == "decimal"}
        ordered = {n.get(W + "numId") for n in num.findall("w:num", NS) if n.find("w:abstractNumId", NS).get(W + "val") in dec}

    def para_text(p):
        parts = []
        for el in p:
            if el.tag == W + "t":
                parts.append(el.text or "")
            elif el.tag == W + "tab":
                parts.append(" ")
            elif el.tag in (W + "br", W + "cr"):
                parts.append("\n")
            elif el.tag == M + "oMathPara":
                parts.append("$$" + " ".join(omml(m) for m in el.iter(M + "oMath")) + "$$")
            elif el.tag == M + "oMath":
                parts.append("$" + omml(el) + "$")
            else:
                parts.append(para_text(el))
        return "".join(parts)

    out = []
    for el in ET.fromstring(z.read("word/document.xml")).find("w:body", NS):
        if el.tag == W + "p":
            text = para_text(el).strip()
            if not text:
                continue
            ps = el.find("w:pPr/w:pStyle", NS)
            style = styles.get(ps.get(W + "val"), ps.get(W + "val").lower()) if ps is not None else ""
            m = re.match(r"heading\s*(\d)", style)
            numpr = el.find("w:pPr/w:numPr/w:numId", NS)
            if m:
                out.append("#" * int(m.group(1)) + " " + text)
            elif numpr is not None:
                out.append(("1. " if numpr.get(W + "val") in ordered else "- ") + text)
            else:
                out.append(text)
        elif el.tag == W + "tbl":
            rows = [[" ".join(para_text(p) for p in tc.findall(".//w:p", NS)) for tc in tr.findall("w:tc", NS)]
                    for tr in el.findall("w:tr", NS)]
            out.append(md_table(rows))
    return "\n\n".join(out)


# ---------------------------------------------------------------- pptx
def _txt(el):
    """text of a txBody: paragraphs joined by newlines, line breaks as spaces (python-pptx shape.text)"""
    paras = []
    for p in el.findall(".//a:p", NS):
        paras.append("".join(t.text or "" if t.tag == A + "t" else " " for t in p.iter() if t.tag in (A + "t", A + "br")))
    return "\n".join(paras)


def _pos(shape):
    off = shape.find("./*/a:xfrm/a:off", NS)
    if off is None:
        off = shape.find("./p:xfrm/a:off", NS)
    inf = float("-inf")
    return (inf, inf) if off is None else (int(off.get("y")), int(off.get("x")))


def _chart_md(z, part):
    root = ET.fromstring(z.read(part))
    title = "".join(t.text or "" for t in root.findall("c:chart/c:title//a:t", NS))
    plot = next(e for e in root.find("c:chart/c:plotArea", NS) if e.tag.endswith("Chart"))
    series = plot.findall("c:ser", NS)

    def pts(el):
        return {int(p.get("idx")): p.find("c:v", NS).text for p in (el.findall(".//c:pt", NS) if el is not None else [])}

    cats = pts(series[0].find("c:cat", NS)) if series else {}
    names = ["".join(v.text or "" for v in s.findall("c:tx//c:v", NS)) for s in series]
    vals = [{i: float(v) for i, v in pts(s.find("c:val", NS)).items()} for s in series]
    rows = [["Category"] + names] + [[cats[i]] + [str(v.get(i)) for v in vals] for i in sorted(cats)]
    return f"### Chart{': ' + title if title else ''}\n\n" + md_table(rows)


def pptx_to_md(path):
    z = zipfile.ZipFile(path)
    pres_rels = _rels(z, "ppt/presentation.xml")
    order = [pres_rels[s.get(R + "id")][1] for s in ET.fromstring(z.read("ppt/presentation.xml")).findall("p:sldIdLst/p:sldId", NS)]
    out = []
    for num, slide in enumerate(order, 1):
        rels = _rels(z, slide)
        md = [f"<!-- Slide number: {num} -->"]

        def shape_md(sh):
            tag = sh.tag.split("}")[1]
            if tag == "pic":
                pr = sh.find("p:nvPicPr/p:cNvPr", NS)
                alt = " ".join(re.sub(r"[\r\n\[\]]", " ", pr.get("descr") or pr.get("name")).split())
                md.append(f"![{alt}]({re.sub(r'[^0-9A-Za-z_]', '', pr.get('name'))}.jpg)")
            elif tag == "graphicFrame":
                tbl = sh.find(".//a:tbl", NS)
                chart = sh.find(".//c:chart", NS)
                if tbl is not None:
                    md.append(md_table([[_txt(tc) for tc in tr.findall("a:tc", NS)] for tr in tbl.findall("a:tr", NS)]))
                elif chart is not None:
                    md.append(_chart_md(z, rels[chart.get(R + "id")][1]))
            elif tag == "sp" and sh.find("p:txBody", NS) is not None:
                text = _txt(sh.find("p:txBody", NS))
                ph = sh.find("p:nvSpPr/p:nvPr/p:ph", NS)
                if ph is not None and ph.get("idx", "0") == "0":
                    if text.strip():
                        md.append("# " + text.lstrip())
                else:
                    md.append(text)
            elif tag == "grpSp":
                for sub in sorted((e for e in sh if e.tag.split("}")[1] in ("sp", "pic", "graphicFrame", "grpSp")), key=_pos):
                    shape_md(sub)

        tree = ET.fromstring(z.read(slide)).find("p:cSld/p:spTree", NS)
        for sh in sorted((e for e in tree if e.tag.split("}")[1] in ("sp", "pic", "graphicFrame", "grpSp")), key=_pos):
            shape_md(sh)
        notes = [t for kind, t in rels.values() if kind == "notesSlide"]
        if notes:
            for sp in ET.fromstring(z.read(notes[0])).iter(P + "sp"):
                ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
                if ph is not None and ph.get("type") == "body" and _txt(sp.find("p:txBody", NS)).strip():
                    md.append("### Notes:\n" + _txt(sp.find("p:txBody", NS)))
        out.append("\n".join(md).strip())
    return "\n\n".join(out)


# ---------------------------------------------------------------- html
class _Html(HTMLParser):
    BLOCK = {"p", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "br", "section", "article", "table", "ul", "ol"}

    def __init__(self):
        super().__init__()
        self.lines, self.cur, self.skip, self.prefix = [], [], 0, ""

    def flush(self):
        text = " ".join("".join(self.cur).split())
        if text:
            self.lines.append(self.prefix + text)
        self.cur, self.prefix = [], ""

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "head", "noscript", "svg"):
            self.skip += 1
        elif tag in self.BLOCK:
            self.flush()
            self.prefix = "#" * int(tag[1]) + " " if re.fullmatch(r"h\d", tag) else "- " if tag == "li" else ""
        elif tag in ("td", "th") and self.cur:
            self.cur.append(" | ")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "head", "noscript", "svg"):
            self.skip = max(0, self.skip - 1)
        elif tag in self.BLOCK:
            self.flush()

    def handle_data(self, data):
        if not self.skip:
            self.cur.append(data)


def html_to_md(text):
    h = _Html()
    h.feed(text)
    h.close()
    h.flush()
    return "\n\n".join(h.lines)


def to_markdown(path):
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".docx":
        return docx_to_md(path)
    if ext == ".pptx":
        return pptx_to_md(path)
    text = path.read_text(encoding="utf-8", errors="replace")
    return html_to_md(text) if ext in (".html", ".htm") else text


# ---------------------------------------------------------------- slide plan
def clip(text, n):
    text = " ".join(text.split())
    if len(text) <= n:
        return text
    cut = text[: n - 1].rsplit(" ", 1)[0] if " " in text[: n - 1] else text[: n - 1]
    return cut.rstrip(",;:.") + "…"


def first_sentence(text):
    m = re.match(r"(.+?[.!?])(\s|$)", " ".join(text.split()))
    return m.group(1) if m else " ".join(text.split())


def _blocks(md):
    """Markdown -> [(kind, payload)]: heading (level, text), para, ol/ul (items), table (rows)"""
    out = []
    md = re.sub(r"^(```|~~~).*?^\1[^\n]*$", "", md, flags=re.S | re.M)  # code blocks do not fit on a slide
    md = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", md)  # inline images, badges
    md = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", md).replace("`", "")  # links keep their text
    md = re.sub(r"^>\s?(\[![A-Z]+\]\s*)?", "", md, flags=re.M)  # quotes and GitHub alerts
    for chunk in re.split(r"\n\s*\n|\n(?=#)|\n(?=<!--)", md.strip()):
        lines = [l for l in chunk.splitlines() if l.strip() and not l.startswith(("<!--", "!["))]
        if not lines:
            continue
        if m := re.match(r"(#{1,6})\s+(.*)", lines[0]):
            out.append(("heading", (len(m.group(1)), m.group(2).strip())))
            lines = lines[1:]
        i = 0
        while i < len(lines):
            line = lines[i]
            if line.lstrip().startswith("|"):
                rows = []
                while i < len(lines) and lines[i].lstrip().startswith("|"):
                    cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                    if not all(set(c) <= set("-: ") for c in cells):
                        rows.append(cells)
                    i += 1
                out.append(("table", rows))
                continue
            kind = "ol" if re.match(r"\s*\d+[.)]\s", line) else "ul" if re.match(r"\s*[-*+]\s", line) else "para"
            if kind == "para":
                out.append(("para", line.strip()))
            elif out and out[-1][0] == kind:
                out[-1][1].append(re.sub(r"^\s*(\d+[.)]|[-*+])\s+", "", line).strip())
            else:
                out.append((kind, [re.sub(r"^\s*(\d+[.)]|[-*+])\s+", "", line).strip()]))
            i += 1
    return out


def _item(text):
    m = re.search(r": | - | — ", text)
    head, note = (text[: m.start()], text[m.end():]) if m else (text, "")
    return [clip(head.replace("**", ""), 44), clip(note.replace("**", ""), 60)]


def plan(md, name, max_slides=10):
    blocks = _blocks(md)
    heads = [b for b in blocks if b[0] == "heading"]
    top = min((lvl for _, (lvl, _) in heads), default=1)
    title, intro, sections = None, [], []
    for kind, val in blocks:
        if kind == "heading" and val[0] == top and title is None and not sections and not intro and top == 1 and \
                sum(1 for k, v in heads if v[0] == 1) == 1:
            title = val[1]  # a single H1 is the document title; its H2s are the sections
            top = min((v[0] for k, v in heads if v[0] > 1), default=2)
        elif kind == "heading" and val[0] <= top:
            sections.append({"title": val[1], "blocks": []})
        elif sections:
            sections[-1]["blocks"].append((kind, val))
        else:
            intro.append((kind, val))
    paras = [v for k, v in intro if k == "para"]
    first = sections[0]["blocks"] if sections else []
    if title is None and not paras and len(sections) > 2 and all(k == "para" and len(v) < 200 for k, v in first) and len(first) <= 2:
        title, paras = sections[0]["title"], [v for _, v in first]  # a deck's title slide becomes the cover
        sections = sections[1:]
    if title is None:
        title = paras.pop(0) if paras else Path(name).stem.replace("-", " ").replace("_", " ")
    sections = [s for s in sections if s["blocks"]] or [{"title": title, "blocks": intro}]
    sections = sections[: max_slides - 2]
    tag = clip(title, 28)
    slides = [{"layout": "cover", "chip": Path(name).suffix.lstrip(".").upper() + " → carousel", "title": clip(title, 64),
               "sub": clip(first_sentence(paras[0]), 90) if paras else f"{len(sections)} parts, one swipe each",
               "items": [clip(s["title"], 40) for s in sections[:5]]}]
    for i, s in enumerate(sections, 1):
        kinds = {k: v for k, v in reversed(s["blocks"])}
        text = " ".join(v for k, v in s["blocks"] if k == "para")
        base = {"chip": f"{i}/{len(sections)} · {tag}", "title": clip(s["title"], 48),
                "sub": clip(first_sentence(text), 80) if text else ""}
        if len(kinds.get("ol", [])) >= 2:
            slides.append({**base, "layout": "steps", "items": [_item(t) for t in kinds["ol"][:5]]})
        elif len(kinds.get("ul", [])) >= 2:
            slides.append({**base, "layout": "checklist", "items": [_item(t) for t in kinds["ul"][:5]]})
        elif "table" in kinds and kinds["table"]:
            rows = [[clip(c, 26) for c in r[:4]] for r in kinds["table"][:7]]
            slides.append({**base, "layout": "table", "rows": rows})
        else:
            big = clip(first_sentence(text or s["title"]), 90)
            rest = " ".join(text.split())[len(first_sentence(text)):].strip() if text else ""
            slides.append({**base, "layout": "statement", "big": big, "note": clip(rest, 180)})
    slides.append({"layout": "end", "chip": tag, "title": "Recap", "sub": clip(title, 60),
                   "items": [clip(s["title"], 40) for s in sections[:4]]})
    return slides


# ---------------------------------------------------------------- html
ICON = {"check": '<path d="M7 16.5L13 22.5L25 10"/>', "info": '<circle cx="16" cy="16" r="11"/><path d="M16 15V22M16 10.5V11"/>',
        "close": '<path d="M8 8L24 24M24 8L8 24"/>', "code": '<path d="M11 10L5 16L11 22M21 10L27 16L21 22"/>'}


def ic(name, cls="ic"):
    return (f'<i class="{cls}"><svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="2.5" '
            f'stroke-linecap="round" stroke-linejoin="round">{ICON[name]}</svg></i>')


def _head(s, center=False, big=False):
    e = html.escape
    sub = f'<p class="sub">{e(s["sub"])}</p>' if s.get("sub") else ""
    return (f'<header class="head{" head--center" if center else ""}"><p class="chip">{e(s["chip"])}</p>'
            f'<h2 class="{"t-display" if big else "t-title"}">{e(s["title"])}</h2>{sub}</header>')


def _rows(items, start):
    e = html.escape
    return "".join(f'<div class="row">{start(i)}<div class="row__text"><p class="row__head">{e(h)}</p>'
                   + (f'<p class="row__note">{e(n)}</p>' if n else "") + "</div></div>" for i, (h, n) in enumerate(items))


def slide_html(n, s):
    e, L = html.escape, s["layout"]
    if L == "cover":
        rows = "".join(f'<div class="row"><span class="badge">{i}</span><div class="row__text"><p class="row__head">{e(t)}</p></div></div>'
                       for i, t in enumerate(s["items"], 1))
        body = (_head(s, center=True, big=True) + f'<div class="stage"><div class="card card--dialog split"><div class="bar">{ic("code")}'
                f'<p class="bar__title">Inside</p>{ic("close")}</div>{rows}'
                f'<div class="foot"><p class="foot__note">Swipe for all {len(s["items"])}</p><span class="btn btn--primary">Start</span></div></div></div>')
    elif L == "steps":
        body = _head(s) + '<div class="stage"><div class="card split">' + _rows(s["items"], lambda i: f'<span class="badge">{i + 1}</span>') + "</div></div>"
    elif L == "checklist":
        rows = _rows(s["items"], lambda i: f'<span class="badge badge--green">{ic("check")}</span>')
        body = (_head(s) + f'<div class="stage"><div class="card split"><div class="row todo__head"><p class="row__head grow">Checklist</p>'
                f'<p class="t-num c-muted">{len(s["items"])} items</p></div>{rows}</div></div>')
    elif L == "table":
        rows = "".join('<div class="setting">' + "".join(f'<p class="{"setting__key" if j == 0 else "setting__value"}">{e(c)}</p>' for j, c in enumerate(r))
                       + "</div>" for r in s["rows"][1:])
        body = (_head(s) + f'<div class="stage settings"><div class="settings__group"><p class="settings__label">{e(" · ".join(s["rows"][0]))}</p>'
                f'<div class="card split">{rows}</div></div></div>')
    elif L == "statement":
        note = f'<div class="card callout">{ic("info")}<p class="t-caption c-body">{e(s["note"])}</p></div>' if s["note"] else ""
        body = (f'<header class="head"><p class="chip">{e(s["chip"])}</p><p class="t-label c-muted">{e(s["title"])}</p>'
                f'<h2 class="t-display"><span class="c-orange">{e(s["big"])}</span></h2></header>{note}')
        return f'<section class="slide slide--center" id="s{n}">{body}</section>'
    else:  # end
        items = "".join(f'<li>{ic("check", "ic c-green")}{e(t)}</li>' for t in s["items"])
        body = (_head(s, center=True, big=True) + f'<div class="stage"><div class="card card--dialog split"><div class="done__top">'
                f'<span class="done__mark">{ic("check", "ic ic--lg")}</span><p class="t-heading">What you just read</p></div>'
                f'<ul class="done__list">{items}</ul><div class="foot foot--base"><p class="foot__note">Save it for later</p>'
                f'<span class="btn btn--primary">Share</span></div></div></div>')
    return f'<section class="slide" id="s{n}">{body}</section>'


CSS = Path(__file__).resolve().parent / "assets" / "carousel.css"
EXTRA_CSS = ".t-display,.t-title{overflow-wrap:anywhere}"  # long words wrap instead of leaving the slide


def render_html(slides, theme="light"):
    css = CSS.read_text(encoding="utf-8")
    dark = ' data-theme="dark"' if theme == "dark" else ""
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8"><title>{html.escape(slides[0]["title"])}</title>'
            f"<style>{css}{EXTRA_CSS}</style></head><body{dark}>\n"
            + "\n".join(slide_html(i, s) for i, s in enumerate(slides, 1)) + "\n</body></html>\n")


# ---------------------------------------------------------------- png
def find_chrome():
    for name in ("chrome", "google-chrome", "chromium", "chromium-browser", "msedge", "microsoft-edge"):
        if p := shutil.which(name):
            return p
    for p in (os.environ.get("CHROME", ""), r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
              "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"):
        if p and Path(p).exists():
            return p
    return None


class Chrome:
    """Just enough of the DevTools protocol (over a stdlib websocket) to load a page and screenshot parts of it."""

    def __init__(self, exe, width=1080, height=1440):
        self.prof = tempfile.mkdtemp()  # own profile: never touches the user's browser
        self.proc = subprocess.Popen([exe, "--headless=new", f"--user-data-dir={self.prof}", "--remote-debugging-port=0",
                                      "--no-first-run", "--hide-scrollbars", f"--window-size={width},{height}", "about:blank"],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        port_file, t0 = Path(self.prof, "DevToolsActivePort"), time.time()
        while not (text := self._port_text(port_file)).count("\n"):
            if time.time() - t0 > 60 or self.proc.poll() is not None:
                self.close()
                raise RuntimeError("Chrome did not start")
            time.sleep(0.05)
        port = int(text.split()[0])
        pages = json.load(urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=30))
        ws = next(p["webSocketDebuggerUrl"] for p in pages if p["type"] == "page")
        self.sock, self.id = socket.create_connection(("127.0.0.1", port), timeout=120), 0
        key = base64.b64encode(os.urandom(16)).decode()
        self.sock.sendall(f"GET {ws.split(str(port), 1)[1]} HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\nUpgrade: websocket\r\n"
                          f"Connection: Upgrade\r\nSec-WebSocket-Key: {key}\r\nSec-WebSocket-Version: 13\r\n\r\n".encode())
        head = b""
        while b"\r\n\r\n" not in head:
            head += self.sock.recv(1)
        if b" 101 " not in head.split(b"\r\n")[0]:
            raise RuntimeError("DevTools handshake failed")
        self.call("Emulation.setDeviceMetricsOverride", width=width, height=height, deviceScaleFactor=1, mobile=False)

    @staticmethod
    def _port_text(path):
        try:
            return path.read_text()
        except OSError:  # missing, or still locked by Chrome while it writes it (Windows)
            return ""

    def _read(self, n):
        buf = b""
        while len(buf) < n:
            chunk = self.sock.recv(n - len(buf))
            if not chunk:
                raise RuntimeError("DevTools connection closed")
            buf += chunk
        return buf

    def call(self, method, **params):
        self.id += 1
        data = json.dumps({"id": self.id, "method": method, "params": params}).encode()
        n, mask = len(data), os.urandom(4)
        size = bytes([0x80 | n]) if n < 126 else bytes([0x80 | 126]) + n.to_bytes(2, "big") if n < 65536 else \
            bytes([0x80 | 127]) + n.to_bytes(8, "big")
        self.sock.sendall(b"\x81" + size + mask + bytes(b ^ mask[i % 4] for i, b in enumerate(data)))
        while True:  # skip events until our answer arrives
            msg = b""
            while True:
                b0, b1 = self._read(2)
                n = b1 & 0x7F
                n = int.from_bytes(self._read(2 if n == 126 else 8), "big") if n >= 126 else n
                msg += self._read(n)
                if b0 & 0x80:
                    break
            reply = json.loads(msg)
            if reply.get("id") == self.id:
                if "error" in reply:
                    raise RuntimeError(f"{method}: {reply['error']}")
                return reply["result"]

    def js(self, expr):
        return self.call("Runtime.evaluate", expression=expr, awaitPromise=True, returnByValue=True)["result"].get("value")

    def open(self, url):
        self.call("Page.navigate", url=url)
        t0 = time.time()
        while self.js("document.readyState") != "complete" and time.time() - t0 < 60:
            time.sleep(0.05)
        self.js("document.fonts.ready.then(() => 1)")

    def shot(self, path, x, y, w, h):
        img = self.call("Page.captureScreenshot", format="png", captureBeyondViewport=True,
                        clip={"x": x, "y": y, "width": w, "height": h, "scale": 1})
        Path(path).write_bytes(base64.b64decode(img["data"]))

    def close(self):
        try:
            self.call("Browser.close")
        except Exception:
            pass
        try:
            self.proc.wait(10)
        except subprocess.TimeoutExpired:
            self.proc.kill()
        shutil.rmtree(self.prof, ignore_errors=True)  # Chrome may still hold a file for a moment


def export_pngs(page, n, out, chrome):
    """one Chrome run: each slide clipped to its own 1080x1440 PNG, then a 3-column overview of the PNGs"""
    c = Chrome(chrome)
    try:
        c.open(page.resolve().as_uri())
        c.js("document.body.style.background='none';document.querySelectorAll('.slide').forEach(s=>s.style.margin='0')")
        tops = c.js("[...document.querySelectorAll('.slide')].map(s => s.getBoundingClientRect().top + scrollY)")
        pngs = [out / f"slide-{i:02d}.png" for i in range(1, n + 1)]
        for p, top in zip(pngs, tops):
            c.shot(p, 0, top, 1080, 1440)
        cols = 3
        rows = (n + cols - 1) // cols
        grid = "".join(f'<img src="{p.name}" style="width:320px;height:427px;border-radius:10px;box-shadow:0 2px 10px #0002">' for p in pngs)
        sheet = out / "overview.html"
        sheet.write_text(f'<!doctype html><body style="margin:0;padding:30px;background:#e4e4e7;display:grid;'
                         f'grid-template-columns:repeat({cols},320px);gap:30px;width:{cols * 350 + 30}px;box-sizing:border-box">{grid}</body>',
                         encoding="utf-8")
        c.open(sheet.resolve().as_uri())
        c.shot(out / "overview.png", 0, 0, cols * 350 + 30, rows * 457 + 30)
        sheet.unlink()
    finally:
        c.close()
    return pngs


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("file", help=".docx, .pptx, .html, .md or .txt")
    ap.add_argument("-o", "--out", default="carousel-out")
    ap.add_argument("--theme", choices=["light", "dark"], default="light")
    ap.add_argument("--max", type=int, default=10, help="most slides, cover and recap included (default 10)")
    ap.add_argument("--md", action="store_true", help="print the Markdown that was read and stop")
    ap.add_argument("--no-png", action="store_true", help="write carousel.html only")
    a = ap.parse_args(argv)
    src = Path(a.file)
    if not src.is_file():
        sys.exit(f"doc2carousel: no such file: {src}")
    md = to_markdown(src)
    if a.md:
        print(md)
        return
    slides = plan(md, src.name, max(3, a.max))
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    page = out / "carousel.html"
    page.write_text(render_html(slides, a.theme), encoding="utf-8")
    print(f"doc2carousel: {src.name} -> {len(slides)} slides ({', '.join(s['layout'] for s in slides)})")
    chrome = None if a.no_png else find_chrome()
    if chrome:
        pngs = export_pngs(page, len(slides), out, chrome)
        print(f"  wrote {len(pngs)} PNGs (1080x1440) to {out}, plus overview.png and carousel.html")
    else:
        print(f"  wrote {page}" + ("" if a.no_png else " (no Chrome or Edge found, so no PNGs; set CHROME=path/to/browser)"))


if __name__ == "__main__":
    main()
