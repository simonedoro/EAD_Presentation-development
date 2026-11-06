#!/usr/bin/env python3
"""
build_deck.py - render an approved storyline into a client PowerPoint template.

    python3 build_deck.py storyline.yaml --profile profiles/ead.json --out deck.pptx

The storyline is the source of truth. This script never invents content and never
invents geometry: every slide is one of a closed set of archetypes with coordinates
fixed by the profile. If a slide does not fit an archetype, it is built by hand in
PowerPoint, not by loosening this script.

Validation is fail-loud. Action-title and source-line rules are hard errors.
"""

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

try:
    import yaml
except ImportError:
    yaml = None

EMU_IN = 914400
ARCHETYPES = (
    "cover", "section", "agenda", "statement", "kpi", "bullets", "columns",
    "cards", "steps", "initiative", "visual", "exhibit", "matrix", "table",
    "chart",
)

# The innovation bootcamp pitch deck. Each entry is (role, allowed archetypes,
# whether the slide is required). Dot N in the participant's storyline becomes
# slide N, so the order is the contract.
BOOTCAMP_SKELETON = [
    ("cover",        ("cover",),                                   True),
    ("agenda",       ("agenda", "bullets", "statement"),           True),
    ("situation",    ("statement", "kpi", "bullets", "chart", "columns",
                     "visual", "exhibit"),                                   True),
    ("complication", ("statement", "kpi", "bullets", "chart", "columns",
                     "visual", "exhibit"),                                   True),
    ("problem",      ("statement",),                               True),
    ("solutions",    ("columns",),                                 True),
    ("initiative",   ("initiative", "columns", "steps", "chart",
                     "visual", "exhibit"),                        True),
    ("recap",        ("table", "cards", "columns", "kpi", "exhibit"), True),
    ("ask",          ("cards", "steps", "bullets", "table",
                     "exhibit"),                                  True),
]

# Titles that announce a topic instead of making a point.
BANNED_TITLE_STARTS = re.compile(
    r"^(overview|introduction|intro|background|context|summary|recap|"
    r"our approach|approach|methodology|key takeaways|takeaways|objectives|"
    r"about|current state|as-?is|to-?be)\b(?!\s*:)",
    re.I,
)


# ---------------------------------------------------------------- utilities

def die(msg):
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(1)


def load_spec(path):
    with open(path, "r", encoding="utf-8") as fh:
        text = fh.read()
    if path.lower().endswith((".yaml", ".yml")):
        if yaml is None:
            die("pyyaml is not installed; use a .json storyline or pip install pyyaml")
        return yaml.safe_load(text)
    return json.loads(text)


def rgb(hexstr):
    return RGBColor.from_string(hexstr.upper().lstrip("#"))


_IMG_CACHE = {}
_SPEC_DIR = None  # relative image paths resolve against the storyline file


def resolve_image(value):
    """Accept a local path or an http(s) URL. Return a local path, or None.

    A URL is downloaded once per run into a temp file. If the network is not
    reachable the slide falls back to a marked placeholder rather than failing
    the build, because a participant may be offline.
    """
    if not value:
        return None
    value = str(value).strip()
    if value in _IMG_CACHE:
        return _IMG_CACHE[value]
    if value.lower().startswith(("http://", "https://")):
        try:
            import urllib.request
            ext = os.path.splitext(value.split("?")[0])[1].lower()
            if ext not in (".png", ".jpg", ".jpeg", ".gif", ".webp"):
                ext = ".png"
            tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
            req = urllib.request.Request(
                value, headers={"User-Agent": "consulting-deck/1.0"})
            with urllib.request.urlopen(req, timeout=20) as resp:
                tmp.write(resp.read())
            tmp.close()
            _IMG_CACHE[value] = tmp.name
            return tmp.name
        except Exception as exc:
            print(f"WARN  could not fetch image {value}: {exc}")
            _IMG_CACHE[value] = None
            return None
    candidates = [value]
    if not os.path.isabs(value) and _SPEC_DIR:
        candidates.insert(0, os.path.join(_SPEC_DIR, value))
    path = next((c for c in candidates if os.path.exists(c)), None)
    if path is None:
        print(f"WARN  image not found: {value}")
    _IMG_CACHE[value] = path
    return path


def place_image(b, slide, value, x, y, w, h, caption="image"):
    """Fit an image inside the box, centred, or draw a marked placeholder."""
    path = resolve_image(value)
    if path:
        try:
            from PIL import Image
            with Image.open(path) as im:
                iw, ih = im.size
            ratio = iw / ih
            box_ratio = w / h
            if ratio > box_ratio:
                dw, dh = w, w / ratio
            else:
                dh, dw = h, h * ratio
            slide.shapes.add_picture(path, Inches(x + (w - dw) / 2),
                                     Inches(y + (h - dh) / 2),
                                     width=Inches(dw), height=Inches(dh))
            return True
        except Exception as exc:
            print(f"WARN  could not place image: {exc}")
    b.rect(slide, x, y, w, h, fill=b.c["white"], line=b.c["rule"])
    b.text(slide, x, y, w, h, caption, size=b.sz("label", 11),
           color=b.c["muted"], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return False


def open_template(path):
    """python-pptx refuses .potx on content type. Patch it in a temp copy."""
    if not os.path.exists(path):
        die(f"template not found: {path}")
    if path.lower().endswith(".potx"):
        tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
        tmp.close()
        zin = zipfile.ZipFile(path)
        zout = zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED)
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"presentationml.template.main+xml",
                    b"presentationml.presentation.main+xml",
                )
            zout.writestr(item, data)
        zout.close()
        zin.close()
        return Presentation(tmp.name), tmp.name
    return Presentation(path), None


def find_layout(prs, name):
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name.strip().lower() == name.strip().lower():
                return layout
    available = [l.name for m in prs.slide_masters for l in m.slide_layouts]
    die(f"layout '{name}' not in template. Available: {available}")


def est_lines(text, width_in, size_pt, bold=False):
    """Rough line count for Arial. Deliberately pessimistic."""
    if not text:
        return 0
    factor = 0.55 if bold else 0.51
    chars_per_line = max(1, int((width_in * 72) / (size_pt * factor)))
    lines = 0
    for para in str(text).split("\n"):
        lines += max(1, -(-len(para) // chars_per_line))
    return lines


def est_height(text, width_in, size_pt, bold=False, leading=1.22):
    return est_lines(text, width_in, size_pt, bold) * size_pt * leading / 72.0


# ---------------------------------------------------------------- primitives

class Builder:
    def __init__(self, prs, profile):
        self.prs = prs
        self.p = profile
        self.c = profile["colors"]
        self.t = profile.get("type", {})
        self.font = profile.get("font", "Arial")
        self.warnings = []

    def sz(self, key, default):
        """Type sizes live in the profile so a client can be tuned without code."""
        return self.t.get(key, default)

    # -- shape helpers ----------------------------------------------------

    def rect(self, slide, x, y, w, h, fill=None, line=None, line_w=0.75,
             shape=MSO_SHAPE.RECTANGLE, radius=None):
        sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            try:
                sh.adjustments[0] = radius
            except Exception:
                pass
        if fill:
            sh.fill.solid()
            sh.fill.fore_color.rgb = rgb(fill)
        else:
            sh.fill.background()
        if line:
            sh.line.color.rgb = rgb(line)
            sh.line.width = Pt(line_w)
        else:
            sh.line.fill.background()
        sh.shadow.inherit = False
        sh.text_frame.word_wrap = True
        return sh

    def text(self, slide, x, y, w, h, runs, size=11, color=None, bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space_after=3,
             line_spacing=1.12):
        """runs: a string, or a list of (text, {overrides}) for mixed paragraphs."""
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0

        items = runs if isinstance(runs, list) else [(runs, {})]
        for i, item in enumerate(items):
            txt, over = item if isinstance(item, (tuple, list)) else (item, {})
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = over.get("align", align)
            para.space_after = Pt(over.get("space_after", space_after))
            para.line_spacing = over.get("line_spacing", line_spacing)
            if over.get("bullet"):
                para.level = 0
            run = para.add_run()
            run.text = str(txt)
            run.font.name = self.font
            run.font.size = Pt(over.get("size", size))
            run.font.bold = over.get("bold", bold)
            run.font.color.rgb = rgb(over.get("color", color or self.c["ink"]))
        return box

    def rich_line(self, slide, x, y, w, h, parts, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE):
        """One paragraph, several runs. parts: [(text, size, color, bold)]."""
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        para = tf.paragraphs[0]
        para.alignment = align
        for txt, size, color, bold in parts:
            run = para.add_run()
            run.text = str(txt)
            run.font.name = self.font
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = rgb(color)
        return box

    def bullet_rows(self, slide, x, y, w, items, size=11, color=None,
                    marker=None, row_gap=0.10, marker_size=0.085):
        """Square-marker bullets. Returns bottom y."""
        color = color or self.c["ink"]
        marker = marker or self.c["primary"]
        cy = y
        for item in items:
            txt = item if isinstance(item, str) else item.get("text", "")
            h = max(0.20, est_height(txt, w - 0.26, size))
            self.rect(slide, x + 0.01, cy + size / 72.0 * 0.38, marker_size,
                      marker_size, fill=marker)
            self.text(slide, x + 0.26, cy, w - 0.26, h, txt, size=size, color=color)
            cy += h + row_gap
        return cy

    # -- chrome -----------------------------------------------------------

    def new_slide(self, layout_key):
        layout = find_layout(self.prs, self.p["layouts"][layout_key])
        slide = self.prs.slides.add_slide(layout)
        return slide

    def strip_placeholders(self, slide, keep_title=True):
        for ph in list(slide.placeholders):
            is_title = ph.placeholder_format.idx == 0
            if is_title and keep_title:
                continue
            ph._element.getparent().remove(ph._element)

    def set_title(self, slide, text):
        t = self.p["title"]
        ph = None
        for cand in slide.placeholders:
            if cand.placeholder_format.idx == 0:
                ph = cand
                break
        if ph is None:
            ph = slide.shapes.add_textbox(
                Inches(t["x"]), Inches(t["y"]), Inches(t["w"]), Inches(t["h"]))
        tf = ph.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.paragraphs[0].text = ""
        run = tf.paragraphs[0].add_run()
        run.text = text
        run.font.name = self.font
        run.font.size = Pt(t["size"])
        run.font.bold = t.get("bold", True)
        run.font.color.rgb = rgb(t["color"])
        tf.paragraphs[0].line_spacing = 1.05
        return ph

    def add_source(self, slide, source):
        s = self.p["source"]
        txt = str(source)
        if not txt.lower().startswith("source"):
            txt = s["prefix"] + txt
        self.text(slide, s["x"], s["y"], s["w"], s["h"], txt,
                  size=s["size"], color=s["color"])

    def content_region(self, has_takeaway=False):
        r = dict(self.p["content"])
        if has_takeaway:
            r["h"] = r["h"] - 0.88
        return r

    def add_takeaway(self, slide, text):
        r = self.p["content"]
        y = r["y"] + r["h"] - 0.68
        self.rect(slide, r["x"], y, r["w"], 0.68, fill=self.c["panel"])
        self.rect(slide, r["x"], y, 0.055, 0.68, fill=self.c["primary"])
        self.text(slide, r["x"] + 0.26, y, r["w"] - 0.52, 0.68, text,
                  size=self.sz("takeaway", 13), bold=True,
                  color=self.c["primary"], anchor=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------- archetypes

def _title_slide(b, slide, spec, cfg):
    b.strip_placeholders(slide, keep_title=True)
    ph = b.set_title(slide, spec["title"])
    run = ph.text_frame.paragraphs[0].runs[0]
    run.font.size = Pt(cfg["title"]["size"])
    run.font.color.rgb = rgb(cfg["title"].get("color", b.c["ink"]))
    if spec.get("subtitle") and cfg.get("subtitle"):
        s = cfg["subtitle"]
        b.text(slide, s["x"], s["y"], s["w"], s["h"], spec["subtitle"],
               size=s["size"], color=s.get("color", b.c["subink"]))
    if spec.get("date") and cfg.get("date"):
        d = cfg["date"]
        b.text(slide, d["x"], d["y"], d["w"], d["h"], spec["date"],
               size=d["size"], color=d.get("color", b.c["muted"]))


def a_cover(b, slide, spec):
    """Cover carries project title, sector and challenge category."""
    cfg = b.p["cover"]
    sub = spec.get("subtitle")
    bits = [x for x in (spec.get("sector"), spec.get("category")) if x]
    if bits:
        sub = " | ".join(bits) if not sub else sub + "\n" + " | ".join(bits)
    spec = dict(spec, subtitle=sub)
    _title_slide(b, slide, spec, cfg)
    if spec.get("team"):
        d = cfg["date"]
        b.text(slide, d["x"], d["y"] + 0.46, d["w"], d["h"], spec["team"],
               size=d["size"], color=d.get("color", b.c["muted"]))


def a_section(b, slide, spec):
    _title_slide(b, slide, spec, b.p.get("section", b.p["cover"]))


def a_agenda(b, slide, spec):
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    items = spec["items"]
    current = spec.get("current")
    n = len(items)
    row_h = min(0.86, (r["h"] - (n - 1) * 0.14) / max(n, 1))
    cy = r["y"]
    for i, item in enumerate(items, start=1):
        label = item if isinstance(item, str) else item.get("text", "")
        live = (current == i)
        b.rect(slide, r["x"], cy, r["w"], row_h,
               fill=b.c["panel"] if live else b.c["white"],
               line=None if live else b.c["rule"])
        if live:
            b.rect(slide, r["x"], cy, 0.055, row_h, fill=b.c["primary"])
        b.text(slide, r["x"] + 0.30, cy, 0.5, row_h, f"{i}",
               size=b.sz("agenda_item", 14) + 1, bold=True,
               anchor=MSO_ANCHOR.MIDDLE,
               color=b.c["primary"] if live else b.c["muted"])
        b.text(slide, r["x"] + 0.90, cy, r["w"] - 1.20, row_h, label,
               size=b.sz("agenda_item", 14), bold=live,
               anchor=MSO_ANCHOR.MIDDLE,
               color=b.c["ink"] if live else b.c["subink"])
        cy += row_h + 0.14


def a_statement(b, slide, spec):
    """Situation, complication, problem statement. One claim, a little proof."""
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    size = b.sz("statement", 20)
    sub = b.sz("statement_sub", 14)
    body = spec.get("statement", "")
    points = spec.get("points") or []
    panel_w = r["w"] if not points else r["w"] * 0.56

    bh = max(0.6, est_height(body, panel_w - 0.72, size)) + 0.80
    bh = min(bh, r["h"])

    pw = r["w"] - panel_w - 0.40
    ph = 0.0
    if points:
        if spec.get("points_heading"):
            ph += 0.42
        for item in points:
            txt = item if isinstance(item, str) else item.get("text", "")
            ph += max(0.22, est_height(txt, pw - 0.26, sub)) + 0.18

    block = max(bh, ph)
    top = r["y"] + max(0.0, (r["h"] - block) / 2)

    b.rect(slide, r["x"], top, panel_w, bh, fill=b.c["panel"])
    b.rect(slide, r["x"], top, 0.07, bh, fill=b.c["primary"])
    b.text(slide, r["x"] + 0.40, top, panel_w - 0.72, bh,
           body, size=size, bold=True, color=b.c["ink"], line_spacing=1.18,
           anchor=MSO_ANCHOR.MIDDLE)

    if points:
        px = r["x"] + panel_w + 0.40
        cy = top + max(0.0, (bh - ph) / 2)
        if spec.get("points_heading"):
            b.text(slide, px, cy, pw, 0.30, spec["points_heading"],
                   size=b.sz("label", 11), bold=True, color=b.c["primary"])
            cy += 0.42
        b.bullet_rows(slide, px, cy, pw, points, size=sub, row_gap=0.18)


def a_kpi(b, slide, spec):
    """Stat tiles. Each may carry a target, rendered as now to target.

    This is the from-to slide: what the number is today and what it needs to be.
    """
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    stats = spec["stats"]
    n = len(stats)
    gap = 0.30
    tw = (r["w"] - gap * (n - 1)) / n
    big = b.sz("kpi_value", 34 if n <= 3 else 28)
    lab = b.sz("col_heading", 14)
    note = b.sz("col_lead", 12.5)
    needed = 0.0
    for st in stats:
        h = 0.30 + big / 72.0 * 1.5 + 0.08
        h += max(0.30, est_height(st.get("label", ""), tw - 0.44, lab, bold=True))
        if st.get("note"):
            h += 0.08 + max(0.24, est_height(st["note"], tw - 0.44, note))
        needed = max(needed, h + 0.34)
    th = min(r["h"], max(1.8, needed))
    top = r["y"] + min(0.45, max(0.0, (r["h"] - th) / 2))

    for i, st in enumerate(stats):
        x = r["x"] + i * (tw + gap)
        hi = bool(st.get("highlight"))
        b.rect(slide, x, top, tw, th,
               fill=b.c["panel"] if hi else b.c["white"],
               line=None if hi else b.c["rule"])
        b.rect(slide, x, top, tw, 0.055, fill=b.c["primary"])

        parts = [(st["value"], big, b.c["primary"], True)]
        if st.get("target"):
            parts.append(("  \u2192  ", big * 0.6, b.c["muted"], False))
            parts.append((st["target"], big, b.c["secondary"], True))
        b.rich_line(slide, x + 0.16, top + 0.30, tw - 0.32, big / 72.0 * 1.5, parts)

        cy = top + 0.30 + big / 72.0 * 1.5 + 0.08
        lh = max(0.30, est_height(st.get("label", ""), tw - 0.44, lab, bold=True))
        b.text(slide, x + 0.22, cy, tw - 0.44, lh, st.get("label", ""),
               size=lab, bold=True, color=b.c["ink"], align=PP_ALIGN.CENTER)
        if st.get("note"):
            b.text(slide, x + 0.22, cy + lh + 0.08, tw - 0.44,
                   max(0.24, est_height(st["note"], tw - 0.44, note)),
                   st["note"], size=note, color=b.c["subink"],
                   align=PP_ALIGN.CENTER)


def a_bullets(b, slide, spec):
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    pts = spec["points"]
    n = len(pts)
    size = b.sz("bullet", 15)
    width = min(r["w"], 10.4)
    row_h = min(0.98, r["h"] / n)
    top = r["y"] + max(0.0, (r["h"] - row_h * n) / 2)
    for i, item in enumerate(pts):
        txt = item if isinstance(item, str) else item.get("text", "")
        y = top + i * row_h
        b.rect(slide, r["x"] + 0.02, y + row_h / 2 - 0.05, 0.10, 0.10,
               fill=b.c["primary"])
        b.text(slide, r["x"] + 0.36, y, width - 0.36, row_h, txt,
               size=size, color=b.c["ink"], anchor=MSO_ANCHOR.MIDDLE)
        if i < n - 1:
            b.rect(slide, r["x"], y + row_h - 0.005, width, 0.009,
                   fill=b.c["rule"])


def a_columns(b, slide, spec):
    """2 or 3 panels. Each may carry an image above its text."""
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    cols = spec["columns"]
    n = len(cols)
    gap = 0.30
    cw = (r["w"] - gap * (n - 1)) / n
    pad = 0.26
    h_sz = b.sz("col_heading", 14)
    l_sz = b.sz("col_lead", 12.5)
    b_sz = b.sz("col_bullet", 12.5)
    img_h = 1.55 if any(c.get("image") for c in cols) else 0.0

    needed = 0.0
    for col in cols:
        h = pad * 2 + 0.12 + img_h
        if col.get("heading"):
            h += max(0.32, est_height(col["heading"], cw - 2 * pad, h_sz, bold=True)) + 0.16
        if col.get("lead"):
            h += max(0.26, est_height(col["lead"], cw - 2 * pad, l_sz)) + 0.14
        for bl in (col.get("bullets") or []):
            txt = bl if isinstance(bl, str) else bl.get("text", "")
            h += max(0.22, est_height(txt, cw - 2 * pad - 0.26, b_sz)) + 0.12
        needed = max(needed, h)
    panel_h = min(r["h"], max(2.0, needed))

    for i, col in enumerate(cols):
        x = r["x"] + i * (cw + gap)
        b.rect(slide, x, r["y"], cw, panel_h, fill=b.c["panel"])
        b.rect(slide, x, r["y"], cw, 0.055, fill=b.c["primary"])
        cy = r["y"] + pad + 0.06
        if img_h:
            place_image(b, slide, col.get("image"), x + pad, cy,
                        cw - 2 * pad, img_h,
                        col.get("image_note") or "image")
            cy += img_h + 0.20
        head = col.get("heading", "")
        if head:
            hh = max(0.32, est_height(head, cw - 2 * pad, h_sz, bold=True))
            b.text(slide, x + pad, cy, cw - 2 * pad, hh, head,
                   size=h_sz, bold=True, color=b.c["primary"])
            cy += hh + 0.16
        if col.get("lead"):
            lh = max(0.26, est_height(col["lead"], cw - 2 * pad, l_sz))
            b.text(slide, x + pad, cy, cw - 2 * pad, lh, col["lead"],
                   size=l_sz, color=b.c["subink"])
            cy += lh + 0.14
        if col.get("bullets"):
            b.bullet_rows(slide, x + pad, cy, cw - 2 * pad, col["bullets"],
                          size=b_sz, row_gap=0.12)


def a_cards(b, slide, spec):
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    cards = spec["cards"]
    n = len(cards)
    gap = 0.18
    ch = (r["h"] - gap * (n - 1)) / n
    h_sz = b.sz("card_heading", 13.5)
    t_sz = b.sz("card_text", 12.5)
    for i, card in enumerate(cards):
        y = r["y"] + i * (ch + gap)
        b.rect(slide, r["x"], y, r["w"], ch, fill=b.c["panel"])
        b.rect(slide, r["x"] + 0.26, y + (ch - 0.46) / 2, 0.46, 0.46,
               fill=b.c["primary"], shape=MSO_SHAPE.OVAL)
        b.text(slide, r["x"] + 0.26, y + (ch - 0.46) / 2, 0.46, 0.46,
               str(card.get("number", i + 1)), size=h_sz, bold=True,
               color=b.c["white"], align=PP_ALIGN.CENTER,
               anchor=MSO_ANCHOR.MIDDLE)
        tx = r["x"] + 0.94
        tw = r["w"] - 1.20
        head = card.get("heading", "")
        body = card.get("text", "")
        if head and body:
            hh = max(0.28, est_height(head, tw, h_sz, bold=True))
            inner = hh + 0.06 + max(0.24, est_height(body, tw, t_sz))
            top = y + max(0.14, (ch - inner) / 2)
            b.text(slide, tx, top, tw, hh, head, size=h_sz, bold=True,
                   color=b.c["ink"])
            b.text(slide, tx, top + hh + 0.06, tw,
                   max(0.24, est_height(body, tw, t_sz)), body,
                   size=t_sz, color=b.c["subink"])
        else:
            b.text(slide, tx, y, tw, ch, head or body, size=h_sz,
                   bold=bool(head), anchor=MSO_ANCHOR.MIDDLE, color=b.c["ink"])


def a_steps(b, slide, spec):
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    steps = spec["steps"]
    n = len(steps)
    arrow = 0.34
    sw = (r["w"] - arrow * (n - 1)) / n
    pad = 0.22
    h_sz = b.sz("step_heading", 12.5)
    t_sz = b.sz("step_text", 11.5)
    needed = 0.0
    for step in steps:
        h = pad * 2 + 0.40
        h += max(0.30, est_height(step.get("heading", ""), sw - 2 * pad, h_sz, bold=True))
        if step.get("text"):
            h += 0.08 + max(0.24, est_height(step["text"], sw - 2 * pad, t_sz))
        needed = max(needed, h)
    sh_h = min(r["h"], max(1.6, needed))
    y = r["y"] + min(0.5, max(0.0, (r["h"] - sh_h) / 2))
    for i, step in enumerate(steps):
        x = r["x"] + i * (sw + arrow)
        b.rect(slide, x, y, sw, sh_h, fill=b.c["panel"])
        b.rect(slide, x, y, sw, 0.055, fill=b.c["primary"])
        b.text(slide, x + pad, y + pad, sw - 2 * pad, 0.28,
               f"{i + 1}", size=b.sz("label", 11), bold=True, color=b.c["primary"])
        head = step.get("heading", "")
        body = step.get("text", "")
        hh = max(0.30, est_height(head, sw - 2 * pad, h_sz, bold=True))
        b.text(slide, x + pad, y + pad + 0.32, sw - 2 * pad, hh, head,
               size=h_sz, bold=True, color=b.c["ink"])
        if body:
            b.text(slide, x + pad, y + pad + 0.32 + hh + 0.08,
                   sw - 2 * pad, sh_h - (pad * 2 + 0.40 + hh), body,
                   size=t_sz, color=b.c["subink"])
        if i < n - 1:
            b.rect(slide, x + sw + 0.07, y + sh_h / 2 - 0.11, 0.20, 0.22,
                   fill=b.c["muted"], shape=MSO_SHAPE.ISOSCELES_TRIANGLE).rotation = 90


def a_initiative(b, slide, spec):
    """One initiative in detail: what it is, how it works, what it needs, impact.

    Left panel holds the description and an optional image. Right rail holds the
    four answers a jury will ask for.
    """
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    h_sz = b.sz("col_heading", 14)
    t_sz = b.sz("col_lead", 12.5)
    lab = b.sz("label", 11)

    lw = r["w"] * 0.46
    b.rect(slide, r["x"], r["y"], lw, r["h"], fill=b.c["panel"])
    b.rect(slide, r["x"], r["y"], lw, 0.055, fill=b.c["primary"])
    pad = 0.28
    cy = r["y"] + pad + 0.06
    if spec.get("name"):
        hh = max(0.32, est_height(spec["name"], lw - 2 * pad, h_sz, bold=True))
        b.text(slide, r["x"] + pad, cy, lw - 2 * pad, hh, spec["name"],
               size=h_sz, bold=True, color=b.c["primary"])
        cy += hh + 0.16
    desc = spec.get("what", "")
    if desc:
        dh = max(0.30, est_height(desc, lw - 2 * pad, t_sz))
        b.text(slide, r["x"] + pad, cy, lw - 2 * pad, dh, desc,
               size=t_sz, color=b.c["ink"])
        cy += dh + 0.22
    img_bottom = r["y"] + r["h"] - pad
    if cy + 0.9 < img_bottom:
        place_image(b, slide, spec.get("image"), r["x"] + pad, cy,
                    lw - 2 * pad, img_bottom - cy,
                    spec.get("image_note") or "image, diagram or mock-up")

    rx = r["x"] + lw + 0.40
    rw = r["w"] - lw - 0.40
    fields = [("How it works", spec.get("how")),
              ("What it needs", spec.get("needs")),
              ("Impact", spec.get("impact")),
              ("Risk and mitigation", spec.get("risk"))]
    fields = [(k, v) for k, v in fields if v]
    gap = 0.16
    bh = (r["h"] - gap * (len(fields) - 1)) / max(len(fields), 1)
    for i, (label, val) in enumerate(fields):
        y = r["y"] + i * (bh + gap)
        b.rect(slide, rx, y, rw, bh, fill=b.c["white"], line=b.c["rule"])
        b.rect(slide, rx, y, 0.055, bh, fill=b.c["primary"])
        b.text(slide, rx + 0.26, y + 0.14, rw - 0.52, 0.26, label.upper(),
               size=lab - 1, bold=True, color=b.c["primary"])
        b.text(slide, rx + 0.26, y + 0.44, rw - 0.52, bh - 0.58, val,
               size=t_sz, color=b.c["ink"])


def a_visual(b, slide, spec):
    """A picture carries the point. Image left, the read on the right."""
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    points = spec.get("points") or []
    iw = r["w"] * 0.60 if points else r["w"]
    cap = spec.get("caption")
    ih = r["h"] - (0.30 if cap else 0.0)
    place_image(b, slide, spec.get("image"), r["x"], r["y"], iw, ih,
                cap or "image")
    if cap:
        b.text(slide, r["x"], r["y"] + ih + 0.06, iw, 0.24, cap,
               size=b.sz("label", 11) - 1, color=b.c["muted"])
    if points:
        px = r["x"] + iw + 0.40
        pw = r["w"] - iw - 0.40
        cy = r["y"] + 0.10
        if spec.get("points_heading"):
            b.text(slide, px, cy, pw, 0.30, spec["points_heading"],
                   size=b.sz("col_heading", 14), bold=True, color=b.c["primary"])
            cy += 0.44
        b.bullet_rows(slide, px, cy, pw, points,
                      size=b.sz("col_lead", 12.5), row_gap=0.18)


def _draw_chart(b, slide, cfg, x, y, w, h):
    """Native PowerPoint chart, consulting defaults. Shared by chart and exhibit."""
    kinds = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "stacked": XL_CHART_TYPE.COLUMN_STACKED,
    }
    kind = kinds.get(cfg.get("chart_type", "bar"))
    if kind is None:
        die(f"unsupported chart_type '{cfg.get('chart_type')}'")
    data = CategoryChartData()
    data.categories = cfg["categories"]
    for series in cfg["series"]:
        data.add_series(series["name"], series["values"])
    top = y + (0.26 if cfg.get("unit") else 0.0)
    gf = slide.shapes.add_chart(kind, Inches(x), Inches(top), Inches(w),
                                Inches(h - (0.26 if cfg.get("unit") else 0.0)), data)
    chart = gf.chart
    chart.font.name = b.font
    chart.font.size = Pt(b.sz("table_body", 11.5))
    chart.has_title = False
    labels_on = cfg.get("data_labels", True)
    try:
        chart.value_axis.has_major_gridlines = False
        chart.value_axis.visible = not labels_on
        chart.value_axis.format.line.color.rgb = rgb(b.c["rule"])
        chart.category_axis.format.line.color.rgb = rgb(b.c["rule"])
        chart.category_axis.has_major_gridlines = False
        chart.plots[0].gap_width = 80
    except Exception:
        pass
    if len(cfg["series"]) > 1:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(b.sz("label", 11))
    else:
        chart.has_legend = False
    palette = b.p.get("chart_palette",
                      [b.c["primary"], b.c["secondary"], b.c["accent"]])
    if len(cfg["series"]) > len(palette):
        die(f"{len(cfg['series'])} series but the palette has {len(palette)} slots")
    for i, ser in enumerate(chart.series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(palette[i])
        ser.format.line.fill.background()
        if labels_on:
            try:
                lb = ser.data_labels
                lb.show_value = True
                lb.number_format = cfg.get("number_format", "General")
                lb.number_format_is_linked = False
                lb.font.size = Pt(b.sz("label", 11) - 1)
                lb.font.name = b.font
                lb.font.bold = True
                lb.font.color.rgb = rgb(b.c["ink"])
            except Exception:
                pass
    if cfg.get("unit"):
        b.text(slide, x, y, w, 0.24, cfg["unit"],
               size=b.sz("label", 11) - 1, color=b.c["muted"])


def _draw_table(b, slide, cfg, x, y, w, h):
    header = cfg["header"]
    rows = cfg["rows"]
    head_h = 0.48
    body_h = min(0.62, max(0.34, (h - head_h) / max(len(rows), 1)))
    shape = slide.shapes.add_table(len(rows) + 1, len(header), Inches(x), Inches(y),
                                   Inches(w), Inches(head_h + body_h * len(rows)))
    tbl = shape.table
    tbl.rows[0].height = Inches(head_h)
    for i in range(1, len(rows) + 1):
        tbl.rows[i].height = Inches(body_h)
    if cfg.get("col_widths"):
        tot = sum(cfg["col_widths"])
        for i, wt in enumerate(cfg["col_widths"]):
            tbl.columns[i].width = Inches(w * wt / tot)
    for j, label in enumerate(header):
        cell = tbl.cell(0, j)
        cell.text = str(label)
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(b.c["primary"])
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        pr = cell.text_frame.paragraphs[0]
        pr.font.name = b.font
        pr.font.size = Pt(b.sz("table_header", 12))
        pr.font.bold = True
        pr.font.color.rgb = rgb(b.c["white"])
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = "" if val is None else str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(b.c["white"] if i % 2 else b.c["panel"])
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            pr = cell.text_frame.paragraphs[0]
            pr.font.name = b.font
            pr.font.size = Pt(b.sz("table_body", 11.5))
            pr.font.color.rgb = rgb(b.c["ink"])


def _draw_stats(b, slide, stats, x, y, w, h):
    """Stat tiles stacked vertically, for the evidence half of an exhibit."""
    n = len(stats)
    gap = 0.18
    th = (h - gap * (n - 1)) / n
    big = b.sz("kpi_value", 34) * (0.8 if n > 2 else 1.0)
    for i, st in enumerate(stats):
        ty = y + i * (th + gap)
        hi = bool(st.get("highlight"))
        b.rect(slide, x, ty, w, th,
               fill=b.c["panel"] if hi else b.c["white"],
               line=None if hi else b.c["rule"])
        b.rect(slide, x, ty, 0.055, th, fill=b.c["primary"])
        parts = [(st["value"], big, b.c["primary"], True)]
        if st.get("target"):
            parts.append(("  \u2192  ", big * 0.6, b.c["muted"], False))
            parts.append((st["target"], big, b.c["secondary"], True))
        b.rich_line(slide, x + 0.30, ty, w * 0.42, th, parts, align=PP_ALIGN.LEFT)
        lx = x + 0.30 + w * 0.42 + 0.20
        lw = w - (lx - x) - 0.26
        lab = b.sz("col_heading", 14)
        note = b.sz("col_lead", 12.5)
        lh = max(0.28, est_height(st.get("label", ""), lw, lab, bold=True))
        nh = max(0.24, est_height(st.get("note", ""), lw, note)) if st.get("note") else 0
        inner = lh + (0.06 + nh if nh else 0)
        ty2 = ty + max(0.10, (th - inner) / 2)
        b.text(slide, lx, ty2, lw, lh, st.get("label", ""), size=lab, bold=True,
               color=b.c["ink"])
        if nh:
            b.text(slide, lx, ty2 + lh + 0.06, lw, nh, st["note"], size=note,
                   color=b.c["subink"])


def _draw_insights(b, slide, spec, x, y, w, h):
    """The right rail: what the evidence means."""
    b.rect(slide, x, y, w, h, fill=b.c["panel"])
    b.rect(slide, x, y, w, 0.055, fill=b.c["primary"])
    pad = 0.26
    lab = b.sz("col_heading", 14)
    body = b.sz("col_lead", 12.5)
    cy = y + pad + 0.06
    heading = spec.get("insight_heading", "What this means")
    b.text(slide, x + pad, cy, w - 2 * pad, 0.30, heading,
           size=lab, bold=True, color=b.c["primary"])
    cy += 0.44

    items = spec.get("insights") or []
    numbered = spec.get("numbered", True)
    items = [{"text": i} if isinstance(i, str) else i for i in items]
    tx = x + pad + (0.46 if numbered else 0.26)
    tw = w - (tx - x) - pad

    def measure(size, gap):
        total = 0.0
        for it in items:
            hh = max(0.26, est_height(it.get("heading"), tw, size, bold=True)) \
                if it.get("heading") else 0
            th_ = max(0.24, est_height(it.get("text", ""), tw, size)) \
                if it.get("text") else 0
            total += hh + (0.05 if hh and th_ else 0) + th_ + gap
        return total

    # Fit the rail to the panel: ease the gap first, then the type, then complain.
    avail = (y + h) - cy - pad
    gap = 0.22
    for g in (0.22, 0.18, 0.14, 0.10):
        if measure(body, g) <= avail:
            gap = g
            break
    else:
        for shrink in (0.5, 1.0, 1.5):
            if measure(body - shrink, 0.10) <= avail:
                body -= shrink
                gap = 0.10
                break
        else:
            die(f"the insight rail on '{spec.get('title', '')[:50]}' does not fit. "
                f"Cut an insight or shorten the text")

    for i, item in enumerate(items, start=1):
        head = item.get("heading")
        text = item.get("text", "")
        hh = max(0.26, est_height(head, tw, body, bold=True)) if head else 0
        th_ = max(0.24, est_height(text, tw, body)) if text else 0
        block_h = hh + (0.05 if hh and th_ else 0) + th_
        if numbered:
            b.rect(slide, x + pad, cy + 0.02, 0.32, 0.32,
                   fill=b.c["primary"], shape=MSO_SHAPE.OVAL)
            b.text(slide, x + pad, cy + 0.02, 0.32, 0.32, str(i),
                   size=body - 1.5, bold=True, color=b.c["white"],
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        else:
            b.rect(slide, x + pad + 0.02, cy + 0.10, 0.09, 0.09, fill=b.c["primary"])
        if head:
            b.text(slide, tx, cy, tw, hh, head, size=body, bold=True, color=b.c["ink"])
        if text:
            b.text(slide, tx, cy + (hh + 0.05 if hh else 0), tw, th_, text,
                   size=body, color=b.c["subink"] if head else b.c["ink"])
        cy += block_h + gap


def a_exhibit(b, slide, spec):
    """Evidence on the left, what it means on the right, action title on top.

    The standard consulting exhibit. The left half takes a chart, a table, stat
    tiles or an image. The right rail carries two to four insights.
    """
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    split = float(spec.get("split", 0.62))
    lw = r["w"] * split
    rx = r["x"] + lw + 0.40
    rw = r["w"] - lw - 0.40

    left = spec.get("left") or {}
    kind = left.get("type", "chart")
    if kind == "chart":
        _draw_chart(b, slide, left, r["x"], r["y"], lw, r["h"])
    elif kind == "table":
        _draw_table(b, slide, left, r["x"], r["y"], lw, r["h"])
    elif kind == "stats":
        _draw_stats(b, slide, left["stats"], r["x"], r["y"], lw, r["h"])
    elif kind == "image":
        cap = left.get("caption")
        ih = r["h"] - (0.30 if cap else 0.0)
        place_image(b, slide, left.get("image"), r["x"], r["y"], lw, ih,
                    left.get("image_note") or "image")
        if cap:
            b.text(slide, r["x"], r["y"] + ih + 0.06, lw, 0.24, cap,
                   size=b.sz("label", 11) - 1, color=b.c["muted"])
    else:
        die(f"exhibit left type '{kind}' is not one of chart, table, stats, image")

    _draw_insights(b, slide, spec, rx, r["y"], rw, r["h"])


def a_matrix(b, slide, spec):
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    axis = 0.42
    gap = 0.16
    gx = r["x"] + axis
    gy = r["y"]
    gw = r["w"] - axis
    gh = r["h"] - axis
    qw = (gw - gap) / 2
    qh = (gh - gap) / 2
    h_sz = b.sz("quad_heading", 13)
    t_sz = b.sz("quad_text", 12)
    quads = spec["quadrants"]
    if len(quads) != 4:
        die("matrix archetype needs exactly 4 quadrants")
    pos = [(gx, gy), (gx + qw + gap, gy),
           (gx, gy + qh + gap), (gx + qw + gap, gy + qh + gap)]
    for (x, y), q in zip(pos, quads):
        hi = bool(q.get("highlight"))
        b.rect(slide, x, y, qw, qh,
               fill=b.c["panel"] if hi else b.c["white"],
               line=b.c["primary"] if hi else b.c["rule"],
               line_w=1.25 if hi else 0.75)
        pad = 0.24
        head = q.get("heading", "")
        hh = max(0.30, est_height(head, qw - 2 * pad, h_sz, bold=True))
        b.text(slide, x + pad, y + pad, qw - 2 * pad, hh, head, size=h_sz,
               bold=True, color=b.c["primary"] if hi else b.c["ink"])
        if q.get("text"):
            b.text(slide, x + pad, y + pad + hh + 0.08, qw - 2 * pad,
                   qh - (2 * pad + hh + 0.08), q["text"], size=t_sz,
                   color=b.c["subink"])
    if spec.get("y_axis"):
        box = b.text(slide, r["x"] - 1.1, gy + gh / 2 - 0.16, 2.4, 0.32,
                     spec["y_axis"], size=b.sz("label", 11), bold=True,
                     color=b.c["muted"], align=PP_ALIGN.CENTER)
        box.rotation = 270
        box.left = Emu(int((r["x"] - 1.04) * EMU_IN))
    if spec.get("x_axis"):
        b.text(slide, gx, r["y"] + r["h"] - 0.28, gw, 0.28, spec["x_axis"],
               size=b.sz("label", 11), bold=True, color=b.c["muted"],
               align=PP_ALIGN.CENTER)


def a_table(b, slide, spec):
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    header = spec["header"]
    rows = spec["rows"]
    ncols = len(header)
    nrows = len(rows) + 1
    head_h = 0.48
    body_h = min(0.62, max(0.36, (r["h"] - head_h) / max(len(rows), 1)))
    total_h = head_h + body_h * len(rows)
    shape = slide.shapes.add_table(nrows, ncols, Inches(r["x"]), Inches(r["y"]),
                                   Inches(r["w"]), Inches(total_h))
    tbl = shape.table
    tbl.rows[0].height = Inches(head_h)
    for i in range(1, nrows):
        tbl.rows[i].height = Inches(body_h)
    widths = spec.get("col_widths")
    if widths:
        tot = sum(widths)
        for i, wgt in enumerate(widths):
            tbl.columns[i].width = Inches(r["w"] * wgt / tot)
    for j, label in enumerate(header):
        cell = tbl.cell(0, j)
        cell.text = str(label)
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(b.c["primary"])
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = cell.text_frame.paragraphs[0]
        para.font.name = b.font
        para.font.size = Pt(b.sz("table_header", 12))
        para.font.bold = True
        para.font.color.rgb = rgb(b.c["white"])
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = "" if val is None else str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(
                b.c["white"] if i % 2 else b.c["panel"])
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = cell.text_frame.paragraphs[0]
            para.font.name = b.font
            para.font.size = Pt(b.sz("table_body", 11.5))
            para.font.color.rgb = rgb(b.c["ink"])


def a_chart(b, slide, spec):
    b.strip_placeholders(slide, keep_title=True)
    b.set_title(slide, spec["title"])
    r = b.content_region(bool(spec.get("takeaway")))
    has_panel = bool(spec.get("read"))
    cw = r["w"] * 0.62 if has_panel else r["w"]
    kinds = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "stacked": XL_CHART_TYPE.COLUMN_STACKED,
    }
    kind = kinds.get(spec.get("chart_type", "bar"))
    if kind is None:
        die(f"unsupported chart_type '{spec.get('chart_type')}'")
    data = CategoryChartData()
    data.categories = spec["categories"]
    for series in spec["series"]:
        data.add_series(series["name"], series["values"])
    gf = slide.shapes.add_chart(kind, Inches(r["x"]), Inches(r["y"]),
                                Inches(cw), Inches(r["h"]), data)
    chart = gf.chart
    chart.font.name = b.font
    chart.font.size = Pt(b.sz("table_body", 11.5))
    chart.has_title = False

    # Direct labels beat an axis in a pitch: the number is read, not measured.
    labels_on = spec.get("data_labels", True)
    try:
        chart.value_axis.has_major_gridlines = False
        chart.value_axis.visible = not labels_on
        chart.value_axis.format.line.color.rgb = rgb(b.c["rule"])
        chart.category_axis.format.line.color.rgb = rgb(b.c["rule"])
        chart.category_axis.has_major_gridlines = False
        chart.plots[0].gap_width = 80
    except Exception:
        pass

    # A legend is always present for two or more series, never for one.
    if len(spec["series"]) > 1:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(b.sz("label", 11))
    else:
        chart.has_legend = False

    # Validated for colour-blind separation. See chart_palette in the profile.
    palette = b.p.get("chart_palette",
                      [b.c["primary"], b.c["secondary"], b.c["accent"]])
    if len(spec["series"]) > len(palette):
        die(f"{len(spec['series'])} series but the palette has {len(palette)} "
            f"slots. Fold the rest into 'Other' or split the chart")
    for i, plot_series in enumerate(chart.series):
        plot_series.format.fill.solid()
        plot_series.format.fill.fore_color.rgb = rgb(palette[i])
        plot_series.format.line.fill.background()
        if labels_on:
            try:
                labels = plot_series.data_labels
                labels.show_value = True
                labels.number_format = spec.get("number_format", "General")
                labels.number_format_is_linked = False
                labels.font.size = Pt(b.sz("label", 11) - 1)
                labels.font.name = b.font
                labels.font.bold = True
                labels.font.color.rgb = rgb(b.c["ink"])
            except Exception:
                pass
    if spec.get("unit"):
        b.text(slide, r["x"], r["y"] - 0.02, cw, 0.24, spec["unit"],
               size=b.sz("label", 11) - 1, color=b.c["muted"])
    if has_panel:
        px = r["x"] + cw + 0.35
        pw = r["w"] - cw - 0.35
        b.rect(slide, px, r["y"], pw, r["h"], fill=b.c["panel"])
        b.rect(slide, px, r["y"], pw, 0.055, fill=b.c["primary"])
        b.text(slide, px + 0.26, r["y"] + 0.32, pw - 0.52, 0.30,
               spec.get("read_heading", "What this says"),
               size=b.sz("col_heading", 14), bold=True, color=b.c["primary"])
        reads = spec["read"] if isinstance(spec["read"], list) else [spec["read"]]
        b.bullet_rows(slide, px + 0.26, r["y"] + 0.78, pw - 0.52, reads,
                      size=b.sz("chart_read", 12), row_gap=0.15)


RENDERERS = {
    "cover": a_cover, "section": a_section, "agenda": a_agenda,
    "statement": a_statement, "kpi": a_kpi, "bullets": a_bullets,
    "columns": a_columns,
    "cards": a_cards, "steps": a_steps, "initiative": a_initiative,
    "visual": a_visual, "exhibit": a_exhibit, "matrix": a_matrix,
    "table": a_table,
    "chart": a_chart,
}

LAYOUT_FOR = {
    # agenda uses the plain content layout: the template's 'Contents Green'
    # layout has a 6.775in title box and a busy watermark, neither of which
    # suits an action title.
    "cover": "cover", "section": "section", "agenda": "content",
    "statement": "content", "kpi": "content", "bullets": "content",
    "columns": "content",
    "cards": "content", "steps": "content", "initiative": "content",
    "visual": "content", "exhibit": "content", "matrix": "content",
    "table": "content",
    "chart": "content",
}


# ---------------------------------------------------------------- validation

def check_action_title(title, profile, idx, errors, warnings):
    t = profile["title"]
    if not title or not str(title).strip():
        errors.append(f"slide {idx}: action title is missing")
        return
    title = str(title).strip()
    words = title.split()
    if len(words) < 6:
        errors.append(f"slide {idx}: title is {len(words)} words. "
                      f"An action title is a full sentence, minimum 6 words: '{title}'")
    if len(title) > t["max_chars"]:
        errors.append(f"slide {idx}: title is {len(title)} chars, max {t['max_chars']}. "
                      f"It will not fit two lines at {t['size']}pt")
    elif len(title) > t["warn_chars"]:
        warnings.append(f"slide {idx}: title is {len(title)} chars and will wrap to two lines")
    if title.endswith("."):
        errors.append(f"slide {idx}: title ends with a full stop. Drop it")
    if BANNED_TITLE_STARTS.match(title):
        errors.append(f"slide {idx}: '{words[0]}' announces a topic instead of making a point. "
                      f"Say the so-what: '{title}'")
    stripped = [w.strip(".,;:!?") for w in words[1:]]
    cap_words = [w for w in stripped if w[:1].isupper() and w.isalpha() and len(w) > 3]
    if len(words) > 4 and len(cap_words) >= len(words) * 0.45:
        warnings.append(f"slide {idx}: title looks like Title Case. Use sentence case")
    if est_lines(title, t["w"], t["size"], bold=True) > 2:
        errors.append(f"slide {idx}: title needs more than two lines at {t['size']}pt")


def validate(spec, profile):
    errors, warnings = [], []
    slides = spec.get("slides") or []
    if not slides:
        die("storyline has no slides")
    for i, s in enumerate(slides, start=1):
        arch = s.get("archetype")
        if arch not in ARCHETYPES:
            errors.append(f"slide {i}: archetype '{arch}' is not in the closed set {list(ARCHETYPES)}")
            continue
        if arch == "cover":
            if not s.get("title"):
                errors.append("slide 1: cover needs a title")
            continue
        if arch == "section":
            if not s.get("title"):
                errors.append(f"slide {i}: section needs a title")
            continue
        check_action_title(s.get("title"), profile, i, errors, warnings)
        if not s.get("source"):
            errors.append(f"slide {i}: no source. Every content slide cites a source")
        # one message per slide
        if arch == "columns":
            cols = s.get("columns") or []
            if not 2 <= len(cols) <= 3:
                errors.append(f"slide {i}: columns archetype takes 2 or 3 columns, got {len(cols)}")
            for c in cols:
                if len(c.get("bullets") or []) > 5:
                    warnings.append(f"slide {i}: a column has {len(c['bullets'])} bullets. Cut to 5 or fewer")
        if arch == "cards":
            n = len(s.get("cards") or [])
            if not 2 <= n <= 4:
                errors.append(f"slide {i}: cards archetype takes 2 to 4 cards, got {n}")
        if arch == "steps":
            n = len(s.get("steps") or [])
            if not 3 <= n <= 5:
                errors.append(f"slide {i}: steps archetype takes 3 to 5 steps, got {n}")
        if arch == "bullets":
            pts = s.get("points") or []
            if not pts:
                errors.append(f"slide {i}: bullets archetype needs points")
            if len(pts) > 6:
                errors.append(f"slide {i}: {len(pts)} bullets. Six is the ceiling; split the slide")
        if arch == "table":
            if not s.get("header") or not s.get("rows"):
                errors.append(f"slide {i}: table needs header and rows")
            elif len(s["rows"]) > 9:
                warnings.append(f"slide {i}: {len(s['rows'])} table rows will crowd the frame")
        if arch == "chart":
            if not s.get("categories") or not s.get("series"):
                errors.append(f"slide {i}: chart needs categories and series")
        if arch == "kpi":
            st = s.get("stats") or []
            if not 2 <= len(st) <= 4:
                errors.append(f"slide {i}: a kpi slide takes 2 to 4 stats, got {len(st)}")
            for t in st:
                if not t.get("value") or not t.get("label"):
                    errors.append(f"slide {i}: every stat needs a value and a label")
        if arch == "exhibit":
            n = len(s.get("insights") or [])
            if not 2 <= n <= 4:
                errors.append(f"slide {i}: an exhibit takes 2 to 4 insights, got {n}")
            if not s.get("left"):
                errors.append(f"slide {i}: an exhibit needs a 'left' block")
        if arch == "visual" and not s.get("image"):
            errors.append(f"slide {i}: a visual slide needs an image")
        if arch == "agenda" and not s.get("items"):
            errors.append(f"slide {i}: agenda needs items")
    return errors, warnings


def validate_skeleton(spec, errors, warnings):
    """Enforce the innovation bootcamp pitch structure."""
    slides = spec["slides"]
    roles = [s.get("role") for s in slides]
    known = {r for r, _, _ in BOOTCAMP_SKELETON} | {"backup", "section"}
    for i, role in enumerate(roles, start=1):
        if role is None:
            errors.append(f"slide {i}: no role. Give it one of {sorted(known)}")
        elif role not in known:
            errors.append(f"slide {i}: role '{role}' is not part of the pitch structure")

    allowed = {r: arch for r, arch, _ in BOOTCAMP_SKELETON}
    for i, s in enumerate(slides, start=1):
        role = s.get("role")
        if role in allowed and s.get("archetype") not in allowed[role]:
            errors.append(
                f"slide {i}: a '{role}' slide uses one of {list(allowed[role])}, "
                f"not '{s.get('archetype')}'")

    # required roles, in order, before any backup section
    main = [r for r in roles if r != "backup"]
    required = [r for r, _, req in BOOTCAMP_SKELETON if req]
    for role in required:
        if role not in main:
            errors.append(f"the deck has no '{role}' slide. The pitch structure needs one")
    order = [r for r in required if r in main]
    seen = [r for r in main if r in required]
    first_at = []
    for role in order:
        first_at.append(seen.index(role))
    if first_at != sorted(first_at):
        errors.append("slides are out of order. The pitch runs cover, agenda, "
                      "situation, complication, problem, solutions, initiatives, "
                      "recap, ask")

    # three solutions, three initiatives
    for i, s in enumerate(slides, start=1):
        if s.get("role") == "solutions":
            n = len(s.get("columns") or [])
            if n != 3:
                errors.append(f"slide {i}: the solutions overview shows three "
                              f"initiatives, got {n}")
    inits = [s for s in slides if s.get("role") == "initiative"]
    names = [s.get("name") or s.get("title") for s in inits]
    distinct = {s.get("name") for s in inits if s.get("name")}
    if len(distinct) and len(distinct) != 3:
        errors.append(f"the detail slides cover {len(distinct)} initiative(s) by "
                      f"name, the pitch needs three")
    if not inits:
        errors.append("no initiative detail slides")
    for name in distinct:
        count = sum(1 for s in inits if s.get("name") == name)
        if count > 2:
            errors.append(f"initiative '{name}' has {count} detail slides. "
                          f"One, or two when it needs depth")
    if len(names) > len(set(names)) and not distinct:
        warnings.append("initiative slides have no 'name' field, so they cannot be "
                        "grouped. Add one per initiative")

    # field completeness where the jury will probe
    for i, s in enumerate(slides, start=1):
        if s.get("archetype") == "statement" and not s.get("statement"):
            errors.append(f"slide {i}: a statement slide needs a 'statement' line")
        if s.get("archetype") == "initiative":
            for field in ("name", "what", "how", "impact"):
                if not s.get(field):
                    errors.append(f"slide {i}: initiative detail is missing '{field}'. "
                                  f"The jury will ask")
            if not s.get("needs"):
                warnings.append(f"slide {i}: no 'needs' on this initiative. State the "
                                f"ask here or on the closing slide")


def check_vertical_logic(spec):
    """Titles read on their own must be the storyline."""
    titles = [s.get("title", "") for s in spec["slides"]
              if s.get("archetype") not in ("cover", "section", "agenda")]
    return titles


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("storyline")
    ap.add_argument("--profile", required=True)
    ap.add_argument("--template", help="override the template path in the profile")
    ap.add_argument("--out", required=True)
    ap.add_argument("--check-only", action="store_true")
    ap.add_argument("--structure", default="bootcamp",
                    choices=["bootcamp", "free"],
                    help="bootcamp enforces the competition pitch skeleton")
    args = ap.parse_args()

    global _SPEC_DIR
    _SPEC_DIR = os.path.dirname(os.path.abspath(args.storyline))

    spec = load_spec(args.storyline)
    profile = load_spec(args.profile)

    errors, warnings = validate(spec, profile)
    if args.structure == "bootcamp":
        validate_skeleton(spec, errors, warnings)
    for w in warnings:
        print(f"WARN  {w}")
    if errors:
        for e in errors:
            print(f"ERROR {e}", file=sys.stderr)
        print(f"\n{len(errors)} blocking issue(s). Nothing was built.", file=sys.stderr)
        sys.exit(2)

    print("\nTitles read on their own (this is the storyline the reader gets):")
    for i, t in enumerate(check_vertical_logic(spec), start=1):
        print(f"  {i}. {t}")

    if args.check_only:
        print("\nCheck only, no file written.")
        return

    tpl = args.template or os.path.join(
        os.path.dirname(os.path.abspath(args.profile)), "..", profile["template"])
    tpl = os.path.normpath(tpl)
    prs, tmpfile = open_template(tpl)

    # a template may ship with example slides; start clean (drop the parts too,
    # or the saved package carries orphaned slideN.xml entries)
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        rid = sld_id.rId
        sld_id_lst.remove(sld_id)
        prs.part.drop_rel(rid)

    b = Builder(prs, profile)
    for s in spec["slides"]:
        arch = s["archetype"]
        slide = b.new_slide(LAYOUT_FOR[arch])
        RENDERERS[arch](b, slide, s)
        if s.get("takeaway"):
            b.add_takeaway(slide, s["takeaway"])
        if s.get("source"):
            b.add_source(slide, s["source"])
        if s.get("notes"):
            slide.notes_slide.notes_text_frame.text = str(s["notes"])

    prs.save(args.out)
    if tmpfile:
        os.unlink(tmpfile)
    print(f"\nBuilt {len(spec['slides'])} slides -> {args.out}")


if __name__ == "__main__":
    main()
