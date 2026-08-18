"""Render a normalized presentation spec into a .pptx file.

Design philosophy: the model chooses *what* goes on each slide; this module
owns *how* it looks.  Every text block is measured and auto-scaled to fit its
box, so variable-length LLM output never overflows or clips — the single most
important thing for making generated decks look intentional.
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any, Iterable

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

try:  # Syntax highlighting is optional; code still renders without it.
    from pygments import lex
    from pygments.lexers import get_lexer_by_name

    _HAVE_PYGMENTS = True
except Exception:  # pragma: no cover - exercised only when pygments is absent
    _HAVE_PYGMENTS = False


# --------------------------------------------------------------------------- #
# Design system
# --------------------------------------------------------------------------- #

EMU_PER_INCH = 914400
SLIDE_W = 13.333
SLIDE_H = 7.5
MARGIN = 0.9  # outer margin, in inches

INK = RGBColor(0x1A, 0x1F, 0x2B)          # primary text
MUTED = RGBColor(0x6B, 0x72, 0x80)        # secondary text
FAINT = RGBColor(0x9A, 0xA1, 0xAD)        # tertiary / footer
PAPER = RGBColor(0xFB, 0xFC, 0xFE)        # slide background
PANEL = RGBColor(0xFF, 0xFF, 0xFF)        # card background
PANEL_ALT = RGBColor(0xF1, 0xF3, 0xF9)    # subtle fill
BORDER = RGBColor(0xE2, 0xE5, 0xEE)       # hairlines
ACCENT = RGBColor(0x4E, 0x5C, 0xDC)       # brand accent
ACCENT_SOFT = RGBColor(0xEC, 0xEE, 0xFF)  # accent tint
ACCENT_INK = RGBColor(0x2C, 0x36, 0x9E)   # text on accent tint

CODE_BG = RGBColor(0x1B, 0x1E, 0x27)
CODE_FG = RGBColor(0xE6, 0xE9, 0xF2)
CODE_DIM = RGBColor(0x8A, 0x93, 0xA8)

FONT_HEAD = "Aptos Display"
FONT_BODY = "Aptos"
FONT_CODE = "Menlo"

# Rough per-point character widths (in points) for line-fit estimation.
_PROP_CHAR_W = 0.50
_MONO_CHAR_W = 0.60


# --------------------------------------------------------------------------- #
# Low-level helpers
# --------------------------------------------------------------------------- #

def _kill_shadow(shape) -> None:
    """Remove the theme-driven preset shadow that muddies auto-shapes.

    python-pptx attaches a ``<p:style>`` whose ``<a:effectRef>`` points at a
    theme effect that includes an outer shadow.  An empty ``<a:effectLst>`` in
    spPr overrides it for PowerPoint, but some renderers still honour the style
    reference — so we also drop the style element entirely.  Fill and line are
    set explicitly on every shape, so nothing of value is lost.
    """
    sp = shape._element
    spPr = sp.spPr
    if spPr.find(qn("a:effectLst")) is None:
        spPr.append(spPr.makeelement(qn("a:effectLst"), {}))
    style = sp.find(qn("p:style"))
    if style is not None:
        sp.remove(style)


def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=0.75, radius=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(
        shape_type, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    if radius:
        # Gentler corner radius than the PowerPoint default.
        try:
            shape.adjustments[0] = 0.06
        except (IndexError, KeyError):
            pass

    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill

    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(line_w)

    shape.shadow.inherit = False
    _kill_shadow(shape)
    return shape


def _est_wrapped_lines(text: str, width_in: float, size_pt: float, mono: bool) -> int:
    """Estimate how many display lines a logical line wraps into."""
    if not text:
        return 1
    char_w = (_MONO_CHAR_W if mono else _PROP_CHAR_W) * size_pt
    chars_per_line = max(1, int((width_in * 72.0) / char_w))
    return max(1, math.ceil(len(text) / chars_per_line))


def _block_height_in(
    lines: Iterable[str],
    width_in: float,
    size_pt: float,
    mono: bool,
    line_spacing: float,
    space_after_pt: float,
) -> float:
    """Estimated rendered height (inches) of a set of paragraphs."""
    lines = list(lines)
    total_display_lines = sum(
        _est_wrapped_lines(ln, width_in, size_pt, mono) for ln in lines
    )
    line_h_pt = size_pt * 1.2 * line_spacing
    gaps_pt = space_after_pt * max(0, len(lines) - 1)
    return (total_display_lines * line_h_pt + gaps_pt) / 72.0


def fit_font_size(
    lines: list[str],
    width_in: float,
    height_in: float,
    max_pt: float,
    min_pt: float,
    mono: bool = False,
    line_spacing: float = 1.12,
    space_after_pt: float = 0.0,
) -> float:
    """Largest point size in [min, max] whose text block fits the box."""
    size = max_pt
    while size > min_pt:
        h = _block_height_in(
            lines, width_in, size, mono, line_spacing, space_after_pt
        )
        # Longest single line must also fit horizontally at this size.
        char_w = (_MONO_CHAR_W if mono else _PROP_CHAR_W) * size
        longest = max((len(ln) for ln in lines), default=0)
        fits_width = longest * char_w <= width_in * 72.0 or not mono
        if h <= height_in and fits_width:
            return round(size, 1)
        size -= 0.5
    return round(min_pt, 1)


def _heading_bottom(text, width_in, size_pt, top_in, line_spacing):
    """Estimated bottom edge (inches) of a bold-display heading block.

    Uses a wider per-char width than body text because the display font wraps
    sooner; over-estimating is safe here — it only pushes the accent rule and
    subtitle a little lower, never up into the title.
    """
    char_w = 0.62 * size_pt
    cpl = max(1, int((width_in * 72.0) / char_w))
    lines = max(1, math.ceil(len(text) / cpl))
    return top_in + (lines * size_pt * 1.2 * line_spacing) / 72.0


def _new_textbox(slide, x, y, w, h, valign=MSO_ANCHOR.TOP, margin=0.0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    return box, tf


def _style_run(run, text, size, color, bold=False, font=FONT_BODY, italic=False):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_text(
    slide, text, x, y, w, h,
    size=18, color=INK, bold=False, italic=False, font=FONT_BODY,
    align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, margin=0.0, line_spacing=1.0,
):
    """Single-block text that honours ``\\n`` as real line breaks."""
    box, tf = _new_textbox(slide, x, y, w, h, valign, margin)
    for i, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing != 1.0:
            p.line_spacing = line_spacing
        _style_run(p.add_run(), line, size, color, bold, font, italic)
    return box


# --------------------------------------------------------------------------- #
# Slide furniture
# --------------------------------------------------------------------------- #

def _blank_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, -0.05, -0.05, SLIDE_W + 0.1, SLIDE_H + 0.1, fill=PAPER)
    return slide


def add_header(slide, title, purpose=None):
    add_rect(slide, MARGIN, 0.62, 0.09, 0.44, fill=ACCENT, radius=False)

    title = title or ""
    size = fit_font_size([title], SLIDE_W - 2 * MARGIN - 0.3, 0.7, 30, 20)
    add_text(
        slide, title, MARGIN + 0.24, 0.5, SLIDE_W - 2 * MARGIN - 0.24, 0.72,
        size=size, color=INK, bold=True, font=FONT_HEAD,
        valign=MSO_ANCHOR.MIDDLE,
    )

    if purpose:
        add_text(
            slide, purpose, MARGIN + 0.24, 1.24, SLIDE_W - 2 * MARGIN - 0.24, 0.42,
            size=13, color=MUTED, line_spacing=1.0,
        )


def add_footer(slide, number, deck_title=""):
    if deck_title:
        add_text(
            slide, deck_title, MARGIN, SLIDE_H - 0.46, 8.0, 0.24,
            size=9, color=FAINT,
        )
    add_text(
        slide, f"{number:02d}", SLIDE_W - MARGIN - 0.6, SLIDE_H - 0.46, 0.6, 0.24,
        size=9, color=FAINT, align=PP_ALIGN.RIGHT,
    )


def add_bullets(slide, items, x, y, w, h, max_pt=20, min_pt=12, color=INK):
    """Bulleted list with a hanging indent, auto-scaled to fit the box."""
    items = [str(i) for i in items if str(i).strip()]
    if not items:
        return None

    space_after = 9.0
    size = fit_font_size(
        items, w - 0.3, h, max_pt, min_pt,
        line_spacing=1.08, space_after_pt=space_after,
    )

    box, tf = _new_textbox(slide, x, y, w, h)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = 1.06
        _apply_hanging_indent(p, bullet_in=0.28)

        dot = p.add_run()
        _style_run(dot, "•", size, ACCENT, bold=True)
        gap = p.add_run()
        _style_run(gap, "  ", size, ACCENT)
        _style_run(p.add_run(), item, size, color)
    return box


def _apply_hanging_indent(paragraph, bullet_in=0.28):
    pPr = paragraph._pPr
    if pPr is None:
        pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(bullet_in * EMU_PER_INCH)))
    pPr.set("indent", str(-int(bullet_in * EMU_PER_INCH)))


# --------------------------------------------------------------------------- #
# Code rendering (syntax highlighted)
# --------------------------------------------------------------------------- #

# Token colour map tuned for the dark code panel.
_CODE_COLORS = {
    "keyword": RGBColor(0xC5, 0x92, 0xFF),   # violet
    "string": RGBColor(0x8F, 0xE3, 0x88),    # green
    "number": RGBColor(0xF7, 0x9E, 0x6C),    # orange
    "comment": RGBColor(0x6B, 0x73, 0x89),   # dim grey
    "name": RGBColor(0x7F, 0xC7, 0xFF),      # blue
    "builtin": RGBColor(0x67, 0xE8, 0xE0),   # teal
    "operator": RGBColor(0xD5, 0xDA, 0xE6),  # light grey
    "text": CODE_FG,
}


def _token_color(tok_type) -> RGBColor:
    if not _HAVE_PYGMENTS:
        return CODE_FG
    t = str(tok_type)
    if "Comment" in t:
        return _CODE_COLORS["comment"]
    if "Keyword" in t:
        return _CODE_COLORS["keyword"]
    if "String" in t:
        return _CODE_COLORS["string"]
    if "Number" in t:
        return _CODE_COLORS["number"]
    if "Name.Builtin" in t or "Name.Function" in t or "Name.Class" in t:
        return _CODE_COLORS["builtin"]
    if "Operator" in t or "Punctuation" in t:
        return _CODE_COLORS["operator"]
    if "Name" in t:
        return _CODE_COLORS["name"]
    return _CODE_COLORS["text"]


def add_code_block(slide, code, language="sql", x=6.75, y=1.75, w=5.68, h=4.55,
                   fit_height=True):
    code = code.rstrip("\n")
    pad = 0.34
    header_h = 0.62
    cw = w - 2 * pad
    lines = code.split("\n")

    # Pick a comfortable size that fits the max box, then (optionally) shrink
    # the panel down to the content so short snippets aren't mostly empty.
    size = fit_font_size(
        lines, cw, h - header_h - pad, max_pt=13.5, min_pt=7.5,
        mono=True, line_spacing=1.18,
    )
    if fit_height:
        needed = header_h + _block_height_in(
            lines, cw, size, mono=True, line_spacing=1.18, space_after_pt=0
        ) + pad
        h = max(1.4, min(h, needed))

    add_rect(slide, x, y, w, h, fill=CODE_BG, radius=True)
    add_text(
        slide, (language or "code").upper(), x + 0.34, y + 0.22, 2.0, 0.24,
        size=9, color=CODE_DIM, bold=True, font=FONT_CODE,
    )

    cx, cy = x + pad, y + header_h
    ch = h - header_h - pad

    box, tf = _new_textbox(slide, cx, cy, cw, ch)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.18
        _render_code_line(p, line, size, language)
    return box


def _render_code_line(paragraph, line, size, language):
    if not line:
        _style_run(paragraph.add_run(), " ", size, CODE_FG, font=FONT_CODE)
        return
    if not _HAVE_PYGMENTS:
        _style_run(paragraph.add_run(), line, size, CODE_FG, font=FONT_CODE)
        return
    try:
        lexer = get_lexer_by_name(language or "sql")
    except Exception:
        lexer = get_lexer_by_name("text")
    for tok_type, value in lex(line + "\n", lexer):
        value = value.rstrip("\n")
        if not value:
            continue
        _style_run(
            paragraph.add_run(), value, size,
            _token_color(tok_type), font=FONT_CODE,
        )


# --------------------------------------------------------------------------- #
# Slide layouts: title + section
# --------------------------------------------------------------------------- #

def render_title(prs, data):
    slide = _blank_slide(prs)
    add_rect(slide, 0, 0, 0.28, SLIDE_H, fill=ACCENT)

    title = data.get("title", "Presentation")
    tw, ty = SLIDE_W - 2.4, 2.55
    size = fit_font_size([title], tw, 1.9, 46, 30, line_spacing=1.05)
    add_text(
        slide, title, 1.1, ty, tw, 1.9,
        size=size, color=INK, bold=True, font=FONT_HEAD, line_spacing=1.02,
    )

    # Place the accent rule and subtitle BELOW the title's real height so a
    # multi-line title never collides with them.
    rule_y = _heading_bottom(title, tw, size, ty, 1.02) + 0.16
    add_rect(slide, 1.14, rule_y, 1.1, 0.06, fill=ACCENT)

    subtitle = data.get("subtitle") or ""
    if subtitle:
        add_text(
            slide, subtitle, 1.14, rule_y + 0.18, tw, 0.9,
            size=18, color=MUTED, line_spacing=1.1,
        )


def render_section(prs, slide_data, number, deck_title):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, -0.05, -0.05, SLIDE_W + 0.1, SLIDE_H + 0.1, fill=INK)

    add_text(
        slide, f"{number - 1:02d}", MARGIN, 2.6, 2.0, 0.6,
        size=16, color=ACCENT_SOFT, bold=True, font=FONT_CODE,
    )

    # Fall back to the content line if the model gave the divider no title.
    content = slide_data.get("content", [])
    title = slide_data.get("title", "")
    tagline = content[0] if content else ""
    if not title:
        title, tagline = tagline, ""

    tw, ty = SLIDE_W - 2 * MARGIN, 3.05
    size = fit_font_size([title], tw, 1.6, 40, 26, line_spacing=1.05)
    add_text(
        slide, title, MARGIN, ty, tw, 1.6,
        size=size, color=PAPER, bold=True, font=FONT_HEAD, line_spacing=1.03,
    )

    # Rule and tagline sit below the title's measured height (no collision).
    rule_y = _heading_bottom(title, tw, size, ty, 1.03) + 0.14
    add_rect(slide, MARGIN + 0.02, rule_y, 1.1, 0.06, fill=ACCENT)

    if tagline:
        add_text(
            slide, tagline, MARGIN + 0.02, rule_y + 0.16, tw, 0.9,
            size=15, color=FAINT, line_spacing=1.15,
        )


# --------------------------------------------------------------------------- #
# Block layout engine
# --------------------------------------------------------------------------- #
#
# A slide body is a list of blocks laid out on a measured grid.  The named slide
# types are just recipes that compile to blocks (see ``slide_to_blocks``), and a
# slide may instead carry an explicit ``blocks`` list for a fully custom layout.
# Because Python still measures and fits every block, custom layouts can't
# overflow any more than the presets can.

BODY_Y = 1.85          # top of the body region, below the header
BODY_BOTTOM = 6.55     # above the footer
BODY_H = BODY_BOTTOM - BODY_Y
BLOCK_KINDS = {"bullets", "code", "columns", "cards", "steps", "callout", "text"}


def _items(value):
    if isinstance(value, list):
        return [str(i) for i in value if str(i).strip()]
    text = str(value).strip() if value is not None else ""
    return [text] if text else []


def _block_weight(b):
    """Relative vertical share a block wants when stacked with others."""
    kind = b.get("kind")
    if kind == "bullets":
        return max(1.0, len(_items(b.get("items"))))
    if kind == "text":
        return 1.0
    if kind == "callout":
        return 1.1
    if kind == "code":
        return max(2.0, len(str(b.get("code", "")).splitlines()) * 0.55)
    if kind == "steps":
        return 2.6
    if kind == "cards":
        return max(2.0, math.ceil(len(_items(b.get("items"))) / 2) * 1.5)
    if kind == "columns":
        return max((_column_weight(c) for c in b.get("columns", [])), default=1.0)
    return 1.0


def _column_weight(col):
    w = sum(_block_weight(b) for b in col.get("blocks", [])) or 1.0
    return w + (0.6 if col.get("heading") else 0.0)


def render_blocks(slide, blocks, x, y, w, h, body_color=INK):
    """Stack blocks vertically, splitting height by weight."""
    blocks = [b for b in blocks if isinstance(b, dict) and b.get("kind") in BLOCK_KINDS]
    if not blocks:
        return
    gap = 0.28
    weights = [_block_weight(b) for b in blocks]
    total = sum(weights) or 1.0
    avail = h - gap * (len(blocks) - 1)
    cy = y
    for b, wt in zip(blocks, weights):
        bh = max(0.4, avail * wt / total)
        render_block(slide, b, x, cy, w, bh, body_color)
        cy += bh + gap


def render_block(slide, b, x, y, w, h, body_color=INK):
    kind = b.get("kind")
    if kind == "bullets":
        add_bullets(slide, _items(b.get("items")), x, y, w, h,
                    max_pt=21, min_pt=12, color=body_color)
    elif kind == "text":
        add_text(slide, " ".join(_items(b.get("text") or b.get("items"))),
                 x, y, w, h, size=16, color=body_color, line_spacing=1.15)
    elif kind == "code":
        add_code_block(slide, str(b.get("code", "")), b.get("language", "sql"),
                       x=x, y=y, w=w, h=h)
    elif kind == "callout":
        _render_callout(slide, " ".join(_items(b.get("text") or b.get("items"))),
                        x, y, w, h)
    elif kind == "steps":
        _render_steps(slide, _items(b.get("items")), x, y, w, h)
    elif kind == "cards":
        _render_cards(slide, _items(b.get("items")), x, y, w, h)
    elif kind == "columns":
        _render_columns(slide, b.get("columns", []), x, y, w, h)


def _render_callout(slide, text, x, y, w, h):
    ch = min(h, 1.5)
    cy = y + max(0.0, (h - ch) / 2)
    add_rect(slide, x, cy, w, ch, fill=ACCENT_SOFT, radius=True)
    size = fit_font_size([text], w - 0.7, ch - 0.2, 19, 12, line_spacing=1.1)
    add_text(slide, text, x + 0.35, cy, w - 0.7, ch,
             size=size, color=ACCENT_INK, bold=True,
             align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, line_spacing=1.1)


def _render_steps(slide, items, x, y, w, h):
    items = items[:5]
    n = max(1, len(items))
    gap = 0.3
    cw = (w - gap * (n - 1)) / n
    ch = min(h, 3.0)
    cy = y + max(0.0, (h - ch) / 2)
    for i, item in enumerate(items):
        cx = x + i * (cw + gap)
        add_rect(slide, cx, cy, cw, ch, fill=PANEL, line=BORDER, radius=True)
        add_rect(slide, cx + 0.24, cy + 0.24, 0.5, 0.5, fill=ACCENT, radius=True)
        add_text(slide, str(i + 1), cx + 0.24, cy + 0.28, 0.5, 0.42,
                 size=16, color=PAPER, bold=True,
                 align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        size = fit_font_size([item], cw - 0.48, ch - 1.1, 15, 10, line_spacing=1.12)
        add_text(slide, item, cx + 0.24, cy + 0.95, cw - 0.48, ch - 1.15,
                 size=size, color=INK, line_spacing=1.12)
        if i < n - 1:
            add_text(slide, "→", cx + cw - 0.04, cy + ch / 2 - 0.2, gap + 0.08, 0.4,
                     size=20, color=ACCENT, bold=True,
                     align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


def _render_cards(slide, items, x, y, w, h, cols=2):
    items = items[:6]
    if not items:
        return
    gap = 0.4
    cw = (w - gap * (cols - 1)) / cols
    rows = max(1, math.ceil(len(items) / cols))
    ch = (h - gap * (rows - 1)) / rows
    for i, item in enumerate(items):
        c, r = i % cols, i // cols
        cx = x + c * (cw + gap)
        cy = y + r * (ch + gap)
        add_rect(slide, cx, cy, cw, ch, fill=PANEL, line=BORDER, radius=True)
        add_rect(slide, cx, cy, 0.09, ch, fill=ACCENT)
        add_text(slide, f"{i + 1:02d}", cx + 0.28, cy, 0.7, ch,
                 size=15, color=ACCENT, bold=True, font=FONT_CODE,
                 valign=MSO_ANCHOR.MIDDLE)
        size = fit_font_size([item], cw - 1.3, ch - 0.2, 16, 11, line_spacing=1.1)
        add_text(slide, item, cx + 0.95, cy, cw - 1.2, ch,
                 size=size, color=INK, valign=MSO_ANCHOR.MIDDLE, line_spacing=1.1)


def _render_columns(slide, columns, x, y, w, h):
    columns = [c for c in columns if isinstance(c, dict)]
    n = max(1, len(columns))
    gap = 0.4
    cw = (w - gap * (n - 1)) / n
    for i, col in enumerate(columns):
        _render_column(slide, col, x + i * (cw + gap), y, cw, h)


def _render_column(slide, col, x, y, w, h):
    heading = col.get("heading")
    accent = bool(col.get("accent"))
    ix, iy, iw, ih = x, y, w, h

    if heading or accent:
        add_rect(slide, x, y, w, h,
                 fill=ACCENT_SOFT if accent else PANEL,
                 line=None if accent else BORDER, radius=True)
        pad = 0.28
        ix, iy, iw, ih = x + pad, y + pad, w - 2 * pad, h - 2 * pad
        if heading:
            add_text(slide, str(heading), ix, iy, iw, 0.34,
                     size=13, color=ACCENT_INK if accent else ACCENT, bold=True)
            iy += 0.46
            ih -= 0.46

    render_blocks(slide, col.get("blocks", []), ix, iy, iw, ih,
                  body_color=ACCENT_INK if accent else INK)


# --------------------------------------------------------------------------- #
# Recipes: named slide types compile to blocks
# --------------------------------------------------------------------------- #

def slide_to_blocks(s):
    """Return the block list for a slide: explicit ``blocks`` or a recipe."""
    explicit = s.get("blocks")
    if isinstance(explicit, list) and explicit:
        return explicit

    t = s.get("type", "content")
    content = s.get("content", [])
    code = s.get("code", "")
    lang = s.get("language", "sql")

    if t == "code":
        if content and code:
            return [{"kind": "columns", "columns": [
                {"blocks": [{"kind": "bullets", "items": content}]},
                {"blocks": [{"kind": "code", "code": code, "language": lang}]},
            ]}]
        if code:
            return [{"kind": "code", "code": code, "language": lang}]
        return [{"kind": "bullets", "items": content}]

    if t == "comparison":
        left = s.get("left")
        right = s.get("right")
        if not left and not right:
            mid = max(1, math.ceil(len(content) / 2))
            left, right = content[:mid], content[mid:]
        return [{"kind": "columns", "columns": [
            {"heading": s.get("left_title") or "Option A", "accent": True,
             "blocks": [{"kind": "bullets", "items": left}]},
            {"heading": s.get("right_title") or "Option B",
             "blocks": [{"kind": "bullets", "items": right}]},
        ]}]

    if t == "process":
        return [{"kind": "steps", "items": content}]

    if t == "summary":
        return [{"kind": "cards", "items": content}]

    if t == "quiz":
        blocks = [{"kind": "bullets", "items": content}]
        if code:
            blocks.append({"kind": "code", "code": code, "language": lang})
        return blocks

    return [{"kind": "bullets", "items": content}]


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #

def render_body_slide(prs, slide_data, number, deck_title):
    slide = _blank_slide(prs)
    add_header(slide, slide_data.get("title", ""), slide_data.get("purpose"))
    render_blocks(
        slide, slide_to_blocks(slide_data),
        MARGIN, BODY_Y, SLIDE_W - 2 * MARGIN, BODY_H,
    )
    add_footer(slide, number, deck_title)


def _add_notes(slide, notes):
    if notes:
        slide.notes_slide.notes_text_frame.text = notes


def render_presentation(data: dict[str, Any], output: Path):
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    render_title(prs, data)

    for number, slide_data in enumerate(data.get("slides", []), start=2):
        if slide_data.get("type") == "section":
            render_section(prs, slide_data, number, data.get("title", ""))
        else:
            render_body_slide(prs, slide_data, number, data.get("title", ""))
        _add_notes(prs.slides[-1], slide_data.get("notes"))

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))
