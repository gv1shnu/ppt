"""Normalize and validate a presentation specification.

The renderer is strict about shapes: every slide is a dict with a known
``type``, ``content`` is always a list of strings, and so on.  Hand-written
JSON and (especially) LLM output are messy, so everything flows through
``normalize_presentation`` first.  That keeps the renderer simple and means a
slightly-malformed AI response degrades gracefully instead of crashing or
silently dropping slides.
"""

from __future__ import annotations

import re
from typing import Any


VALID_TYPES = {
    "title",
    "content",
    "code",
    "comparison",
    "process",
    "summary",
    "quiz",
    "section",
}

# Sensible headings to fall back to when the model omits a slide title, so a
# slide never renders with a blank header.
DEFAULT_TITLES = {
    "summary": "Summary",
    "quiz": "Practice",
    "comparison": "Comparison",
    "process": "How it works",
}


def _strip_inline_markdown(text: str) -> str:
    """Remove inline markdown so it doesn't render literally on a slide.

    Small models like to wrap terms in `backticks` or **bold**.  Slides carry
    their own typography, so we strip the markers (this is formatting cleanup,
    not rewriting — the words are untouched).

    Only *paired* emphasis markers are removed.  A lone ``*`` is left alone so
    real content like ``SELECT * FROM t`` survives intact.
    """
    text = text.replace("`", "")
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)          # **bold**
    text = re.sub(r"\*(\S(?:.*?\S)?)\*", r"\1", text)      # *italic* (paired)
    return text.strip()


def _as_raw_text(value: Any) -> str:
    """Coerce a scalar to a string WITHOUT touching its characters.

    Used for the code field, where ``*`` and backticks are meaningful.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    return str(value).strip()


def _as_text(value: Any) -> str:
    """Coerce a scalar to a clean, slide-ready string (markdown stripped)."""
    return _strip_inline_markdown(_as_raw_text(value))


def _as_list(value: Any) -> list[str]:
    """Coerce content into a flat list of non-empty strings.

    Accepts a string (split on newlines), a list, or a dict of {key: value}
    pairs (rendered as ``key: value``), which small models like to emit.
    """
    if value is None or value == "":
        return []

    if isinstance(value, str):
        parts = [p.strip() for p in value.split("\n")]
        return [p for p in parts if p]

    if isinstance(value, dict):
        items = []
        for key, val in value.items():
            key = _as_text(key)
            val = _as_text(val)
            items.append(f"{key}: {val}" if val else key)
        return [i for i in items if i]

    if isinstance(value, list):
        items = []
        for entry in value:
            if isinstance(entry, (dict, list)):
                items.extend(_as_list(entry))
            else:
                text = _as_text(entry)
                if text:
                    items.append(text)
        return items

    text = _as_text(value)
    return [text] if text else []


BLOCK_KINDS = {"bullets", "code", "columns", "cards", "steps", "callout", "text"}


def normalize_block(b: Any) -> dict[str, Any] | None:
    """Validate and clean one layout block, or return None if unusable."""
    if not isinstance(b, dict):
        return None
    kind = _as_text(b.get("kind")).lower()
    if kind not in BLOCK_KINDS:
        return None

    if kind in ("bullets", "cards", "steps"):
        items = _as_list(b.get("items"))
        return {"kind": kind, "items": items} if items else None

    if kind in ("text", "callout"):
        text = _as_text(b.get("text")) or " ".join(_as_list(b.get("items")))
        return {"kind": kind, "text": text} if text.strip() else None

    if kind == "code":
        code = _as_raw_text(b.get("code"))
        if not code:
            return None
        return {"kind": "code", "code": code,
                "language": _as_text(b.get("language")) or "sql"}

    if kind == "columns":
        cols = []
        for c in b.get("columns") or []:
            if not isinstance(c, dict):
                continue
            child = normalize_blocks(c.get("blocks"))
            if not child:
                continue
            col: dict[str, Any] = {"blocks": child}
            if c.get("heading"):
                col["heading"] = _as_text(c.get("heading"))
            if c.get("accent"):
                col["accent"] = True
            cols.append(col)
        return {"kind": "columns", "columns": cols} if cols else None

    return None


def normalize_blocks(value: Any) -> list[dict[str, Any]]:
    """Clean a list of layout blocks, dropping any that are unusable."""
    if not isinstance(value, list):
        return []
    return [nb for nb in (normalize_block(b) for b in value) if nb]


def _infer_type(slide: dict[str, Any]) -> str:
    """Pick a layout when the spec omits ``type``."""
    declared = _as_text(slide.get("type")).lower()
    if declared in VALID_TYPES:
        return declared

    title = _as_text(slide.get("title")).lower()
    has_code = bool(_as_text(slide.get("code")))
    n_items = len(_as_list(slide.get("content")))

    if "quiz" in title or "practice" in title or "exercise" in title:
        return "quiz"
    if "summary" in title or "recap" in title or "takeaway" in title:
        return "summary"
    if "vs" in title or "versus" in title or "compare" in title:
        return "comparison"
    if has_code and n_items:
        return "code"
    if has_code:
        return "code"
    return "content"


def normalize_slide(slide: Any, index: int) -> dict[str, Any]:
    """Return a slide dict with every field the renderer expects."""
    if not isinstance(slide, dict):
        # A bare string/list becomes a content slide.
        slide = {"content": slide}

    slide_type = _infer_type(slide)
    code = _as_raw_text(slide.get("code"))

    # A slide that carries code must use a layout that shows it.  Models often
    # mislabel a code slide as "content", which would silently hide the code.
    if code and slide_type not in ("code", "quiz"):
        slide_type = "code"

    title = _as_text(slide.get("title"))
    if not title and slide_type in DEFAULT_TITLES:
        title = DEFAULT_TITLES[slide_type]

    normalized = {
        "slide_number": slide.get("slide_number", index),
        "type": slide_type,
        "title": title,
        "purpose": _as_text(slide.get("purpose") or slide.get("subtitle")),
        "content": _as_list(slide.get("content") or slide.get("bullets")),
        "code": code,
        "language": _as_text(slide.get("language")) or "sql",
        "notes": _as_text(slide.get("notes") or slide.get("speaker_notes")),
    }

    # Comparison slides may carry structured columns; keep them if present.
    for key in ("left", "right", "left_title", "right_title"):
        if key in slide:
            normalized[key] = (
                _as_list(slide[key])
                if key in ("left", "right")
                else _as_text(slide[key])
            )

    # A slide may carry an explicit block tree for a fully custom layout.
    blocks = normalize_blocks(slide.get("blocks"))
    if blocks:
        normalized["blocks"] = blocks

    return normalized


def normalize_presentation(data: Any) -> dict[str, Any]:
    """Return a clean presentation dict: ``title``, ``subtitle``, ``slides``.

    Raises ``ValueError`` only when the input is unusable (no slides at all),
    so callers can fall back to the original hand-written spec.
    """
    if not isinstance(data, dict):
        raise ValueError("Presentation must be a JSON object.")

    raw_slides = data.get("slides")
    if not isinstance(raw_slides, list) or not raw_slides:
        raise ValueError('Presentation must contain a non-empty "slides" array.')

    slides = [normalize_slide(s, i) for i, s in enumerate(raw_slides, start=1)]
    # Drop slides that are completely empty after normalization.
    slides = [
        s
        for s in slides
        if s["title"] or s["content"] or s["code"] or s.get("blocks")
    ]

    if not slides:
        raise ValueError("No renderable slides after normalization.")

    return {
        "title": _as_text(data.get("title")) or "Presentation",
        "subtitle": _as_text(data.get("subtitle")),
        "slides": slides,
    }
