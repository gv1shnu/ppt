"""Draft a full deck from a plain-English brief, using the local model.

Two stages, both on the local model via Ollama:

1. Outline — one small, fast call turns the brief into an ordered list of slide
   stubs (title + type + one-line focus).  If you describe structure in the
   brief, it's followed; if you only give a topic, a sensible flow is designed.
2. Write — each stub is expanded into a real slide (bullets, optional code,
   speaker notes) one at a time, with per-slide timeout and fallback.

Splitting it this way keeps every model call small, which is what makes this
reliable on a 3B model.  A single "generate the whole deck as JSON" call
truncates and drops slides; a short outline plus bounded per-slide calls does
not, and it shows progress as it goes.
"""

from __future__ import annotations

import json
from typing import Any

import requests

from model import chat, merge_slide
from schema import normalize_presentation


OUTLINE_SYSTEM = r"""You plan the structure of a slide deck.

You are given a brief in plain English. Return ONLY JSON, no prose, no fences:

{
  "title": string,
  "subtitle": string,
  "slides": [
    { "type": string, "title": string, "content": [ "one line: what this slide covers" ] }
  ]
}

Rules:
- If the brief describes sections, an order, or specific slides, FOLLOW it.
- If it only gives a topic, design a clear 6-9 slide flow that teaches it.
- Pick "type" per slide: "content" | "code" | "comparison" | "process" |
  "summary" | "quiz" | "section".
- Most slides should TEACH (content/code/comparison/process/summary). Use
  "section" ONLY as an occasional divider — at most one or two in the whole
  deck, never for a slide that carries real material.
- Use "code" when a slide should show code, "summary" for the closing recap.
- Give each slide ONE short line of focus in "content" (a seed, not the final
  bullets). Keep 4-12 slides total.
"""

SLIDE_GEN_SYSTEM = r"""You write ONE slide of a presentation.

You are given the deck topic, and this slide's title, type, and focus.
Return ONLY JSON for this one slide, no prose, no fences:

{
  "type": string,
  "title": string,
  "purpose": one line describing what the slide teaches,
  "content": [ 3-5 short strings, each under ~12 words ],
  "code": string (only if the slide should show code; use \n for newlines),
  "language": string (e.g. "sql", "python"; only with code),
  "notes": 1-3 sentences of speaker notes
}

Advanced (optional): if a standard layout does not fit, instead of "content"
you may return "blocks": an ordered list of layout blocks. Each block is one of:
{"kind":"bullets","items":[...]}, {"kind":"code","code":"...","language":"..."},
{"kind":"callout","text":"..."}, {"kind":"cards","items":[...]},
{"kind":"columns","columns":[{"heading":"...","blocks":[...]}, ...]}.
Use blocks ONLY when the composition clearly beats a plain layout; otherwise
prefer the simple fields above.

Rules:
- Write real, teachable bullets for the focus you were given.
- Add "code" only when a code example genuinely helps; keep it short and
  correct, and set "language".
- Be accurate. Do NOT fabricate specific numbers, names, dates, or citations
  you are not sure of — keep those bullets conceptual instead.
- Keep the deck's running example/domain consistent across slides.
- Always include "notes".
"""


def _outline(brief: str, model: str, ollama_url: str, timeout: int) -> dict[str, Any]:
    user = "Plan a deck from this brief:\n\n" + brief.strip()
    return chat(OUTLINE_SYSTEM, user, model, ollama_url, timeout)


def _cap_sections(slides: list[dict[str, Any]], max_sections: int) -> None:
    """Demote runaway "section" dividers to real content slides, in place.

    Small models sometimes label most of a deck as section dividers, which
    renders as a wall of title-only slides.  We keep the first few as dividers
    and turn the rest into content — done before slides are written, so the
    demoted ones get fleshed out with real bullets.
    """
    kept = 0
    for s in slides:
        if s.get("type") == "section":
            if kept < max_sections:
                kept += 1
            else:
                s["type"] = "content"


def _write_slide(stub: dict[str, Any], deck_title: str, model: str,
                 ollama_url: str, timeout: int) -> dict[str, Any]:
    focus = " ".join(stub.get("content", [])) or stub.get("title", "")
    user = (
        f"Deck topic: {deck_title}\n"
        f"Slide title: {stub.get('title', '')}\n"
        f"Slide type: {stub.get('type', 'content')}\n"
        f"Focus: {focus}\n\n"
        "Write this slide as JSON."
    )
    written = chat(SLIDE_GEN_SYSTEM, user, model, ollama_url, timeout)
    return merge_slide(stub, written)


def deck_from_brief(
    brief: str,
    model: str = "qwen2.5:3b",
    ollama_url: str = "http://localhost:11434",
    outline_timeout: int = 120,
    per_slide_timeout: int = 150,
) -> dict[str, Any]:
    """Return a rendered-ready deck spec drafted from a plain-English brief."""
    if not brief.strip():
        raise ValueError("The brief is empty — write a topic or outline first.")

    print("Planning the outline ...")
    try:
        outline = _outline(brief, model, ollama_url, outline_timeout)
        spec = normalize_presentation(outline)
    except (requests.RequestException, ValueError, json.JSONDecodeError, KeyError) as e:
        raise SystemExit(
            f"Could not plan a deck from the brief ({e}).\n"
            "Try a clearer brief, or check that Ollama is running."
        )

    deck_title = spec["title"]
    stubs = spec["slides"]
    total = len(stubs)
    _cap_sections(stubs, max_sections=max(1, total // 5))
    print(f"Outline: {total} slides — writing each ...")

    written = []
    ok = 0
    for i, stub in enumerate(stubs, start=1):
        label = stub.get("title") or f"slide {i}"
        try:
            slide = _write_slide(stub, deck_title, model, ollama_url, per_slide_timeout)
            written.append(slide)
            ok += 1
            print(f"  [{i}/{total}] {label}  ✓")
        except (requests.RequestException, ValueError, json.JSONDecodeError, KeyError):
            written.append(stub)  # keep the stub so the slide still appears
            print(f"  [{i}/{total}] {label}  ↺ (kept outline stub)")

    print(f"  wrote {ok}/{total} slides")
    spec["slides"] = written
    return spec
