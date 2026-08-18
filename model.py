"""Talk to a local model through Ollama and fold its output back into a slide.

Model-agnostic: anything Ollama can serve works — pass its tag with ``--model``.
Three small helpers shared by the deck generator:

- ``chat`` asks the model for a JSON object and returns it parsed.
- ``extract_json`` pulls that object out of a noisy response.
- ``merge_slide`` folds a model's slide onto a stub without losing information
  (existing code is preserved, empty fields fall back to the original).
"""

from __future__ import annotations

import json
import re
from typing import Any

import requests

from schema import VALID_TYPES, normalize_slide


def extract_json(text: str) -> dict[str, Any]:
    """Pull a JSON object out of a model response, tolerating stray text."""
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()
    fence = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, flags=re.DOTALL)
    if fence:
        text = fence.group(1)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start:
        return json.loads(text[start:end + 1])
    raise ValueError("No JSON object found in model response.")


def chat(system: str, user: str, model: str, ollama_url: str,
         timeout: int) -> dict[str, Any]:
    """Send one system+user turn and return the model's JSON reply."""
    url = ollama_url.rstrip("/") + "/api/chat"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "stream": False,
        "format": "json",
        # Disable "thinking" on reasoning models (e.g. qwen3) for faster,
        # cleaner JSON. Ollama ignores this for models that don't think.
        "think": False,
        "options": {"temperature": 0.3, "num_ctx": 4096},
    }
    r = requests.post(url, json=payload, timeout=timeout)
    r.raise_for_status()
    return extract_json(r.json()["message"]["content"])


def merge_slide(original: dict[str, Any], written: dict[str, Any]) -> dict[str, Any]:
    """Fold a model-written slide onto a stub, keeping the stub as a backstop.

    Content is only replaced when the model returns a non-empty list, and code
    is preserved from the stub when the model drops it — the model fills a
    slide in, it never empties one out.
    """
    out = dict(original)

    wtype = str(written.get("type", "")).lower()
    if wtype in VALID_TYPES and wtype not in ("title", "section"):
        out["type"] = wtype

    for key in ("title", "purpose", "language", "notes"):
        val = written.get(key)
        if isinstance(val, str) and val.strip():
            out[key] = val.strip()

    content = written.get("content")
    if isinstance(content, list) and any(str(c).strip() for c in content):
        out["content"] = content

    # A model may return a custom layout as "blocks"; pass it through for
    # schema.normalize_slide to validate. Malformed blocks are dropped there
    # and the slide falls back to its preset recipe, so this can't break a deck.
    wblocks = written.get("blocks")
    if isinstance(wblocks, list) and wblocks:
        out["blocks"] = wblocks

    new_code = written.get("code")
    if original.get("code") and not (isinstance(new_code, str) and new_code.strip()):
        out["code"] = original["code"]
    elif isinstance(new_code, str) and new_code.strip():
        out["code"] = new_code

    return normalize_slide(out, original.get("slide_number", 1))
