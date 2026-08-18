#!/usr/bin/env python3
"""Plain-English brief -> local model -> styled .pptx.

Write a topic (and optionally the structure you want) in a plain text file, and
the local model drafts a deck the renderer draws.
"""

import argparse
import subprocess
from pathlib import Path

from generate import deck_from_brief
from renderer import render_presentation


def main():
    parser = argparse.ArgumentParser(
        description="Turn a plain-English brief into a styled PowerPoint, locally."
    )
    parser.add_argument("input", type=Path,
                        help="A plain-text brief describing the deck")
    parser.add_argument("-o", "--output", type=Path,
                        help="Output .pptx path (default: same name as input)")
    parser.add_argument("--model", default="qwen2.5:3b",
                        help="Ollama model name (default: qwen2.5:3b)")
    parser.add_argument("--ollama-url", default="http://localhost:11434",
                        help="Ollama API URL")
    parser.add_argument("--open", action="store_true",
                        help="Open the .pptx when done (macOS)")
    args = parser.parse_args()

    try:
        brief = args.input.read_text(encoding="utf-8")
    except FileNotFoundError:
        raise SystemExit(f"Input file not found: {args.input}")

    output = args.output or args.input.with_suffix(".pptx")

    print(f"Drafting deck from brief with local model: {args.model} ...")
    deck = deck_from_brief(brief, model=args.model, ollama_url=args.ollama_url)

    render_presentation(deck, output)
    print(f"Created: {output.resolve()}  ({len(deck.get('slides', [])) + 1} slides)")

    if args.open:
        subprocess.run(["open", str(output)], check=False)


if __name__ == "__main__":
    main()
