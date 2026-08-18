# Plain English → PowerPoint, with a local model

Describe a deck in plain English and get a styled `.pptx` back, generated on your
own machine. A local model (via [Ollama](https://ollama.com)) drafts the slides;
Python does all the drawing. Nothing leaves your laptop.

You write the topic and, if you want, the structure — "open with a section on X,
then compare A and B, then a code example, close with a summary." The model turns
that into slides. It only decides *what goes on each slide* (bullets, layout,
speaker notes); it never touches fonts, spacing, colors, or the PowerPoint file
itself. That split is why the output looks consistent even though a 3B model
wrote the words.

```
your brief  ──►  model plans an outline  ──►  model writes each slide  ──►  renderer  ──►  deck.pptx
(plain text)     (one fast call)              (one call per slide,          fit text,
                                               with progress + fallback)     highlight code
```

## Requirements

- macOS (the renderer is cross-platform; `--open` is macOS-only)
- Python 3.10+
- [Ollama](https://ollama.com) running locally

Pull the default model:

```bash
ollama pull qwen2.5:3b
```

## Choosing a model

Any model Ollama can serve works — pass its tag with `--model`. The one rule
that matters on a laptop: the model has to fit in memory *with room to spare*, or
it swaps and crawls. Bigger isn't better if it doesn't fit.

Recommended models, smallest to largest (RAM and speed measured/estimated on an
8 GB M1 with `llmfit`):

| model | params | needs ~RAM | notes |
|-------|--------|-----------|-------|
| `qwen2.5:3b` *(default)* | 3B | 3.7 GB | Best fit for 8 GB machines; unusually good at clean JSON. ~12 tok/s. |
| `llama3.2:3b` | 3B | 4.2 GB | Comparable alternative in the 3B class. |
| `qwen3:4b` | 4B | 5.0 GB | A step smarter; still fine on 8 GB if little else is open. |
| `qwen2.5:7b` / `qwen3:8b` | 7–8B | 5–5.6 GB* | Better writing — but wants **≥16 GB RAM** in practice. |

*`llmfit` rates the 7–8B models a "perfect" fit by weight alone, but on an 8 GB
M1 with the OS and apps already using ~7 GB, the 8B swaps and drops to ~2.5 tok/s
— slow enough that a whole-deck request times out. Below ~16 GB, stay at 3–4B.

```bash
ollama pull qwen3:4b
python main.py examples/brief.txt --model qwen3:4b
```

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Make a deck

Write a brief in a plain text file (see `examples/brief.txt`):

```text
Topic: Introduction to database indexing for backend engineers.

Structure I want:
- Open with a section divider for "Why indexes matter".
- What a database index actually is (analogy to a book index).
- Compare a full table scan vs an index lookup.
- A code slide showing CREATE INDEX and a query that uses it.
- The trade-offs (writes get slower, storage cost).
- Close with a short summary of when to add an index.

Keep it practical and beginner-friendly.
```

Then generate:

```bash
python main.py examples/brief.txt
```

You'll see it plan an outline, then write each slide with live progress, and
write the `.pptx` next to your input. Other flags:

```bash
python main.py examples/brief.txt -o talk.pptx     # choose the output path
python main.py examples/brief.txt --open           # open it when done (macOS)
python main.py examples/brief.txt --model llama3.2:3b
```

The more structure your brief gives, the more the deck follows it. Give it only a
topic and it designs a sensible flow on its own.

## How much can go wrong?

Not much, by design. The model works in small steps — a short outline, then one
slide at a time — so a slow or malformed response only affects that one slide,
and it falls back to the outline stub instead of failing the whole run. Every
slide is validated before it's drawn. The model won't invent code it wasn't given
and can't drop a slide out of the deck.

It's still a 3B model drafting a first pass, so read what it wrote before you
present it.

## Slide types

The model picks one type per slide. Each is a *recipe* over the block engine
below — a convenient name for a common composition:

| type | looks like |
|------|-----------|
| `content` | title + bullet list |
| `code` | bullets on the left, highlighted code on the right |
| `comparison` | two labeled columns |
| `process` | numbered steps left to right |
| `summary` | a grid of key takeaways |
| `quiz` | questions, with optional code |
| `section` | a full-bleed divider between parts of the deck |

A slide with code always renders a code layout, so a snippet never silently
disappears.

## Custom layouts (blocks)

The seven types are shortcuts. Under them is a small **block engine**: a slide
body is a list of blocks the renderer arranges on a measured grid, so you can
compose a layout no preset offers — and it still can't overflow, because every
block is measured and auto-fit like the presets.

Give a slide a `blocks` list instead of `content` (the model may do this too):

```json
{
  "title": "Offense and defense are one skill",
  "blocks": [
    { "kind": "callout", "text": "Every attack is taught with its detection." },
    { "kind": "columns", "columns": [
      { "heading": "Attack", "accent": true,
        "blocks": [{ "kind": "bullets", "items": ["Enumerate", "Exploit", "Escalate"] }] },
      { "heading": "Defend",
        "blocks": [{ "kind": "code", "code": "alert tcp any any -> $HOME_NET 22 ...", "language": "bash" }] }
    ]}
  ]
}
```

Block kinds: `bullets`, `code`, `callout`, `text`, `cards`, `steps`, and
`columns` (which nests blocks side by side). Blocks stack vertically and split
the height by weight; `columns` splits the width. Anything malformed is dropped
and the slide falls back to its preset, so a bad block can't break the deck.

## How the drawing works

The renderer measures every text block and scales the font to fit its box, so
variable-length model output doesn't overflow or clip — the main reason generated
decks usually look off. Code is highlighted with [Pygments](https://pygments.org)
on a dark panel and shrinks to fit. Speaker notes ride along in the `.pptx` notes
pane. (Without Pygments, code still renders, just in one color.)

## Files

- `main.py` — CLI entry point
- `generate.py` — drafts a deck from a brief: outline, then each slide
- `model.py` — talks to any Ollama model and merges its output into a slide
- `schema.py` — normalizes and validates a deck into what the renderer expects
- `renderer.py` — the design system and all layouts
- `examples/brief.txt` — a sample brief
