# Live Office PPT

Render JSON slide specs into native, editable PowerPoint decks through desktop
Microsoft PowerPoint.

Live Office PPT uses the Windows PowerPoint COM object model to draw text boxes,
shapes, connectors, and pictures directly in PowerPoint. When run with
PowerPoint visible, the deck is built step by step, which makes the workflow
useful for agent-assisted slide authoring and live visual review.

It is not a macro generator, not a headless server renderer, and not an
SVG-to-PPTX converter. The goal is a conservative local renderer that produces
ordinary `.pptx` files with editable PowerPoint objects.

## Features

- JSON-to-PPTX rendering through PowerPoint COM
- Native PowerPoint text, shapes, connectors, and image objects
- Visible step-by-step drawing for live review
- Image assets referenced by ID from a top-level `assets` list
- Image fit modes: `cover`, `contain`, `stretch`, and `original`
- Speaker notes support
- COM-free validation for schema errors, canvas bounds, and asset references
- Windows validation mode that checks specs before opening PowerPoint
- Agent Skill integration for Codex and Claude Code
- OpenClaw-oriented adapter for Windows desktop workflows

## Requirements

- Windows
- Desktop Microsoft PowerPoint
- PowerShell 5.1 or newer

Rendering uses local Office automation and runs on Windows. Spec validation
with `validate-deck.py` also runs on Linux and in CI. That command needs
Python 3 and the packages in `requirements-validate.txt`.

## Quick Start

Clone the repository and run an example:

```powershell
git clone https://github.com/binle11/live-office-ppt.git
cd live-office-ppt

powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\network-diagram.json `
  -StepDelayMs 80 `
  -Overwrite
```

The generated deck is written to `exports/`.

Validate a spec without opening PowerPoint. Agents and CI should use the pure
validator. It loads the JSON schema, checks canvas bounds and asset ids, and
prints `path: reason` when a spec fails:

```bash
python3 -m pip install -r requirements-validate.txt
python3 validate-deck.py examples/network-diagram.json
```

On Windows, the renderer can also preflight a spec before it opens PowerPoint.
That check confirms slides and image files. Field-level schema results come
from `validate-deck.py`. See [docs/validation.md](docs/validation.md).

```powershell
powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\network-diagram.json `
  -ValidateOnly
```

Close PowerPoint automatically after a smoke test:

```powershell
powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\network-diagram.json `
  -StepDelayMs 20 `
  -CloseAfterSave `
  -Overwrite
```

## JSON Spec

A deck spec defines page settings, optional assets, and ordered slide elements.
Coordinates use PowerPoint points. The default 16:9 canvas is `960 x 540`.

```json
{
  "page": {
    "width": 960,
    "height": 540,
    "background": "#FFFFFF",
    "theme": {
      "font": "Aptos"
    }
  },
  "output": "exports/example.pptx",
  "slides": [
    {
      "background": "#F8FAFC",
      "elements": [
        {
          "type": "title",
          "text": "Live Office PPT",
          "x": 64,
          "y": 48,
          "w": 720,
          "h": 40,
          "fontSize": 24,
          "align": "left"
        }
      ]
    }
  ]
}
```

See [docs/json-format.md](docs/json-format.md),
[docs/validation.md](docs/validation.md),
[docs/agent-workflow.md](docs/agent-workflow.md), and
[schema/deck.schema.json](schema/deck.schema.json) for the full format,
validation commands, and agent-oriented workflow.

## Image Assets

The renderer inserts local image files as native PowerPoint picture objects. It
does not call image generation APIs. Agent workflows should generate or source
images first, save them into the deck project, and reference them from the spec.

```json
{
  "assets": [
    {
      "id": "hero",
      "type": "image",
      "source": "agent-image-generation",
      "path": "assets/images/hero.png",
      "prompt": "Clean professional slide illustration, no text."
    }
  ],
  "slides": [
    {
      "elements": [
        {
          "type": "image",
          "asset": "hero",
          "x": 480,
          "y": 0,
          "w": 480,
          "h": 540,
          "fit": "cover"
        }
      ]
    }
  ]
}
```

Use editable PowerPoint text for labels and slide copy. Avoid baking dense text
into generated images unless that text is intentionally part of the artwork.

## Agent Integrations

Live Office PPT includes a canonical Agent Skill and adapter docs:

- [Codex](integrations/codex/README.md)
- [Claude Code](integrations/claude-code/README.md)
- [OpenClaw](integrations/openclaw/README.md)

The canonical skill lives at:

```text
integrations/agent-skills/office-live-ppt
```

Codex and Claude Code can use that same skill folder. The OpenClaw adapter adds
runtime metadata for Windows desktop PowerPoint usage.

## Working With Your Own Decks

For real work, keep each deck's JSON spec and assets in a separate project
folder, then point `render-live-ppt.ps1` at that spec.

```powershell
powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec ..\my-deck-project\deck.json `
  -StepDelayMs 80 `
  -Overwrite
```

Generated PowerPoint files are written to `exports/`, which is ignored by Git.
Keep reusable templates and sanitized examples in this repository; keep
project-specific specs, private assets, and generated exports in the consuming
deck project unless you intentionally want to publish them.

## Examples

Choose a scene, edit the JSON, and validate the spec before rendering. The
[examples index](docs/examples-index.md) maps each scene to its file and the
fields it demonstrates.

- [examples/network-diagram.json](examples/network-diagram.json): native shapes, arrows, and text.
- [examples/image-asset-demo.json](examples/image-asset-demo.json): generated image asset inserted through COM.
- [examples/consulting-deck.json](examples/consulting-deck.json): Chinese consulting deck with agenda, comparison, and recommendation slides.
- [examples/research-brief.json](examples/research-brief.json): Chinese research brief with summary, findings, a citations area, and speaker notes.
- [examples/product-overview.json](examples/product-overview.json): Chinese product overview with feature points and a local image asset.

Field reference, checks, and the edit loop: [docs/json-format.md](docs/json-format.md),
[docs/validation.md](docs/validation.md), and
[docs/agent-workflow.md](docs/agent-workflow.md).

## Safety Model

Live Office PPT runs local PowerShell and automates desktop PowerPoint. Treat
deck specs and scripts as code:

- run specs from trusted sources
- review local image paths before rendering
- keep secrets and private project files out of reusable examples
- read [SECURITY.md](SECURITY.md) before using untrusted specs

Generated `.pptx` files do not contain auto-run macros.

## Roadmap

- More native chart primitives
- Template and theme support
- Optional grouped element abstractions

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). For agent-specific repository guidance,
see [AGENTS.md](AGENTS.md).

## License

MIT. See [LICENSE](LICENSE).
