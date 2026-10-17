# Examples

Use these specs as starting points. Pick a scene, edit the JSON, and run
`validate-deck.py` before rendering.

- [JSON format](json-format.md) lists fields, colors, and coordinates.
- [Validation](validation.md) explains the pure checker and the Windows preflight.
- [Agent workflow](agent-workflow.md) describes the edit and review loop.

Slide copy in the consulting, research, and product specs is Chinese. The
network diagram and the image asset demo are English. Treat the wording,
figures, and citations as sample layout content, and replace them for a real
deck. Specs stay short so they can be read in one pass. A deck can contain as
many slides as the scene needs.

| Scene | File | Demo fields |
| --- | --- | --- |
| Consulting: agenda, comparison, recommendation | [consulting-deck.json](../examples/consulting-deck.json) | `title`, `text`, `shape`, `line` (`from`, `to`, `endArrow`), `fill`, `notes` |
| Research brief: summary, findings, citations | [research-brief.json](../examples/research-brief.json) | `notes`, `title`, `text`, `shape`, `line` |
| Product overview: feature points and image | [product-overview.json](../examples/product-overview.json) | `assets`, `image` (`asset`, `fit`, `frame`), `shape`, `text`, `notes` |
| Network diagram | [network-diagram.json](../examples/network-diagram.json) | `title`, `text`, `box`, `line`, `notes` |
| Image asset | [image-asset-demo.json](../examples/image-asset-demo.json) | `assets`, `image` (`fit`), `shape`, `box`, `line`, `notes` |

## Consulting deck

[examples/consulting-deck.json](../examples/consulting-deck.json) is a three-slide
consulting narrative:

1. Agenda, with three steps joined by arrows.
2. Comparison of three approaches, laid out as shape cards.
3. A recommendation sentence and three next steps.

Edit the card copy, the comparison rows, and the recommendation sentence.
Connectors are `line` elements with `from`, `to`, and `endArrow`.

## Research brief

[examples/research-brief.json](../examples/research-brief.json) is a three-slide
brief:

1. Summary of the question and three takeaways.
2. Three findings, each with one statement and one supporting line.
3. A citations panel.

Every slide sets `notes`. Citation lines are marked 演示条目. Replace the panel
text before using the brief outside a demo.

## Product overview

[examples/product-overview.json](../examples/product-overview.json) is a
two-slide product page:

1. A short description, three feature points, and one picture.
2. Four feature points in a grid.

The picture uses asset id `productHero`. Its `path` is
`../assets/images/biometric-network-demo.png` relative to the spec, which
resolves to `assets/images/biometric-network-demo.png` in this repository.
`fit` is `contain`, and `frame` draws a rectangle behind the picture. To use
another local image, change that asset `path` to a file that exists beside the
spec.

## Check a spec

From the repository root:

```bash
python3 -m pip install -r requirements-validate.txt
python3 validate-deck.py examples/consulting-deck.json
python3 validate-deck.py examples/research-brief.json
python3 validate-deck.py examples/product-overview.json
```

The command exits 0 when the spec matches the schema and stays inside the page.
Rendering still uses `render-live-ppt.ps1` on Windows with desktop PowerPoint.
See the [README](../README.md).
