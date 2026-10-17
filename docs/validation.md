# Validation

Deck specs can be checked before PowerPoint opens. Use the pure validator for
schema errors, canvas bounds, and asset references. Use the Windows renderer
preflight when you are about to render on a machine that has PowerPoint.

## Pure validator

`validate-deck.py` is the command for agents and CI, including Linux. It is a
Python program. It reads the spec, loads
[schema/deck.schema.json](../schema/deck.schema.json), then applies the
cross-field rules below. PowerPoint stays closed, and the command does not
write `exports/*.pptx`.

```bash
python3 -m pip install -r requirements-validate.txt
python3 validate-deck.py examples/network-diagram.json
python3 validate-deck.py examples/image-asset-demo.json
python3 validate-deck.py examples/consulting-deck.json
python3 validate-deck.py examples/research-brief.json
python3 validate-deck.py examples/product-overview.json
```

Several specs can be passed in one invocation. The process exits `0` when every
spec passes. It exits `1` when any spec fails, and prints one `path: reason`
line per failure to stderr:

```text
/slides/0/elements/0/text: missing required property
/slides/0/elements/0/w: 0 must be greater than 0
/slides/0/elements/0/x: rectangle is outside the canvas (x=1200 is past page width 960)
/slides/0/elements/0/asset: unknown asset id 'missing-hero'
```

Paths are JSON Pointers. A missing required field is reported on that field.
Schema failures name the same kind of path.

The checker stops after schema errors. Canvas and asset rules run once the
schema accepts the file, so a size error and a bounds error in the same element
are fixed in that order.

Run the fixture suite from the repository root:

```bash
python3 -m unittest tests/test_validation.py
```

`examples/network-diagram.json`, `examples/image-asset-demo.json`,
`examples/consulting-deck.json`, `examples/research-brief.json`, and
`examples/product-overview.json` are the passing specs. The scene map is in
[examples-index.md](examples-index.md). `tests/fixtures/invalid/` holds specs
that must fail:

| Fixture | Failure |
| --- | --- |
| `missing-required.json` | Text element omits `text` |
| `wrong-element-type.json` | Element `type` is not a supported kind |
| `non-positive-size.json` | `w` is 0 and `h` is negative |
| `outside-canvas.json` | Rectangle lies past the page width and height |
| `unknown-asset.json` | Image `asset` names an id that is not declared |
| `bad-color.json` | Slide background is not `#RRGGBB` |

## Windows renderer preflight

On Windows, this check still runs inside the renderer and returns before
PowerPoint is created:

```powershell
powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\network-diagram.json `
  -ValidateOnly
```

`-ValidateOnly` confirms the spec has slides and that image paths exist. It
does not load `schema/deck.schema.json`, so schema mistakes do not get a field
path from that command. Use `validate-deck.py` for schema, canvas, and asset-id
failures. Keep `-ValidateOnly` when you want the renderer’s own preflight on
Windows.

## Schema and validator split

JSON Schema can constrain a value by itself. It cannot compare an element with
the page size, because width and height live on a different object and change
from deck to deck. Identity and the filesystem are outside the schema for the
same reason.

The schema enforces:

- `slides` is present and contains at least one slide
- each element has a supported `type`
- `text` and `title` include `text`
- `text`, `title`, `box`, `shape`, and `image` include `x`, `y`, `w`, and `h`
- `w`, `h`, `fontSize`, and line `width` are greater than 0 when present
- `x`, `y`, and line coordinates are greater than or equal to 0
- page `width` and `height` are greater than 0 when present
- colors are `#RRGGBB`
- `fill` is `"none"`, a color, or a fill object with transparency from 0 to 1
- `line` is `"none"` or a line object
- an image includes `asset` or `path`
- a line includes `from`/`to` or `x1`, `y1`, `x2`, and `y2`
- `sleep.ms` is an integer greater than or equal to 0
- asset `id` values are non-empty strings

`validate-deck.py` then enforces:

- omitted page width defaults to `960`, and omitted page height defaults to `540`
- a positioned element stays inside the page, edges included: `x + w <= width` and `y + h <= height`
- line endpoints stay inside that same canvas
- asset ids are unique
- an image `asset` value matches a declared asset id
- a referenced asset has `path` or `file`
- those files, and a direct image `path`, exist on disk

Relative image paths are resolved from the spec file’s directory. Absolute paths
are used as written. An asset map may be an array of objects with `id`, or an
object whose keys are ids. When a map entry also sets `id`, that explicit id is
the reference name.
