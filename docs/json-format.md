# JSON Format

The renderer consumes one JSON file per deck. Coordinates use PowerPoint points.
The default 16:9 canvas is `960 x 540`, with `(0, 0)` at the top-left.

## Top-Level Structure

```json
{
  "version": "0.2.0",
  "page": {
    "width": 960,
    "height": 540,
    "background": "#FFFFFF",
    "theme": {
      "font": "Aptos"
    }
  },
  "output": "exports/deck.pptx",
  "assets": [],
  "slides": []
}
```

## Assets

Assets define reusable local files, especially generated images.

```json
{
  "id": "hero",
  "type": "image",
  "source": "agent-image-generation",
  "path": "assets/images/hero.png",
  "prompt": "Clean professional slide illustration, no text."
}
```

Rules:

- `id` must be unique.
- `path` is resolved relative to the JSON spec file.
- `prompt` is documentation for the agent workflow; the renderer does not call
  image generation APIs.
- Missing image files fail validation before PowerPoint opens.

## Slides

Each slide has a background and ordered elements:

```json
{
  "background": "#F8FAFC",
  "notes": "Speaker notes.",
  "elements": []
}
```

Elements are drawn in order. Later elements appear above earlier elements.

## Text and Titles

```json
{
  "type": "title",
  "text": "Slide Title",
  "x": 64,
  "y": 48,
  "w": 720,
  "h": 40,
  "fontSize": 24,
  "bold": true,
  "align": "left",
  "color": "#0F172A"
}
```

`type: "text"` uses the same fields. Supported alignment values are `left`,
`center`, and `right`. Supported vertical alignment values are `top`, `middle`,
and `bottom`.

## Shapes

```json
{
  "type": "box",
  "shape": "roundRect",
  "text": {
    "text": "Encoder",
    "fontSize": 13,
    "bold": true,
    "color": "#1D4ED8"
  },
  "x": 120,
  "y": 160,
  "w": 128,
  "h": 56,
  "fill": "#EAF3FF",
  "line": {
    "color": "#93C5FD",
    "width": 1.2
  }
}
```

Supported shape names:

- `rect`
- `roundRect`
- `oval`
- `diamond`
- `triangle`
- `hexagon`
- `parallelogram`

Set `fill` or `line` to `"none"` to hide it.

## Lines and Connectors

```json
{
  "type": "line",
  "connector": "straight",
  "from": [240, 188],
  "to": [360, 188],
  "line": {
    "color": "#64748B",
    "width": 1.5,
    "endArrow": "triangle"
  }
}
```

Supported connector values:

- `straight`
- `elbow`
- `curve`

Supported arrow values:

- `triangle`
- `stealth`
- `oval`
- `diamond`
- `none`

## Images

```json
{
  "type": "image",
  "asset": "hero",
  "x": 480,
  "y": 0,
  "w": 480,
  "h": 540,
  "fit": "cover"
}
```

Use `asset` for reusable files or `path` for one-off local files.

Fit modes:

- `cover`: crop to fill the frame.
- `contain`: preserve aspect ratio and center inside the frame.
- `stretch`: fill the frame exactly.
- `original`: insert at native size.

Optional frame:

```json
{
  "type": "image",
  "asset": "hero",
  "x": 100,
  "y": 100,
  "w": 320,
  "h": 180,
  "fit": "contain",
  "frame": {
    "shape": "rect",
    "fill": "#FFFFFF",
    "line": {
      "color": "#CBD5E1",
      "width": 1
    }
  }
}
```

## Pauses

Use `sleep` to slow the visible drawing flow:

```json
{
  "type": "sleep",
  "ms": 300
}
```

Most decks should instead use the command-line `-StepDelayMs` option.
