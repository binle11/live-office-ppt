# Deck Authoring

## Spec Skeleton

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
  "slides": [
    {
      "background": "#F8FAFC",
      "notes": "",
      "elements": []
    }
  ]
}
```

## Common Elements

Text:

```json
{
  "type": "text",
  "text": "Short body copy",
  "x": 64,
  "y": 120,
  "w": 360,
  "h": 64,
  "fontSize": 16,
  "align": "left",
  "color": "#334155"
}
```

Shape:

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
  "x": 280,
  "y": 160,
  "w": 120,
  "h": 56,
  "fill": "#EAF3FF",
  "line": {
    "color": "#93C5FD",
    "width": 1.2
  }
}
```

Connector:

```json
{
  "type": "line",
  "from": [400, 188],
  "to": [500, 188],
  "line": {
    "color": "#64748B",
    "width": 1.5,
    "endArrow": "triangle"
  }
}
```

Image:

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

## Quality Checks

Before rendering:

```powershell
powershell -ExecutionPolicy Bypass -File <tool>\render-live-ppt.ps1 `
  -Spec <deck-spec.json> `
  -ValidateOnly
```

After rendering, inspect:

- Text is editable and not rasterized.
- Long text fits its box.
- Connectors point to the intended shapes.
- Image crops are intentional.
- Speaker notes are present when requested.
