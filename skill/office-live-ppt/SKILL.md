---
name: office-live-ppt
description: >
  Create, edit, and debug PowerPoint decks with the Live Office PPT workflow:
  Codex writes a JSON slide spec, optionally generates bitmap assets with image
  generation, then runs a PowerShell COM renderer that opens desktop PowerPoint
  and draws native editable shapes, text, connectors, and pictures live. Use
  when the user asks for PPT/PPTX creation with visible Office drawing,
  PowerPoint COM automation, live deck review, native PowerPoint objects, or
  replacing ppt-master/ppt-skill with an interactive Office-rendered workflow.
---

# Office Live PPT

Use this skill to make PowerPoint decks through the local `live-office-ppt`
project instead of SVG-first PPT pipelines.

## Default Workflow

1. Locate the `live-office-ppt` project. In this machine it is usually in the
   active workspace as `<workspace>\live-office-ppt`; if absent, ask the user
   where it is.
2. Create or update a deck JSON spec in the user's project directory, not inside
   the reusable tool unless the user is editing the tool itself.
3. If the deck needs bitmap visuals, use the built-in image generation tool,
   then copy the selected image files into the deck project's `assets/images/`.
4. Add generated images to top-level `assets` and reference them from image
   elements by `asset` ID.
5. Run the renderer with `-ValidateOnly` before opening PowerPoint.
6. Render with desktop PowerPoint visible:
   ```powershell
   powershell -ExecutionPolicy Bypass -File <live-office-ppt>\render-live-ppt.ps1 `
     -Spec <deck-spec.json> `
     -StepDelayMs 80 `
     -Overwrite
   ```
   Use `-CloseAfterSave` for smoke tests where the PowerPoint window should not
   remain open.
7. Inspect the live PowerPoint result, revise JSON, and rerun.

## Image Assets

The renderer does not call image APIs. Codex handles image generation before
rendering:

1. Generate images with the built-in image generation tool unless the user
   explicitly requests another image backend.
2. Move or copy final selected images into the workspace. Never leave a
   project-referenced asset only under Codex's generated image cache.
3. Record the prompt in `assets[].prompt`.
4. Use editable PowerPoint text for labels; do not ask image generation to
   render dense slide copy.

Example asset:

```json
{
  "id": "hero",
  "type": "image",
  "source": "codex-imagegen",
  "path": "assets/images/hero.png",
  "prompt": "Clean professional slide illustration, no text, no logo."
}
```

Example image element:

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

## Authoring Rules

- Prefer native PowerPoint text, shapes, and connectors over rasterized slide
  screenshots.
- Keep coordinates deterministic. Default 16:9 canvas is `960 x 540`.
- Draw elements in z-order: backgrounds first, foreground content last.
- Use `fit: "cover"` for full-bleed visuals and `fit: "contain"` for diagrams
  or screenshots that must not crop.
- Validate image paths before opening PowerPoint.
- Do not put confidential source files, client content, or local absolute paths
  into reusable examples.

## References

Load only when needed:

- `references/deck-authoring.md`: JSON element patterns and quality checks.
- `references/project-release.md`: GitHub cleanup and publication checklist.
