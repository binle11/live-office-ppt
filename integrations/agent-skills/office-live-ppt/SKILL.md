---
name: office-live-ppt
description: >
  Create, edit, and debug PowerPoint decks with the Live Office PPT workflow:
  write a JSON slide spec, optionally generate bitmap assets first, then run a
  PowerShell COM renderer that opens desktop PowerPoint and draws native
  editable shapes, text, connectors, and pictures live. Use when the user asks
  for PPT/PPTX creation with visible Office drawing, PowerPoint COM automation,
  live deck review, native PowerPoint objects, or an interactive alternative to
  SVG-first PPT generation.
---

# Office Live PPT

Use this skill to produce PowerPoint decks through the local `live-office-ppt`
renderer. The renderer turns JSON into native PowerPoint objects through Office
COM automation.

## Workflow

1. Locate the `live-office-ppt` project. If it is not in the active workspace,
   ask the user for its path.
2. Create the deck JSON spec in the user's project directory, not inside the
   reusable renderer repository unless the user is editing the renderer itself.
3. If bitmap visuals are needed, generate them first, save the selected files
   under the deck project's `assets/images/`, and add them to top-level
   `assets`.
4. Validate before opening PowerPoint. Schema, canvas bounds, and asset ids:
   ```bash
   python3 <live-office-ppt>/validate-deck.py <deck-spec.json>
   ```
   On Windows, the renderer preflight is also available. It checks slides and
   image files and does not load the JSON schema:
   ```powershell
   powershell -ExecutionPolicy Bypass -File <live-office-ppt>\render-live-ppt.ps1 `
     -Spec <deck-spec.json> `
     -ValidateOnly
   ```
5. Render with PowerPoint visible:
   ```powershell
   powershell -ExecutionPolicy Bypass -File <live-office-ppt>\render-live-ppt.ps1 `
     -Spec <deck-spec.json> `
     -StepDelayMs 80 `
     -Overwrite
   ```
   Use `-CloseAfterSave` for smoke tests.
6. Inspect the live PowerPoint result, revise JSON, and rerun.

## Image Assets

The renderer does not call image APIs. The agent should generate images before
rendering, then reference local files from the deck JSON.

```json
{
  "id": "hero",
  "type": "image",
  "source": "agent-image-generation",
  "path": "assets/images/hero.png",
  "prompt": "Clean professional slide illustration, no text, no logo."
}
```

Use editable PowerPoint text for labels and slide copy. Do not ask image models
to render dense slide text that should remain editable.

## Authoring Rules

- Prefer native PowerPoint text, shapes, and connectors over rasterized slide
  screenshots.
- Keep coordinates deterministic. The default 16:9 canvas is `960 x 540`.
- Draw elements in z-order: backgrounds first, foreground content last.
- Use `fit: "cover"` for full-bleed visuals and `fit: "contain"` for diagrams
  or screenshots that must not crop.
- Keep project-specific deck specs and private assets outside reusable examples.
- Run validation and a local render smoke test before handing off a deck.

## References

Load only when needed:

- `references/deck-authoring.md`: JSON element patterns and quality checks.
- `references/repository-quality.md`: repository maintenance and release checks.
