---
name: office-live-ppt
description: >
  Create, edit, and debug PowerPoint decks with the Live Office PPT workflow:
  write a JSON slide spec, optionally generate bitmap assets first, then run a
  PowerShell COM renderer that opens desktop PowerPoint and draws native
  editable shapes, text, connectors, and pictures live. Use for PPT/PPTX tasks
  that require visible Office drawing, PowerPoint COM automation, live deck
  review, or native editable PowerPoint objects.
metadata:
  openclaw:
    required_bins:
      - powershell
    requires_desktop: true
    requires_powerpoint: true
---

# Office Live PPT

Use this skill only on Windows machines with desktop Microsoft PowerPoint
available. The renderer is a local Office automation workflow, not a headless
server renderer.

## Workflow

1. Locate the `live-office-ppt` repository.
2. Create the deck JSON spec in the user's project directory.
3. If bitmap visuals are needed, generate or source them first, save them under
   the deck project's `assets/images/`, and add them to top-level `assets`.
4. Validate before opening PowerPoint. Schema, canvas bounds, and asset ids:
   ```bash
   python3 <live-office-ppt>/validate-deck.py <deck-spec.json>
   ```
   The renderer preflight checks slides and image files. It does not load the
   JSON schema:
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
6. Inspect the live result, revise JSON, and rerun.

## Authoring Rules

- Prefer native PowerPoint text, shapes, and connectors over rasterized slides.
- Use editable PowerPoint text for labels and slide copy.
- Keep coordinates deterministic. The default 16:9 canvas is `960 x 540`.
- Use `fit: "cover"` for full-bleed images and `fit: "contain"` for diagrams
  or screenshots that must not crop.
- Keep project-specific materials outside reusable repository examples.
