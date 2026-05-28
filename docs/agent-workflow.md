# Agent Workflow

Live Office PPT works well with coding agents because the deck is split into two
editable layers:

1. A deterministic JSON slide spec.
2. A PowerPoint COM renderer that turns the spec into native PPT objects.

## Standard Flow

1. Plan the deck structure and visual direction.
2. Write a JSON spec under the deck project workspace.
3. If the deck needs bitmap visuals, generate or source them first and copy the
   final files into the deck project's `assets/images/`.
4. Add each image to top-level `assets`.
5. Validate the spec with `-ValidateOnly`.
6. Render with PowerPoint visible and a moderate `-StepDelayMs`.
7. Inspect the live PowerPoint window.
8. Revise the JSON and rerun until the deck is acceptable.

## Image Generation Contract

The renderer never calls image APIs directly. Agents should:

1. Produce a clear image prompt.
2. Generate or source the image with the user's chosen tool.
3. Copy the selected result into the deck project.
4. Reference the file from `assets[].path`.

This keeps the renderer portable and avoids embedding private API details in
the open-source project.

Recommended asset entry:

```json
{
  "id": "marketMapHero",
  "type": "image",
  "source": "agent-image-generation",
  "path": "assets/images/market-map-hero.png",
  "prompt": "Editorial business illustration of ..."
}
```

## Prompting Guidance For Slide Images

Prefer images that support the slide layout instead of competing with it:

- Use `no text, no logo, no watermark` unless exact text is required.
- Ask for clean negative space when text overlays are planned.
- For business decks, request restrained editorial or product imagery.
- For technical decks, request simple scientific illustration rather than dense
  poster-style infographics.
- Avoid asking image models to render small labels that should remain editable
  PowerPoint text.

## Review Loop

If the live render looks wrong:

- Change coordinates, sizes, colors, or copy in JSON first.
- Change the renderer only when the JSON needs a reusable primitive that does
  not exist yet.
- Avoid manual PowerPoint edits unless the final deliverable is intentionally a
  one-off hand-polished deck. Manual changes are not reproducible.

## When To Use A Different Tool

Use a different workflow when:

- The deck must be generated on a server without Windows PowerPoint.
- The output should be one static image per slide.
- The user explicitly needs macro-enabled behavior.
- The deck requires complex chart automation already handled by another mature
  charting pipeline.
