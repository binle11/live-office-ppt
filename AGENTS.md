# Agent Guidance

This repository contains a local PowerPoint renderer and agent integration
files. Keep changes reproducible and easy to audit.

## Repository Rules

- Do not commit generated `exports/*.pptx` files.
- Do not commit private deck specs, client assets, API keys, or local cache
  paths.
- Keep the renderer independent of any specific image generation API.
- Keep the canonical Agent Skill at
  `integrations/agent-skills/office-live-ppt`.
- Keep Codex and Claude Code documentation pointed at the canonical skill.
- Keep OpenClaw-specific metadata isolated under `integrations/openclaw`.

## Validation

After schema, example, or spec changes, run the pure validator. It loads
`schema/deck.schema.json`, checks canvas bounds, and runs on Linux without
PowerPoint:

```bash
python3 -m pip install -r requirements-validate.txt
python3 validate-deck.py examples/network-diagram.json
python3 validate-deck.py examples/image-asset-demo.json
python3 validate-deck.py examples/consulting-deck.json
python3 validate-deck.py examples/research-brief.json
python3 validate-deck.py examples/product-overview.json
python3 -m unittest tests/test_validation.py
```

On Windows, `-ValidateOnly` remains the renderer preflight. It checks slides
and image files and does not load the JSON schema:

```powershell
powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\network-diagram.json `
  -ValidateOnly

powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\image-asset-demo.json `
  -ValidateOnly

powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\consulting-deck.json `
  -ValidateOnly

powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\research-brief.json `
  -ValidateOnly

powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\product-overview.json `
  -ValidateOnly
```

For skill changes, validate the canonical skill with the local skill validator
when available.

## Documentation Tone

Write public docs as product documentation for users. Avoid internal warnings,
private process notes, exaggerated claims, or promises that the renderer can
replace every presentation workflow.
