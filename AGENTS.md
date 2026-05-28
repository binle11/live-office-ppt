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

Run these after renderer, schema, example, or skill changes:

```powershell
powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\network-diagram.json `
  -ValidateOnly

powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\image-asset-demo.json `
  -ValidateOnly
```

For skill changes, validate the canonical skill with the local skill validator
when available.

## Documentation Tone

Write public docs as product documentation for users. Avoid internal warnings,
private process notes, exaggerated claims, or promises that the renderer can
replace every presentation workflow.
