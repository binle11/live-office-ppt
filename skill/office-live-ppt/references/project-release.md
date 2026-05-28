# Repository Quality Checklist

Use this checklist when preparing a reusable `live-office-ppt` repository or
refreshing public examples.

## Files That Should Exist

- `README.md`
- `LICENSE`
- `.gitignore`
- `render-live-ppt.ps1`
- `schema/deck.schema.json`
- `docs/json-format.md`
- `docs/codex-workflow.md`
- `examples/*.json`
- `skill/office-live-ppt/SKILL.md`

## Content Scope

Keep the repository focused on reusable tooling:

- renderer scripts
- documentation
- schema files
- sanitized examples
- small demo assets
- the Codex skill

Keep generated decks, private project materials, and local workspace artifacts
in the deck project that consumes this renderer.

Before a public refresh, scan the repository for local paths, temporary cache
paths, API key names, and token prefixes used on the machine where the examples
were prepared.

## Validation

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\network-diagram.json `
  -ValidateOnly

powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\image-asset-demo.json `
  -ValidateOnly
```

For a local manual smoke test, render at least one example with PowerPoint
visible.
