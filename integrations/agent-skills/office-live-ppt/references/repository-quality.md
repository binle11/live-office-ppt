# Repository Quality

Use this checklist when refreshing public examples or preparing a release.

## Content Scope

Keep the repository focused on reusable tooling:

- renderer scripts
- documentation
- schema files
- sanitized examples
- small demo assets
- agent integration files

Keep generated decks, private project materials, and local workspace artifacts
in the deck project that consumes this renderer.

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
