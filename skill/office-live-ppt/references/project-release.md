# Project Release Checklist

Use this checklist before publishing `live-office-ppt` to GitHub.

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

## Cleanup

Do not commit:

- User project decks unless explicitly intended as sanitized examples.
- Client or confidential content.
- Absolute local user-profile or drive-specific workspace paths.
- Agent-generated image cache paths.
- API keys or environment files.
- Rendered `exports/*.pptx` files unless a release artifact is intentionally
  attached outside source control.

Search before publishing for local paths, temporary cache paths, API key names,
and token prefixes used on the machine where the repository was prepared.

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
