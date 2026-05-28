# Contributing

Contributions are welcome when they keep the renderer predictable, local, and
easy to inspect.

## Development Setup

Requirements:

- Windows
- Desktop Microsoft PowerPoint for render smoke tests
- PowerShell 5.1 or newer
- Git

No package install is required for the current renderer.

## Validation

Run spec validation before opening a pull request:

```powershell
powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\network-diagram.json `
  -ValidateOnly

powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\image-asset-demo.json `
  -ValidateOnly
```

When changing render behavior, also run at least one visible PowerPoint smoke
test:

```powershell
powershell -ExecutionPolicy Bypass -File .\render-live-ppt.ps1 `
  -Spec .\examples\network-diagram.json `
  -StepDelayMs 20 `
  -CloseAfterSave `
  -Overwrite
```

## Design Principles

- Keep output editable in PowerPoint.
- Prefer conservative COM calls over version-specific Office tricks.
- Keep image generation outside the renderer.
- Add JSON schema coverage for new spec fields.
- Update examples when adding a user-facing capability.
- Keep public examples sanitized and small.

## Pull Requests

Good pull requests include:

- a clear description of the change
- updated documentation when behavior changes
- validation commands that were run
- screenshots or sample output when visual behavior changes

Do not commit generated `exports/*.pptx` files unless maintainers explicitly ask
for release artifacts.
