# Contributing

Contributions are welcome when they keep the renderer predictable, local, and
easy to inspect.

## Development Setup

Requirements:

- Windows with desktop Microsoft PowerPoint for render smoke tests
- PowerShell 5.1 or newer for the renderer
- Python 3 for COM-free spec validation
- Git

Install the validation dependency before running the pure checker:

```bash
python3 -m pip install -r requirements-validate.txt
```

No package install is required for the PowerPoint renderer itself.

## Validation

Run the pure validator before opening a pull request. It is the schema and
canvas check, and it runs without PowerPoint:

```bash
python3 validate-deck.py examples/network-diagram.json
python3 validate-deck.py examples/image-asset-demo.json
python3 validate-deck.py examples/consulting-deck.json
python3 validate-deck.py examples/research-brief.json
python3 validate-deck.py examples/product-overview.json
python3 -m unittest tests/test_validation.py
```

On Windows, the renderer preflight is still available. It does not load the
JSON schema:

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
