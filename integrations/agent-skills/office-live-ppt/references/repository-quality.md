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

Run the pure validator. It loads the schema and checks the public examples
without PowerPoint:

```bash
python3 -m pip install -r requirements-validate.txt
python3 validate-deck.py examples/network-diagram.json
python3 validate-deck.py examples/image-asset-demo.json
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
```

For a local manual smoke test, render at least one example with PowerPoint
visible.
