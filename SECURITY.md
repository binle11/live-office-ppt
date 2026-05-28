# Security

Live Office PPT runs local PowerShell and automates desktop Microsoft
PowerPoint through COM. Treat deck specs and scripts as code.

## Supported Use

- Run the renderer only from repositories and JSON specs you trust.
- Review image paths before rendering specs from other people.
- Keep private deck content and generated outputs in your own project workspace.
- Use normal Git secret-scanning practices before publishing examples.

## What The Renderer Does

- Opens desktop PowerPoint on the local machine.
- Creates or edits a presentation from a JSON spec.
- Reads local image files referenced by `path` or top-level `assets`.
- Saves a `.pptx` file to the configured output path.

## What The Renderer Does Not Do

- It does not execute macros.
- It does not embed auto-run macros in generated `.pptx` files.
- It does not call image generation APIs.
- It does not require API keys.
- It does not upload files.

## Reporting Issues

If you find a security issue, open a private report through GitHub Security
Advisories when available. If advisories are not enabled, contact the repository
owner privately before opening a public issue.
