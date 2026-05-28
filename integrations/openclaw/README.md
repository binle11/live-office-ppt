# OpenClaw Integration

OpenClaw can use the dedicated skill folder in this directory. It contains the
same workflow as the canonical Agent Skill plus OpenClaw-oriented runtime
metadata.

## Install

Copy `integrations/openclaw/office-live-ppt` into your OpenClaw skills location
according to your OpenClaw setup.

```powershell
$repo = "C:\path\to\live-office-ppt"
$target = "C:\path\to\openclaw\skills"
New-Item -ItemType Directory -Force -Path $target | Out-Null
Copy-Item -Recurse -Force `
  "$repo\integrations\openclaw\office-live-ppt" `
  "$target\office-live-ppt"
```

## Example Prompt

```text
Use office-live-ppt to create a native PowerPoint deck. Validate the JSON spec,
then run the PowerShell renderer with desktop PowerPoint visible.
```

## Notes

- The renderer depends on `powershell` and desktop Microsoft PowerPoint.
- OpenClaw should not run this workflow in a headless Linux container.
- Keep the OpenClaw skill behavior aligned with
  `integrations/agent-skills/office-live-ppt`.
