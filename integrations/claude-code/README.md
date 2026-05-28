# Claude Code Integration

Claude Code can use the same canonical Agent Skill included in this repository.

## Install For One Project

Copy the skill into the target project's `.claude/skills` directory:

```powershell
$repo = "C:\path\to\live-office-ppt"
$project = "C:\path\to\your-project"
New-Item -ItemType Directory -Force -Path "$project\.claude\skills" | Out-Null
Copy-Item -Recurse -Force `
  "$repo\integrations\agent-skills\office-live-ppt" `
  "$project\.claude\skills\office-live-ppt"
```

## Install For Your User

Copy the skill into your user-level Claude skills directory:

```powershell
$repo = "C:\path\to\live-office-ppt"
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.claude\skills" | Out-Null
Copy-Item -Recurse -Force `
  "$repo\integrations\agent-skills\office-live-ppt" `
  "$env:USERPROFILE\.claude\skills\office-live-ppt"
```

Restart Claude Code after installing or updating the skill.

## Example Prompt

```text
Use the office-live-ppt skill to build a native PowerPoint architecture deck.
Keep labels editable, generate any needed image assets separately, validate the
spec, and render through PowerPoint COM.
```

## Notes

- The renderer requires Windows desktop PowerPoint.
- Claude Code should generate or edit the JSON spec, then call
  `render-live-ppt.ps1`.
- The live drawing behavior happens while the PowerShell renderer runs; normal
  `.pptx` files do not redraw themselves when opened later.
