# Codex Integration

Live Office PPT ships an Agent Skill that Codex can use directly.

## Install

Copy the canonical skill folder into your Codex skills directory:

```powershell
$repo = "C:\path\to\live-office-ppt"
$codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { "$env:USERPROFILE\.codex" }
New-Item -ItemType Directory -Force -Path "$codexHome\skills" | Out-Null
Copy-Item -Recurse -Force `
  "$repo\integrations\agent-skills\office-live-ppt" `
  "$codexHome\skills\office-live-ppt"
```

Restart Codex after installing or updating the skill.

## Example Prompt

```text
Use office-live-ppt to create a 6-slide technical project update. Generate any
needed bitmap illustrations first, save them as deck assets, validate the JSON,
then render with PowerPoint visible.
```

## Notes

- Codex should keep project deck specs outside this reusable repository.
- Generated images used by a deck should be copied into the deck project's
  `assets/images/` directory before rendering.
- Use `-CloseAfterSave` for smoke tests and omit it when the user wants to
  inspect PowerPoint live.
