# Changelog

All notable changes to this project will be documented here.

This project follows a lightweight form of Keep a Changelog. Versioning starts
at `0.1.0` while the JSON spec and renderer are still evolving.

## Unreleased

- Added `validate-deck.py`, a COM-free checker for Linux and CI. It loads the deck schema and prints each failure as a field path plus a reason.
- Required schema colors to use `#RRGGBB`, and required element width and height to be greater than zero.
- The pure validator checks that positioned elements and line endpoints stay inside the page canvas, and that image elements reference a declared asset id.
- Added Agent Skill integrations for Codex and Claude Code.
- Added an OpenClaw-oriented skill adapter.
- Added repository security, contribution, and agent maintenance guidance.

## 0.1.0 - 2026-05-28

- Added PowerShell COM renderer for native PowerPoint generation.
- Added JSON deck spec support for text, shapes, connectors, images, and pauses.
- Added top-level image assets and image fit modes.
- Added JSON schema and deck authoring documentation.
- Added example network diagram and image asset demo.
