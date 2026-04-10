# AGENTS.md

## Project Status

This repository is a greenfield library. Assume there are no backwards-compatibility requirements unless the user explicitly asks for them.
Assume persisted data, snapshots, and serialized engine state are produced by the current version of the code unless the user explicitly says otherwise.

## Change Guidelines

- Refactors may change APIs, data shapes, file layouts, and behavior when that improves the design.
- Update in-repo callers, tests, and docs in the same change instead of adding compatibility layers.
- Fix the current code path rather than adding compatibility for older APIs, snapshots, serialized data, exports, or previous behavior unless explicitly requested.
- Do not add deprecation shims, migration wrappers, versioned fallbacks, or legacy-format handling unless they are explicitly requested.
