# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added
- Added an API map to `README.md` for public and advanced integration entrypoints.
- Added/expanded docstrings across core modules to improve maintainability and release readability.
- Added coverage tests for edge cases in link rewriting and conditional-formatting class application.

### Changed
- Standardized public transform API naming to `create_xlsx_transform`.
- Standardized internal `cova_` naming consistency in patch/render helpers.
- Centralized shared type aliases and typed render payloads in `src/xx2html/core/types.py`.

### Fixed
- Fixed local-sheet link rewriting for sheet names that contain dots (for example `#Q1.2026.A1`).
- Fixed local-sheet link rewriting for Excel-style internal anchors (for example `#'My Sheet'!A1`).
- Fixed unresolved local-sheet mappings to preserve original anchors instead of generating empty `href` values.
- Fixed conditional-formatting class merging to avoid duplicate classes and keep deterministic class ordering.
