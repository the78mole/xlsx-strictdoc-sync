---
home: true
title: xlsx-strictdoc-sync
heroText: xlsx-strictdoc-sync
tagline: Synchronize requirements between Microsoft Excel and StrictDoc (.sdoc) files — with bidirectional sync and per-field direction overrides.
actions:
  - text: Get Started
    link: /getting-started
    type: primary
  - text: CLI Reference
    link: /cli-reference
    type: secondary
features:
  - title: Bidirectional Sync
    details: Propagate new UIDs in both directions. Control which side is authoritative per-field via field_directions overrides.
  - title: Excel Table & Legacy Mode
    details: Target named Excel ListObjects by display name (table mode) or plain worksheets with A1 column references (legacy mode).
  - title: Custom Grammars
    details: Map any Excel column to a StrictDoc grammar field — go beyond the standard UID / TITLE / STATEMENT fields.
  - title: Non-destructive & Idempotent
    details: The sync never writes back unless the merged value actually differs. Repeated runs produce no spurious changes.
  - title: Auto-config Generation
    details: Run `reqsync init-config requirements.xlsx` to generate a reqsync.toml template from an existing workbook.
  - title: Installable via uv
    details: Ship as a proper Python package — install with `uv tool install xlsx-strictdoc-sync` and use from any project.
footer: MIT License | Copyright © the78mole
---

## Quick start

```bash
# Install
pip install xlsx-strictdoc-sync

# Generate a config template from an existing Excel workbook
reqsync init-config requirements.xlsx

# Sync Excel → SDoc (default)
reqsync sync

# Bidirectional sync with SDoc winning title changes
reqsync sync --direction both --prefer-sdoc
```

## Bidirectional sync with field exceptions

```toml
[SYS_REQS]
sync_direction      = "both"
conflict_resolution = "excel"      # Excel wins for unspecified fields

[SYS_REQS.field_directions]
TITLE     = "sdoc_to_excel"        # SDoc owns the approved title
STATEMENT = "excel_to_sdoc"        # Engineers write statements in Excel
```

## Pages

| Page | Description |
|------|-------------|
| [Getting Started](/xlsx-strictdoc-sync/getting-started) | Installation and first sync |
| [Configuration](/xlsx-strictdoc-sync/configuration) | Full `reqsync.toml` reference |
| [Bidirectional Sync](/xlsx-strictdoc-sync/bidirectional-sync) | Field-level direction overrides |
| [CLI Reference](/xlsx-strictdoc-sync/cli-reference) | All commands and flags |
| [Architecture](/xlsx-strictdoc-sync/architecture) | Module design and data flow |
