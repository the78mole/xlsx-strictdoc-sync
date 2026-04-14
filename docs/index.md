# xlsx-strictdoc-sync

**Version {{VERSION}}** — Synchronize requirements between Microsoft Excel and StrictDoc (`.sdoc`) files.

[![CI](https://github.com/the78mole/xlsx-strictdoc-sync/actions/workflows/ci.yml/badge.svg)](https://github.com/the78mole/xlsx-strictdoc-sync/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/xlsx-strictdoc-sync)](https://pypi.org/project/xlsx-strictdoc-sync/)

---

## What is it?

`xlsx-strictdoc-sync` (`reqsync`) is a Python command-line tool that bridges the gap between
flexible Excel-based requirement tracking and formal [StrictDoc](https://strictdoc.readthedocs.io/)
`.sdoc` files. It supports:

- **Excel → SDoc** — migrate or update `.sdoc` files from an Excel workbook.
- **SDoc → Excel** — push changes back from SDoc into Excel.
- **Bidirectional sync** — propagate new UIDs to both sides; control which system
  is authoritative for each field via `field_directions` overrides.
- **Table mode** — target named Excel tables (ListObjects) by display name.
- **Legacy mode** — target plain worksheets with A1-style column references.
- **Custom grammars** — map any Excel column to a custom SDoc grammar field.

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

## Navigation

| Page | Description |
|------|-------------|
| [Getting Started](getting-started.md) | Installation and first sync |
| [Configuration](configuration.md) | Full `reqsync.toml` reference |
| [Bidirectional Sync](bidirectional-sync.md) | Field-level direction overrides |
| [CLI Reference](cli-reference.md) | All commands and flags |
| [Architecture](architecture.md) | Module design and data flow |
