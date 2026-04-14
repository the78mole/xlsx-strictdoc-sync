# xlsx-strictdoc-sync

[![CI](https://github.com/the78mole/xlsx-strictdoc-sync/actions/workflows/ci.yml/badge.svg)](https://github.com/the78mole/xlsx-strictdoc-sync/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/xlsx-strictdoc-sync)](https://pypi.org/project/xlsx-strictdoc-sync/)
[![Docs](https://img.shields.io/badge/docs-GitHub%20Pages-blue)](https://the78mole.github.io/xlsx-strictdoc-sync/)
[![Python](https://img.shields.io/pypi/pyversions/xlsx-strictdoc-sync)](https://pypi.org/project/xlsx-strictdoc-sync/)

A tool to synchronize requirements between Microsoft Excel and StrictDoc (`.sdoc`) files —
mainly for migrating from Excel to Git-based StrictDoc requirements management.

## Features

- **Excel → SDoc** and **SDoc → Excel** (bidirectional)
- **Bidirectional field-level control** — decide per-field which side is authoritative
- **Table mode** — targets named Excel ListObjects (formatted tables)
- **Legacy mode** — targets plain sheets with A1 column references
- **Custom grammars** — map any Excel column to a StrictDoc grammar field
- **`init-config`** — auto-generate `reqsync.toml` from a workbook
- **`generate-grammar`** — create `.sdoc` grammar files from config

## Quick start

```bash
pip install xlsx-strictdoc-sync

# Generate config from an Excel file
reqsync init-config requirements.xlsx

# Sync Excel → SDoc
reqsync sync

# Bidirectional sync
reqsync sync --direction both

# Preview without writing
reqsync sync --direction both --dry-run
```

## Bidirectional sync with field exceptions

```toml
[SYS_REQS]
sync_direction    = "both"
conflict_resolution = "excel"      # Excel wins for unspecified fields

[SYS_REQS.field_directions]
TITLE     = "sdoc_to_excel"        # SDoc owns the approved title
STATEMENT = "excel_to_sdoc"        # Engineers write statements in Excel
```

## Documentation

Full documentation: **https://the78mole.github.io/xlsx-strictdoc-sync/**

## Development

```bash
git clone https://github.com/the78mole/xlsx-strictdoc-sync
cd xlsx-strictdoc-sync
pip install -e ".[dev]"
pytest
```

## License

MIT
