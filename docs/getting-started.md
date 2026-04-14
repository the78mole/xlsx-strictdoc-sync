# Getting Started

## Prerequisites

- Python 3.11 or 3.12
- `uv` (recommended) or `pip`

## Installation

=== "uv (recommended)"

    ```bash
    uv tool install xlsx-strictdoc-sync
    ```

=== "pip"

    ```bash
    pip install xlsx-strictdoc-sync
    ```

## First sync in 3 steps

### 1. Generate a config template

```bash
reqsync init-config requirements.xlsx
```

This scans your workbook and writes `reqsync.toml` with one section per table
or sheet it detects.

### 2. Review and edit `reqsync.toml`

```toml
[global]
excel_file = "requirements.xlsx"

[SYS_REQS]
sdoc_file         = "docs/sys_reqs.sdoc"
mode              = "table"          # use the Excel formatted table
anchor            = "SYS_Reqs"       # display name of the table
uid_col           = "UID"
title_col         = "Title"
statement_col     = "Statement"
relations_col     = "Parent Req"
sync_direction    = "excel_to_sdoc"  # default: Excel is authoritative
```

### 3. Run the sync

```bash
reqsync sync
```

The `.sdoc` file is created (or updated) at the path you specified.

## Verifying the output

```bash
cat docs/sys_reqs.sdoc
```

You should see a document with a `[GRAMMAR]` block and one `[REQUIREMENT]`
node per Excel row.
