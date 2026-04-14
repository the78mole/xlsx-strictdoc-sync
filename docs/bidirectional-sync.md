# Bidirectional Sync

`reqsync` supports full bidirectional synchronization between Excel and SDoc.
This page explains exactly how it works and how to configure fine-grained
control over which system is authoritative for each field.

## Overview

```
┌──────────┐  new UIDs →   ┌──────────┐
│  Excel   │ ◄──────────── │   SDoc   │
│ workbook │ ──────────── ►│   file   │
└──────────┘  ← new UIDs   └──────────┘
          field-level control
```

When `sync_direction = "both"` the sync engine:

1. **Reads** both sources.
2. **Propagates new UIDs** – a UID found only in Excel is added to SDoc and
   vice-versa.
3. **Merges existing UIDs** – for requirements that already exist on both
   sides, each field is resolved according to the rules below.
4. **Writes** the merged state back to each target.

## Field resolution rules

For each field (TITLE, STATEMENT, RELATIONS, custom fields):

| Condition | Winner |
|-----------|--------|
| `field_directions[FIELD] = "excel_to_sdoc"` | Excel value → SDoc; Excel unchanged |
| `field_directions[FIELD] = "sdoc_to_excel"` | SDoc value → Excel; SDoc unchanged |
| No field override, `conflict_resolution = "excel"` (default) | Excel wins both sides |
| No field override, `conflict_resolution = "sdoc"` | SDoc wins both sides |

!!! tip "No unnecessary writes"
    If the merged value equals what a target already has, no write is
    performed for that requirement.  The sync is idempotent.

## CLI overrides

You can override the direction and conflict resolution at runtime:

```bash
# Force bidirectional for all sections (ignores per-section sync_direction)
reqsync sync --direction both

# SDoc wins for unspecified fields
reqsync sync --direction both --prefer-sdoc

# Inspect without writing anything
reqsync sync --direction both --dry-run
```

## Worked example

### Config

```toml
[SYS_REQS]
sdoc_file      = "docs/sys_reqs.sdoc"
mode           = "table"
anchor         = "SYS_Requirements"
uid_col        = "UID"
title_col      = "Title"
statement_col  = "Statement"
sync_direction = "both"
conflict_resolution = "excel"   # Excel wins for fields without overrides

[SYS_REQS.field_directions]
TITLE = "sdoc_to_excel"    # Titles are refined in SDoc by the reviewer
```

### Scenario

| UID     | Excel title       | SDoc title       | Excel stmt          | SDoc stmt |
|---------|-------------------|------------------|---------------------|-----------|
| SYS-001 | "Boot"            | "System Boot"    | "Shall boot in 5s." | (same)    |
| SYS-002 | (new, not in SDoc)| —                | "Shall log errors." | —         |
| SYS-003 | —                 | (new, not Excel) | —                  | "Auth."   |

### Result after sync

- **SYS-001** in SDoc: title = `"System Boot"` (SDoc wins for TITLE), stmt = `"Shall boot in 5s."` (Excel wins, no change).
- **SYS-001** in Excel: title = `"System Boot"` (pushed back from SDoc).
- **SYS-002** added to SDoc with title/stmt from Excel.
- **SYS-003** added to Excel with title/stmt from SDoc.

## Conflict resolution summary

```
For field F in a matched pair (same UID exists in both Excel and SDoc):

if field_directions[F] == "excel_to_sdoc":
    sdoc[F] = excel[F]    # Excel value overwrites SDoc
    excel[F] = excel[F]   # Excel unchanged

elif field_directions[F] == "sdoc_to_excel":
    excel[F] = sdoc[F]    # SDoc value overwrites Excel
    sdoc[F]  = sdoc[F]    # SDoc unchanged

else:  # no override
    winner = excel if conflict_resolution == "excel" else sdoc
    sdoc[F]  = winner[F]
    excel[F] = winner[F]
```
