# Configuration Reference

All configuration lives in a single `reqsync.toml` file.

## `[global]`

| Key | Required | Description |
|-----|----------|-------------|
| `excel_file` | ✅ | Path to the Excel workbook |

## `[SECTION_NAME]`

Each section maps one requirement layer (e.g. system, software, hardware) to
one `.sdoc` file and one Excel table or sheet.

| Key | Required | Default | Description |
|-----|----------|---------|-------------|
| `sdoc_file` | ✅ | — | Output `.sdoc` file path |
| `mode` | ✅ | — | `"table"` or `"legacy"` |
| `anchor` | ✅ | — | Table display-name (table mode) or sheet name (legacy) |
| `uid_col` | ✅ | — | Header name (table) or column letter (legacy) for UID |
| `title_col` | | `""` | Header/letter for TITLE |
| `statement_col` | | `""` | Header/letter for STATEMENT |
| `relations_col` | | `""` | Header/letter for parent UIDs (`;`-separated) |
| `grammar_tag` | | `"REQUIREMENT"` | SDoc grammar element tag |
| `sync_direction` | | `"excel_to_sdoc"` | Default sync direction for this section |
| `conflict_resolution` | | `"excel"` | Tiebreaker when `sync_direction = "both"` and no field override applies: `"excel"` or `"sdoc"` |
| `last_updated_col` | | `""` | Excel column header/letter that stores a per-row last-updated timestamp. When set and `sync_direction = "both"`, the side with the newer timestamp wins for fields without an explicit `field_directions` entry. Falls back to `conflict_resolution` when timestamps are equal, missing, or unparseable. |

### `[SECTION_NAME.extra_cols]`

Maps Excel column header/letter → SDoc grammar field name for any fields
beyond the standard three.

```toml
[SYS_REQS.extra_cols]
Safety_Level = "SAFETY_LEVEL"
Rationale    = "RATIONALE"
```

### `[SECTION_NAME.field_directions]`

Per-field direction overrides (only relevant when `sync_direction = "both"`).
Keys are **SDoc grammar field names**; values are `"excel_to_sdoc"` or
`"sdoc_to_excel"`.

```toml
[SYS_REQS.field_directions]
TITLE       = "sdoc_to_excel"   # SDoc is authoritative for titles
STATEMENT   = "excel_to_sdoc"   # Excel is authoritative for statements
SAFETY_LEVEL = "sdoc_to_excel"  # Safety classification managed in SDoc
```

### Timestamp-based sync direction (`last_updated_col`)

When `last_updated_col` is set and `sync_direction = "both"`, the sync engine
compares the per-row timestamps stored in that column and lets the **more
recently updated side win** for every field that has no explicit
`field_directions` entry.

**Priority order** (highest → lowest):

1. `field_directions[FIELD]` — explicit per-field override always wins.
2. `last_updated_col` timestamp comparison — newer side wins.
3. `conflict_resolution` — static tiebreaker (`"excel"` or `"sdoc"`), used
   when timestamps are equal, missing, or cannot be parsed.

**Setup:** declare the timestamp column in `extra_cols` so it is mapped to a
SDoc grammar field, then reference the Excel column header in `last_updated_col`:

```toml
[SYS_REQS]
sync_direction   = "both"
last_updated_col = "Last Updated"   # Excel column header

[SYS_REQS.extra_cols]
"Last Updated" = "LAST_UPDATED"     # Excel header → SDoc field name
```

Timestamps are compared as `datetime.fromisoformat()` values.  Plain
ISO-8601 date strings (`2024-06-15`) and full datetime strings
(`2024-06-15T10:30:00`) are both supported.  If parsing fails, a
lexicographic string comparison is used as a fallback (which still gives
correct ordering for ISO date strings).



### Table mode (`mode = "table"`)

Targets a formatted Excel table (ListObject) by its **display name** as shown
in the Excel "Table Design" ribbon tab.

- `uid_col`, `title_col`, etc. must match **column header names** exactly.
- New rows appended by `reqsync sync` are written inside the table and the
  table reference is expanded automatically.

### Legacy mode (`mode = "legacy"`)

Targets a plain worksheet by its **sheet name**.

- `uid_col`, `title_col`, etc. are **A1-style column letters** (`A`, `B`, `AA`, …).
- Row 1 is treated as the header row; data starts from row 2.

## Full example

```toml
[global]
excel_file = "requirements.xlsx"

# ── System Requirements ──────────────────────────────────────────────────
[SYS_REQS]
sdoc_file         = "docs/sys_reqs.sdoc"
mode              = "table"
anchor            = "SYS_Requirements"
uid_col           = "UID"
title_col         = "Title"
statement_col     = "Statement"
relations_col     = "Derived From"
sync_direction    = "both"
conflict_resolution = "excel"

[SYS_REQS.field_directions]
TITLE = "sdoc_to_excel"   # titles are edited in SDoc (e.g. after review)

# ── Software Requirements ────────────────────────────────────────────────
[SW_REQS]
sdoc_file         = "docs/sw_reqs.sdoc"
mode              = "table"
anchor            = "SW_Requirements"
uid_col           = "UID"
title_col         = "Title"
statement_col     = "Statement"
relations_col     = "Parent SYS"
sync_direction    = "excel_to_sdoc"

[SW_REQS.extra_cols]
Safety_Level = "SAFETY_LEVEL"

# ── Hardware Requirements (legacy sheet) ─────────────────────────────────
[HW_REQS]
sdoc_file         = "docs/hw_reqs.sdoc"
mode              = "legacy"
anchor            = "HW Requirements"
uid_col           = "A"
title_col         = "B"
statement_col     = "C"
relations_col     = "D"
sync_direction    = "excel_to_sdoc"
```
