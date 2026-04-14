# Architecture

## Module overview

```
src/xlsx_strictdoc_sync/
├── __init__.py          # package version
├── cli.py               # argparse entry point – dispatches sub-commands
├── models.py            # Requirement dataclass (format-agnostic)
├── config_manager.py    # TOML config loading / validation
├── excel_engine.py      # openpyxl backend (table mode + legacy mode)
├── sdoc_engine.py       # StrictDoc backend (read / write / merge)
└── sync_engine.py       # Bidirectional merge logic (pure function)
```

## Data flow

```
reqsync.toml
     │
     ▼
 Config / SectionMapping
     │
     ├──────────────────────────────────────────────┐
     │                                              │
     ▼                                              ▼
ExcelEngine.read_requirements()          SDocEngine.read_sdoc()
     │                                              │
     ▼                                              ▼
list[Requirement]  ◄──────── Requirement ──────► list[Requirement]
     │                      (models.py)              │
     └────────────────────┬─────────────────────────┘
                          │
                          ▼
              sync_engine.compute_sync()
              ┌────────────────────────┐
              │  Applies direction +   │
              │  field_directions      │
              │  conflict_resolution   │
              └────────────────────────┘
                    │             │
                    ▼             ▼
         excel_updates      sdoc_updates
                    │             │
                    ▼             ▼
     ExcelEngine.write()   SDocEngine.write_sdoc()
```

## Key design decisions

### `Requirement` as the exchange format

The `Requirement` dataclass in `models.py` is entirely format-agnostic.
Neither `excel_engine` nor `sdoc_engine` knows about the other; they
communicate exclusively through lists of `Requirement` objects.

### Pure merge function

`sync_engine.compute_sync()` is a **pure function** – it takes two lists of
requirements and returns two lists of requirements.  It has no side effects and
makes no I/O calls.  This makes it trivially testable and easy to reason about.

### Field direction override precedence

```
CLI --direction flag
  └─► overrides sync_direction in reqsync.toml
        └─► field_directions[FIELD] overrides section-level direction
              └─► conflict_resolution is the final tiebreaker
```

### Excel table expansion

When new rows are appended to a formatted table in table-mode, `ExcelEngine`
updates the table's `ref` attribute so Excel recognises the new rows as part
of the table.  This avoids the need to delete and re-create the table.

### StrictDoc integration

`sdoc_engine` uses `SDReader.read()` and `SDWriter.write()` directly from the
`strictdoc` package.  Grammar elements are built programmatically from the
section mapping so that custom fields appear correctly in the written `.sdoc`.
