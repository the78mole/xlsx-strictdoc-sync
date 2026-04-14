# CLI Reference

## `reqsync sync`

Synchronize requirements between an Excel workbook and `.sdoc` files.

```
reqsync sync [OPTIONS] [CONFIG]
```

| Argument / Option | Default | Description |
|-------------------|---------|-------------|
| `CONFIG` | `reqsync.toml` | Path to the config file |
| `--section NAME` | all | Sync only the named section (repeatable) |
| `--direction DIR` | from config | Override direction for all sections: `excel_to_sdoc`, `sdoc_to_excel`, or `both` |
| `--prefer-sdoc` | false | When `--direction both`, SDoc wins for fields without a `field_directions` override (default: Excel wins) |
| `--dry-run` | false | Print what would change without modifying files |

### Examples

```bash
# Default: sync all sections using per-section directions from reqsync.toml
reqsync sync

# Use a different config file
reqsync sync path/to/my.toml

# Only sync the SYS_REQS section
reqsync sync --section SYS_REQS

# Force bidirectional for all sections
reqsync sync --direction both

# Preview bidirectional changes without writing anything
reqsync sync --direction both --dry-run

# Bidirectional with SDoc winning unspecified fields
reqsync sync --direction both --prefer-sdoc
```

---

## `reqsync init-config`

Scan an Excel file and generate a `reqsync.toml` template.

```
reqsync init-config EXCEL_FILE [-o OUTPUT]
```

| Argument / Option | Default | Description |
|-------------------|---------|-------------|
| `EXCEL_FILE` | — | Path to the Excel workbook to inspect |
| `-o OUTPUT` | `reqsync.toml` | Output path for the generated config |

The command detects named tables (preferred) or plain sheets and creates one
section per table/sheet with guessed column mappings and commented-out
`sync_direction` and `field_directions` blocks.

---

## `reqsync generate-grammar`

Generate `.sdoc` grammar files from a `reqsync.toml`.

```
reqsync generate-grammar [CONFIG] [--output-dir DIR]
```

| Argument / Option | Default | Description |
|-------------------|---------|-------------|
| `CONFIG` | `reqsync.toml` | Path to the config file |
| `--output-dir DIR` | `.` | Directory where grammar files are written |

One grammar file `<section_name>_grammar.sdoc` is written per section.

---

## `reqsync --version`

Print the installed version and exit.
