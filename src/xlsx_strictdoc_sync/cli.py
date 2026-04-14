"""Command-line interface for xlsx-strictdoc-sync (``reqsync`` command).

Sub-commands
------------
``sync``
    Synchronize requirements between an Excel workbook and ``.sdoc`` files.
    Supports three directions (``--direction`` flag):

    * ``excel_to_sdoc`` *(default)* – Excel → SDoc only.
    * ``sdoc_to_excel`` – SDoc → Excel only.
    * ``both`` – Bidirectional: new UIDs propagate to both sides; existing
      UIDs are merged field-by-field according to
      ``[SECTION.field_directions]`` in the config.

``init-config``
    Scan an Excel file and generate a ``reqsync.toml`` template.

``generate-grammar``
    Read a ``reqsync.toml`` and write ``.sdoc`` grammar files for each
    section.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .config_manager import VALID_DIRECTIONS, generate_config_template, load_config
from .excel_engine import ExcelEngine, ExcelEngineError
from .models import Requirement
from .sdoc_engine import (
    SDocEngineError,
    document_to_requirements,
    generate_grammar_sdoc,
    read_sdoc,
    requirements_to_document,
    update_document,
    write_sdoc,
)
from .sync_engine import SyncResult, compute_sync


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    """Parse CLI arguments and dispatch to the appropriate sub-command.

    Returns:
        Exit code (``0`` = success, ``1`` = error).
    """
    parser = _build_parser()
    args = parser.parse_args(argv)

    if not hasattr(args, "func"):
        parser.print_help()
        return 0

    try:
        return args.func(args)
    except (FileNotFoundError, ExcelEngineError, SDocEngineError, KeyError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


# ---------------------------------------------------------------------------
# Argument parser
# ---------------------------------------------------------------------------


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="reqsync",
        description="Synchronize requirements between Excel and StrictDoc (.sdoc) files.",
    )
    parser.add_argument("--version", action="version", version=f"reqsync {__version__}")

    sub = parser.add_subparsers(title="commands", metavar="<command>")

    # ---- sync ----
    sync_p = sub.add_parser(
        "sync",
        help="Synchronize requirements (Excel <-> SDoc).",
        description=(
            "Read requirements from Excel and/or SDoc and synchronize them "
            "according to the configured direction and field-level overrides."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
direction values:
  excel_to_sdoc  Excel is the source of truth (default when not in config).
  sdoc_to_excel  SDoc is the source of truth.
  both           Bidirectional: new UIDs propagate to both sides; field-level
                 overrides in [SECTION.field_directions] control per-field
                 authority.  Use --prefer-sdoc to make SDoc win for fields
                 with no specific override.

examples:
  reqsync sync                          # uses reqsync.toml, section default directions
  reqsync sync --direction both         # force bidirectional for all sections
  reqsync sync --direction both --prefer-sdoc
  reqsync sync --section SYS_REQS --dry-run
""",
    )
    sync_p.add_argument(
        "config",
        nargs="?",
        default="reqsync.toml",
        help="Path to reqsync.toml (default: %(default)s).",
    )
    sync_p.add_argument(
        "--section",
        dest="sections",
        metavar="NAME",
        action="append",
        default=None,
        help="Sync only this section (repeatable). Defaults to all sections.",
    )
    sync_p.add_argument(
        "--direction",
        choices=sorted(VALID_DIRECTIONS),
        default=None,
        metavar="DIR",
        help=(
            "Override sync direction for all selected sections "
            "(both | excel_to_sdoc | sdoc_to_excel)."
        ),
    )
    sync_p.add_argument(
        "--prefer-sdoc",
        action="store_true",
        default=False,
        help=(
            "When --direction both is used and a field has no specific "
            "field_directions override, prefer the SDoc value over Excel "
            "(default: Excel wins)."
        ),
    )
    sync_p.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be written without modifying any files.",
    )
    sync_p.set_defaults(func=cmd_sync)

    # ---- init-config ----
    init_p = sub.add_parser(
        "init-config",
        help="Scan an Excel file and generate a reqsync.toml template.",
        description=(
            "Inspects an Excel workbook and writes a reqsync.toml whose sections "
            "correspond to the detected sheets / tables."
        ),
    )
    init_p.add_argument("excel_file", help="Path to the Excel workbook to inspect.")
    init_p.add_argument(
        "-o",
        "--output",
        default="reqsync.toml",
        help="Output path for the generated config (default: %(default)s).",
    )
    init_p.set_defaults(func=cmd_init_config)

    # ---- generate-grammar ----
    grammar_p = sub.add_parser(
        "generate-grammar",
        help="Generate .sdoc grammar files from a reqsync.toml.",
        description="Create empty .sdoc files containing GRAMMAR definitions.",
    )
    grammar_p.add_argument(
        "config",
        nargs="?",
        default="reqsync.toml",
        help="Path to reqsync.toml (default: %(default)s).",
    )
    grammar_p.add_argument(
        "--output-dir",
        default=".",
        help="Directory where grammar files are written (default: %(default)s).",
    )
    grammar_p.set_defaults(func=cmd_generate_grammar)

    return parser


# ---------------------------------------------------------------------------
# Sub-command implementations
# ---------------------------------------------------------------------------


def cmd_sync(args: argparse.Namespace) -> int:
    """Synchronize requirements using the configured or overridden direction."""
    config = load_config(args.config)

    sections = config.sections
    if args.sections:
        requested = set(args.sections)
        sections = [s for s in sections if s.name in requested]
        if not sections:
            print(
                f"No sections matched: {args.sections}. "
                f"Available: {[s.name for s in config.sections]}",
                file=sys.stderr,
            )
            return 1

    results: list[SyncResult] = []

    with ExcelEngine(config.excel_file) as eng:
        for mapping in sections:
            # Apply --prefer-sdoc at runtime by patching conflict_resolution
            if args.prefer_sdoc:
                from dataclasses import replace
                mapping = replace(mapping, conflict_resolution="sdoc")

            eff_dir = args.direction or mapping.sync_direction
            print(f"  [{mapping.name}] Syncing (direction={eff_dir}) ...")

            # --- Read current state ---
            excel_reqs: list[Requirement] = eng.read_requirements(mapping)
            sdoc_reqs: list[Requirement] = []
            sdoc_path = Path(mapping.sdoc_file)
            if sdoc_path.exists():
                doc = read_sdoc(sdoc_path)
                sdoc_reqs = document_to_requirements(doc, mapping.grammar_tag)

            # --- Compute desired post-sync state ---
            excel_updates, sdoc_updates, sync_result = compute_sync(
                excel_reqs=excel_reqs,
                sdoc_reqs=sdoc_reqs,
                mapping=mapping,
                direction_override=args.direction,
            )

            if args.dry_run:
                _print_dry_run(mapping, excel_updates, sdoc_updates)
            else:
                # --- Write SDoc ---
                if sdoc_updates:
                    if sdoc_path.exists():
                        doc = read_sdoc(sdoc_path)
                        doc = update_document(doc, sdoc_updates, mapping)
                    else:
                        doc = requirements_to_document(
                            sdoc_updates,
                            title=mapping.name.replace("_", " ").title(),
                            mapping=mapping,
                        )
                    write_sdoc(doc, sdoc_path)

                # --- Write Excel ---
                if excel_updates:
                    eng.write_requirements(excel_updates, mapping)

            results.append(sync_result)

        # Save Excel once after all sections if any Excel changes were made
        if not args.dry_run and any(r.excel_added + r.excel_updated > 0 for r in results):
            eng.save()

    # Print summary
    for r in results:
        print(f"  {r.summary()}")

    return 0


def cmd_init_config(args: argparse.Namespace) -> int:
    """Scan an Excel file and write a reqsync.toml template."""
    excel_path = Path(args.excel_file)
    output_path = Path(args.output)

    with ExcelEngine(excel_path) as eng:
        tables = eng.list_tables()
        sheets = eng.list_sheets()

    section_descriptors: list[dict[str, object]] = []

    if tables:
        for table_name in tables:
            safe_name = _make_section_name(table_name)
            section_descriptors.append(
                {
                    "name": safe_name,
                    "sdoc_file": f"docs/{safe_name.lower()}.sdoc",
                    "mode": "table",
                    "anchor": table_name,
                    "uid_col": "UID",
                    "title_col": "Title",
                    "statement_col": "Statement",
                }
            )
    else:
        for sheet_name in sheets:
            safe_name = _make_section_name(sheet_name)
            section_descriptors.append(
                {
                    "name": safe_name,
                    "sdoc_file": f"docs/{safe_name.lower()}.sdoc",
                    "mode": "legacy",
                    "anchor": sheet_name,
                    "uid_col": "A",
                    "title_col": "B",
                    "statement_col": "C",
                }
            )

    toml_content = generate_config_template(
        excel_file=str(excel_path),
        sections=section_descriptors,
    )

    if output_path.exists():
        print(f"Warning: '{output_path}' already exists - overwriting.", file=sys.stderr)

    output_path.write_text(toml_content, encoding="utf-8")
    print(f"Config written to: {output_path}")
    print(f"Detected {len(section_descriptors)} section(s).")
    return 0


def cmd_generate_grammar(args: argparse.Namespace) -> int:
    """Generate .sdoc grammar files for all configured sections."""
    config = load_config(args.config)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    for mapping in config.sections:
        title = f"{mapping.name.replace('_', ' ').title()} Grammar"
        content = generate_grammar_sdoc(title=title, mapping=mapping)
        out_file = output_dir / f"{mapping.name.lower()}_grammar.sdoc"
        out_file.write_text(content, encoding="utf-8")
        print(f"  [{mapping.name}] Grammar written -> {out_file}")

    return 0


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _print_dry_run(
    mapping: object,
    excel_updates: list[Requirement],
    sdoc_updates: list[Requirement],
) -> None:
    from strictdoc.backend.sdoc.writer import SDWriter
    from strictdoc.core.project_config import ProjectConfig

    name = getattr(mapping, "name", "?")
    if sdoc_updates:
        print(f"  [{name}] [dry-run] Would write {len(sdoc_updates)} req(s) to SDoc:")
        doc = requirements_to_document(
            sdoc_updates,
            title=f"[dry-run] {name}",
            mapping=mapping,
        )
        print(SDWriter(ProjectConfig()).write(doc))
    if excel_updates:
        print(
            f"  [{name}] [dry-run] Would write {len(excel_updates)} req(s) to Excel "
            f"(UIDs: {[r.uid for r in excel_updates]})."
        )
    if not sdoc_updates and not excel_updates:
        print(f"  [{name}] [dry-run] Nothing to sync.")


def _make_section_name(raw: str) -> str:
    """Convert an arbitrary string to a safe TOML section name."""
    cleaned = "".join(c if c.isalnum() or c in "_-" else "_" for c in raw)
    return cleaned.upper().strip("_") or "SECTION"
