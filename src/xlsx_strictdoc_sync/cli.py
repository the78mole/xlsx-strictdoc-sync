"""Command-line interface for xlsx-strictdoc-sync (``reqsync`` command).

Sub-commands
------------
``sync``
    Synchronize requirements between an Excel workbook and ``.sdoc`` files.
    By default the Excel workbook is the authoritative source (Excel → SDoc).

``init-config``
    Scan an Excel file and generate a ``reqsync.toml`` template that guesses
    the section mappings from sheet/table names and header rows.

``generate-grammar``
    Read a ``reqsync.toml`` and write standalone ``.sdoc`` grammar files for
    each section.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .config_manager import generate_config_template, load_config
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
    parser.add_argument(
        "--version", action="version", version=f"reqsync {__version__}"
    )

    sub = parser.add_subparsers(title="commands", metavar="<command>")

    # ---- sync ----
    sync_p = sub.add_parser(
        "sync",
        help="Synchronize requirements (Excel → SDoc by default).",
        description="Read requirements from Excel and write them to .sdoc files.",
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
    """Synchronize requirements from Excel into .sdoc files."""
    config = load_config(args.config)
    excel_path = Path(config.excel_file)

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

    with ExcelEngine(excel_path) as eng:
        for mapping in sections:
            print(f"  [{mapping.name}] Reading from Excel …")
            requirements: list[Requirement] = eng.read_requirements(mapping)
            print(f"  [{mapping.name}] Found {len(requirements)} requirement(s).")

            sdoc_path = Path(mapping.sdoc_file)
            if sdoc_path.exists():
                print(f"  [{mapping.name}] Updating existing {sdoc_path} …")
                doc = read_sdoc(sdoc_path)
                doc = update_document(doc, requirements, mapping)
            else:
                print(f"  [{mapping.name}] Creating new {sdoc_path} …")
                doc = requirements_to_document(
                    requirements,
                    title=mapping.name.replace("_", " ").title(),
                    mapping=mapping,
                )

            if args.dry_run:
                from strictdoc.backend.sdoc.writer import SDWriter
                from strictdoc.core.project_config import ProjectConfig

                print(f"  [{mapping.name}] [dry-run] Output for {sdoc_path}:")
                writer = SDWriter(ProjectConfig())
                print(writer.write(doc))
            else:
                write_sdoc(doc, sdoc_path)
                print(f"  [{mapping.name}] Written → {sdoc_path}")

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
        # Prefer tables over plain sheets when both exist
        for table_name, sheet_name in tables.items():
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
        print(
            f"Warning: '{output_path}' already exists – overwriting.",
            file=sys.stderr,
        )

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
        print(f"  [{mapping.name}] Grammar written → {out_file}")

    return 0


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_section_name(raw: str) -> str:
    """Convert an arbitrary string to a safe TOML section name."""
    cleaned = "".join(c if c.isalnum() or c in "_-" else "_" for c in raw)
    return cleaned.upper().strip("_") or "SECTION"
