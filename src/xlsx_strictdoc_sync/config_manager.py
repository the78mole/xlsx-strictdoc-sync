"""Configuration management: load and validate ``reqsync.toml``."""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path

if sys.version_info >= (3, 11):
    import tomllib
else:
    try:
        import tomllib  # type: ignore[no-redef]
    except ImportError:
        import tomli as tomllib  # type: ignore[no-redef]

# Valid direction string constants
DIRECTION_EXCEL_TO_SDOC = "excel_to_sdoc"
DIRECTION_SDOC_TO_EXCEL = "sdoc_to_excel"
DIRECTION_BOTH = "both"
VALID_DIRECTIONS = frozenset([DIRECTION_EXCEL_TO_SDOC, DIRECTION_SDOC_TO_EXCEL, DIRECTION_BOTH])

CONFLICT_EXCEL = "excel"
CONFLICT_SDOC = "sdoc"
VALID_CONFLICT_RESOLUTIONS = frozenset([CONFLICT_EXCEL, CONFLICT_SDOC])


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------


@dataclass
class SectionMapping:
    """Configuration for one requirement section (e.g. SYS_REQS, SW_REQS).

    Attributes:
        name: Section name as used in the TOML file.
        sdoc_file: Path to the corresponding ``.sdoc`` file.
        mode: Access mode – ``"table"`` (Excel ListObject) or
            ``"legacy"`` (sheet + A1 column references).
        anchor: Table display-name (table mode) or sheet name (legacy mode).
        uid_col: Column header (table) or letter (legacy) for the UID field.
        title_col: Column header / letter for the TITLE field.
        statement_col: Column header / letter for the STATEMENT field.
        relations_col: Optional column for parent UID references.
            Multiple values in a single cell must be separated by ``;``.
        grammar_tag: SDoc grammar element tag (default ``"REQUIREMENT"``).
        extra_cols: Mapping of Excel column header/letter → SDoc field name
            for fields beyond the standard three.
        sync_direction: Overall sync direction for this section.

            * ``"excel_to_sdoc"`` *(default)* – Excel is the source of truth.
            * ``"sdoc_to_excel"`` – SDoc is the source of truth.
            * ``"both"`` – Bidirectional: new items flow from each side; for
              fields that exist on both sides the winner is determined by
              :attr:`conflict_resolution` unless overridden in
              :attr:`field_directions`.

        conflict_resolution: Tiebreaker when ``sync_direction="both"`` and a
            field has no entry in :attr:`field_directions`.

            * ``"excel"`` *(default)* – Excel value wins.
            * ``"sdoc"`` – SDoc value wins.

        field_directions: Per-field direction overrides (only meaningful when
            ``sync_direction="both"``).  Keys are **SDoc grammar field names**
            (e.g. ``"STATEMENT"``, ``"TITLE"``, ``"SAFETY_LEVEL"``).  Values
            are ``"excel_to_sdoc"`` or ``"sdoc_to_excel"``.

            Example in TOML::

                [SYS_REQS.field_directions]
                TITLE     = "sdoc_to_excel"    # SDoc is authoritative for titles
                STATEMENT = "excel_to_sdoc"    # Excel is authoritative for text
    """

    name: str
    sdoc_file: str
    mode: str
    anchor: str
    uid_col: str
    title_col: str = ""
    statement_col: str = ""
    relations_col: str = ""
    grammar_tag: str = "REQUIREMENT"
    extra_cols: dict[str, str] = field(default_factory=dict)
    sync_direction: str = DIRECTION_EXCEL_TO_SDOC
    conflict_resolution: str = CONFLICT_EXCEL
    field_directions: dict[str, str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if self.mode not in ("table", "legacy"):
            raise ValueError(
                f"[{self.name}] 'mode' must be 'table' or 'legacy', got '{self.mode}'."
            )
        if not self.anchor:
            raise ValueError(f"[{self.name}] 'anchor' must not be empty.")
        if not self.uid_col:
            raise ValueError(f"[{self.name}] 'uid_col' must not be empty.")
        if self.sync_direction not in VALID_DIRECTIONS:
            raise ValueError(
                f"[{self.name}] 'sync_direction' must be one of "
                f"{sorted(VALID_DIRECTIONS)}, got '{self.sync_direction}'."
            )
        if self.conflict_resolution not in VALID_CONFLICT_RESOLUTIONS:
            raise ValueError(
                f"[{self.name}] 'conflict_resolution' must be 'excel' or 'sdoc', "
                f"got '{self.conflict_resolution}'."
            )
        for fname, fdir in self.field_directions.items():
            if fdir not in (DIRECTION_EXCEL_TO_SDOC, DIRECTION_SDOC_TO_EXCEL):
                raise ValueError(
                    f"[{self.name}] field_directions['{fname}'] must be "
                    f"'excel_to_sdoc' or 'sdoc_to_excel', got '{fdir}'."
                )


@dataclass
class Config:
    """Top-level configuration loaded from ``reqsync.toml``.

    Attributes:
        excel_file: Path to the Excel workbook.
        sections: Ordered list of requirement sections to synchronize.
    """

    excel_file: str
    sections: list[SectionMapping]

    @property
    def excel_path(self) -> Path:
        """Return the Excel path as a :class:`~pathlib.Path` object."""
        return Path(self.excel_file)


# ---------------------------------------------------------------------------
# Loader
# ---------------------------------------------------------------------------


def load_config(path: str | Path) -> Config:
    """Load and validate a ``reqsync.toml`` file.

    Args:
        path: Filesystem path to the TOML configuration file.

    Returns:
        A validated :class:`Config` instance.

    Raises:
        FileNotFoundError: If the configuration file does not exist.
        KeyError: If a required key is missing from the file.
        ValueError: If a value fails semantic validation.
    """
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    with config_path.open("rb") as fh:
        data = tomllib.load(fh)

    global_section = data.get("global", {})
    excel_file: str = global_section.get("excel_file", "")
    if not excel_file:
        raise KeyError("Missing required key 'excel_file' under [global].")

    sections: list[SectionMapping] = []
    reserved = {"global"}
    for section_name, section_data in data.items():
        if section_name in reserved:
            continue
        if not isinstance(section_data, dict):
            continue
        try:
            extra_cols: dict[str, str] = section_data.get("extra_cols", {})
            field_directions: dict[str, str] = section_data.get("field_directions", {})
            mapping = SectionMapping(
                name=section_name,
                sdoc_file=section_data["sdoc_file"],
                mode=section_data["mode"],
                anchor=section_data["anchor"],
                uid_col=section_data["uid_col"],
                title_col=section_data.get("title_col", ""),
                statement_col=section_data.get("statement_col", ""),
                relations_col=section_data.get("relations_col", ""),
                grammar_tag=section_data.get("grammar_tag", "REQUIREMENT"),
                extra_cols=extra_cols,
                sync_direction=section_data.get("sync_direction", DIRECTION_EXCEL_TO_SDOC),
                conflict_resolution=section_data.get("conflict_resolution", CONFLICT_EXCEL),
                field_directions=field_directions,
            )
        except KeyError as exc:
            raise KeyError(f"[{section_name}] Missing required key: {exc}") from exc
        sections.append(mapping)

    return Config(excel_file=excel_file, sections=sections)


def generate_config_template(
    excel_file: str,
    sections: list[dict[str, object]],
) -> str:
    """Generate a ``reqsync.toml`` template string.

    Args:
        excel_file: Path to the Excel workbook.
        sections: List of section descriptors, each a dict with keys
            ``name``, ``sdoc_file``, ``mode``, ``anchor``, ``uid_col``,
            ``title_col``, ``statement_col``.

    Returns:
        TOML-formatted string ready to be written to a file.
    """
    lines: list[str] = [
        "# reqsync.toml – generated by reqsync init-config",
        "",
        "[global]",
        f'excel_file = "{excel_file}"',
        "",
    ]

    for sec in sections:
        sname = sec['name']
        lines += [
            f"[{sname}]",
            f"sdoc_file         = \"{sec.get('sdoc_file', sname.lower() + '.sdoc')}\"",
            f"mode              = \"{sec.get('mode', 'legacy')}\"",
            f"anchor            = \"{sec.get('anchor', '')}\"",
            f"uid_col           = \"{sec.get('uid_col', 'A')}\"",
            f"title_col         = \"{sec.get('title_col', 'B')}\"",
            f"statement_col     = \"{sec.get('statement_col', 'C')}\"",
            "# relations_col   = \"D\"",
            "# grammar_tag     = \"REQUIREMENT\"",
            "# sync_direction  = \"excel_to_sdoc\"  # excel_to_sdoc | sdoc_to_excel | both",
            "# conflict_resolution = \"excel\"      # excel | sdoc (tiebreaker for 'both' mode)",
            "",
            f"# [{sname}.extra_cols]",
            "# Safety_Level = \"SAFETY_LEVEL\"",
            "",
            f"# [{sname}.field_directions]",
            "# TITLE     = \"sdoc_to_excel\"   # SDoc is authoritative for this field",
            "# STATEMENT = \"excel_to_sdoc\"   # Excel is authoritative for this field",
            "",
        ]

    return "\n".join(lines)
