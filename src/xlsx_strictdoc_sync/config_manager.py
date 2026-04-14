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

    def __post_init__(self) -> None:
        if self.mode not in ("table", "legacy"):
            raise ValueError(
                f"[{self.name}] 'mode' must be 'table' or 'legacy', got '{self.mode}'."
            )
        if not self.anchor:
            raise ValueError(f"[{self.name}] 'anchor' must not be empty.")
        if not self.uid_col:
            raise ValueError(f"[{self.name}] 'uid_col' must not be empty.")


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
        lines += [
            f"[{sec['name']}]",
            f"sdoc_file  = \"{sec.get('sdoc_file', sec['name'].lower() + '.sdoc')}\"",
            f"mode       = \"{sec.get('mode', 'legacy')}\"",
            f"anchor     = \"{sec.get('anchor', '')}\"",
            f"uid_col    = \"{sec.get('uid_col', 'A')}\"",
            f"title_col  = \"{sec.get('title_col', 'B')}\"",
            f"statement_col = \"{sec.get('statement_col', 'C')}\"",
            "# relations_col = \"D\"",
            "# grammar_tag   = \"REQUIREMENT\"",
            "# [SYS_REQS.extra_cols]",
            "# Safety_Level = \"SAFETY_LEVEL\"",
            "",
        ]

    return "\n".join(lines)
