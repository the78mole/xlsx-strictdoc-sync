"""Sync engine: orchestrate bidirectional synchronization between Excel and SDoc.

This module implements the **merge logic** that sits above the individual
:mod:`excel_engine` and :mod:`sdoc_engine` backends.  It answers the
question: *given a set of requirements from Excel and a set from SDoc, what
should each side look like after the sync?*

Sync directions
---------------
Three overall directions are supported (set per-section in ``reqsync.toml``
or overridden at runtime via the ``--direction`` CLI flag):

``excel_to_sdoc``
    Excel is the authoritative source.  The SDoc file is updated to match
    Excel.  The workbook is never modified.

``sdoc_to_excel``
    SDoc is the authoritative source.  The Excel workbook is updated to match
    the SDoc file.  The SDoc file is never modified.

``both``
    Bidirectional sync.  New UIDs found only on one side are propagated to the
    other.  For UIDs that already exist on both sides, individual **field
    direction overrides** (``[SECTION.field_directions]`` in the config)
    determine which side wins field-by-field.  When a field has no specific
    override, the section-level ``conflict_resolution`` setting (``"excel"``
    or ``"sdoc"``) is used as the tiebreaker.

Field direction overrides
-------------------------
Only relevant when ``sync_direction = "both"``::

    [SYS_REQS.field_directions]
    TITLE     = "sdoc_to_excel"   # SDoc is authoritative for titles
    STATEMENT = "excel_to_sdoc"   # Excel is authoritative for statements
    SAFETY    = "sdoc_to_excel"   # SDoc owns safety classification

Keys are **SDoc grammar field names** (matching the grammar element defined
for the section).  Values are ``"excel_to_sdoc"`` or ``"sdoc_to_excel"``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .config_manager import (
    CONFLICT_EXCEL,
    DIRECTION_BOTH,
    DIRECTION_EXCEL_TO_SDOC,
    DIRECTION_SDOC_TO_EXCEL,
    VALID_DIRECTIONS,
    SectionMapping,
)
from .models import Requirement

# ---------------------------------------------------------------------------
# Public constants & types
# ---------------------------------------------------------------------------

#: Standard SDoc field names that receive special direction handling.
_STANDARD_FIELDS = ("TITLE", "STATEMENT", "RELATIONS")


# ---------------------------------------------------------------------------
# Result container
# ---------------------------------------------------------------------------


@dataclass
class SyncResult:
    """Summary statistics for one synced section.

    Attributes:
        section: Section name from the config.
        sdoc_added: Number of requirements added to the SDoc file.
        sdoc_updated: Number of requirements updated in the SDoc file.
        excel_added: Number of requirements added to the Excel workbook.
        excel_updated: Number of requirements updated in the Excel workbook.
        skipped: Number of requirements whose UID appeared in both sources but
            all fields were identical (no write needed).
    """

    section: str
    sdoc_added: int = 0
    sdoc_updated: int = 0
    excel_added: int = 0
    excel_updated: int = 0
    skipped: int = 0

    @property
    def total_changes(self) -> int:
        """Total number of modifications across both targets."""
        return self.sdoc_added + self.sdoc_updated + self.excel_added + self.excel_updated

    def summary(self) -> str:
        """Return a one-line human-readable summary."""
        parts = []
        if self.sdoc_added or self.sdoc_updated:
            parts.append(f"SDoc: +{self.sdoc_added} / ~{self.sdoc_updated}")
        if self.excel_added or self.excel_updated:
            parts.append(f"Excel: +{self.excel_added} / ~{self.excel_updated}")
        if self.skipped:
            parts.append(f"skipped: {self.skipped}")
        return f"[{self.section}] " + (", ".join(parts) if parts else "no changes")


# ---------------------------------------------------------------------------
# Core merge API
# ---------------------------------------------------------------------------


def compute_sync(
    excel_reqs: list[Requirement],
    sdoc_reqs: list[Requirement],
    mapping: SectionMapping,
    direction_override: str | None = None,
) -> tuple[list[Requirement], list[Requirement], SyncResult]:
    """Compute what each target should look like after the sync.

    This is the **pure** part of the sync – it does not touch the filesystem.
    It consumes the current state of both sides and produces the desired
    post-sync state for each target.

    Args:
        excel_reqs: Requirements currently in the Excel workbook for this
            section.
        sdoc_reqs: Requirements currently in the SDoc file for this section.
        mapping: Section configuration (direction, field overrides, …).
        direction_override: Runtime direction override, e.g. from the
            ``--direction`` CLI flag.  When supplied it takes precedence over
            :attr:`~.config_manager.SectionMapping.sync_direction`.

    Returns:
        A 3-tuple of:

        * ``excel_updates`` – complete list of requirements that should be
          written to Excel (empty list = no Excel writes needed).
        * ``sdoc_updates`` – complete list of requirements that should be
          written to SDoc (empty list = no SDoc writes needed).
        * :class:`SyncResult` – statistics about the operation.
    """
    effective_dir = _resolve_direction(mapping.sync_direction, direction_override)

    excel_by_uid: dict[str, Requirement] = {r.uid: r for r in excel_reqs}
    sdoc_by_uid: dict[str, Requirement] = {r.uid: r for r in sdoc_reqs}

    # Preserve insertion order: Excel-first, then SDoc-only additions.
    all_uids: list[str] = list(
        dict.fromkeys(list(excel_by_uid) + list(sdoc_by_uid))
    )

    excel_updates: list[Requirement] = []
    sdoc_updates: list[Requirement] = []
    result = SyncResult(section=mapping.name)

    for uid in all_uids:
        e_req = excel_by_uid.get(uid)
        s_req = sdoc_by_uid.get(uid)

        if e_req is not None and s_req is None:
            # UID only in Excel
            if effective_dir in (DIRECTION_EXCEL_TO_SDOC, DIRECTION_BOTH):
                sdoc_updates.append(e_req)
                result.sdoc_added += 1

        elif s_req is not None and e_req is None:
            # UID only in SDoc
            if effective_dir in (DIRECTION_SDOC_TO_EXCEL, DIRECTION_BOTH):
                excel_updates.append(s_req)
                result.excel_added += 1

        else:
            # UID in both – apply field-level direction rules
            assert e_req is not None and s_req is not None
            merged_for_sdoc, merged_for_excel = _merge_pair(
                e_req, s_req, mapping, effective_dir
            )

            if effective_dir in (DIRECTION_EXCEL_TO_SDOC, DIRECTION_BOTH):
                if not _requirements_equal(merged_for_sdoc, s_req):
                    sdoc_updates.append(merged_for_sdoc)
                    result.sdoc_updated += 1
                else:
                    result.skipped += 1

            if effective_dir in (DIRECTION_SDOC_TO_EXCEL, DIRECTION_BOTH):
                if not _requirements_equal(merged_for_excel, e_req):
                    excel_updates.append(merged_for_excel)
                    result.excel_updated += 1

    return excel_updates, sdoc_updates, result


# ---------------------------------------------------------------------------
# Field-level merge
# ---------------------------------------------------------------------------


def _merge_pair(
    excel_req: Requirement,
    sdoc_req: Requirement,
    mapping: SectionMapping,
    effective_dir: str,
) -> tuple[Requirement, Requirement]:
    """Merge a matched pair (same UID) applying per-field direction rules.

    Returns:
        ``(merged_for_sdoc, merged_for_excel)`` – the desired state on each
        target for this requirement.
    """
    fd = mapping.field_directions
    prefer_excel = mapping.conflict_resolution == CONFLICT_EXCEL

    def _for_sdoc(field_name: str, excel_val: str, sdoc_val: str) -> str:
        """Value that the SDoc side should have for *field_name*."""
        override = fd.get(field_name)
        if override == DIRECTION_SDOC_TO_EXCEL:
            # SDoc is authoritative: SDoc keeps its own value.
            return sdoc_val
        if override == DIRECTION_EXCEL_TO_SDOC:
            return excel_val
        # No field-level override: apply tiebreaker.
        if effective_dir == DIRECTION_EXCEL_TO_SDOC:
            return excel_val
        if effective_dir == DIRECTION_SDOC_TO_EXCEL:
            return sdoc_val
        # both + no override → conflict_resolution
        return excel_val if prefer_excel else sdoc_val

    def _for_excel(field_name: str, excel_val: str, sdoc_val: str) -> str:
        """Value that the Excel side should have for *field_name*."""
        override = fd.get(field_name)
        if override == DIRECTION_EXCEL_TO_SDOC:
            # Excel is authoritative: Excel keeps its own value.
            return excel_val
        if override == DIRECTION_SDOC_TO_EXCEL:
            return sdoc_val
        if effective_dir == DIRECTION_SDOC_TO_EXCEL:
            return sdoc_val
        if effective_dir == DIRECTION_EXCEL_TO_SDOC:
            return excel_val
        # both + no override → conflict_resolution
        return excel_val if prefer_excel else sdoc_val

    # --- TITLE ---
    title_s = _for_sdoc("TITLE", excel_req.title, sdoc_req.title)
    title_e = _for_excel("TITLE", excel_req.title, sdoc_req.title)

    # --- STATEMENT ---
    stmt_s = _for_sdoc("STATEMENT", excel_req.statement, sdoc_req.statement)
    stmt_e = _for_excel("STATEMENT", excel_req.statement, sdoc_req.statement)

    # --- RELATIONS ---
    # Relations are compared as semicolon-joined sorted strings for direction logic
    excel_rels_str = "; ".join(sorted(excel_req.relations))
    sdoc_rels_str = "; ".join(sorted(sdoc_req.relations))
    rels_s_str = _for_sdoc("RELATIONS", excel_rels_str, sdoc_rels_str)
    rels_e_str = _for_excel("RELATIONS", excel_rels_str, sdoc_rels_str)
    rels_s = [r.strip() for r in rels_s_str.split(";") if r.strip()] if rels_s_str else []
    rels_e = [r.strip() for r in rels_e_str.split(";") if r.strip()] if rels_e_str else []

    # --- CUSTOM FIELDS ---
    all_custom = set(excel_req.custom_fields) | set(sdoc_req.custom_fields)
    custom_s: dict[str, str] = {}
    custom_e: dict[str, str] = {}
    for fname in all_custom:
        e_val = excel_req.custom_fields.get(fname, "")
        s_val = sdoc_req.custom_fields.get(fname, "")
        custom_s[fname] = _for_sdoc(fname, e_val, s_val)
        custom_e[fname] = _for_excel(fname, e_val, s_val)

    merged_for_sdoc = Requirement(
        uid=excel_req.uid,
        title=title_s,
        statement=stmt_s,
        custom_fields=custom_s,
        relations=rels_s,
    )
    merged_for_excel = Requirement(
        uid=excel_req.uid,
        title=title_e,
        statement=stmt_e,
        custom_fields=custom_e,
        relations=rels_e,
    )
    return merged_for_sdoc, merged_for_excel


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _resolve_direction(section_dir: str, override: str | None) -> str:
    """Return the effective sync direction.

    The *override* (e.g. from the ``--direction`` CLI flag) wins when it is a
    valid direction string; otherwise *section_dir* is used.

    Raises:
        ValueError: If *override* is not a recognised direction.
    """
    if override is None:
        return section_dir
    if override not in VALID_DIRECTIONS:
        raise ValueError(
            f"Invalid direction '{override}'. Must be one of {sorted(VALID_DIRECTIONS)}."
        )
    return override


def _requirements_equal(a: Requirement, b: Requirement) -> bool:
    """Return ``True`` if two requirements are semantically identical."""
    return (
        a.uid == b.uid
        and a.title == b.title
        and a.statement == b.statement
        and a.custom_fields == b.custom_fields
        and sorted(a.relations) == sorted(b.relations)
    )
