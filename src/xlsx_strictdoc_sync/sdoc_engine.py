"""SDoc engine: read/write StrictDoc ``.sdoc`` documents.

This module wraps ``strictdoc.backend.sdoc.reader.SDReader`` and
``strictdoc.backend.sdoc.writer.SDWriter`` to provide a clean API for
converting between :class:`~.models.Requirement` objects and SDoc documents.

Grammar conventions
-------------------
Each section produces its own ``.sdoc`` file with an inline GRAMMAR block.
The grammar element tag defaults to ``"REQUIREMENT"`` and can be overridden
via :attr:`~.config_manager.SectionMapping.grammar_tag`.

Standard fields written/read:

``UID`` → ``requirement.uid``
``TITLE`` → ``requirement.title``   (omitted if empty)
``STATEMENT`` → ``requirement.statement``  (omitted if empty)

Any extra SDoc fields are stored in ``requirement.custom_fields``.

Parent relations
----------------
If a requirement has :attr:`~.models.Requirement.relations`, each parent UID
is written as a ``RELATIONS: TYPE: Parent`` block.  The grammar element must
declare ``RELATIONS: - TYPE: Parent`` for the writer to include them.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from strictdoc.backend.sdoc.models.document import SDocDocument
from strictdoc.backend.sdoc.models.document_grammar import DocumentGrammar
from strictdoc.backend.sdoc.models.grammar_element import (
    GrammarElement,
    GrammarElementFieldString,
    GrammarElementRelationParent,
)
from strictdoc.backend.sdoc.models.node import (
    ParentReqReference,
    SDocNode,
    SDocNodeField,
)
from strictdoc.backend.sdoc.reader import SDReader
from strictdoc.backend.sdoc.writer import SDWriter
from strictdoc.core.project_config import ProjectConfig

from .models import Requirement

if TYPE_CHECKING:
    from .config_manager import SectionMapping

# Shared project config used by SDWriter (no project-level config needed).
_PROJECT_CONFIG = ProjectConfig()


class SDocEngineError(Exception):
    """Raised when an SDoc operation fails."""


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def read_sdoc(path: str | Path) -> SDocDocument:
    """Read an ``.sdoc`` file and return the parsed :class:`SDocDocument`.

    Args:
        path: Path to the ``.sdoc`` file.

    Returns:
        Parsed document.

    Raises:
        FileNotFoundError: If *path* does not exist.
        SDocEngineError: If the file cannot be parsed.
    """
    sdoc_path = Path(path)
    if not sdoc_path.exists():
        raise FileNotFoundError(f"SDoc file not found: {sdoc_path}")
    try:
        return SDReader.read(sdoc_path.read_text(encoding="utf-8"), file_path=str(sdoc_path))
    except Exception as exc:
        raise SDocEngineError(f"Failed to parse '{sdoc_path}': {exc}") from exc


def write_sdoc(document: SDocDocument, path: str | Path) -> None:
    """Serialise *document* and write it to *path*.

    Args:
        document: The document to write.
        path: Destination path (parent directories are created if absent).
    """
    sdoc_path = Path(path)
    sdoc_path.parent.mkdir(parents=True, exist_ok=True)
    writer = SDWriter(_PROJECT_CONFIG)
    sdoc_path.write_text(writer.write(document), encoding="utf-8")


def document_to_requirements(
    document: SDocDocument,
    grammar_tag: str = "REQUIREMENT",
) -> list[Requirement]:
    """Extract :class:`~.models.Requirement` objects from an SDoc document.

    Only nodes whose :attr:`~strictdoc.backend.sdoc.models.node.SDocNode.node_type`
    matches *grammar_tag* are considered.

    Args:
        document: Parsed SDoc document.
        grammar_tag: Grammar element tag to filter on.

    Returns:
        Ordered list of requirements as found in the document.
    """
    requirements: list[Requirement] = []

    for node in document.section_contents:
        if not isinstance(node, SDocNode):
            continue
        if node.node_type != grammar_tag:
            continue

        uid = _get_field_value(node, "UID")
        if not uid:
            continue

        title = _get_field_value(node, "TITLE")
        statement = _get_field_value(node, "STATEMENT")

        # Collect custom fields (anything other than the standard three)
        standard = {"UID", "TITLE", "STATEMENT"}
        custom_fields: dict[str, str] = {}
        for field_name, field_list in node.ordered_fields_lookup.items():
            if field_name in standard:
                continue
            values = [f.get_text_value() for f in field_list if f.get_text_value()]
            if values:
                custom_fields[field_name] = values[0]

        # Parent relations
        relations: list[str] = [
            ref.ref_uid
            for ref in node.relations
            if hasattr(ref, "ref_uid")
        ]

        requirements.append(
            Requirement(
                uid=uid,
                title=title,
                statement=statement,
                custom_fields=custom_fields,
                relations=relations,
            )
        )

    return requirements


def requirements_to_document(
    requirements: list[Requirement],
    title: str,
    mapping: "SectionMapping",
) -> SDocDocument:
    """Create a new :class:`SDocDocument` from a list of requirements.

    Args:
        requirements: Requirements to include.
        title: Document title.
        mapping: Section configuration (grammar tag, extra fields).

    Returns:
        A fully constructed SDoc document ready to be written.
    """
    grammar = _build_grammar(mapping)
    document = SDocDocument(
        mid=None,
        title=title,
        config=None,
        view=None,
        grammar=grammar,
        section_contents=[],
    )
    grammar.parent = document
    for elem in grammar.elements:
        elem.parent = grammar

    for req in requirements:
        node = _requirement_to_node(req, document, mapping)
        document.section_contents.append(node)

    return document


def update_document(
    document: SDocDocument,
    requirements: list[Requirement],
    mapping: "SectionMapping",
) -> SDocDocument:
    """Merge *requirements* into an existing *document*.

    * Existing nodes whose UID matches an incoming requirement are updated.
    * New requirements (UIDs not yet in the document) are appended.
    * Nodes whose UIDs are *not* in *requirements* are left unchanged.

    Args:
        document: Existing parsed document.
        requirements: Incoming requirements (typically from Excel).
        mapping: Section configuration.

    Returns:
        The modified document (mutated in-place for convenience).
    """
    grammar_tag = mapping.grammar_tag
    req_by_uid = {r.uid: r for r in requirements}

    # Update existing nodes
    existing_uids: set[str] = set()
    for node in document.section_contents:
        if not isinstance(node, SDocNode) or node.node_type != grammar_tag:
            continue
        uid = _get_field_value(node, "UID")
        if uid and uid in req_by_uid:
            existing_uids.add(uid)
            _update_node_fields(node, req_by_uid[uid], mapping)

    # Append new requirements
    for req in requirements:
        if req.uid not in existing_uids:
            node = _requirement_to_node(req, document, mapping)
            document.section_contents.append(node)

    return document


def generate_grammar_sdoc(
    title: str,
    mapping: "SectionMapping",
) -> str:
    """Generate a standalone ``.sdoc`` grammar file content.

    Args:
        title: Document title for the grammar file.
        mapping: Section configuration.

    Returns:
        SDoc-formatted string for the grammar file.
    """
    grammar = _build_grammar(mapping)
    document = SDocDocument(
        mid=None,
        title=title,
        config=None,
        view=None,
        grammar=grammar,
        section_contents=[],
    )
    grammar.parent = document
    for elem in grammar.elements:
        elem.parent = grammar

    writer = SDWriter(_PROJECT_CONFIG)
    return writer.write(document)


# ---------------------------------------------------------------------------
# Internals
# ---------------------------------------------------------------------------


def _build_grammar(mapping: "SectionMapping") -> DocumentGrammar:
    """Construct a :class:`DocumentGrammar` from *mapping*."""
    standard_fields: list[GrammarElementFieldString] = [
        GrammarElementFieldString(parent=None, title="UID", human_title=None, required="False"),
    ]
    if mapping.title_col:
        standard_fields.append(
            GrammarElementFieldString(parent=None, title="TITLE", human_title=None, required="False")
        )
    if mapping.statement_col:
        standard_fields.append(
            GrammarElementFieldString(
                parent=None, title="STATEMENT", human_title=None, required="False"
            )
        )

    # Extra custom fields
    extra_fields = [
        GrammarElementFieldString(
            parent=None, title=sdoc_field, human_title=None, required="False"
        )
        for sdoc_field in mapping.extra_cols.values()
    ]

    relations: list[GrammarElementRelationParent] = []
    if mapping.relations_col:
        relations.append(
            GrammarElementRelationParent(parent=None, relation_type="Parent", relation_role=None)
        )

    element = GrammarElement(
        parent=None,
        tag=mapping.grammar_tag,
        property_is_composite="",
        property_prefix="",
        property_view_style="",
        fields=standard_fields + extra_fields,
        relations=relations,
    )

    return DocumentGrammar(parent=None, elements=[element])


def _requirement_to_node(
    req: Requirement,
    document: SDocDocument,
    mapping: "SectionMapping",
) -> SDocNode:
    """Convert a :class:`~.models.Requirement` to an :class:`SDocNode`."""
    fields: list[SDocNodeField] = [
        SDocNodeField(parent=None, field_name="UID", parts=[req.uid], multiline__=None),
    ]
    if mapping.title_col and req.title:
        fields.append(
            SDocNodeField(parent=None, field_name="TITLE", parts=[req.title], multiline__=None)
        )
    if mapping.statement_col and req.statement:
        fields.append(
            SDocNodeField(
                parent=None,
                field_name="STATEMENT",
                parts=[req.statement],
                multiline__=None,
            )
        )
    for sdoc_field in mapping.extra_cols.values():
        value = req.custom_fields.get(sdoc_field, "")
        if value:
            fields.append(
                SDocNodeField(
                    parent=None,
                    field_name=sdoc_field,
                    parts=[value],
                    multiline__=None,
                )
            )

    relations: list[ParentReqReference] = []
    if mapping.relations_col:
        relations = [
            ParentReqReference(parent=None, ref_uid=parent_uid, role=None)
            for parent_uid in req.relations
        ]

    node = SDocNode(
        parent=document,
        node_type=mapping.grammar_tag,
        fields=fields,
        relations=relations,
    )
    for f in fields:
        f.parent = node
    for r in relations:
        r.parent = node

    return node


def _update_node_fields(
    node: SDocNode,
    req: Requirement,
    mapping: "SectionMapping",
) -> None:
    """Overwrite the fields of *node* with values from *req*."""
    def _set(field_name: str, value: str) -> None:
        if not value:
            node.ordered_fields_lookup.pop(field_name, None)
            return
        field = SDocNodeField(
            parent=node,
            field_name=field_name,
            parts=[value],
            multiline__=None,
        )
        node.ordered_fields_lookup[field_name] = [field]

    _set("UID", req.uid)
    if mapping.title_col:
        _set("TITLE", req.title)
    if mapping.statement_col:
        _set("STATEMENT", req.statement)
    for sdoc_field in mapping.extra_cols.values():
        _set(sdoc_field, req.custom_fields.get(sdoc_field, ""))

    if mapping.relations_col:
        node.relations = [
            ParentReqReference(parent=node, ref_uid=uid, role=None)
            for uid in req.relations
        ]


def _get_field_value(node: SDocNode, field_name: str) -> str:
    """Return the text value of the first occurrence of *field_name*."""
    fields = node.ordered_fields_lookup.get(field_name, [])
    if not fields:
        return ""
    return fields[0].get_text_value()
