"""Tests for the SDoc engine."""

from __future__ import annotations

import pytest

from xlsx_strictdoc_sync.config_manager import SectionMapping
from xlsx_strictdoc_sync.models import Requirement
from xlsx_strictdoc_sync.sdoc_engine import (
    SDocEngineError,
    document_to_requirements,
    generate_grammar_sdoc,
    read_sdoc,
    requirements_to_document,
    update_document,
    write_sdoc,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _basic_mapping(**kwargs) -> SectionMapping:
    defaults = dict(
        name="SYS_REQS",
        sdoc_file="sys.sdoc",
        mode="table",
        anchor="SYS_Reqs",
        uid_col="UID",
        title_col="Title",
        statement_col="Statement",
        relations_col="Parent",
    )
    defaults.update(kwargs)
    return SectionMapping(**defaults)


def _reqs() -> list[Requirement]:
    return [
        Requirement(uid="SYS-001", title="Boot", statement="Shall boot."),
        Requirement(uid="SYS-002", title="Log", statement="Shall log.", relations=["SYS-001"]),
    ]


# ---------------------------------------------------------------------------
# requirements_to_document
# ---------------------------------------------------------------------------

class TestRequirementsToDocument:
    def test_creates_document(self):
        doc = requirements_to_document(_reqs(), "System Reqs", _basic_mapping())
        assert doc.title == "System Reqs"
        assert len(doc.section_contents) == 2

    def test_uids_preserved(self):
        doc = requirements_to_document(_reqs(), "T", _basic_mapping())
        reqs = document_to_requirements(doc)
        assert [r.uid for r in reqs] == ["SYS-001", "SYS-002"]

    def test_relations_preserved(self):
        doc = requirements_to_document(_reqs(), "T", _basic_mapping())
        reqs = document_to_requirements(doc)
        assert reqs[1].relations == ["SYS-001"]

    def test_titles_preserved(self):
        doc = requirements_to_document(_reqs(), "T", _basic_mapping())
        reqs = document_to_requirements(doc)
        assert reqs[0].title == "Boot"

    def test_custom_fields_preserved(self):
        mapping = _basic_mapping(extra_cols={"Safety": "SAFETY_LEVEL"})
        reqs = [Requirement(uid="SYS-001", custom_fields={"SAFETY_LEVEL": "ASIL-B"})]
        doc = requirements_to_document(reqs, "T", mapping)
        out = document_to_requirements(doc)
        assert out[0].custom_fields.get("SAFETY_LEVEL") == "ASIL-B"

    def test_grammar_tag_used(self):
        mapping = _basic_mapping(grammar_tag="SW_REQUIREMENT")
        doc = requirements_to_document(_reqs(), "T", mapping)
        from strictdoc.backend.sdoc.models.node import SDocNode
        tags = [n.node_type for n in doc.section_contents if isinstance(n, SDocNode)]
        assert all(t == "SW_REQUIREMENT" for t in tags)

    def test_no_title_col_omits_title_field(self):
        mapping = _basic_mapping(title_col="")
        doc = requirements_to_document(_reqs(), "T", mapping)
        reqs = document_to_requirements(doc)
        assert all(r.title == "" for r in reqs)

    def test_empty_list(self):
        doc = requirements_to_document([], "Empty", _basic_mapping())
        assert doc.section_contents == []


# ---------------------------------------------------------------------------
# document_to_requirements
# ---------------------------------------------------------------------------

class TestDocumentToRequirements:
    def test_round_trip(self):
        original = _reqs()
        doc = requirements_to_document(original, "T", _basic_mapping())
        result = document_to_requirements(doc)
        assert len(result) == len(original)
        for r_in, r_out in zip(original, result):
            assert r_in.uid == r_out.uid
            assert r_in.statement == r_out.statement

    def test_filters_by_grammar_tag(self):
        mapping = _basic_mapping(grammar_tag="REQUIREMENT")
        doc = requirements_to_document(_reqs(), "T", mapping)
        # Ask for a different tag – should get nothing
        result = document_to_requirements(doc, grammar_tag="SW_REQ")
        assert result == []


# ---------------------------------------------------------------------------
# update_document
# ---------------------------------------------------------------------------

class TestUpdateDocument:
    def test_updates_existing(self):
        original = [Requirement(uid="SYS-001", statement="Old")]
        doc = requirements_to_document(original, "T", _basic_mapping())
        update = [Requirement(uid="SYS-001", statement="New")]
        doc = update_document(doc, update, _basic_mapping())
        result = document_to_requirements(doc)
        assert result[0].statement == "New"

    def test_appends_new(self):
        original = [Requirement(uid="SYS-001", statement="S")]
        doc = requirements_to_document(original, "T", _basic_mapping())
        new_req = [Requirement(uid="SYS-002", statement="New")]
        doc = update_document(doc, new_req, _basic_mapping())
        result = document_to_requirements(doc)
        uids = [r.uid for r in result]
        assert "SYS-001" in uids
        assert "SYS-002" in uids

    def test_untouched_nodes_preserved(self):
        original = [
            Requirement(uid="SYS-001", statement="Untouched"),
            Requirement(uid="SYS-002", statement="Will be updated"),
        ]
        doc = requirements_to_document(original, "T", _basic_mapping())
        update = [Requirement(uid="SYS-002", statement="Updated")]
        doc = update_document(doc, update, _basic_mapping())
        result = document_to_requirements(doc)
        by_uid = {r.uid: r for r in result}
        assert by_uid["SYS-001"].statement == "Untouched"
        assert by_uid["SYS-002"].statement == "Updated"


# ---------------------------------------------------------------------------
# read_sdoc / write_sdoc
# ---------------------------------------------------------------------------

class TestReadWriteSdoc:
    def test_write_then_read(self, tmp_path):
        sdoc_path = tmp_path / "test.sdoc"
        doc = requirements_to_document(_reqs(), "System Reqs", _basic_mapping())
        write_sdoc(doc, sdoc_path)
        assert sdoc_path.exists()

        doc2 = read_sdoc(sdoc_path)
        reqs = document_to_requirements(doc2)
        assert len(reqs) == 2
        assert reqs[0].uid == "SYS-001"

    def test_read_nonexistent_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            read_sdoc(tmp_path / "missing.sdoc")

    def test_write_creates_parent_dirs(self, tmp_path):
        sdoc_path = tmp_path / "sub" / "dir" / "out.sdoc"
        doc = requirements_to_document([], "Empty", _basic_mapping())
        write_sdoc(doc, sdoc_path)
        assert sdoc_path.exists()

    def test_relations_survive_round_trip(self, tmp_path):
        sdoc_path = tmp_path / "rels.sdoc"
        doc = requirements_to_document(_reqs(), "T", _basic_mapping())
        write_sdoc(doc, sdoc_path)
        doc2 = read_sdoc(sdoc_path)
        reqs = document_to_requirements(doc2)
        assert reqs[1].relations == ["SYS-001"]


# ---------------------------------------------------------------------------
# generate_grammar_sdoc
# ---------------------------------------------------------------------------

class TestGenerateGrammarSdoc:
    def test_returns_string(self):
        content = generate_grammar_sdoc("My Grammar", _basic_mapping())
        assert isinstance(content, str)
        assert "REQUIREMENT" in content

    def test_custom_fields_included(self):
        mapping = _basic_mapping(extra_cols={"Safety": "SAFETY_LEVEL"})
        content = generate_grammar_sdoc("G", mapping)
        assert "SAFETY_LEVEL" in content

    def test_custom_tag_included(self):
        mapping = _basic_mapping(grammar_tag="HW_REQUIREMENT")
        content = generate_grammar_sdoc("G", mapping)
        assert "HW_REQUIREMENT" in content
