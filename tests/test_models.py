"""Tests for the Requirement model."""

from __future__ import annotations

import pytest

from xlsx_strictdoc_sync.models import Requirement


class TestRequirementCreation:
    def test_minimal_creation(self):
        req = Requirement(uid="SYS-001")
        assert req.uid == "SYS-001"
        assert req.title == ""
        assert req.statement == ""
        assert req.custom_fields == {}
        assert req.relations == []

    def test_full_creation(self):
        req = Requirement(
            uid="SW-042",
            title="Boot time",
            statement="The system shall boot within 5 s.",
            custom_fields={"SAFETY": "ASIL-B"},
            relations=["SYS-001", "SYS-002"],
        )
        assert req.uid == "SW-042"
        assert req.title == "Boot time"
        assert req.statement == "The system shall boot within 5 s."
        assert req.custom_fields == {"SAFETY": "ASIL-B"}
        assert req.relations == ["SYS-001", "SYS-002"]

    def test_empty_uid_raises(self):
        with pytest.raises(ValueError, match="UID must not be empty"):
            Requirement(uid="")

    def test_whitespace_uid_raises(self):
        with pytest.raises(ValueError, match="UID must not be empty"):
            Requirement(uid="   ")

    def test_has_relations_false(self):
        req = Requirement(uid="X-1")
        assert req.has_relations is False

    def test_has_relations_true(self):
        req = Requirement(uid="X-2", relations=["X-1"])
        assert req.has_relations is True


class TestRequirementSerialization:
    def test_to_dict_round_trip(self):
        req = Requirement(
            uid="HW-007",
            title="Power supply",
            statement="Shall deliver 12 V.",
            custom_fields={"PRIORITY": "HIGH"},
            relations=["SYS-003"],
        )
        d = req.to_dict()
        assert d["uid"] == "HW-007"
        assert d["title"] == "Power supply"
        assert d["statement"] == "Shall deliver 12 V."
        assert d["custom_fields"] == {"PRIORITY": "HIGH"}
        assert d["relations"] == ["SYS-003"]

    def test_from_dict(self):
        d = {
            "uid": "SYS-010",
            "title": "Safety",
            "statement": "…",
            "custom_fields": {"FOO": "bar"},
            "relations": ["SYS-001"],
        }
        req = Requirement.from_dict(d)
        assert req.uid == "SYS-010"
        assert req.custom_fields == {"FOO": "bar"}

    def test_from_dict_minimal(self):
        req = Requirement.from_dict({"uid": "MIN-001"})
        assert req.uid == "MIN-001"
        assert req.title == ""
        assert req.relations == []

    def test_to_dict_is_copy(self):
        req = Requirement(uid="X-1", relations=["A"])
        d = req.to_dict()
        d["relations"].append("B")
        # Original should be untouched
        assert req.relations == ["A"]
