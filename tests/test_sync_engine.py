"""Tests for the bidirectional sync engine."""

from __future__ import annotations

import pytest

from xlsx_strictdoc_sync.config_manager import SectionMapping
from xlsx_strictdoc_sync.models import Requirement
from xlsx_strictdoc_sync.sync_engine import (
    SyncResult,
    _compare_timestamps,
    _requirements_equal,
    _resolve_direction,
    compute_sync,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _mapping(**kwargs) -> SectionMapping:
    defaults = dict(
        name="SYS_REQS",
        sdoc_file="sys.sdoc",
        mode="table",
        anchor="SYS_Reqs",
        uid_col="UID",
        title_col="Title",
        statement_col="Statement",
        relations_col="Parent",
        sync_direction="both",
        conflict_resolution="excel",
        field_directions={},
    )
    defaults.update(kwargs)
    return SectionMapping(**defaults)


def _req(uid, title="", statement="", relations=None, custom_fields=None) -> Requirement:
    return Requirement(
        uid=uid,
        title=title,
        statement=statement,
        relations=relations or [],
        custom_fields=custom_fields or {},
    )


# ---------------------------------------------------------------------------
# _resolve_direction
# ---------------------------------------------------------------------------


class TestResolveDirection:
    def test_no_override_uses_section(self):
        assert _resolve_direction("excel_to_sdoc", None) == "excel_to_sdoc"

    def test_override_wins(self):
        assert _resolve_direction("excel_to_sdoc", "sdoc_to_excel") == "sdoc_to_excel"

    def test_invalid_override_raises(self):
        with pytest.raises(ValueError, match="Invalid direction"):
            _resolve_direction("both", "wrong")


# ---------------------------------------------------------------------------
# _requirements_equal
# ---------------------------------------------------------------------------


class TestRequirementsEqual:
    def test_identical(self):
        a = _req("SYS-001", title="T", statement="S", relations=["P-1"])
        b = _req("SYS-001", title="T", statement="S", relations=["P-1"])
        assert _requirements_equal(a, b)

    def test_different_statement(self):
        a = _req("SYS-001", statement="A")
        b = _req("SYS-001", statement="B")
        assert not _requirements_equal(a, b)

    def test_relations_order_irrelevant(self):
        a = _req("X", relations=["B", "A"])
        b = _req("X", relations=["A", "B"])
        assert _requirements_equal(a, b)


# ---------------------------------------------------------------------------
# compute_sync – excel_to_sdoc
# ---------------------------------------------------------------------------


class TestComputeSyncExcelToSdoc:
    def test_new_in_excel_added_to_sdoc(self):
        e = [_req("SYS-001", statement="From Excel")]
        s = []
        ex_upd, sdoc_upd, result = compute_sync(e, s, _mapping(sync_direction="excel_to_sdoc"))
        assert len(sdoc_upd) == 1
        assert sdoc_upd[0].uid == "SYS-001"
        assert ex_upd == []
        assert result.sdoc_added == 1

    def test_new_in_sdoc_not_pushed_to_excel(self):
        e = []
        s = [_req("SYS-001", statement="SDoc only")]
        ex_upd, sdoc_upd, _ = compute_sync(e, s, _mapping(sync_direction="excel_to_sdoc"))
        assert ex_upd == []
        assert sdoc_upd == []

    def test_existing_excel_wins(self):
        e = [_req("SYS-001", statement="Excel value")]
        s = [_req("SYS-001", statement="SDoc value")]
        _, sdoc_upd, result = compute_sync(e, s, _mapping(sync_direction="excel_to_sdoc"))
        assert len(sdoc_upd) == 1
        assert sdoc_upd[0].statement == "Excel value"
        assert result.sdoc_updated == 1

    def test_identical_no_update(self):
        req = _req("SYS-001", statement="Same")
        _, sdoc_upd, result = compute_sync([req], [req], _mapping(sync_direction="excel_to_sdoc"))
        assert sdoc_upd == []
        assert result.skipped == 1


# ---------------------------------------------------------------------------
# compute_sync – sdoc_to_excel
# ---------------------------------------------------------------------------


class TestComputeSyncSdocToExcel:
    def test_new_in_sdoc_added_to_excel(self):
        e = []
        s = [_req("SYS-001", statement="From SDoc")]
        ex_upd, sdoc_upd, result = compute_sync(e, s, _mapping(sync_direction="sdoc_to_excel"))
        assert len(ex_upd) == 1
        assert ex_upd[0].uid == "SYS-001"
        assert sdoc_upd == []
        assert result.excel_added == 1

    def test_new_in_excel_not_pushed_to_sdoc(self):
        e = [_req("SYS-001")]
        s = []
        ex_upd, sdoc_upd, _ = compute_sync(e, s, _mapping(sync_direction="sdoc_to_excel"))
        assert ex_upd == []
        assert sdoc_upd == []

    def test_existing_sdoc_wins(self):
        e = [_req("SYS-001", statement="Excel value")]
        s = [_req("SYS-001", statement="SDoc value")]
        ex_upd, _, result = compute_sync(e, s, _mapping(sync_direction="sdoc_to_excel"))
        assert len(ex_upd) == 1
        assert ex_upd[0].statement == "SDoc value"
        assert result.excel_updated == 1


# ---------------------------------------------------------------------------
# compute_sync – both (bidirectional)
# ---------------------------------------------------------------------------


class TestComputeSyncBoth:
    def test_new_in_excel_added_to_sdoc(self):
        e = [_req("SYS-001")]
        s = []
        _, sdoc_upd, result = compute_sync(e, s, _mapping())
        assert len(sdoc_upd) == 1
        assert result.sdoc_added == 1

    def test_new_in_sdoc_added_to_excel(self):
        e = []
        s = [_req("SYS-001")]
        ex_upd, _, result = compute_sync(e, s, _mapping())
        assert len(ex_upd) == 1
        assert result.excel_added == 1

    def test_excel_wins_by_default(self):
        e = [_req("SYS-001", statement="Excel")]
        s = [_req("SYS-001", statement="SDoc")]
        _, sdoc_upd, _ = compute_sync(e, s, _mapping(conflict_resolution="excel"))
        assert sdoc_upd[0].statement == "Excel"

    def test_sdoc_wins_when_conflict_sdoc(self):
        e = [_req("SYS-001", statement="Excel")]
        s = [_req("SYS-001", statement="SDoc")]
        ex_upd, _, _ = compute_sync(e, s, _mapping(conflict_resolution="sdoc"))
        assert ex_upd[0].statement == "SDoc"

    def test_field_direction_sdoc_to_excel_overrides_title(self):
        """TITLE marked sdoc_to_excel: SDoc value used on both sides for TITLE."""
        e = [_req("SYS-001", title="Excel Title", statement="Excel Stmt")]
        s = [_req("SYS-001", title="SDoc Title", statement="SDoc Stmt")]
        mapping = _mapping(
            conflict_resolution="excel",
            field_directions={"TITLE": "sdoc_to_excel"},
        )
        _, sdoc_upd, _ = compute_sync(e, s, mapping)
        # SDoc TITLE should stay as SDoc's (sdoc_to_excel means SDoc is authoritative)
        assert sdoc_upd[0].title == "SDoc Title"
        # SDoc STATEMENT should come from Excel (excel wins by default)
        assert sdoc_upd[0].statement == "Excel Stmt"

    def test_field_direction_excel_to_sdoc_overrides_statement(self):
        """STATEMENT marked excel_to_sdoc: Excel value used in SDoc for STATEMENT."""
        e = [_req("SYS-001", title="Excel Title", statement="Excel Stmt")]
        s = [_req("SYS-001", title="SDoc Title", statement="SDoc Stmt")]
        mapping = _mapping(
            conflict_resolution="sdoc",
            field_directions={"STATEMENT": "excel_to_sdoc"},
        )
        _, sdoc_upd, _ = compute_sync(e, s, mapping)
        # STATEMENT: excel_to_sdoc override → Excel value goes to SDoc
        assert sdoc_upd[0].statement == "Excel Stmt"
        # TITLE: no override, conflict_resolution=sdoc → SDoc value
        assert sdoc_upd[0].title == "SDoc Title"

    def test_relations_direction_override(self):
        e = [_req("SYS-002", relations=["SYS-001"])]
        s = [_req("SYS-002", relations=["SYS-000"])]
        mapping = _mapping(field_directions={"RELATIONS": "sdoc_to_excel"})
        ex_upd, sdoc_upd, _ = compute_sync(e, s, mapping)
        # RELATIONS marked sdoc_to_excel: Excel should get SDoc's relations
        assert sorted(ex_upd[0].relations) == ["SYS-000"]
        # SDoc already has "SYS-000" – no update needed (value unchanged)
        assert sdoc_upd == []

    def test_custom_field_direction_override(self):
        e = [_req("SYS-001", custom_fields={"SAFETY": "ASIL-A"})]
        s = [_req("SYS-001", custom_fields={"SAFETY": "ASIL-B"})]
        mapping = _mapping(field_directions={"SAFETY": "sdoc_to_excel"})
        ex_upd, sdoc_upd, _ = compute_sync(e, s, mapping)
        # SDoc is authoritative for SAFETY → Excel should receive ASIL-B
        assert ex_upd[0].custom_fields["SAFETY"] == "ASIL-B"
        # SDoc already has ASIL-B – no SDoc update needed
        assert sdoc_upd == []

    def test_direction_override_parameter_wins(self):
        """CLI --direction flag overrides section config."""
        e = [_req("SYS-001", statement="Excel")]
        s = []
        # Section says sdoc_to_excel, but override says excel_to_sdoc
        _, sdoc_upd, _ = compute_sync(
            e, s,
            _mapping(sync_direction="sdoc_to_excel"),
            direction_override="excel_to_sdoc",
        )
        assert len(sdoc_upd) == 1

    def test_no_changes_when_identical(self):
        req = _req("SYS-001", title="T", statement="S")
        ex_upd, sdoc_upd, result = compute_sync([req], [req], _mapping())
        assert ex_upd == []
        assert sdoc_upd == []
        assert result.skipped == 1


# ---------------------------------------------------------------------------
# _compare_timestamps
# ---------------------------------------------------------------------------


class TestCompareTimestamps:
    def test_excel_newer_iso(self):
        assert _compare_timestamps("2024-06-15T10:00:00", "2024-06-14T10:00:00") is True

    def test_sdoc_newer_iso(self):
        assert _compare_timestamps("2024-06-13", "2024-06-15") is False

    def test_equal_timestamps(self):
        assert _compare_timestamps("2024-01-01", "2024-01-01") is None

    def test_both_empty(self):
        assert _compare_timestamps("", "") is None

    def test_only_excel_has_timestamp(self):
        assert _compare_timestamps("2024-01-01", "") is True

    def test_only_sdoc_has_timestamp(self):
        assert _compare_timestamps("", "2024-01-01") is False

    def test_unparseable_falls_back_to_lexicographic(self):
        # "2024-12-01" > "2024-06-01" lexicographically → Excel newer
        assert _compare_timestamps("2024-12-01", "2024-06-01") is True

    def test_both_unparseable_equal(self):
        assert _compare_timestamps("bad-date", "bad-date") is None


# ---------------------------------------------------------------------------
# compute_sync – last_updated timestamp-based direction
# ---------------------------------------------------------------------------


def _mapping_with_ts(**kwargs) -> SectionMapping:
    """Build a mapping with last_updated_col configured."""
    defaults = dict(
        name="SYS_REQS",
        sdoc_file="sys.sdoc",
        mode="table",
        anchor="SYS_Reqs",
        uid_col="UID",
        title_col="Title",
        statement_col="Statement",
        sync_direction="both",
        conflict_resolution="excel",  # static fallback
        field_directions={},
        extra_cols={"Last Updated": "LAST_UPDATED"},
        last_updated_col="Last Updated",
    )
    defaults.update(kwargs)
    return SectionMapping(**defaults)


class TestComputeSyncLastUpdated:
    def test_excel_newer_wins(self):
        """Excel has a newer timestamp → Excel value propagates to SDoc."""
        e = [_req("SYS-001", statement="Excel stmt",
                  custom_fields={"LAST_UPDATED": "2024-06-15"})]
        s = [_req("SYS-001", statement="SDoc stmt",
                  custom_fields={"LAST_UPDATED": "2024-06-10"})]
        _, sdoc_upd, _ = compute_sync(e, s, _mapping_with_ts())
        assert sdoc_upd[0].statement == "Excel stmt"

    def test_sdoc_newer_wins(self):
        """SDoc has a newer timestamp → SDoc value propagates to Excel."""
        e = [_req("SYS-001", statement="Excel stmt",
                  custom_fields={"LAST_UPDATED": "2024-06-10"})]
        s = [_req("SYS-001", statement="SDoc stmt",
                  custom_fields={"LAST_UPDATED": "2024-06-15"})]
        ex_upd, _, _ = compute_sync(e, s, _mapping_with_ts())
        assert ex_upd[0].statement == "SDoc stmt"

    def test_equal_timestamps_fall_back_to_conflict_resolution(self):
        """Equal timestamps → fall back to conflict_resolution='excel'."""
        ts = "2024-06-10"
        e = [_req("SYS-001", statement="Excel stmt",
                  custom_fields={"LAST_UPDATED": ts})]
        s = [_req("SYS-001", statement="SDoc stmt",
                  custom_fields={"LAST_UPDATED": ts})]
        _, sdoc_upd, _ = compute_sync(e, s, _mapping_with_ts(conflict_resolution="excel"))
        assert sdoc_upd[0].statement == "Excel stmt"

    def test_missing_timestamps_fall_back_to_conflict_resolution(self):
        """Missing timestamps → fall back to conflict_resolution='sdoc'."""
        e = [_req("SYS-001", statement="Excel stmt")]
        s = [_req("SYS-001", statement="SDoc stmt")]
        ex_upd, _, _ = compute_sync(e, s, _mapping_with_ts(conflict_resolution="sdoc"))
        assert ex_upd[0].statement == "SDoc stmt"

    def test_field_direction_overrides_timestamp(self):
        """field_directions override beats timestamp-based direction."""
        e = [_req("SYS-001", title="Excel Title", statement="Excel stmt",
                  custom_fields={"LAST_UPDATED": "2024-06-15"})]
        s = [_req("SYS-001", title="SDoc Title", statement="SDoc stmt",
                  custom_fields={"LAST_UPDATED": "2024-06-10"})]
        # Excel is newer → Excel should win for statement, but TITLE is forced sdoc_to_excel
        mapping = _mapping_with_ts(field_directions={"TITLE": "sdoc_to_excel"})
        _, sdoc_upd, _ = compute_sync(e, s, mapping)
        # TITLE: field_directions override → SDoc value kept
        assert sdoc_upd[0].title == "SDoc Title"
        # STATEMENT: no override, Excel newer → Excel value
        assert sdoc_upd[0].statement == "Excel stmt"

    def test_last_updated_not_configured_uses_conflict_resolution(self):
        """Without last_updated_col, conflict_resolution is used as before."""
        e = [_req("SYS-001", statement="Excel stmt")]
        s = [_req("SYS-001", statement="SDoc stmt")]
        mapping = _mapping(conflict_resolution="sdoc")
        ex_upd, _, _ = compute_sync(e, s, mapping)
        assert ex_upd[0].statement == "SDoc stmt"


# ---------------------------------------------------------------------------
# SyncResult
# ---------------------------------------------------------------------------


class TestSyncResult:
    def test_total_changes(self):
        r = SyncResult(section="S", sdoc_added=2, sdoc_updated=1, excel_added=0, excel_updated=3)
        assert r.total_changes == 6

    def test_summary_no_changes(self):
        r = SyncResult(section="S")
        assert "no changes" in r.summary()

    def test_summary_with_changes(self):
        r = SyncResult(section="S", sdoc_added=1, excel_updated=2)
        s = r.summary()
        assert "SDoc" in s
        assert "Excel" in s
