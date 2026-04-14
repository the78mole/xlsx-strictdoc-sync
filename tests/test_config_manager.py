"""Tests for the config manager."""

from __future__ import annotations

import pytest

from xlsx_strictdoc_sync.config_manager import (
    DIRECTION_BOTH,
    DIRECTION_EXCEL_TO_SDOC,
    DIRECTION_SDOC_TO_EXCEL,
    Config,
    SectionMapping,
    generate_config_template,
    load_config,
)


# ---------------------------------------------------------------------------
# SectionMapping validation
# ---------------------------------------------------------------------------


class TestSectionMapping:
    def test_minimal_creation(self):
        m = SectionMapping(
            name="SYS",
            sdoc_file="sys.sdoc",
            mode="table",
            anchor="SYS_Reqs",
            uid_col="UID",
        )
        assert m.sync_direction == DIRECTION_EXCEL_TO_SDOC
        assert m.conflict_resolution == "excel"
        assert m.field_directions == {}

    def test_invalid_mode(self):
        with pytest.raises(ValueError, match="mode"):
            SectionMapping(
                name="S", sdoc_file="s.sdoc", mode="invalid",
                anchor="A", uid_col="UID",
            )

    def test_invalid_sync_direction(self):
        with pytest.raises(ValueError, match="sync_direction"):
            SectionMapping(
                name="S", sdoc_file="s.sdoc", mode="table",
                anchor="A", uid_col="UID", sync_direction="unknown",
            )

    def test_invalid_conflict_resolution(self):
        with pytest.raises(ValueError, match="conflict_resolution"):
            SectionMapping(
                name="S", sdoc_file="s.sdoc", mode="table",
                anchor="A", uid_col="UID", conflict_resolution="other",
            )

    def test_invalid_field_direction_value(self):
        with pytest.raises(ValueError, match="field_directions"):
            SectionMapping(
                name="S", sdoc_file="s.sdoc", mode="table",
                anchor="A", uid_col="UID",
                sync_direction="both",
                field_directions={"TITLE": "wrong_value"},
            )

    def test_valid_field_directions(self):
        m = SectionMapping(
            name="S", sdoc_file="s.sdoc", mode="table",
            anchor="A", uid_col="UID",
            sync_direction="both",
            field_directions={
                "TITLE": "sdoc_to_excel",
                "STATEMENT": "excel_to_sdoc",
            },
        )
        assert m.field_directions["TITLE"] == "sdoc_to_excel"

    def test_empty_anchor_raises(self):
        with pytest.raises(ValueError, match="anchor"):
            SectionMapping(name="S", sdoc_file="s.sdoc", mode="table",
                           anchor="", uid_col="UID")

    def test_last_updated_col_not_in_extra_cols_defaults_to_LAST_UPDATED(self):
        m = SectionMapping(
            name="S", sdoc_file="s.sdoc", mode="table",
            anchor="A", uid_col="UID",
            last_updated_col="Modified",
            extra_cols={},  # column not mapped in extra_cols
        )
        assert m.last_updated_sdoc_field == "LAST_UPDATED"




# ---------------------------------------------------------------------------
# load_config
# ---------------------------------------------------------------------------


class TestLoadConfig:
    def _write_toml(self, tmp_path, content: str) -> str:
        p = tmp_path / "reqsync.toml"
        p.write_text(content, encoding="utf-8")
        return str(p)

    def test_minimal_valid(self, tmp_path):
        path = self._write_toml(tmp_path, """
[global]
excel_file = "reqs.xlsx"

[SYS_REQS]
sdoc_file = "sys.sdoc"
mode = "table"
anchor = "SYS_Reqs"
uid_col = "UID"
""")
        cfg = load_config(path)
        assert cfg.excel_file == "reqs.xlsx"
        assert len(cfg.sections) == 1
        assert cfg.sections[0].name == "SYS_REQS"

    def test_with_bidirectional_fields(self, tmp_path):
        path = self._write_toml(tmp_path, """
[global]
excel_file = "reqs.xlsx"

[SYS_REQS]
sdoc_file = "sys.sdoc"
mode = "table"
anchor = "SYS_Reqs"
uid_col = "UID"
sync_direction = "both"
conflict_resolution = "sdoc"

[SYS_REQS.field_directions]
TITLE = "sdoc_to_excel"
STATEMENT = "excel_to_sdoc"
""")
        cfg = load_config(path)
        m = cfg.sections[0]
        assert m.sync_direction == DIRECTION_BOTH
        assert m.conflict_resolution == "sdoc"
        assert m.field_directions["TITLE"] == DIRECTION_SDOC_TO_EXCEL
        assert m.field_directions["STATEMENT"] == DIRECTION_EXCEL_TO_SDOC

    def test_missing_excel_file_key(self, tmp_path):
        path = self._write_toml(tmp_path, "[global]\n")
        with pytest.raises(KeyError, match="excel_file"):
            load_config(path)

    def test_missing_section_key(self, tmp_path):
        path = self._write_toml(tmp_path, """
[global]
excel_file = "x.xlsx"

[SYS_REQS]
sdoc_file = "sys.sdoc"
mode = "table"
anchor = "T"
""")
        with pytest.raises(KeyError, match="uid_col"):
            load_config(path)

    def test_file_not_found(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            load_config(tmp_path / "missing.toml")

    def test_extra_cols_parsed(self, tmp_path):
        path = self._write_toml(tmp_path, """
[global]
excel_file = "x.xlsx"

[SYS_REQS]
sdoc_file = "sys.sdoc"
mode = "legacy"
anchor = "Sheet1"
uid_col = "A"

[SYS_REQS.extra_cols]
Safety = "SAFETY_LEVEL"
""")
        cfg = load_config(path)
        assert cfg.sections[0].extra_cols == {"Safety": "SAFETY_LEVEL"}

    def test_last_updated_col_parsed(self, tmp_path):
        path = self._write_toml(tmp_path, """
[global]
excel_file = "x.xlsx"

[SYS_REQS]
sdoc_file = "sys.sdoc"
mode = "table"
anchor = "SYS_Reqs"
uid_col = "UID"
last_updated_col = "Last Updated"

[SYS_REQS.extra_cols]
"Last Updated" = "LAST_UPDATED"
""")
        cfg = load_config(path)
        m = cfg.sections[0]
        assert m.last_updated_col == "Last Updated"
        assert m.last_updated_sdoc_field == "LAST_UPDATED"

    def test_last_updated_col_absent_gives_empty(self, tmp_path):
        path = self._write_toml(tmp_path, """
[global]
excel_file = "x.xlsx"

[SYS_REQS]
sdoc_file = "sys.sdoc"
mode = "table"
anchor = "SYS_Reqs"
uid_col = "UID"
""")
        cfg = load_config(path)
        m = cfg.sections[0]
        assert m.last_updated_col == ""
        assert m.last_updated_sdoc_field == ""

    def test_last_updated_col_not_in_extra_cols_defaults_to_LAST_UPDATED(self):
        m = SectionMapping(
            name="S", sdoc_file="s.sdoc", mode="table",
            anchor="A", uid_col="UID",
            last_updated_col="Modified",
            extra_cols={},  # column not mapped
        )
        assert m.last_updated_sdoc_field == "LAST_UPDATED"

    def test_multiple_sections(self, tmp_path):
        path = self._write_toml(tmp_path, """
[global]
excel_file = "x.xlsx"

[SYS_REQS]
sdoc_file = "sys.sdoc"
mode = "table"
anchor = "SYS"
uid_col = "UID"

[SW_REQS]
sdoc_file = "sw.sdoc"
mode = "legacy"
anchor = "SW"
uid_col = "A"
""")
        cfg = load_config(path)
        assert len(cfg.sections) == 2
        assert cfg.sections[0].name == "SYS_REQS"
        assert cfg.sections[1].name == "SW_REQS"


# ---------------------------------------------------------------------------
# generate_config_template
# ---------------------------------------------------------------------------


class TestGenerateConfigTemplate:
    def test_contains_excel_file(self):
        tmpl = generate_config_template("reqs.xlsx", [])
        assert 'excel_file = "reqs.xlsx"' in tmpl

    def test_section_names_included(self):
        sections = [
            {"name": "SYS_REQS", "mode": "table", "anchor": "SYS", "uid_col": "UID",
             "title_col": "Title", "statement_col": "Statement"},
        ]
        tmpl = generate_config_template("reqs.xlsx", sections)
        assert "[SYS_REQS]" in tmpl
        assert "sync_direction" in tmpl
        assert "field_directions" in tmpl

    def test_round_trip_parse(self, tmp_path):
        """Generated template should parse without errors after minimal edits."""
        sections = [
            {"name": "SYS_REQS", "mode": "table", "anchor": "SYS", "uid_col": "UID",
             "title_col": "Title", "statement_col": "Statement"},
        ]
        tmpl = generate_config_template("reqs.xlsx", sections)
        # Strip comment lines that would cause TOML parse errors (they start with #)
        # and write the mandatory keys
        path = tmp_path / "reqsync.toml"
        path.write_text(tmpl, encoding="utf-8")
        # load_config strips comment lines via TOML parser naturally
        cfg = load_config(str(path))
        assert cfg.excel_file == "reqs.xlsx"
