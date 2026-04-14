"""Internal data model for a requirement, decoupling Excel from SDoc."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Requirement:
    """A single requirement, independent of source format.

    This class serves as the intermediate representation when synchronizing
    between Excel and StrictDoc, keeping the two backends fully decoupled.

    Attributes:
        uid: Unique identifier (e.g. ``SYS-001``).
        title: Short human-readable title.
        statement: Full requirement statement text.
        custom_fields: Extra named fields beyond UID/TITLE/STATEMENT, mapping
            SDoc field name → value.
        relations: Parent requirement UIDs that this requirement traces to.
    """

    uid: str
    title: str = ""
    statement: str = ""
    custom_fields: dict[str, str] = field(default_factory=dict)
    relations: list[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        if not self.uid or not self.uid.strip():
            raise ValueError("Requirement UID must not be empty.")

    @property
    def has_relations(self) -> bool:
        """Return ``True`` if this requirement references any parent UIDs."""
        return bool(self.relations)

    def to_dict(self) -> dict[str, object]:
        """Serialize the requirement to a plain dictionary."""
        return {
            "uid": self.uid,
            "title": self.title,
            "statement": self.statement,
            "custom_fields": dict(self.custom_fields),
            "relations": list(self.relations),
        }

    @classmethod
    def from_dict(cls, data: dict[str, object]) -> "Requirement":
        """Deserialize a requirement from a plain dictionary."""
        return cls(
            uid=str(data["uid"]),
            title=str(data.get("title", "")),
            statement=str(data.get("statement", "")),
            custom_fields=dict(data.get("custom_fields", {})),  # type: ignore[arg-type]
            relations=list(data.get("relations", [])),  # type: ignore[arg-type]
        )
