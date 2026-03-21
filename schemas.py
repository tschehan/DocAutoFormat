"""Project data structures for the demo project."""

from dataclasses import dataclass, field


STRUCTURE_ROLES = [
    "body",
    "paper_title",
    "subtitle",
    "heading_1",
    "heading_2",
    "heading_3",
    "keyword_line",
    "figure_title",
    "table_title",
    "table_cell",
]

REQUIREMENT_ROLES = [
    "paper_title",
    "subtitle",
    "heading_1",
    "heading_2",
    "heading_3",
    "body",
]

REQUIREMENT_SCRIPTS = ["zh", "en"]

SUPPORTED_FORMAT_FIELDS = [
    "font_name",
    "western_font_name",
    "font_size",
    "bold",
    "line_spacing",
    "alignment",
]


@dataclass
class DocumentBlock:
    """One extracted block from the document."""

    block_id: str
    text: str
    role: str | None = None
    table_id: str | None = None
    row: int | None = None
    col: int | None = None


@dataclass
class StructureResult:
    """Structure-recognition result."""

    blocks: list[DocumentBlock] = field(default_factory=list)


@dataclass
class FormatRule:
    """Formatting rule for one role/script pair."""

    role: str
    script: str | None = None
    font_name: str | None = None
    western_font_name: str | None = None
    font_size: str | None = None
    bold: bool | None = None
    line_spacing: str | None = None
    alignment: str | None = None


@dataclass
class RequirementResult:
    """Requirement-understanding result."""

    rules: list[FormatRule] = field(default_factory=list)
    ignored_requirements: list[str] = field(default_factory=list)


@dataclass
class PlanAction:
    """One concrete action in the execution plan."""

    target_block_id: str
    role: str
    font_name: str | None = None
    western_font_name: str | None = None
    font_size: str | None = None
    bold: bool | None = None
    line_spacing: str | None = None
    alignment: str | None = None


@dataclass
class ExecutionPlan:
    """Full execution plan."""

    actions: list[PlanAction] = field(default_factory=list)
