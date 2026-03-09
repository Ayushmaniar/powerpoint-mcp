"""
Shared data types for cross-platform PowerPoint backend abstraction.

These dataclasses provide platform-agnostic data exchange between backends and tools.
"""

from dataclasses import dataclass, field
from typing import Optional


class UnsupportedFeatureError(Exception):
    """Raised when a feature is not supported on the current platform."""
    pass


@dataclass
class PresentationInfo:
    name: str
    full_path: str
    slide_count: int


@dataclass
class SlideInfo:
    slide_number: int
    total_slides: int
    name: str
    layout_name: str
    shape_count: int


@dataclass
class TextRun:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    color_hex: Optional[str] = None  # e.g. "#ff0000"
    font_name: Optional[str] = None
    font_size: Optional[float] = None


@dataclass
class ShapeInfo:
    name: str
    id: int
    type_name: str
    type_value: int
    left: float
    top: float
    width: float
    height: float
    text: Optional[str] = None
    html_text: Optional[str] = None
    text_runs: Optional[list[TextRun]] = None
    font_info: Optional[str] = None
    hyperlinks: Optional[list[dict]] = None
    # Table data
    is_table: bool = False
    table_info: Optional[str] = None
    table_content: Optional[list[list[str]]] = None
    table_content_html: Optional[list[list[str]]] = None
    table_hyperlinks: Optional[list[dict]] = None
    container_type: Optional[str] = None
    # Chart data
    chart_info: Optional[str] = None
    chart_title: Optional[str] = None
    chart_data: Optional[dict] = None
    chart_error: Optional[str] = None
    table_error: Optional[str] = None


@dataclass
class TableData:
    rows: int
    columns: int
    cells_plain: list[list[str]]
    cells_html: list[list[str]]
    hyperlinks: Optional[list[dict]] = None


@dataclass
class ChartData:
    chart_type: int
    has_title: bool
    title: str
    series: Optional[list[dict]] = None
    categories: Optional[list[str]] = None
    axes: Optional[dict] = None
    legend: Optional[dict] = None


@dataclass
class CommentInfo:
    text: str
    author: str
    date: str
    position: str
    associated_object: Optional[dict] = None


@dataclass
class LayoutInfo:
    index: int
    name: str


@dataclass
class PlaceholderInfo:
    index: int
    type_value: int
    type_name: str
    name: str
    position: str
    size: str


@dataclass
class TemplateDir:
    path: str
    dir_type: str  # 'personal', 'user', 'system', 'other'


@dataclass
class FeatureSupport:
    latex_equations: bool = False
    animations: bool = False
    animation_by_paragraph: bool = False
    raw_evaluate: bool = False
    hidden_presentations: bool = False
    character_formatting: bool = False
