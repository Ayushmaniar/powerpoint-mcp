"""
Tool descriptions for PowerPoint MCP Server.

Each tool has its description in a separate module for better organization.
"""

from .manage_presentation import DESCRIPTION as MANAGE_PRESENTATION
from .slide_snapshot import DESCRIPTION as SLIDE_SNAPSHOT
from .switch_slide import DESCRIPTION as SWITCH_SLIDE
from .add_speaker_notes import DESCRIPTION as ADD_SPEAKER_NOTES
from .list_templates import DESCRIPTION as LIST_TEMPLATES
from .analyze_template import DESCRIPTION as ANALYZE_TEMPLATE
from .add_slide_with_layout import DESCRIPTION as ADD_SLIDE_WITH_LAYOUT
from .populate_placeholder import DESCRIPTION as POPULATE_PLACEHOLDER
from .manage_slide import DESCRIPTION as MANAGE_SLIDE
from .powerpoint_evaluate import DESCRIPTION as POWERPOINT_EVALUATE

__all__ = [
    'MANAGE_PRESENTATION',
    'SLIDE_SNAPSHOT',
    'SWITCH_SLIDE',
    'ADD_SPEAKER_NOTES',
    'LIST_TEMPLATES',
    'ANALYZE_TEMPLATE',
    'ADD_SLIDE_WITH_LAYOUT',
    'POPULATE_PLACEHOLDER',
    'MANAGE_SLIDE',
    'POWERPOINT_EVALUATE',
]
