"""Skills library - provides MCP tool access within powerpoint_evaluate execution context."""

from typing import Optional
from .presentation import manage_presentation
from .snapshot import powerpoint_snapshot
from .switch_slide import powerpoint_switch_slide
from .add_speaker_notes import powerpoint_add_speaker_notes
from .list_templates import powerpoint_list_templates
from .analyze_template import powerpoint_analyze_template
from .add_slide_with_layout import powerpoint_add_slide_with_layout
from .populate_placeholder import powerpoint_populate_placeholder
from .manage_slide import powerpoint_manage_slide


class Skills:
    """Wrapper providing access to all PowerPoint MCP tools within evaluate code."""
    
    @staticmethod
    def populate_placeholder(placeholder_name: str, content: str, content_type: str = "auto", slide_number: Optional[int] = None) -> dict:
        return powerpoint_populate_placeholder(placeholder_name, content, content_type, slide_number)
    
    @staticmethod
    def snapshot(slide_number: Optional[int] = None, include_screenshot: bool = True, screenshot_filename: Optional[str] = None) -> dict:
        return powerpoint_snapshot(slide_number, include_screenshot, screenshot_filename)
    
    @staticmethod
    def switch_slide(slide_number: int) -> dict:
        return powerpoint_switch_slide(slide_number)
    
    @staticmethod
    def add_speaker_notes(slide_number: int, notes_text: str) -> dict:
        return powerpoint_add_speaker_notes(slide_number, notes_text)
    
    @staticmethod
    def list_templates() -> dict:
        return powerpoint_list_templates()
    
    @staticmethod
    def analyze_template(source: str = "current") -> dict:
        return powerpoint_analyze_template(source)
    
    @staticmethod
    def add_slide_with_layout(template_name: str, layout_name: str, after_slide: int) -> dict:
        return powerpoint_add_slide_with_layout(template_name, layout_name, after_slide)
    
    @staticmethod
    def manage_presentation(action: str, file_path: Optional[str] = None, save_path: Optional[str] = None, template_path: Optional[str] = None, presentation_name: Optional[str] = None) -> str:
        return manage_presentation(action, file_path, save_path, template_path, presentation_name)
    
    @staticmethod
    def manage_slide(operation: str, slide_number: int, target_position: Optional[int] = None) -> dict:
        return powerpoint_manage_slide(operation, slide_number, target_position)


# Singleton instance available in evaluate execution context
skills = Skills()
