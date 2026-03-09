"""
Abstract base class for PowerPoint backends.

Defines the interface that platform-specific backends must implement.
Methods map to "what I need to do with PowerPoint", not individual COM/JXA property accesses.
"""

from abc import ABC, abstractmethod
from contextlib import contextmanager
from typing import Optional

from .types import (
    PresentationInfo, SlideInfo, ShapeInfo, CommentInfo,
    LayoutInfo, PlaceholderInfo, TemplateDir, FeatureSupport,
    UnsupportedFeatureError,
)


class PowerPointBackend(ABC):
    """Abstract base class for PowerPoint automation backends."""

    # --- Application lifecycle ---

    @abstractmethod
    def connect(self):
        """Connect to or launch PowerPoint. Called before any operation."""
        ...

    @abstractmethod
    def get_presentation_count(self) -> int:
        """Return the number of open presentations."""
        ...

    @abstractmethod
    def get_active_presentation_info(self) -> PresentationInfo:
        """Return info about the active presentation."""
        ...

    # --- Presentation management ---

    @abstractmethod
    def open_presentation(self, file_path: str) -> PresentationInfo:
        """Open a presentation file. Returns info about the opened presentation."""
        ...

    @abstractmethod
    def close_presentation(self, presentation_name: Optional[str] = None) -> str:
        """Close a presentation. If name is None, close the active one."""
        ...

    @abstractmethod
    def create_presentation(self, file_path: Optional[str] = None, template_path: Optional[str] = None) -> PresentationInfo:
        """Create a new presentation, optionally from a template, optionally saving immediately."""
        ...

    @abstractmethod
    def save_presentation(self) -> str:
        """Save the active presentation at its current location."""
        ...

    @abstractmethod
    def save_presentation_as(self, save_path: str) -> str:
        """Save the active presentation to a new location."""
        ...

    # --- Slide navigation ---

    @abstractmethod
    def get_current_slide_index(self) -> Optional[int]:
        """Get the 1-based index of the currently active slide, or None."""
        ...

    @abstractmethod
    def get_slide_count(self) -> int:
        """Return the number of slides in the active presentation."""
        ...

    @abstractmethod
    def goto_slide(self, slide_number: int) -> SlideInfo:
        """Switch the view to the specified slide. Returns info about the slide."""
        ...

    # --- Slide management ---

    @abstractmethod
    def duplicate_slide(self, slide_number: int, target_position: Optional[int] = None) -> dict:
        """Duplicate a slide. Returns dict with operation details."""
        ...

    @abstractmethod
    def delete_slide(self, slide_number: int) -> dict:
        """Delete a slide. Returns dict with operation details."""
        ...

    @abstractmethod
    def move_slide(self, slide_number: int, target_position: int) -> dict:
        """Move a slide to a new position. Returns dict with operation details."""
        ...

    # --- Shape/content reading ---

    @abstractmethod
    def get_slide_info(self, slide_number: int) -> SlideInfo:
        """Get basic info about a slide."""
        ...

    @abstractmethod
    def get_shapes(self, slide_number: int) -> list[ShapeInfo]:
        """Get all shapes on a slide with full detail (text, tables, charts)."""
        ...

    @abstractmethod
    def get_speaker_notes(self, slide_number: int) -> Optional[str]:
        """Get speaker notes for a slide (HTML formatted). Returns None if no notes."""
        ...

    @abstractmethod
    def get_speaker_notes_plain(self, slide_number: int) -> Optional[str]:
        """Get speaker notes as plain text. Returns None if no notes."""
        ...

    @abstractmethod
    def get_comments(self, slide_number: int) -> list[CommentInfo]:
        """Get all comments on a slide."""
        ...

    @abstractmethod
    def export_slide_image(self, slide_number: int, output_path: str):
        """Export a slide as a PNG image to the specified path."""
        ...

    @abstractmethod
    def get_slide_dimensions(self) -> tuple[float, float]:
        """Return (width, height) of slides in the active presentation."""
        ...

    # --- Content writing ---

    @abstractmethod
    def set_text(self, slide_number: int, shape_name: str, text: str):
        """Set the text of a shape, replacing any existing text."""
        ...

    @abstractmethod
    def apply_character_formatting(self, slide_number: int, shape_name: str, segments: list[dict]):
        """Apply formatting segments to a shape's text range.
        Each segment has 'start' (1-based), 'length', and 'formatting' dict."""
        ...

    @abstractmethod
    def clear_bullets(self, slide_number: int, shape_name: str):
        """Remove default bullet formatting from a shape."""
        ...

    @abstractmethod
    def insert_image(self, slide_number: int, shape_name: str, image_path: str,
                     matplotlib_code: Optional[str] = None) -> dict:
        """Replace a shape with an image, maintaining aspect ratio within the shape's bounds.
        Returns dict with new shape info."""
        ...

    @abstractmethod
    def set_speaker_notes(self, slide_number: int, notes_text: str):
        """Set the speaker notes for a slide."""
        ...

    # --- LaTeX ---

    def convert_latex_to_equation(self, slide_number: int, shape_name: str, latex_segments: list[dict]):
        """Convert LaTeX segments in a shape's text to native equation objects.
        Default implementation raises UnsupportedFeatureError."""
        raise UnsupportedFeatureError("LaTeX equation conversion is not supported on this platform.")

    # --- Animation ---

    @abstractmethod
    def add_animation_effect(self, slide_number: int, shape_name: str, effect_id: int,
                             level: int = 0, trigger: int = 1, duration: float = 0.5) -> int:
        """Add an animation effect to a shape. Returns the total animation count."""
        ...

    @abstractmethod
    def remove_shape_animations(self, slide_number: int, shape_name: str) -> int:
        """Remove all animations for a shape. Returns count of removed animations."""
        ...

    @abstractmethod
    def get_paragraph_count(self, slide_number: int, shape_name: str) -> int:
        """Get the number of paragraphs in a shape's text."""
        ...

    # --- Templates ---

    @abstractmethod
    def get_template_directories(self) -> list[TemplateDir]:
        """Return platform-appropriate template directories that exist."""
        ...

    @abstractmethod
    def get_layouts(self) -> list[LayoutInfo]:
        """Get available layouts from the active presentation's slide master."""
        ...

    @abstractmethod
    def add_slide_with_layout(self, template_path: str, layout_name: str, after_slide: int) -> dict:
        """Add a slide using a template layout. Returns dict with new slide info."""
        ...

    # --- Hidden presentations (for template analysis) ---

    @contextmanager
    def hidden_presentation(self, template_path: str):
        """Context manager that yields a HiddenPresentation helper for template analysis.
        Default implementation raises UnsupportedFeatureError."""
        raise UnsupportedFeatureError("Hidden presentations are not supported on this platform.")
        yield  # pragma: no cover

    # --- Raw evaluate support ---

    def get_feature_support(self) -> FeatureSupport:
        """Return which features are supported on this platform."""
        return FeatureSupport()

    def get_raw_context(self, slide_number: Optional[int] = None, shape_ref: Optional[str] = None) -> dict:
        """Return raw platform objects for evaluate tool. Returns {} by default (evaluate disabled)."""
        return {}


class HiddenPresentation:
    """Helper for operations on a hidden/temporary presentation."""

    @abstractmethod
    def get_layouts(self) -> list[LayoutInfo]:
        ...

    @abstractmethod
    def add_slide(self, layout_index: int):
        ...

    @abstractmethod
    def get_placeholders(self, slide_index: int) -> list[PlaceholderInfo]:
        ...

    @abstractmethod
    def populate_placeholder_defaults(self, slide_index: int):
        ...

    @abstractmethod
    def export_slide(self, slide_index: int, output_path: str):
        ...

    @abstractmethod
    def get_dimensions(self) -> tuple[float, float]:
        ...
