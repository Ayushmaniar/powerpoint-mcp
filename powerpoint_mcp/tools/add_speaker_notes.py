"""
PowerPoint speaker notes addition tool.
"""

from typing import Optional

from ..backends import get_backend


def get_current_slide_index(ppt_app=None):
    """Get the index of the currently selected/active slide.

    If ppt_app is provided (legacy COM object), delegates to backend.
    """
    backend = get_backend()
    return backend.get_current_slide_index()


def powerpoint_add_speaker_notes(slide_number: Optional[int] = None, notes_text: str = "") -> dict:
    """
    Add speaker notes to a specific slide in the active PowerPoint presentation.

    Args:
        slide_number: Slide number to add notes to (1-based). If None, uses current active slide.
        notes_text: Text content to add as speaker notes

    Returns:
        Dictionary with success status and slide information or error message
    """
    try:
        backend = get_backend()
        backend.connect()

        if not backend.get_presentation_count():
            return {"error": "No PowerPoint presentation is open"}

        # Determine slide to use
        if slide_number is None:
            slide_number = backend.get_current_slide_index()
            if slide_number is None:
                slide_number = 1

        # Validate slide number
        slide_count = backend.get_slide_count()
        if slide_number < 1 or slide_number > slide_count:
            return {"error": f"Invalid slide number {slide_number}. Presentation has {slide_count} slides."}

        backend.set_speaker_notes(slide_number, notes_text)

        return {
            "success": True,
            "slide_number": slide_number,
            "total_slides": slide_count,
            "notes_length": len(notes_text),
            "message": f"Added speaker notes to slide {slide_number}",
        }

    except Exception as e:
        return {"error": f"Failed to add speaker notes: {str(e)}"}
