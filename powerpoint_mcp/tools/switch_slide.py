"""
PowerPoint slide switching tool.
"""

from ..backends import get_backend


def powerpoint_switch_slide(slide_number: int) -> dict:
    """
    Switch to a specific slide in the active PowerPoint presentation.

    Args:
        slide_number: Slide number to switch to (1-based)

    Returns:
        Dictionary with success status and slide information or error message
    """
    try:
        backend = get_backend()
        backend.connect()

        if not backend.get_presentation_count():
            return {"error": "No PowerPoint presentation is open"}

        slide_info = backend.goto_slide(slide_number)

        return {
            "success": True,
            "slide_number": slide_info.slide_number,
            "total_slides": slide_info.total_slides,
            "slide_name": slide_info.name,
            "message": f"Switched to slide {slide_info.slide_number} of {slide_info.total_slides}"
        }

    except ValueError as e:
        return {"error": str(e)}
    except Exception as e:
        return {"error": f"Failed to switch slide: {str(e)}"}
