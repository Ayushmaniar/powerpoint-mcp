"""
PowerPoint slide management tool for MCP server.
Provides comprehensive slide operations: duplicate, delete, and move.
"""

from typing import Optional

from ..backends import get_backend


def powerpoint_manage_slide(operation: str, slide_number: int, target_position: Optional[int] = None) -> dict:
    """
    Manage slides in the active PowerPoint presentation.

    Args:
        operation: The operation to perform ("duplicate", "delete", or "move")
        slide_number: The slide number to operate on (1-based index)
        target_position: For 'move' operation - where to move the slide (1-based index)
                        For 'duplicate' operation - where to place the duplicate (optional, defaults to after original)

    Returns:
        Dictionary with success status and operation details or error message
    """
    try:
        backend = get_backend()
        backend.connect()

        if not backend.get_presentation_count():
            return {"error": "No PowerPoint presentation is open. Please open a presentation first."}

        slide_count = backend.get_slide_count()

        # Validate slide_number
        if slide_number < 1 or slide_number > slide_count:
            return {"error": f"Invalid slide number {slide_number}. Must be between 1 and {slide_count}."}

        # Validate operation
        valid_operations = ["duplicate", "delete", "move"]
        if operation not in valid_operations:
            return {"error": f"Invalid operation '{operation}'. Must be one of: {', '.join(valid_operations)}."}

        if operation == "duplicate":
            return backend.duplicate_slide(slide_number, target_position)
        elif operation == "delete":
            return backend.delete_slide(slide_number)
        elif operation == "move":
            if target_position is None:
                return {"error": "target_position is required for 'move' operation."}
            return backend.move_slide(slide_number, target_position)

    except (ValueError, RuntimeError) as e:
        return {"error": str(e)}
    except Exception as e:
        return {"error": f"Failed to manage slide: {str(e)}"}


def generate_mcp_response(result):
    """Generate the MCP tool response for the LLM."""
    if not result.get('success'):
        return f"Failed to manage slide: {result.get('error')}"

    operation = result['operation']

    if operation == "duplicate":
        response_lines = [
            f"Duplicated slide {result['original_slide']} to position {result['new_slide']}",
            f"Total slides: {result['new_slide_count']} (increased from {result['original_slide_count']})",
            f"Currently viewing the duplicated slide at position {result['new_slide']}"
        ]

    elif operation == "delete":
        response_lines = [
            f"Deleted slide {result['deleted_slide']}",
            f"Total slides: {result['new_slide_count']} (decreased from {result['original_slide_count']})",
            f"Currently viewing slide {result['current_slide']}"
        ]

    elif operation == "move":
        if result.get('original_position') == result.get('new_position'):
            response_lines = [f"Slide {result['slide_number']} was already at position {result['target_position']}"]
        else:
            response_lines = [
                f"Moved slide from position {result['original_position']} to position {result['new_position']}",
                f"Total slides: {result['total_slides']}",
                f"Currently viewing the moved slide at position {result['new_position']}"
            ]

    return "\n".join(response_lines)
