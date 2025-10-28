"""
PowerPoint slide management tool for MCP server.
Provides comprehensive slide operations: duplicate, delete, and move.
"""

import win32com.client
from typing import Optional


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
        # 1. Connect to PowerPoint (following established pattern)
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        if not ppt_app.Presentations.Count:
            return {"error": "No PowerPoint presentation is open. Please open a presentation first."}

        active_presentation = ppt_app.ActivePresentation
        original_slide_count = active_presentation.Slides.Count

        # 2. Validate slide_number parameter
        if slide_number < 1 or slide_number > original_slide_count:
            return {"error": f"Invalid slide number {slide_number}. Must be between 1 and {original_slide_count}."}

        # 3. Validate operation parameter
        valid_operations = ["duplicate", "delete", "move"]
        if operation not in valid_operations:
            return {"error": f"Invalid operation '{operation}'. Must be one of: {', '.join(valid_operations)}."}

        # 4. Execute the requested operation
        if operation == "duplicate":
            return _duplicate_slide(active_presentation, slide_number, target_position, ppt_app)

        elif operation == "delete":
            return _delete_slide(active_presentation, slide_number, ppt_app)

        elif operation == "move":
            if target_position is None:
                return {"error": "target_position is required for 'move' operation."}
            return _move_slide(active_presentation, slide_number, target_position, ppt_app)

    except Exception as e:
        return {"error": f"Failed to manage slide: {str(e)}"}


def _duplicate_slide(presentation, slide_number: int, target_position: Optional[int], ppt_app) -> dict:
    """Duplicate a slide using PowerPoint COM automation."""
    try:
        original_count = presentation.Slides.Count

        # Get the slide to duplicate
        source_slide = presentation.Slides(slide_number)

        # Duplicate the slide (creates it immediately after the original)
        duplicated_slide = source_slide.Duplicate()

        # If target_position is specified and different from default position
        if target_position is not None:
            # The duplicate was created at slide_number + 1
            current_position = slide_number + 1

            # Validate target position
            new_count = presentation.Slides.Count
            if target_position < 1 or target_position > new_count:
                return {"error": f"Invalid target_position {target_position}. Must be between 1 and {new_count}."}

            # Move the duplicated slide to target position if different
            if target_position != current_position:
                # Get the duplicated slide and move it
                slide_to_move = presentation.Slides(current_position)
                slide_to_move.MoveTo(target_position)
                final_position = target_position
            else:
                final_position = current_position
        else:
            # Default position (immediately after original)
            final_position = slide_number + 1

        # Switch to the duplicated slide
        _switch_to_slide(ppt_app, final_position)

        return {
            "success": True,
            "operation": "duplicate",
            "original_slide": slide_number,
            "new_slide": final_position,
            "original_slide_count": original_count,
            "new_slide_count": presentation.Slides.Count,
            "message": f"Duplicated slide {slide_number} to position {final_position}"
        }

    except Exception as e:
        return {"error": f"Failed to duplicate slide: {str(e)}"}


def _delete_slide(presentation, slide_number: int, ppt_app) -> dict:
    """Delete a slide using PowerPoint COM automation."""
    try:
        original_count = presentation.Slides.Count

        # Cannot delete the last remaining slide
        if original_count == 1:
            return {"error": "Cannot delete the last remaining slide in the presentation."}

        # Get the slide to delete
        slide_to_delete = presentation.Slides(slide_number)

        # Delete the slide
        slide_to_delete.Delete()

        new_count = presentation.Slides.Count

        # Determine which slide to switch to after deletion
        if slide_number <= new_count:
            # Switch to the slide that took the deleted slide's position
            next_slide = slide_number
        else:
            # If we deleted the last slide, switch to the new last slide
            next_slide = new_count

        # Switch to the appropriate slide
        _switch_to_slide(ppt_app, next_slide)

        return {
            "success": True,
            "operation": "delete",
            "deleted_slide": slide_number,
            "current_slide": next_slide,
            "original_slide_count": original_count,
            "new_slide_count": new_count,
            "message": f"Deleted slide {slide_number}. Now viewing slide {next_slide}."
        }

    except Exception as e:
        return {"error": f"Failed to delete slide: {str(e)}"}


def _move_slide(presentation, slide_number: int, target_position: int, ppt_app) -> dict:
    """Move a slide to a new position using PowerPoint COM automation."""
    try:
        slide_count = presentation.Slides.Count

        # Validate target position
        if target_position < 1 or target_position > slide_count:
            return {"error": f"Invalid target_position {target_position}. Must be between 1 and {slide_count}."}

        # No need to move if already in target position
        if slide_number == target_position:
            return {
                "success": True,
                "operation": "move",
                "slide_number": slide_number,
                "target_position": target_position,
                "message": f"Slide {slide_number} is already at position {target_position}. No move needed."
            }

        # Get the slide to move
        slide_to_move = presentation.Slides(slide_number)

        # Move the slide to target position
        slide_to_move.MoveTo(target_position)

        # Switch to the moved slide at its new position
        _switch_to_slide(ppt_app, target_position)

        return {
            "success": True,
            "operation": "move",
            "slide_number": slide_number,
            "original_position": slide_number,
            "new_position": target_position,
            "total_slides": slide_count,
            "message": f"Moved slide from position {slide_number} to position {target_position}"
        }

    except Exception as e:
        return {"error": f"Failed to move slide: {str(e)}"}


def _switch_to_slide(ppt_app, slide_number: int):
    """Switch to a specific slide in the presentation."""
    try:
        if hasattr(ppt_app, 'ActiveWindow') and ppt_app.ActiveWindow:
            active_window = ppt_app.ActiveWindow
            if hasattr(active_window, 'View'):
                view = active_window.View
                if hasattr(view, 'GotoSlide'):
                    view.GotoSlide(slide_number)
                elif hasattr(view, 'Slide'):
                    # Alternative method for some PowerPoint versions
                    view.Slide = ppt_app.ActivePresentation.Slides(slide_number)
    except Exception:
        # Don't fail the operation if slide switching fails
        pass


def generate_mcp_response(result):
    """Generate the MCP tool response for the LLM."""
    if not result.get('success'):
        return f"Failed to manage slide: {result.get('error')}"

    operation = result['operation']

    if operation == "duplicate":
        response_lines = [
            f"✅ Duplicated slide {result['original_slide']} to position {result['new_slide']}",
            f"Total slides: {result['new_slide_count']} (increased from {result['original_slide_count']})",
            f"Currently viewing the duplicated slide at position {result['new_slide']}"
        ]

    elif operation == "delete":
        response_lines = [
            f"✅ Deleted slide {result['deleted_slide']}",
            f"Total slides: {result['new_slide_count']} (decreased from {result['original_slide_count']})",
            f"Currently viewing slide {result['current_slide']}"
        ]

    elif operation == "move":
        if result.get('original_position') == result.get('new_position'):
            response_lines = [f"✅ Slide {result['slide_number']} was already at position {result['target_position']}"]
        else:
            response_lines = [
                f"✅ Moved slide from position {result['original_position']} to position {result['new_position']}",
                f"Total slides: {result['total_slides']}",
                f"Currently viewing the moved slide at position {result['new_position']}"
            ]

    return "\n".join(response_lines)