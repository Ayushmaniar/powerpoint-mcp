"""
PowerPoint presentation management tools.

This module provides tools for comprehensive PowerPoint presentation management.
"""

import os
from typing import Optional

from ..backends import get_backend


def manage_presentation(
    action: str,
    file_path: Optional[str] = None,
    save_path: Optional[str] = None,
    template_path: Optional[str] = None,
    presentation_name: Optional[str] = None
) -> str:
    """
    Comprehensive PowerPoint presentation management tool.

    Args:
        action: Action to perform - "open", "close", "create", "save", or "save_as"
        file_path: Path for open/create operations (required for open/create)
        save_path: New path for save_as operation (required for save_as)
        template_path: Template file for create operation (optional)
        presentation_name: Specific presentation name for close operation (optional)

    Returns:
        Success message with operation details, or error message
    """
    try:
        backend = get_backend()
        backend.connect()

        if action == "open":
            if not file_path:
                return "Error: file_path is required for open action"

            abs_file_path = os.path.abspath(file_path)
            if not os.path.exists(abs_file_path):
                return f"Error: File not found: {abs_file_path}"

            info = backend.open_presentation(abs_file_path)
            return f"Successfully opened '{info.name}' with {info.slide_count} slides"

        elif action == "close":
            if not backend.get_presentation_count():
                return "No presentations are currently open"

            try:
                msg = backend.close_presentation(presentation_name)
                return msg
            except ValueError as e:
                return f"Error: {e}"

        elif action == "create":
            if template_path:
                abs_template_path = os.path.abspath(template_path)
                if not os.path.exists(abs_template_path):
                    return f"Error: Template file not found: {abs_template_path}"
            else:
                abs_template_path = None

            abs_file_path = os.path.abspath(file_path) if file_path else None

            info = backend.create_presentation(
                file_path=abs_file_path,
                template_path=abs_template_path
            )

            if file_path:
                return f"Created presentation '{info.name}' at {abs_file_path}. Has 0 slides. Use add_slide_with_layout to add slides before populating content."
            else:
                return f"Created presentation '{info.name}' (not saved). Has 0 slides. Use add_slide_with_layout to add slides before populating content."

        elif action == "save":
            if not backend.get_presentation_count():
                return "Error: No presentations are open to save"

            try:
                return backend.save_presentation()
            except RuntimeError as e:
                return f"Error: {e}"

        elif action == "save_as":
            if not save_path:
                return "Error: save_path is required for save_as action"

            if not backend.get_presentation_count():
                return "Error: No presentations are open to save"

            return backend.save_presentation_as(save_path)

        else:
            return f"Error: Unknown action '{action}'. Valid actions: open, close, create, save, save_as"

    except Exception as e:
        return f"Error performing {action} action: {str(e)}"


# Keep the old function for backward compatibility (if needed)
def open_presentation(file_path: str) -> str:
    """Legacy function - use manage_presentation with action="open" instead."""
    return manage_presentation("open", file_path=file_path)
