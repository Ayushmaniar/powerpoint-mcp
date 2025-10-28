"""
PowerPoint presentation management tools.

This module provides tools for comprehensive PowerPoint presentation management.
"""

import os
from typing import Optional
import win32com.client


def manage_presentation(
    action: str,
    file_path: Optional[str] = None,
    save_path: Optional[str] = None,
    template_path: Optional[str] = None,
    presentation_name: Optional[str] = None
) -> str:
    """
    Comprehensive PowerPoint presentation management tool.

    This tool works on Windows only. Use Windows path format with double backslashes.

    Args:
        action: Action to perform - "open", "close", "create", "save", or "save_as"
        file_path: Path for open/create operations (required for open/create)
        save_path: New path for save_as operation (required for save_as)
        template_path: Template file for create operation (optional)
        presentation_name: Specific presentation name for close operation (optional)

    Actions:
        - "open": Opens an existing presentation (requires file_path)
        - "close": Closes a presentation (optional presentation_name, closes active if not specified)
        - "create": Creates new presentation (optional file_path for immediate save, optional template_path)
        - "save": Saves current presentation at its current location
        - "save_as": Saves current presentation to new location (requires save_path)

    Examples:
        manage_presentation("open", file_path="C:\\docs\\slides.pptx")
        manage_presentation("create", file_path="C:\\new\\presentation.pptx")
        manage_presentation("save")
        manage_presentation("save_as", save_path="C:\\backup\\slides_v2.pptx")
        manage_presentation("close")

    Returns:
        Success message with operation details, or error message
    """
    try:
        # Get or create PowerPoint application instance
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        ppt_app.Visible = True

        if action == "open":
            if not file_path:
                return "Error: file_path is required for open action"

            # Convert to absolute path
            abs_file_path = os.path.abspath(file_path)

            # Check if file exists
            if not os.path.exists(abs_file_path):
                return f"Error: File not found: {abs_file_path}"

            # Check if presentation is already open
            for presentation in ppt_app.Presentations:
                try:
                    if os.path.samefile(presentation.FullName, abs_file_path):
                        slide_count = presentation.Slides.Count
                        return f"Presentation '{presentation.Name}' is already open with {slide_count} slides"
                except (OSError, AttributeError):
                    continue

            # Open the presentation
            presentation = ppt_app.Presentations.Open(abs_file_path)
            slide_count = presentation.Slides.Count
            presentation_name = presentation.Name

            return f"Successfully opened '{presentation_name}' with {slide_count} slides"

        elif action == "close":
            if ppt_app.Presentations.Count == 0:
                return "No presentations are currently open"

            if presentation_name:
                # Close specific presentation by name
                for presentation in ppt_app.Presentations:
                    if presentation.Name == presentation_name:
                        presentation.Close()
                        return f"Successfully closed presentation '{presentation_name}'"
                return f"Error: Presentation '{presentation_name}' not found"
            else:
                # Close active presentation
                if ppt_app.ActivePresentation:
                    name = ppt_app.ActivePresentation.Name
                    ppt_app.ActivePresentation.Close()
                    return f"Successfully closed active presentation '{name}'"
                else:
                    return "Error: No active presentation to close"

        elif action == "create":
            # Create new presentation
            if template_path:
                # Create from template
                abs_template_path = os.path.abspath(template_path)
                if not os.path.exists(abs_template_path):
                    return f"Error: Template file not found: {abs_template_path}"
                presentation = ppt_app.Presentations.Open(abs_template_path)
            else:
                # Create blank presentation
                presentation = ppt_app.Presentations.Add()

            # Save immediately if file_path provided
            if file_path:
                abs_file_path = os.path.abspath(file_path)
                # Ensure directory exists
                os.makedirs(os.path.dirname(abs_file_path), exist_ok=True)
                presentation.SaveAs(abs_file_path)
                return f"Created presentation '{presentation.Name}' at {abs_file_path}. Has 0 slides. Use add_slide_with_layout to add slides before populating content."
            else:
                return f"Created presentation '{presentation.Name}' (not saved). Has 0 slides. Use add_slide_with_layout to add slides before populating content."

        elif action == "save":
            if ppt_app.Presentations.Count == 0:
                return "Error: No presentations are open to save"

            if not ppt_app.ActivePresentation:
                return "Error: No active presentation to save"

            presentation = ppt_app.ActivePresentation

            # Check if presentation has been saved before
            if hasattr(presentation, 'FullName') and presentation.FullName:
                presentation.Save()
                return f"Successfully saved '{presentation.Name}' to {presentation.FullName}"
            else:
                return f"Error: Presentation '{presentation.Name}' has never been saved. Use save_as action with save_path to save to a location."

        elif action == "save_as":
            if not save_path:
                return "Error: save_path is required for save_as action"

            if ppt_app.Presentations.Count == 0:
                return "Error: No presentations are open to save"

            if not ppt_app.ActivePresentation:
                return "Error: No active presentation to save"

            presentation = ppt_app.ActivePresentation
            abs_save_path = os.path.abspath(save_path)

            # Ensure directory exists
            os.makedirs(os.path.dirname(abs_save_path), exist_ok=True)
            presentation.SaveAs(abs_save_path)

            return f"Successfully saved '{presentation.Name}' to {abs_save_path}"

        else:
            return f"Error: Unknown action '{action}'. Valid actions: open, close, create, save, save_as"

    except Exception as e:
        return f"Error performing {action} action: {str(e)}"


# Keep the old function for backward compatibility (if needed)
def open_presentation(file_path: str) -> str:
    """Legacy function - use manage_presentation with action="open" instead."""
    return manage_presentation("open", file_path=file_path)