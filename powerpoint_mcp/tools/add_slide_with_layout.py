"""
PowerPoint slide creation tool for MCP server.
Creates slides with specific template layouts at specified positions.
"""

import win32com.client
from typing import Optional

# Import reusable functions from existing tools
from .analyze_template import get_template_directories, find_template_by_name


def find_layout_in_template(template_path, layout_name):
    """
    Find a specific layout by name within a template.

    Args:
        template_path: Path to the template file
        layout_name: Name of the layout to find

    Returns:
        Layout object and temp presentation, or (None, None) if not found
    """
    try:
        # Get or create PowerPoint application
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        # Create hidden temporary presentation to access template layouts
        temp_presentation = ppt_app.Presentations.Add(WithWindow=False)
        temp_presentation.ApplyTemplate(template_path)

        # Search for layout by name (case-insensitive)
        for i in range(1, temp_presentation.SlideMaster.CustomLayouts.Count + 1):
            layout = temp_presentation.SlideMaster.CustomLayouts(i)
            if layout.Name.lower() == layout_name.lower():
                return layout, temp_presentation

        # Layout not found - clean up temp presentation
        temp_presentation.Close()
        return None, None

    except Exception as e:
        return None, None


def powerpoint_add_slide_with_layout(template_name: str, layout_name: str, after_slide: int) -> dict:
    """
    Add a slide with a specific template layout at the specified position.

    Args:
        template_name: Name of the template (e.g., "Pitchbook", "Training")
        layout_name: Name of the layout within the template (e.g., "Title", "Agenda")
        after_slide: Insert slide after this position (creates slide at after_slide + 1)

    Returns:
        Dictionary with success status and slide information or error message
    """
    temp_presentation = None

    try:
        # 1. Connect to PowerPoint (reusing pattern from add_speaker_notes.py)
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        if not ppt_app.Presentations.Count:
            return {"error": "No PowerPoint presentation is open. Please open a presentation first."}

        active_presentation = ppt_app.ActivePresentation
        original_slide_count = active_presentation.Slides.Count

        # 2. Validate after_slide parameter
        if after_slide < 0 or after_slide > original_slide_count:
            return {"error": f"Invalid after_slide position {after_slide}. Must be between 0 and {original_slide_count}."}

        # 3. Resolve template name to file path (reusing function from analyze_template.py)
        template_path = find_template_by_name(template_name)
        if not template_path:
            return {"error": f"Template '{template_name}' not found. Use list_templates() to see available templates."}

        # 4. Find layout in template
        target_layout, temp_presentation = find_layout_in_template(template_path, layout_name)
        if not target_layout:
            return {"error": f"Layout '{layout_name}' not found in template '{template_name}'. Use analyze_template(source='{template_name}') to see available layouts."}

        # Store layout name before closing temp presentation (to avoid COM object reference issues)
        actual_layout_name = target_layout.Name

        # 5. Create slide from template without affecting existing slides (safe copy/paste approach)
        temp_slide = temp_presentation.Slides.AddSlide(1, target_layout)

        # Copy the slide from temp presentation
        temp_slide.Copy()

        # Calculate new slide position
        new_slide_position = after_slide + 1

        # Paste at the correct position in active presentation
        if new_slide_position <= active_presentation.Slides.Count:
            # Insert at specific position
            active_presentation.Slides.Paste(new_slide_position)
        else:
            # Add at the end
            active_presentation.Slides.Paste()

        new_slide_count = active_presentation.Slides.Count

        # 6. Switch to the newly created slide
        try:
            # Get the active window and switch to the new slide
            if hasattr(ppt_app, 'ActiveWindow') and ppt_app.ActiveWindow:
                active_window = ppt_app.ActiveWindow
                if hasattr(active_window, 'View'):
                    view = active_window.View
                    if hasattr(view, 'GotoSlide'):
                        view.GotoSlide(new_slide_position)
                    elif hasattr(view, 'Slide'):
                        # Alternative method for some PowerPoint versions
                        view.Slide = active_presentation.Slides(new_slide_position)
        except Exception:
            # Don't fail the whole operation if slide switching fails
            pass

        # 7. Clean up temp presentation
        if temp_presentation:
            temp_presentation.Close()
            temp_presentation = None

        # 8. Return success result (following pattern from add_speaker_notes.py)
        return {
            "success": True,
            "new_slide_number": new_slide_position,
            "layout_name": actual_layout_name,
            "template_name": template_name,
            "original_slide_count": original_slide_count,
            "new_slide_count": new_slide_count,
            "total_slides": new_slide_count,
            "message": f"Added slide {new_slide_position} using '{actual_layout_name}' layout from '{template_name}' template and switched to it"
        }

    except Exception as e:
        # Always clean up temp presentation on error
        if temp_presentation:
            try:
                temp_presentation.Close()
            except:
                pass

        return {"error": f"Failed to add slide with layout: {str(e)}"}


def generate_mcp_response(result):
    """Generate the MCP tool response for the LLM."""
    if not result.get('success'):
        return f"Failed to add slide: {result.get('error')}"

    # Create clean response for LLM
    response_lines = [
        f"✅ Added slide {result['new_slide_number']} using '{result['layout_name']}' layout from '{result['template_name']}' template",
        f"Position: Inserted after slide {result['new_slide_number'] - 1}",
        f"Total slides: {result['new_slide_count']} (increased from {result['original_slide_count']})",
        f"Ready for content population using populate_slide_content(slide_number={result['new_slide_number']})"
    ]

    return "\n".join(response_lines)