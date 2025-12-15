"""
PowerPoint slide creation tool for MCP server.
Creates slides with specific template layouts at specified positions.
"""

import win32com.client
from typing import Optional

# Import reusable functions from existing tools
from .analyze_template import get_template_directories, find_template_by_name


def powerpoint_add_slide_with_layout(template_name: str, layout_name: str, after_slide: int) -> dict:
    """
    Add a slide with a specific template layout at the specified position.
    Properly imports the template design to preserve all styling (backgrounds, fonts, colors, etc.).

    Args:
        template_name: Name of the template (e.g., "Pitchbook", "Training")
        layout_name: Name of the layout within the template (e.g., "Title", "Agenda")
        after_slide: Insert slide after this position (creates slide at after_slide + 1)

    Returns:
        Dictionary with success status and slide information or error message
    """
    try:
        # 1. Connect to PowerPoint
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

        # 3. Resolve template name to file path
        template_path = find_template_by_name(template_name)
        if not template_path:
            return {"error": f"Template '{template_name}' not found. Use list_templates() to see available templates."}

        # 4. Import the template design into the active presentation
        # This ensures all styling (backgrounds, fonts, colors) is preserved
        design = active_presentation.Designs.Load(template_path)

        # Get the slide master from the imported design
        slide_master = design.SlideMaster

        # 5. Find the layout by name in the imported design
        target_layout = None
        for i in range(1, slide_master.CustomLayouts.Count + 1):
            layout = slide_master.CustomLayouts(i)
            if layout.Name.lower() == layout_name.lower():
                target_layout = layout
                break

        if not target_layout:
            return {"error": f"Layout '{layout_name}' not found in template '{template_name}'. Use analyze_template(source='{template_name}') to see available layouts."}

        # 6. Add the slide with the imported layout
        new_slide_position = after_slide + 1

        # PowerPoint's Slides.AddSlide method: AddSlide(Index, pCustomLayout)
        # Index is 1-based, so position 1 means "first slide"
        # Note: Use AddSlide (not Add) for custom template layouts
        new_slide = active_presentation.Slides.AddSlide(new_slide_position, target_layout)

        new_slide_count = active_presentation.Slides.Count

        # 7. Switch to the newly created slide
        try:
            if hasattr(ppt_app, 'ActiveWindow') and ppt_app.ActiveWindow:
                active_window = ppt_app.ActiveWindow
                if hasattr(active_window, 'View'):
                    view = active_window.View
                    if hasattr(view, 'GotoSlide'):
                        view.GotoSlide(new_slide_position)
                    elif hasattr(view, 'Slide'):
                        view.Slide = active_presentation.Slides(new_slide_position)
        except Exception:
            # Don't fail the whole operation if slide switching fails
            pass

        # 8. Return success result
        return {
            "success": True,
            "new_slide_number": new_slide_position,
            "layout_name": target_layout.Name,
            "template_name": template_name,
            "original_slide_count": original_slide_count,
            "new_slide_count": new_slide_count,
            "total_slides": new_slide_count,
            "message": f"Added slide {new_slide_position} using '{target_layout.Name}' layout from '{template_name}' template with full styling preserved"
        }

    except Exception as e:
        return {"error": f"Failed to add slide with layout: {str(e)}"}


def generate_mcp_response(result):
    """Generate the MCP tool response for the LLM."""
    if not result.get('success'):
        return f"Failed to add slide: {result.get('error')}"

    # Create clean response for LLM
    response_lines = [
        f"Added slide {result['new_slide_number']} using '{result['layout_name']}' layout from '{result['template_name']}' template",
        f"Position: Inserted after slide {result['new_slide_number'] - 1}",
        f"Total slides: {result['new_slide_count']} (increased from {result['original_slide_count']})",
        f"Ready for content population using populate_slide_content(slide_number={result['new_slide_number']})"
    ]

    return "\n".join(response_lines)