"""
PowerPoint slide creation tool for MCP server.
Creates slides with specific template layouts at specified positions.
"""

from typing import Optional

from ..backends import get_backend
from .analyze_template import find_template_by_name


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
        backend = get_backend()
        backend.connect()

        if not backend.get_presentation_count():
            return {"error": "No PowerPoint presentation is open. Please open a presentation first."}

        slide_count = backend.get_slide_count()

        # Validate after_slide parameter
        if after_slide < 0 or after_slide > slide_count:
            return {"error": f"Invalid after_slide position {after_slide}. Must be between 0 and {slide_count}."}

        # Resolve template name to file path
        template_path = find_template_by_name(template_name)
        if not template_path:
            return {"error": f"Template '{template_name}' not found. Use list_templates() to see available templates."}

        result = backend.add_slide_with_layout(template_path, layout_name, after_slide)

        if not result.get("success"):
            return {"error": result.get("error", "Failed to add slide")}

        result["template_name"] = template_name
        result["total_slides"] = result["new_slide_count"]
        result["message"] = f"Added slide {result['new_slide_number']} using '{result['layout_name']}' layout from '{template_name}' template with full styling preserved"

        return result

    except ValueError as e:
        return {"error": str(e)}
    except Exception as e:
        return {"error": f"Failed to add slide with layout: {str(e)}"}


def generate_mcp_response(result):
    """Generate the MCP tool response for the LLM."""
    if not result.get('success'):
        return f"Failed to add slide: {result.get('error')}"

    response_lines = [
        f"Added slide {result['new_slide_number']} using '{result['layout_name']}' layout from '{result['template_name']}' template",
        f"Position: Inserted after slide {result['new_slide_number'] - 1}",
        f"Total slides: {result['new_slide_count']} (increased from {result['original_slide_count']})",
        f"Ready for content population using populate_slide_content(slide_number={result['new_slide_number']})"
    ]

    return "\n".join(response_lines)
