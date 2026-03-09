"""
PowerPoint animation tool for MCP server.
Adds animations to shapes with support for paragraph-level text animation.
"""

from typing import Optional

from ..backends import get_backend


def powerpoint_add_animation(
    shape_name: str,
    effect: str = "fade",
    animate_text: str = "all_at_once",
    slide_number: Optional[int] = None
) -> dict:
    """
    Add animation to a shape in PowerPoint.

    Args:
        shape_name: Name of the shape to animate (e.g., "Title 1", "Content Placeholder 2")
        effect: Animation effect - "fade", "appear", "fly", "wipe", or "zoom" (default: "fade")
        animate_text: How to animate text - "all_at_once" or "by_paragraph" (default: "all_at_once")
        slide_number: Target slide number (1-based). If None, uses current active slide

    Returns:
        Dictionary with success status and animation details
    """
    try:
        backend = get_backend()
        backend.connect()

        if not backend.get_presentation_count():
            return {"error": "No PowerPoint presentation is open. Please open a presentation first."}

        # Determine target slide
        if slide_number is None:
            slide_number = backend.get_current_slide_index()
            if slide_number is None:
                slide_number = 1

        slide_count = backend.get_slide_count()
        if slide_number < 1 or slide_number > slide_count:
            return {"error": f"Invalid slide number {slide_number}. Must be between 1 and {slide_count}."}

        # Map effect names to PowerPoint constants
        effect_map = {
            "fade": 10,      # msoAnimEffectFade
            "appear": 1,     # msoAnimEffectAppear
            "fly": 2,        # msoAnimEffectFly
            "wipe": 22,      # msoAnimEffectWipe
            "zoom": 23,      # msoAnimEffectZoom
        }

        if effect.lower() not in effect_map:
            return {"error": f"Invalid effect '{effect}'. Must be one of: {', '.join(effect_map.keys())}"}

        effect_id = effect_map[effect.lower()]

        # Validate animate_text parameter
        if animate_text.lower() not in ["all_at_once", "by_paragraph"]:
            return {"error": f"Invalid animate_text '{animate_text}'. Must be one of: all_at_once, by_paragraph"}

        # Remove existing animations for this shape (to replace, not duplicate)
        backend.remove_shape_animations(slide_number, shape_name)

        # Count paragraphs for by_paragraph animation
        paragraph_count = None
        level = 0  # msoAnimateLevelNone

        if animate_text.lower() == "by_paragraph":
            feature_support = backend.get_feature_support()
            if feature_support.animation_by_paragraph:
                paragraph_count = backend.get_paragraph_count(slide_number, shape_name)
                if paragraph_count and paragraph_count > 0:
                    level = 2  # msoAnimateTextByFirstLevel
            else:
                paragraph_count = backend.get_paragraph_count(slide_number, shape_name)

        # Add animation effect
        total_animations = backend.add_animation_effect(
            slide_number, shape_name, effect_id,
            level=level, trigger=1, duration=0.5
        )

        result = {
            "success": True,
            "shape_name": shape_name,
            "effect": effect.lower(),
            "animate_text": animate_text.lower(),
            "animation_number": total_animations,
            "slide_number": slide_number,
            "total_animations": total_animations
        }

        if paragraph_count is not None:
            result["paragraph_count"] = paragraph_count

        return result

    except ValueError as e:
        error_msg = str(e)
        if "not found" in error_msg.lower():
            # Extract available shapes from error message if possible
            return {"error": f"Shape '{shape_name}' not found on slide {slide_number}"}
        return {"error": error_msg}
    except Exception as e:
        return {"error": f"Failed to add animation to '{shape_name}': {str(e)}"}


def generate_mcp_response(result):
    """Generate the MCP tool response for the LLM."""
    if not result.get('success'):
        error_msg = result.get('error', 'Unknown error')

        if 'available_shapes' in result:
            available = '\n  - '.join(result['available_shapes'])
            return f"Failed: {error_msg}\n\nAvailable shapes:\n  - {available}"

        return f"Failed to add animation: {error_msg}"

    lines = [
        f"Added '{result['effect']}' animation to '{result['shape_name']}'"
    ]

    if result['animate_text'] == 'by_paragraph':
        if result.get('paragraph_count'):
            lines.append(f"   Text will animate by paragraph ({result['paragraph_count']} paragraphs detected)")
        else:
            lines.append(f"   Text will animate by paragraph")
    else:
        lines.append(f"   Text will animate all at once")

    lines.append(f"   This is animation #{result['animation_number']} on slide {result['slide_number']}")
    lines.append(f"   Total animations on this slide: {result['total_animations']}")

    return '\n'.join(lines)
