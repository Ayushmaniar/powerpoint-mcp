"""
PowerPoint animation tool for MCP server.
Adds animations to shapes with support for paragraph-level text animation.
"""

import win32com.client
from typing import Optional


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
        # Connect to PowerPoint
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        if not ppt_app.Presentations.Count:
            return {"error": "No PowerPoint presentation is open. Please open a presentation first."}

        active_presentation = ppt_app.ActivePresentation

        # Determine target slide
        if slide_number is not None:
            if slide_number < 1 or slide_number > active_presentation.Slides.Count:
                return {
                    "error": f"Invalid slide number {slide_number}. Must be between 1 and {active_presentation.Slides.Count}."
                }
            target_slide = active_presentation.Slides(slide_number)
        else:
            # Use current active slide
            try:
                active_window = ppt_app.ActiveWindow
                if hasattr(active_window, 'View') and hasattr(active_window.View, 'Slide'):
                    target_slide = active_window.View.Slide
                else:
                    target_slide = active_presentation.Slides(1)
            except:
                target_slide = active_presentation.Slides(1)

        # Find the shape by name (case-insensitive)
        target_shape = None
        for shape in target_slide.Shapes:
            if shape.Name.lower() == shape_name.lower():
                target_shape = shape
                break

        if not target_shape:
            available_shapes = [shape.Name for shape in target_slide.Shapes]
            return {
                "error": f"Shape '{shape_name}' not found on slide {target_slide.SlideIndex}",
                "available_shapes": available_shapes
            }

        # Map effect names to PowerPoint constants
        effect_map = {
            "fade": 10,      # msoAnimEffectFade
            "appear": 1,     # msoAnimEffectAppear
            "fly": 2,        # msoAnimEffectFly
            "wipe": 22,      # msoAnimEffectWipe
            "zoom": 23,      # msoAnimEffectZoom
        }

        if effect.lower() not in effect_map:
            return {
                "error": f"Invalid effect '{effect}'. Must be one of: {', '.join(effect_map.keys())}"
            }

        effect_id = effect_map[effect.lower()]

        # Validate animate_text parameter
        if animate_text.lower() not in ["all_at_once", "by_paragraph"]:
            return {
                "error": f"Invalid animate_text '{animate_text}'. Must be one of: all_at_once, by_paragraph"
            }

        # Remove existing animations for this shape (to replace, not duplicate)
        main_sequence = target_slide.TimeLine.MainSequence
        effects_to_remove = []

        for i in range(1, main_sequence.Count + 1):
            effect_obj = main_sequence.Item(i)
            # Check if this effect belongs to our target shape
            if effect_obj.Shape.Name == target_shape.Name:
                effects_to_remove.append(i)

        # Remove in reverse order to avoid index shifting
        for idx in reversed(effects_to_remove):
            main_sequence.Item(idx).Delete()

        # Count paragraphs for by_paragraph animation
        paragraph_count = None
        if animate_text.lower() == "by_paragraph" and hasattr(target_shape, 'TextFrame'):
            try:
                text_range = target_shape.TextFrame.TextRange
                paragraph_count = text_range.Paragraphs().Count
            except:
                paragraph_count = 0

        # Add animation effect(s)
        if animate_text.lower() == "by_paragraph" and paragraph_count and paragraph_count > 0:
            # Create one Effect per paragraph (PowerPoint's approach for "by paragraph")
            for para_index in range(1, paragraph_count + 1):
                new_effect = main_sequence.AddEffect(
                    Shape=target_shape,
                    effectId=effect_id,
                    Level=2,  # msoAnimateTextByFirstLevel
                    trigger=1,  # msoAnimTriggerOnPageClick
                    Index=-1
                )
                # Set which paragraph this effect applies to
                new_effect.Paragraph = para_index
                # Set timing
                new_effect.Timing.Duration = 0.5
                new_effect.Timing.TriggerDelayTime = 0.0
        else:
            # Animate entire shape at once
            new_effect = main_sequence.AddEffect(
                Shape=target_shape,
                effectId=effect_id,
                Level=0,  # msoAnimateLevelNone
                trigger=1,  # msoAnimTriggerOnPageClick
                Index=-1
            )
            # Set timing
            new_effect.Timing.Duration = 0.5
            new_effect.Timing.TriggerDelayTime = 0.0

        # Get current animation count
        total_animations = main_sequence.Count

        # Build result
        result = {
            "success": True,
            "shape_name": target_shape.Name,
            "effect": effect.lower(),
            "animate_text": animate_text.lower(),
            "animation_number": total_animations,
            "slide_number": target_slide.SlideIndex,
            "total_animations": total_animations
        }

        if paragraph_count is not None:
            result["paragraph_count"] = paragraph_count

        return result

    except Exception as e:
        return {"error": f"Failed to add animation to '{shape_name}': {str(e)}"}


def generate_mcp_response(result):
    """Generate the MCP tool response for the LLM."""
    if not result.get('success'):
        error_msg = result.get('error', 'Unknown error')

        # If shape not found, show available shapes
        if 'available_shapes' in result:
            available = '\n  - '.join(result['available_shapes'])
            return f"Failed: {error_msg}\n\nAvailable shapes:\n  - {available}"

        return f"Failed to add animation: {error_msg}"

    # Build success message
    lines = [
        f"✅ Added '{result['effect']}' animation to '{result['shape_name']}'"
    ]

    # Add text animation info
    if result['animate_text'] == 'by_paragraph':
        if result.get('paragraph_count'):
            lines.append(f"   Text will animate by paragraph ({result['paragraph_count']} paragraphs detected)")
        else:
            lines.append(f"   Text will animate by paragraph")
    else:
        lines.append(f"   Text will animate all at once")

    # Add sequence info
    lines.append(f"   This is animation #{result['animation_number']} on slide {result['slide_number']}")
    lines.append(f"   Total animations on this slide: {result['total_animations']}")

    return '\n'.join(lines)
