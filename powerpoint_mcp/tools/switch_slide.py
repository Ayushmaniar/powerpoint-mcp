"""
PowerPoint slide switching tool.
"""

import win32com.client


def powerpoint_switch_slide(slide_number: int) -> dict:
    """
    Switch to a specific slide in the active PowerPoint presentation.

    Args:
        slide_number: Slide number to switch to (1-based)

    Returns:
        Dictionary with success status and slide information or error message
    """
    try:
        # Connect to PowerPoint
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        if not ppt_app.Presentations.Count:
            return {"error": "No PowerPoint presentation is open"}

        presentation = ppt_app.ActivePresentation

        # Validate slide number
        if slide_number < 1 or slide_number > presentation.Slides.Count:
            return {"error": f"Invalid slide number {slide_number}. Presentation has {presentation.Slides.Count} slides."}

        # Switch to the specified slide
        try:
            slide = presentation.Slides(slide_number)

            # Try to get the active window and switch view
            if hasattr(ppt_app, 'ActiveWindow') and ppt_app.ActiveWindow:
                active_window = ppt_app.ActiveWindow

                # Set the view to the specified slide
                if hasattr(active_window, 'View'):
                    view = active_window.View
                    if hasattr(view, 'GotoSlide'):
                        view.GotoSlide(slide_number)
                    elif hasattr(view, 'Slide'):
                        # Alternative method for some PowerPoint versions
                        view.Slide = slide

            return {
                "success": True,
                "slide_number": slide_number,
                "total_slides": presentation.Slides.Count,
                "slide_name": getattr(slide, 'Name', f"Slide {slide_number}"),
                "message": f"Switched to slide {slide_number} of {presentation.Slides.Count}"
            }

        except Exception as switch_error:
            return {"error": f"Failed to switch to slide {slide_number}: {str(switch_error)}"}

    except Exception as e:
        return {"error": f"Failed to switch slide: {str(e)}"}