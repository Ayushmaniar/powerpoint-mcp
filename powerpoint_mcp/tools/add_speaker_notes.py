"""
PowerPoint speaker notes addition tool.
"""

import win32com.client
from typing import Optional


def get_current_slide_index(ppt_app):
    """Get the index of the currently selected/active slide."""
    try:
        if not ppt_app:
            return None

        active_window = ppt_app.ActiveWindow

        # Method 1: Try to get from the current view (most reliable)
        try:
            if hasattr(active_window, 'View') and hasattr(active_window.View, 'Slide'):
                slide_index = active_window.View.Slide.SlideIndex
                if slide_index > 0:
                    return slide_index
        except:
            pass

        # Method 2: Try to get from selection (works in slide sorter view)
        try:
            if (hasattr(active_window, 'Selection') and
                hasattr(active_window.Selection, 'SlideRange') and
                active_window.Selection.SlideRange.Count > 0):
                return active_window.Selection.SlideRange[0].SlideIndex
        except:
            pass

        # Method 3: Try SlideShowWindow if in slideshow mode
        try:
            if hasattr(ppt_app, 'SlideShowWindows') and ppt_app.SlideShowWindows.Count > 0:
                slide_show = ppt_app.SlideShowWindows(1)
                if hasattr(slide_show, 'View') and hasattr(slide_show.View, 'CurrentShowPosition'):
                    return slide_show.View.CurrentShowPosition
        except:
            pass

        # Fallback: return 1 if presentation exists
        presentation = ppt_app.ActivePresentation
        if presentation and presentation.Slides.Count > 0:
            return 1

        return None

    except Exception:
        return 1  # Safe fallback


def powerpoint_add_speaker_notes(slide_number: Optional[int] = None, notes_text: str = "") -> dict:
    """
    Add speaker notes to a specific slide in the active PowerPoint presentation.

    Args:
        slide_number: Slide number to add notes to (1-based). If None, uses current active slide.
        notes_text: Text content to add as speaker notes

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

        # Determine slide to use
        if slide_number is None:
            slide_number = get_current_slide_index(ppt_app)
            if slide_number is None:
                slide_number = 1

        # Validate slide number
        if slide_number < 1 or slide_number > presentation.Slides.Count:
            return {"error": f"Invalid slide number {slide_number}. Presentation has {presentation.Slides.Count} slides."}

        # Get the specified slide
        slide = presentation.Slides(slide_number)

        # Add speaker notes using the CORRECT Microsoft-documented approach
        try:
            # Access the notes page for this slide
            notes_page = slide.NotesPage

            # Use the official Microsoft VBA approach: find ppPlaceholderBody placeholder
            # Based on: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
            notes_shape = None
            debug_info = f"Slide {slide_number} debug: Shapes={notes_page.Shapes.Count}"

            # Iterate through all shapes to find the ppPlaceholderBody (notes text area)
            for i in range(1, notes_page.Shapes.Count + 1):
                shape = notes_page.Shapes(i)
                try:
                    # Check if shape is a placeholder (msoPlaceholder = 14)
                    if hasattr(shape, 'Type') and shape.Type == 14:
                        # Check if it's the body placeholder (ppPlaceholderBody = 2)
                        if (hasattr(shape, 'PlaceholderFormat') and
                            hasattr(shape.PlaceholderFormat, 'Type') and
                            shape.PlaceholderFormat.Type == 2):
                            notes_shape = shape
                            debug_info += f", FoundNotesPlaceholder=Shape{i}"
                            break
                except:
                    continue

            if notes_shape and hasattr(notes_shape, 'TextFrame') and notes_shape.TextFrame:
                # Simple, direct assignment - this should work reliably with the correct placeholder
                try:
                    notes_shape.TextFrame.TextRange.Text = notes_text

                    return {
                        "success": True,
                        "slide_number": slide_number,
                        "total_slides": presentation.Slides.Count,
                        "notes_length": len(notes_text),
                        "message": f"Added speaker notes to slide {slide_number}",
                        "debug_info": debug_info
                    }
                except Exception as e:
                    return {"error": f"Failed to set notes text for slide {slide_number}: {str(e)}. Debug: {debug_info}"}
            else:
                return {"error": f"Could not find notes placeholder (ppPlaceholderBody) for slide {slide_number}. Debug: {debug_info}"}

        except Exception as notes_error:
            return {"error": f"Failed to add notes to slide {slide_number}: {str(notes_error)}"}

    except Exception as e:
        return {"error": f"Failed to add speaker notes: {str(e)}"}