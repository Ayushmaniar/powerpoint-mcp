"""
Simple test script to take a screenshot of the current active slide.
Uses the same screenshot functionality as the slide_snapshot tool.
"""

import os
import tempfile
import win32com.client
from pathlib import Path
from PIL import Image
from datetime import datetime

def get_current_slide_index(ppt_app):
    """Get the index of the currently selected/active slide."""
    try:
        if not ppt_app:
            return None

        active_window = ppt_app.ActiveWindow

        # Try to get from the current view
        try:
            if hasattr(active_window, 'View') and hasattr(active_window.View, 'Slide'):
                slide_index = active_window.View.Slide.SlideIndex
                if slide_index > 0:
                    return slide_index
        except:
            pass

        # Fallback: return 1 if presentation exists
        presentation = ppt_app.ActivePresentation
        if presentation and presentation.Slides.Count > 0:
            return 1

        return None

    except Exception:
        return 1


def take_screenshot():
    """Take a screenshot of the current active slide and save it."""
    try:
        # Connect to PowerPoint
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        if not ppt_app.Presentations.Count:
            print("Error: No PowerPoint presentation is open")
            return

        presentation = ppt_app.ActivePresentation

        # Get current slide number
        slide_number = get_current_slide_index(ppt_app)
        if slide_number is None:
            slide_number = 1

        slide = presentation.Slides(slide_number)

        # Create output filename
        script_dir = Path(__file__).parent
        output_filename = f"slide_{slide_number}.png"
        output_path = script_dir / output_filename

        # Export slide to temporary file first
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
            temp_path = temp_file.name

        # Export slide as PNG
        slide.Export(temp_path, "PNG")

        # Load image and save to final location
        image = Image.open(temp_path)
        image.save(str(output_path), "PNG")

        # Clean up temp file
        os.unlink(temp_path)

        print(f"Screenshot saved: {output_path}")
        print(f"Slide: {slide_number} of {presentation.Slides.Count}")
        print(f"Image size: {image.size[0]}x{image.size[1]}")

    except Exception as e:
        print(f"Error taking screenshot: {str(e)}")


if __name__ == "__main__":
    take_screenshot()