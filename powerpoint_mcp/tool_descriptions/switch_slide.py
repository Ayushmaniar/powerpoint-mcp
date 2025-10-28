"""Description for switch_slide tool"""

DESCRIPTION = """
    Switch to a specific slide in the active PowerPoint presentation.

    Changes the current active slide to the specified slide number, allowing you
    to navigate through the presentation programmatically.

    Args:
        slide_number: Slide number to switch to (1-based). Must be between 1 and total slides.

    Returns:
        Success message with slide information, or error message
    """
