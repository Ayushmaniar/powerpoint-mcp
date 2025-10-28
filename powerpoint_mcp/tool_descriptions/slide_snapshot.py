"""Description for slide_snapshot tool"""

DESCRIPTION = """
    Capture comprehensive context of a PowerPoint slide with optional screenshot.

    This tool provides detailed information about the current (or specified) slide
    including all objects, text content with HTML formatting, tables, charts, and
    layout details.

    Includes optional screenshot functionality with green bounding boxes and yellow
    ID labels overlaid on all objects. The screenshot is saved to a file and the LLM
    is informed of the location for visual reference.

    The tool automatically detects the current active slide if no slide number is
    specified. It returns formatted slide context including object positions, IDs,
    text content with HTML formatting, and structural information.

    Args:
        slide_number: Slide number to capture (1-based). If None, uses current active slide
        include_screenshot: Whether to save a screenshot with bounding boxes. Default True.
        screenshot_filename: Optional custom filename for screenshot. If None, generates slide-{timestamp}.png

    Returns:
        Comprehensive slide context with all objects and their properties, plus screenshot info if enabled, or error message
    """
