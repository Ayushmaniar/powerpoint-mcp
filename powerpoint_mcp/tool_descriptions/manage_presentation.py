"""Description for manage_presentation_tool"""

DESCRIPTION = """
    Comprehensive PowerPoint presentation management tool.

    This tool works on Windows only. Use Windows path format with double backslashes.

    Args:
        action: Action to perform - "open", "close", "create", "save", or "save_as"
        file_path: Path for open/create operations (required for open/create)
        save_path: New path for save_as operation (required for save_as)
        template_path: Template file for create operation (optional)
        presentation_name: Specific presentation name for close operation (optional)

    Actions:
        - "open": Opens an existing presentation (requires file_path)
          Example: action="open", file_path="C:\\Users\\Name\\slides.pptx"

        - "close": Closes a presentation (optional presentation_name, closes active if not specified)
          Example: action="close" or action="close", presentation_name="MyPresentation.pptx"

        - "create": Creates new presentation (optional file_path for immediate save, optional template_path)
          Example: action="create", file_path="C:\\new\\presentation.pptx"
          Example: action="create", template_path="C:\\templates\\corporate.potx", file_path="C:\\new\\slides.pptx"

        - "save": Saves current presentation at its current location
          Example: action="save"

        - "save_as": Saves current presentation to new location (requires save_path)
          Example: action="save_as", save_path="C:\\backup\\slides_v2.pptx"

    Use double backslashes (\\\\) in Windows paths.

    Returns:
        Success message with operation details, or error message
    """
