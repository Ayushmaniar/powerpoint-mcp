"""Description for manage_slide tool"""

DESCRIPTION = """
    Manage slides in the active PowerPoint presentation.

    Provides comprehensive slide operations for duplicating, deleting, and moving slides.
    All operations automatically switch to the relevant slide after completion.

    Args:
        operation: The operation to perform ("duplicate", "delete", or "move")
        slide_number: The slide number to operate on (1-based index)
        target_position: For 'move' operation - where to move the slide (required)
                        For 'duplicate' operation - where to place the duplicate (optional, defaults to after original)

    Operations:
        - "duplicate": Creates a copy of the specified slide
          Example: manage_slide("duplicate", 3)  # Duplicates slide 3 to position 4
          Example: manage_slide("duplicate", 3, 7)  # Duplicates slide 3 to position 7

        - "delete": Removes the specified slide from the presentation
          Example: manage_slide("delete", 5)  # Deletes slide 5

        - "move": Moves a slide to a new position
          Example: manage_slide("move", 2, 8)  # Moves slide 2 to position 8

    Returns:
        Success message with operation details, or error message

    Notes:
        - Cannot delete the last remaining slide in a presentation
        - All slide numbers are 1-based (first slide is 1, not 0)
        - After any operation, the tool automatically switches to the relevant slide
        - For move operation, target_position is required
        - For duplicate operation, target_position is optional (defaults to after original)
    """
