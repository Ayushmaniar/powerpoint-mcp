"""Description for add_speaker_notes tool"""

DESCRIPTION = """
    Add speaker notes to a specific slide in the active PowerPoint presentation.

    Adds or replaces the speaker notes content for the specified slide with the
    provided text. Speaker notes are visible in presenter view and when printing
    notes pages, but not during the actual slideshow.

    Args:
        slide_number: Slide number to add notes to (1-based). Must be between 1 and total slides.
        notes_text: Text content to add as speaker notes. Can be a long text string.

    Returns:
        Success message with slide information, or error message
    """
