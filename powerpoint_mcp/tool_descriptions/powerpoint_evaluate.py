"""Description for powerpoint_evaluate_tool"""

DESCRIPTION = """
    Execute arbitrary Python code in PowerPoint automation context.

    CRITICAL: ALWAYS use 'skills' methods for content operations. Only use direct COM for styling.

    PREFERRED - Use skills for content, then COM for styling:
        # Step 1: Use skills to add/modify content
        skills.populate_placeholder("Title 1", "<b>My Title</b>")

        # Step 2: Fine-tune styling with COM if needed
        for shape in slide.Shapes:
            if "Title 1" in shape.Name:
                shape.TextFrame.TextRange.Font.Size = 48
                shape.TextFrame.TextRange.Font.Name = "Arial"

    WRONG - Don't use COM for content operations:
        shape.TextFrame.TextRange.Text = "text"  # NO! Use skills.populate_placeholder()
        slide.NotesPage.Shapes(2).TextFrame.TextRange.Text = "notes"  # NO! Use skills.add_speaker_notes()

    Available in execution context:
        - skills: All MCP tools (populate_placeholder, add_speaker_notes, manage_slide, etc.)
        - ppt, presentation, slide, shape: PowerPoint COM objects
        - math: Python math module

    Common patterns:
        1. Batch operations: Loop with skills calls
        2. Content + Styling: skills for content, then COM for font/colors
        3. Geometric layouts: Create shapes with COM, populate with skills

    Args:
        code: Python code to execute
        slide_number: Target slide (1-based). If None, uses current slide
        shape_ref: Optional shape ID/Name to operate on
        description: Human-readable description of operation intent

    Returns:
        Execution result with success/error status and optional return data

    Example - Skills + styling:
        code = '''
        # Use skills to add content
        skills.populate_placeholder("Title 1", "Welcome")
        skills.populate_placeholder("Subtitle 2", "Introduction")

        # Then style with COM
        for shape in slide.Shapes:
            if "Title 1" in shape.Name:
                shape.TextFrame.TextRange.Font.Size = 54
                shape.TextFrame.TextRange.Font.Color.RGB = 255  # Red
        '''

    Example - Batch with skills:
        code = '''
        for i in range(1, 4):
            skills.add_speaker_notes(i, f"Slide {i} notes")
            skills.populate_placeholder(f"Title {i}", f"<b>Section {i}</b>")
        '''
    """
