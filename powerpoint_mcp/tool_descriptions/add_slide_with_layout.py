"""Description for add_slide_with_layout tool"""

DESCRIPTION = """
    Add a slide with a specific template layout at the specified position.

    Args:
        template_name: Name of the template (use list_templates() to discover available templates)
        layout_name: Name of the layout within the template (use analyze_template() to see layouts)
        after_slide: Insert the new slide after this position (new slide becomes after_slide + 1)

    Returns:
        Success message with slide details, or error message

    Examples:
        add_slide_with_layout(template_name="Training", layout_name="Title", after_slide=0)
        add_slide_with_layout(template_name="Pitchbook", layout_name="2-Up", after_slide=5)
    """
