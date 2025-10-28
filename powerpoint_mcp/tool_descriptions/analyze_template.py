"""Description for analyze_template tool"""

DESCRIPTION = """
    Analyze PowerPoint template layouts with comprehensive placeholder analysis and screenshots.

    Creates a hidden temporary presentation to analyze template layouts without interfering
    with the user's active presentation. Generates screenshots with green bounding boxes
    and yellow ID labels for all placeholders, and provides detailed placeholder analysis.

    Screenshots are saved to ~/.powerpoint-mcp/ directory (same as slide_snapshot tool)
    and can be viewed using the Read tool for visual reference.

    Args:
        source: Template source - can be:
            - "current": Use the active presentation as template
            - Template name: e.g., "Training", "Pitchbook" (use list_templates() to discover)
            - Full path: e.g., "C:/path/to/template.potx"
        detailed: If True, include position and size information for each placeholder.
                 If False (default), show compact output without coordinates.

    Returns:
        Comprehensive template analysis with layout details, placeholder information,
        and screenshot locations. Screenshots show green bounding boxes with yellow ID
        labels for each placeholder.

    Examples:
        analyze_template(source="current")  # Compact output
        analyze_template(source="Training", detailed=True)  # Detailed with coordinates
        analyze_template(source="C:/Templates/Corporate.potx")
    """
