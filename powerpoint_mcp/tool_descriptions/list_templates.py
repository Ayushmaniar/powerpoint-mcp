"""Description for list_templates tool"""

DESCRIPTION = """
    Discover and list available PowerPoint templates.

    Scans common template directories (Personal, User, System) to find available
    PowerPoint template files (.potx, .potm, .pot). Returns a clean list of
    template names that can be used with the analyze_template tool.

    The tool searches in:
    - Personal Templates: Custom Office Templates folder
    - User Templates: AppData/Roaming/Microsoft/Templates
    - System Templates: Program Files/Microsoft Office/Templates

    Returns:
        Organized list of available templates grouped by location, with usage instructions
    """
