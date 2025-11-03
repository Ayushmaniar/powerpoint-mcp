"""
PowerPoint MCP Server

A Model Context Protocol server for automating Microsoft PowerPoint using pywin32.
"""

from typing import Optional
from mcp.server.fastmcp import FastMCP

from .tools.snapshot import powerpoint_snapshot
from .tools.presentation import manage_presentation
from .tools.switch_slide import powerpoint_switch_slide
from .tools.add_speaker_notes import powerpoint_add_speaker_notes
from .tools.list_templates import powerpoint_list_templates, generate_mcp_response
from .tools.analyze_template import powerpoint_analyze_template, generate_mcp_response as generate_analyze_response
from .tools.add_slide_with_layout import powerpoint_add_slide_with_layout, generate_mcp_response as generate_add_slide_response
from .tools.populate_placeholder import powerpoint_populate_placeholder, generate_mcp_response as generate_populate_response
from .tools.manage_slide import powerpoint_manage_slide, generate_mcp_response as generate_manage_slide_response
from .tools.evaluate import powerpoint_evaluate, generate_mcp_response as generate_evaluate_response

# Create the MCP server instance
mcp = FastMCP("PowerPoint MCP Server")

@mcp.tool()
def manage_presentation_tool(
    action: str,
    file_path: Optional[str] = None,
    save_path: Optional[str] = None,
    template_path: Optional[str] = None,
    presentation_name: Optional[str] = None
) -> str:
    """
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
    return manage_presentation(action, file_path, save_path, template_path, presentation_name)



@mcp.tool()
def slide_snapshot(slide_number: Optional[str] = None,
                  include_screenshot: Optional[bool] = True,
                  screenshot_filename: Optional[str] = None) -> str:
    """
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
    # Convert string to int if provided
    if slide_number is not None:
        try:
            slide_number = int(slide_number)
        except ValueError:
            return f"Error: slide_number must be a valid integer, got '{slide_number}'"

    # Convert boolean if needed (handles JSON boolean type)
    if include_screenshot is None:
        include_screenshot = True

    result = powerpoint_snapshot(slide_number, include_screenshot, screenshot_filename)

    if "error" in result:
        return f"Error: {result['error']}"

    response_parts = [
        f"Slide context captured: {result['slide_number']} of {result['total_slides']}",
        f"Objects found: {result['object_count']}"
    ]

    # Add screenshot information if included
    if include_screenshot:
        if result.get('screenshot_saved'):
            response_parts.extend([
                "",
                f"Screenshot saved: {result['screenshot_path']}",
                f"Image size: {result['image_size']}",
                f"Objects annotated: {result['objects_annotated']} (green boxes with yellow ID labels)",
                f"{result['screenshot_message']}",
                "",
                "The screenshot file has been saved and can be viewed using the Read tool for visual reference."
            ])
        else:
            response_parts.extend([
                "",
                f"Screenshot failed: {result.get('screenshot_error', 'Unknown error')}"
            ])

    response_parts.extend(["", result['context']])

    return "\n".join(response_parts)


@mcp.tool()
def switch_slide(slide_number: str) -> str:
    """
    Switch to a specific slide in the active PowerPoint presentation.

    Changes the current active slide to the specified slide number, allowing you
    to navigate through the presentation programmatically.

    Args:
        slide_number: Slide number to switch to (1-based). Must be between 1 and total slides.

    Returns:
        Success message with slide information, or error message
    """
    # Convert string to int
    try:
        slide_number = int(slide_number)
    except ValueError:
        return f"Error: slide_number must be a valid integer, got '{slide_number}'"

    result = powerpoint_switch_slide(slide_number)

    if "error" in result:
        return f"Error: {result['error']}"

    return f"Successfully switched to slide {result['slide_number']} of {result['total_slides']}"


@mcp.tool()
def add_speaker_notes(slide_number: str, notes_text: str) -> str:
    """
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
    # Convert string to int
    try:
        slide_number = int(slide_number)
    except ValueError:
        return f"Error: slide_number must be a valid integer, got '{slide_number}'"

    result = powerpoint_add_speaker_notes(slide_number, notes_text)

    if "error" in result:
        return f"Error: {result['error']}"

    return (f"Successfully added speaker notes to slide {result['slide_number']} of {result['total_slides']}\n"
            f"Notes length: {result['notes_length']} characters")


@mcp.tool()
def list_templates() -> str:
    """
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
    result = powerpoint_list_templates()
    return generate_mcp_response(result)


@mcp.tool()
def analyze_template(source: str = "current", detailed: bool = False) -> str:
    """
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
    result = powerpoint_analyze_template(source)
    return generate_analyze_response(result, detailed)


@mcp.tool()
def add_slide_with_layout(template_name: str, layout_name: str, after_slide: int) -> str:
    """
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
    result = powerpoint_add_slide_with_layout(template_name, layout_name, after_slide)
    return generate_add_slide_response(result)


@mcp.tool()
def populate_placeholder(
    placeholder_name: str,
    content: str,
    content_type: str = "auto",
    slide_number: Optional[str] = None
) -> str:
    """
    Populate a PowerPoint placeholder with content including HTML formatting and LaTeX equations.

    Supports semantic placeholder names and auto-detects content type (text/image/plot).
    Handles simplified HTML formatting: <b>bold</b>, <i>italic</i>, <u>underline</u>,
    colors like <red>text</red>, lists <ul><li>items</li></ul>, and LaTeX equations <latex>equation</latex>.

    Args:
        placeholder_name: Name of the placeholder (e.g., "Title 1", "Subtitle 2")
        content: Text with HTML/LaTeX formatting, image file path, or matplotlib code
        content_type: "text", "image", "plot", or "auto" (auto-detect based on content)
        slide_number: Target slide number (1-based). If None, uses current active slide

    Matplotlib plots:
        - Use content_type="plot" for matplotlib code
        - DO NOT include plt.savefig() or plt.close() - these are handled automatically
        - Imports available: numpy as np, matplotlib.pyplot as plt

    Returns:
        Success message with operation details, or error message

    Examples:
        # Basic text
        populate_placeholder("Title 1", "My Presentation Title")

        # HTML formatting
        populate_placeholder("Content Placeholder 2", "<b>Bold</b> and <red>red text</red>")

        # LaTeX equations (simple)
        populate_placeholder("Equation1", "Pythagorean theorem: <latex>a^2+b^2=c^2</latex>")

        # LaTeX equations (complex fractions)
        populate_placeholder("Equation2", "Quadratic formula: <latex>x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}</latex>")

        # LaTeX equations (integrals)
        populate_placeholder("Equation3", "Integration: <latex>\\int_a^b f(x)dx</latex>")

        # Mixed content: HTML formatting + LaTeX (positions adjust automatically!)
        populate_placeholder("Mixed1",
            "<b>Einstein's famous equation:</b> <latex>E=mc^2</latex> <i>where c is the speed of light</i>")

        # Colored text with fractions
        populate_placeholder("Mixed2",
            "<red>Important:</red> The derivative <latex>\\frac{dy}{dx}</latex> represents the <b>rate of change</b>")

        # Multiple equations with formatting
        populate_placeholder("Mixed3",
            "Wave equation: <latex>c=\\lambda f</latex> and energy: <latex>E=hf</latex> are <b><blue>fundamental</blue></b>")

        # Lists with LaTeX equations
        populate_placeholder("List1",
            "Key formulas:<ul><li><b>Area:</b> <latex>A=\\pi r^2</latex></li><li><b>Circumference:</b> <latex>C=2\\pi r</latex></li><li><b>Volume:</b> <latex>V=\\frac{4}{3}\\pi r^3</latex></li></ul>")

        # Numbered lists with equations
        populate_placeholder("List2",
            "Steps:<ol><li>Start with <latex>f(x)=x^2</latex></li><li>Take derivative: <latex>f'(x)=2x</latex></li><li><green>Result is linear!</green></li></ol>")

        # Image
        populate_placeholder("Picture Placeholder 7", "C:\\Images\\chart.png", "image")

        # Matplotlib plot (simple)
        populate_placeholder("Picture Placeholder 2",
            "plt.plot([1,2,3,4], [1,4,9,16])\\nplt.title('Square Numbers')\\nplt.grid(True)", "plot")

        # Matplotlib plot (educational - quadratic with roots)
        populate_placeholder("Picture Placeholder 2",
            '''import numpy as np
x = np.linspace(-1, 5, 200)
y = x**2 - 4*x + 3
plt.figure(figsize=(10, 7))
plt.plot(x, y, 'b-', linewidth=3, label=r'$f(x) = x^2 - 4x + 3$')
plt.plot([1, 3], [0, 0], 'ro', markersize=12, label='Roots')
plt.axhline(y=0, color='k', linewidth=1)
plt.axvline(x=0, color='k', linewidth=1)
plt.grid(True, alpha=0.3)
plt.xlabel('x', fontsize=14)
plt.ylabel('f(x)', fontsize=14)
plt.title('Quadratic Equation', fontsize=16, weight='bold')
plt.legend()''', "plot")
    """
    # Convert string to int if provided
    if slide_number is not None:
        try:
            slide_number = int(slide_number)
        except ValueError:
            return f"Error: slide_number must be a valid integer, got '{slide_number}'"

    result = powerpoint_populate_placeholder(placeholder_name, content, content_type, slide_number)
    return generate_populate_response(result)


@mcp.tool()
def manage_slide(
    operation: str,
    slide_number: str,
    target_position: Optional[int] = None
) -> str:
    """
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
    # Convert string to int
    try:
        slide_number = int(slide_number)
    except ValueError:
        return f"Error: slide_number must be a valid integer, got '{slide_number}'"

    result = powerpoint_manage_slide(operation, slide_number, target_position)
    return generate_manage_slide_response(result)


@mcp.tool()
def powerpoint_evaluate_tool(
    code: str,
    slide_number: Optional[str] = None,
    shape_ref: Optional[str] = None,
    description: Optional[str] = None
) -> str:
    """
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
    # Convert string to int if provided
    if slide_number is not None:
        try:
            slide_number = int(slide_number)
        except ValueError:
            return f"Error: slide_number must be a valid integer, got '{slide_number}'"

    result = powerpoint_evaluate(code, slide_number, shape_ref, description)
    return generate_evaluate_response(result)
  

def main():
    """Main entry point for the PowerPoint MCP server."""
    mcp.run()