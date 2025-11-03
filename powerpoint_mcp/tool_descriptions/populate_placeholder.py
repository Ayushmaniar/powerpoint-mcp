"""Description for populate_placeholder tool"""

DESCRIPTION = """
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
