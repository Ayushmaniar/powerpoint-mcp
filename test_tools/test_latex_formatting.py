"""
Test LaTeX equations and HTML formatting with populate_placeholder tool.
Tests: LaTeX equations, inline equations with formatting, complex math, mixed content.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'python'))

from powerpoint_mcp.tools.populate_placeholder import powerpoint_populate_placeholder, generate_mcp_response
from powerpoint_mcp.tools.presentation import manage_presentation
import win32com.client


def test_basic_latex():
    """Test 1: Basic LaTeX equations."""
    print("\n" + "="*60)
    print("TEST 1: Basic LaTeX Equations")
    print("="*60)

    # Create test presentation
    manage_presentation("create", file_path=r"C:\Users\Rajat\test_latex_formatting.pptx")

    # Activate PowerPoint window
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        ppt.Activate()
        # Add blank slide
        if ppt.ActivePresentation and ppt.ActivePresentation.Slides.Count == 0:
            ppt.ActivePresentation.Slides.Add(Index=1, Layout=12)  # ppLayoutBlank
        if ppt.ActivePresentation and ppt.ActivePresentation.Windows.Count > 0:
            ppt.ActivePresentation.Windows(1).Activate()
    except:
        pass

    # Add title and content placeholders
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        slide = ppt.ActivePresentation.Slides(1)

        # Add title textbox
        title_shape = slide.Shapes.AddTextbox(
            Orientation=1,  # msoTextOrientationHorizontal
            Left=50,
            Top=50,
            Width=600,
            Height=60
        )
        title_shape.Name = "Title1"

        # Add content textboxes for equations
        textbox1 = slide.Shapes.AddTextbox(1, 50, 150, 600, 80)
        textbox1.Name = "Equation1"

        textbox2 = slide.Shapes.AddTextbox(1, 50, 250, 600, 80)
        textbox2.Name = "Equation2"

        textbox3 = slide.Shapes.AddTextbox(1, 50, 350, 600, 80)
        textbox3.Name = "Equation3"
    except Exception as e:
        print(f"Error creating shapes: {e}")

    # Test 1a: Simple quadratic formula
    result = powerpoint_populate_placeholder(
        "Title1",
        "<b><blue>Basic LaTeX Equations</blue></b>",
        slide_number=1
    )
    print("\nTitle Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "Title should populate"

    result = powerpoint_populate_placeholder(
        "Equation1",
        "Quadratic formula: <latex>x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}</latex>",
        slide_number=1
    )
    print("\nEquation 1 Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "Quadratic formula should populate"
    assert result.get("latex_equations_applied", 0) == 1, "Should apply 1 LaTeX equation"

    # Test 1b: Pythagorean theorem
    result = powerpoint_populate_placeholder(
        "Equation2",
        "Pythagorean theorem: <latex>a^2+b^2=c^2</latex>",
        slide_number=1
    )
    print("\nEquation 2 Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "Pythagorean theorem should populate"

    # Test 1c: Integral
    result = powerpoint_populate_placeholder(
        "Equation3",
        "Integration: <latex>\\int_a^b f(x)dx</latex>",
        slide_number=1
    )
    print("\nEquation 3 Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "Integral should populate"

    print("\nPASS: Test 1 - Basic LaTeX equations")
    input("\nPress ENTER to continue to next test...")


def test_inline_latex_with_formatting():
    """Test 2: Inline LaTeX with HTML formatting."""
    print("\n" + "="*60)
    print("TEST 2: Inline LaTeX with HTML Formatting")
    print("="*60)

    # Add new slide
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        # Add blank slide (layout 12 = ppLayoutBlank)
        slide = ppt.ActivePresentation.Slides.Add(Index=2, Layout=12)

        # Add title and content textboxes
        title_shape = slide.Shapes.AddTextbox(1, 50, 50, 600, 60)
        title_shape.Name = "Title2"

        textbox1 = slide.Shapes.AddTextbox(1, 50, 150, 600, 100)
        textbox1.Name = "Mixed1"

        textbox2 = slide.Shapes.AddTextbox(1, 50, 270, 600, 100)
        textbox2.Name = "Mixed2"

        textbox3 = slide.Shapes.AddTextbox(1, 50, 390, 600, 100)
        textbox3.Name = "Mixed3"
    except Exception as e:
        print(f"Error creating slide 2 shapes: {e}")

    # Title
    result = powerpoint_populate_placeholder(
        "Title2",
        "<b><green>Mixed Content: Text + LaTeX + Formatting</green></b>",
        slide_number=2
    )
    print("\nTitle Result:")
    print(generate_mcp_response(result))

    # Test 2a: Bold text with inline equation
    result = powerpoint_populate_placeholder(
        "Mixed1",
        "<b>Einstein's famous equation:</b> <latex>E=mc^2</latex> <i>where c is the speed of light</i>",
        slide_number=2
    )
    print("\nMixed 1 Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "Mixed content should populate"
    assert result.get("latex_equations_applied", 0) == 1, "Should apply 1 equation"

    # Test 2b: Colored text with equation
    result = powerpoint_populate_placeholder(
        "Mixed2",
        "<red>Important:</red> The derivative <latex>\\frac{dy}{dx}</latex> represents the <b>rate of change</b>",
        slide_number=2
    )
    print("\nMixed 2 Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "Colored mixed content should populate"

    # Test 2c: Multiple equations in one text
    result = powerpoint_populate_placeholder(
        "Mixed3",
        "Wave equation: <latex>c=\\lambda f</latex> and energy: <latex>E=hf</latex> are <b><blue>fundamental</blue></b>",
        slide_number=2
    )
    print("\nMixed 3 Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "Multiple equations should populate"
    assert result.get("latex_equations_applied", 0) == 2, "Should apply 2 equations"

    print("\nPASS: Test 2 - Inline LaTeX with formatting")
    input("\nPress ENTER to continue to next test...")


def test_complex_equations():
    """Test 3: Complex mathematical equations."""
    print("\n" + "="*60)
    print("TEST 3: Complex Mathematical Equations")
    print("="*60)

    # Add new slide
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        # Add blank slide (layout 12 = ppLayoutBlank)
        slide = ppt.ActivePresentation.Slides.Add(Index=3, Layout=12)

        title_shape = slide.Shapes.AddTextbox(1, 50, 30, 600, 50)
        title_shape.Name = "Title3"

        textbox1 = slide.Shapes.AddTextbox(1, 50, 100, 600, 80)
        textbox1.Name = "Complex1"

        textbox2 = slide.Shapes.AddTextbox(1, 50, 200, 600, 80)
        textbox2.Name = "Complex2"

        textbox3 = slide.Shapes.AddTextbox(1, 50, 300, 600, 80)
        textbox3.Name = "Complex3"

        textbox4 = slide.Shapes.AddTextbox(1, 50, 400, 600, 80)
        textbox4.Name = "Complex4"
    except Exception as e:
        print(f"Error creating slide 3 shapes: {e}")

    # Title
    result = powerpoint_populate_placeholder(
        "Title3",
        "<b><purple>Complex Mathematical Equations</purple></b>",
        slide_number=3
    )
    print("\nTitle Result:")
    print(generate_mcp_response(result))

    # Test 3a: Matrix
    result = powerpoint_populate_placeholder(
        "Complex1",
        "Matrix notation: <latex>\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}</latex>",
        slide_number=3
    )
    print("\nComplex 1 (Matrix) Result:")
    print(generate_mcp_response(result))

    # Test 3b: Summation
    result = powerpoint_populate_placeholder(
        "Complex2",
        "Summation: <latex>\\sum_{i=1}^{n}i=\\frac{n(n+1)}{2}</latex>",
        slide_number=3
    )
    print("\nComplex 2 (Summation) Result:")
    print(generate_mcp_response(result))

    # Test 3c: Greek letters and subscripts
    result = powerpoint_populate_placeholder(
        "Complex3",
        "Greek letters: <latex>\\alpha+\\beta=\\gamma</latex> and subscripts: <latex>x_1+x_2=x_3</latex>",
        slide_number=3
    )
    print("\nComplex 3 (Greek & subscripts) Result:")
    print(generate_mcp_response(result))

    # Test 3d: Limits
    result = powerpoint_populate_placeholder(
        "Complex4",
        "Limit definition: <latex>\\lim_{x\\to\\infty}\\frac{1}{x}=0</latex>",
        slide_number=3
    )
    print("\nComplex 4 (Limits) Result:")
    print(generate_mcp_response(result))

    print("\nPASS: Test 3 - Complex equations")
    input("\nPress ENTER to continue to next test...")


def test_lists_with_equations():
    """Test 4: Lists with equations and formatting."""
    print("\n" + "="*60)
    print("TEST 4: Lists with Equations")
    print("="*60)

    # Add new slide
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        # Add blank slide (layout 12 = ppLayoutBlank)
        slide = ppt.ActivePresentation.Slides.Add(Index=4, Layout=12)

        title_shape = slide.Shapes.AddTextbox(1, 50, 30, 600, 50)
        title_shape.Name = "Title4"

        textbox1 = slide.Shapes.AddTextbox(1, 50, 100, 600, 150)
        textbox1.Name = "List1"

        textbox2 = slide.Shapes.AddTextbox(1, 50, 270, 600, 150)
        textbox2.Name = "List2"
    except Exception as e:
        print(f"Error creating slide 4 shapes: {e}")

    # Title
    result = powerpoint_populate_placeholder(
        "Title4",
        "<b><orange>Equations in Lists</orange></b>",
        slide_number=4
    )
    print("\nTitle Result:")
    print(generate_mcp_response(result))

    # Test 4a: Bullet list with equations
    result = powerpoint_populate_placeholder(
        "List1",
        "Key formulas:<ul><li><b>Area:</b> <latex>A=\\pi r^2</latex></li><li><b>Circumference:</b> <latex>C=2\\pi r</latex></li><li><b>Volume:</b> <latex>V=\\frac{4}{3}\\pi r^3</latex></li></ul>",
        slide_number=4
    )
    print("\nList 1 (Bullet) Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "List with equations should populate"

    # Test 4b: Numbered list with colored equations
    result = powerpoint_populate_placeholder(
        "List2",
        "Steps:<ol><li>Start with <latex>f(x)=x^2</latex></li><li>Take derivative: <latex>f'(x)=2x</latex></li><li><green>Result is linear!</green></li></ol>",
        slide_number=4
    )
    print("\nList 2 (Numbered) Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "Numbered list with equations should populate"

    print("\nPASS: Test 4 - Lists with equations")
    input("\nPress ENTER to continue to next test...")


def test_comprehensive_mixed():
    """Test 5: Comprehensive test with all features."""
    print("\n" + "="*60)
    print("TEST 5: Comprehensive Mixed Content")
    print("="*60)

    # Add new slide
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        # Add blank slide (layout 12 = ppLayoutBlank)
        slide = ppt.ActivePresentation.Slides.Add(Index=5, Layout=12)

        title_shape = slide.Shapes.AddTextbox(1, 50, 20, 600, 60)
        title_shape.Name = "Title5"

        content = slide.Shapes.AddTextbox(1, 50, 100, 600, 380)
        content.Name = "Content5"
    except Exception as e:
        print(f"Error creating slide 5 shapes: {e}")

    # Title
    result = powerpoint_populate_placeholder(
        "Title5",
        "<b><red>Comprehensive Test:</red> <blue>LaTeX</blue> + <green>Formatting</green> + <purple>Lists</purple></b>",
        slide_number=5
    )
    print("\nTitle Result:")
    print(generate_mcp_response(result))

    # Comprehensive content with everything
    content_text = """<b><blue>Physics Equations Summary</blue></b>

<b>Classical Mechanics:</b><ul><li>Newton's 2nd law: <latex>F=ma</latex></li><li>Kinetic energy: <latex>KE=\\frac{1}{2}mv^2</latex></li><li><red>Important:</red> Momentum <latex>p=mv</latex></li></ul>

<b>Electromagnetism:</b><ol><li>Coulomb's law: <latex>F=k\\frac{q_1q_2}{r^2}</latex></li><li>Ohm's law: <latex>V=IR</latex></li><li><green>Maxwell's equations</green> govern all EM phenomena</li></ol>

<i>Remember: <latex>\\hbar=\\frac{h}{2\\pi}</latex> is the reduced Planck constant</i>"""

    result = powerpoint_populate_placeholder(
        "Content5",
        content_text,
        slide_number=5
    )
    print("\nComprehensive Content Result:")
    print(generate_mcp_response(result))
    assert result.get("success") == True, "Comprehensive content should populate"
    print(f"Format segments applied: {result.get('format_segments_applied', 0)}")
    print(f"LaTeX equations applied: {result.get('latex_equations_applied', 0)}")

    print("\nPASS: Test 5 - Comprehensive mixed content")
    input("\nPress ENTER to cleanup...")


def cleanup():
    """Cleanup: Close the test presentation."""
    print("\n" + "="*60)
    print("CLEANUP")
    print("="*60)

    try:
        # Save before closing so user can review
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        if ppt.ActivePresentation:
            ppt.ActivePresentation.Save()
            print("INFO: Presentation saved")
    except:
        pass

    manage_presentation("close", presentation_name="test_latex_formatting.pptx")
    print("INFO: Test presentation closed")


if __name__ == "__main__":
    try:
        test_basic_latex()
        test_inline_latex_with_formatting()
        test_complex_equations()
        test_lists_with_equations()
        test_comprehensive_mixed()

        print("\n" + "="*60)
        print("ALL LATEX FORMATTING TESTS PASSED!")
        print("="*60)

    except AssertionError as e:
        print(f"\nFAIL: TEST FAILED: {e}")
    except Exception as e:
        print(f"\nFAIL: UNEXPECTED ERROR: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
    finally:
        cleanup()
