"""
Test basic functionality of powerpoint_evaluate tool.
Tests: basic code execution, error handling, return values.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'python'))

from powerpoint_mcp.tools.evaluate import powerpoint_evaluate, generate_mcp_response
from powerpoint_mcp.tools.presentation import manage_presentation
import win32com.client


def test_basic_execution():
    """Test 1: Basic code execution with skills integration."""
    print("\n" + "="*60)
    print("TEST 1: Basic Code Execution with Skills")
    print("="*60)

    # Create a test presentation
    manage_presentation("create", file_path=r"C:\Users\Rajat\test_evaluate_basic.pptx")
    
    # Activate PowerPoint window so user can see it
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        ppt.Activate()
        # Ensure there's at least one slide
        if ppt.ActivePresentation and ppt.ActivePresentation.Slides.Count == 0:
            ppt.ActivePresentation.Slides.Add(Index=1, Layout=12)  # ppLayoutBlank
        if ppt.ActivePresentation and ppt.ActivePresentation.Windows.Count > 0:
            ppt.ActivePresentation.Windows(1).Activate()
    except:
        pass

    code = """
# Simple test - add a textbox
textbox = slide.Shapes.AddTextbox(
    Orientation=1,
    Left=100,
    Top=100,
    Width=200,
    Height=50
)
textbox.Name = "WelcomeBox"

# Use skills to populate with formatted text
skills.populate_placeholder("WelcomeBox", "<blue><b>Hello from evaluate with skills!</b></blue>")
"""

    result = powerpoint_evaluate(code=code, description="Add textbox and populate with skills")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Basic execution should succeed"
    print("\nPASS: Test 1 - Basic execution with skills works")
    input("\nPress ENTER to continue to next test...")


def test_with_return_value():
    """Test 2: Code execution with return value."""
    print("\n" + "="*60)
    print("TEST 2: Code Execution with Return Value")
    print("="*60)

    code = """
# Count shapes on slide
shape_count = slide.Shapes.Count
result = {
    "total_shapes": shape_count,
    "slide_width": slide.Master.Width,
    "slide_height": slide.Master.Height
}
"""

    result = powerpoint_evaluate(code=code, description="Count shapes and get slide dimensions")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Execution should succeed"
    assert result.get("result") is not None, "Should have return value"
    assert "total_shapes" in result["result"], "Should have shape count"
    print("\nPASS: Test 2 - Return values work")
    input("\nPress ENTER to continue to next test...")


def test_math_operations():
    """Test 3: Math module availability."""
    print("\n" + "="*60)
    print("TEST 3: Math Operations")
    print("="*60)

    code = """
import math
# Calculate some values using math module
result = {
    "pi": math.pi,
    "sqrt_2": math.sqrt(2),
    "sin_90": math.sin(math.pi/2)
}
"""

    result = powerpoint_evaluate(code=code, description="Test math module")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Math operations should succeed"
    assert abs(result["result"]["pi"] - 3.14159) < 0.001, "Pi value should be correct"
    print("\nPASS: Test 3 - Math module works")
    input("\nPress ENTER to continue to next test...")


def test_numpy_availability():
    """Test 4: Numpy availability (if installed)."""
    print("\n" + "="*60)
    print("TEST 4: Numpy Availability")
    print("="*60)

    code = """
if has_numpy:
    # Use numpy for calculations
    angles = [np.pi/4, np.pi/2, np.pi]
    result = {
        "numpy_available": True,
        "cos_values": [float(np.cos(a)) for a in angles]
    }
else:
    result = {"numpy_available": False}
"""

    result = powerpoint_evaluate(code=code, description="Test numpy availability")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Numpy check should succeed"
    if result["result"]["numpy_available"]:
        print("INFO: Numpy is available")
    else:
        print("WARN: Numpy not available (this is OK)")
    print("\nPASS: Test 4 - Numpy check works")
    input("\nPress ENTER to continue to next test...")


def test_error_handling():
    """Test 5: Error handling for invalid code."""
    print("\n" + "="*60)
    print("TEST 5: Error Handling")
    print("="*60)

    code = """
# This will cause an error - undefined variable
invalid_variable.do_something()
"""

    result = powerpoint_evaluate(code=code, description="Test error handling")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("error") is not None, "Should return error"
    assert "NameError" in result["error"], "Should identify NameError"
    print("\nPASS: Test 5 - Error handling works")
    input("\nPress ENTER to continue to next test...")


def test_slide_number_parameter():
    """Test 6: Using slide_number parameter."""
    print("\n" + "="*60)
    print("TEST 6: Slide Number Parameter")
    print("="*60)

    # Add a second slide
    code_add_slide = """
presentation.Slides.Add(Index=2, Layout=12)  # ppLayoutBlank
"""
    powerpoint_evaluate(code=code_add_slide, description="Add second slide")

    # Now target slide 2 specifically
    code = """
slide.Shapes.AddTextbox(
    Orientation=1,
    Left=150,
    Top=150,
    Width=300,
    Height=80
)
result = {"slide_number": slide.SlideNumber}
"""

    result = powerpoint_evaluate(
        code=code,
        slide_number=2,
        description="Add textbox to slide 2"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should succeed"
    assert result["result"]["slide_number"] == 2, "Should operate on slide 2"
    print("\nPASS: Test 6 - Slide number parameter works")
    input("\nPress ENTER to continue to next test...")


def test_invalid_slide_number():
    """Test 7: Invalid slide number handling."""
    print("\n" + "="*60)
    print("TEST 7: Invalid Slide Number")
    print("="*60)

    code = "pass"
    result = powerpoint_evaluate(code=code, slide_number=999)
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("error") is not None, "Should return error"
    assert "out of range" in result["error"].lower(), "Should indicate out of range"
    print("\nPASS: Test 7 - Invalid slide number handled correctly")
    input("\nPress ENTER to continue to next test...")


def test_presentation_context():
    """Test 8: Access to presentation context and skills."""
    print("\n" + "="*60)
    print("TEST 8: Presentation Context Access with Skills")
    print("="*60)

    code = """
# Use skills to get snapshot
snapshot = skills.snapshot(include_screenshot=False)

result = {
    "presentation_name": presentation.Name,
    "total_slides": presentation.Slides.Count,
    "has_ppt_object": ppt is not None,
    "has_slide_object": slide is not None,
    "has_skills_object": skills is not None,
    "snapshot_success": snapshot.get("success", False),
    "object_count_from_snapshot": snapshot.get("object_count", 0)
}
"""

    result = powerpoint_evaluate(code=code, description="Check context objects including skills")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should succeed"
    assert result["result"]["has_ppt_object"] == True, "Should have ppt object"
    assert result["result"]["has_slide_object"] == True, "Should have slide object"
    assert result["result"]["has_skills_object"] == True, "Should have skills object"
    assert result["result"]["snapshot_success"] == True, "Snapshot should succeed"
    print("\nPASS: Test 8 - Presentation context with skills accessible")
    input("\nPress ENTER to continue to cleanup...")


def cleanup():
    """Cleanup: Close the test presentation."""
    print("\n" + "="*60)
    print("CLEANUP")
    print("="*60)
    manage_presentation("close", presentation_name="test_evaluate_basic.pptx")
    print("INFO: Test presentation closed")


if __name__ == "__main__":
    try:
        test_basic_execution()
        test_with_return_value()
        test_math_operations()
        test_numpy_availability()
        test_error_handling()
        test_slide_number_parameter()
        test_invalid_slide_number()
        test_presentation_context()

        print("\n" + "="*60)
        print("ALL BASIC TESTS PASSED! PASS")
        print("="*60)

    except AssertionError as e:
        print(f"\nFAIL: TEST FAILED: {e}")
    except Exception as e:
        print(f"\nFAIL: UNEXPECTED ERROR: {type(e).__name__}: {e}")
    finally:
        cleanup()