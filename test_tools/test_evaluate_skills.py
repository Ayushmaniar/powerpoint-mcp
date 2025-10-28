"""
Test skills library integration with powerpoint_evaluate tool.
Tests: Using all PowerPoint MCP tools within evaluate code via the 'skills' object.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'python'))

from powerpoint_mcp.tools.evaluate import powerpoint_evaluate, generate_mcp_response
from powerpoint_mcp.tools.presentation import manage_presentation
import win32com.client


def setup_test_presentation():
    """Setup: Create presentation for skills testing."""
    print("\n" + "="*60)
    print("SETUP: Creating Test Presentation")
    print("="*60)

    manage_presentation("create", file_path=r"C:\Users\Rajat\test_evaluate_skills.pptx")
    
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
    
    print("INFO: Test presentation created")
    input("\nPress ENTER to continue to next test...")


def test_circular_layout_with_skills():
    """Test 1: Create circular layout and use skills.populate_placeholder."""
    print("\n" + "="*60)
    print("TEST 1: Circular Layout with Skills Integration")
    print("="*60)

    code = """
# Get slide dimensions
slide_width = slide.Master.Width
slide_height = slide.Master.Height

# Calculate center of slide
center_x = slide_width / 2
center_y = slide_height / 2

# Circular layout parameters
n_circles = 5
radius = min(slide_width, slide_height) * 0.25
circle_size = min(slide_width, slide_height) * 0.10

circles_created = []

for i in range(n_circles):
    # Calculate angle for this circle
    theta = 2 * math.pi * i / n_circles
    
    # Convert polar to cartesian coordinates
    x = center_x + radius * math.cos(theta) - circle_size/2
    y = center_y + radius * math.sin(theta) - circle_size/2
    
    # Create circle shape
    circle = slide.Shapes.AddShape(
        Type=9,  # msoShapeOval
        Left=x,
        Top=y,
        Width=circle_size,
        Height=circle_size
    )
    
    # Give it a unique name
    shape_name = f"Circle_{i+1}"
    circle.Name = shape_name
    circle.Fill.ForeColor.RGB = 0x4472C4  # Blue
    
    # Use skills to populate with formatted text!
    text_content = f"<blue><b>Step {i+1}</b></blue>"
    populate_result = skills.populate_placeholder(shape_name, text_content)
    
    circles_created.append({
        "index": i,
        "name": shape_name,
        "populated": populate_result.get("success", False)
    })

result = {
    "circles_created": len(circles_created),
    "all_populated": all(c["populated"] for c in circles_created),
    "circles": circles_created
}
"""

    result = powerpoint_evaluate(code=code, description="Create circular layout with skills.populate_placeholder")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should succeed"
    assert result["result"]["circles_created"] == 5, "Should create 5 circles"
    assert result["result"]["all_populated"] == True, "All circles should be populated using skills"
    print("\nPASS: Test 1 - Circular layout with skills integration works")
    input("\nPress ENTER to continue to next test...")


def test_snapshot_integration():
    """Test 2: Use skills.snapshot() to get slide info."""
    print("\n" + "="*60)
    print("TEST 2: Using skills.snapshot()")
    print("="*60)

    code = """
# Use skills to get snapshot of current slide
snapshot_result = skills.snapshot(include_screenshot=False)

# Extract information from snapshot
objects_info = []
if snapshot_result.get("success") and "objects" in snapshot_result:
    for obj in snapshot_result["objects"]:
        objects_info.append({
            "name": obj.get("name"),
            "type": obj.get("type")
        })

result = {
    "snapshot_success": snapshot_result.get("success", False),
    "object_count": snapshot_result.get("object_count", 0),
    "objects": objects_info[:3]  # First 3 for brevity
}
"""

    result = powerpoint_evaluate(code=code, description="Use skills.snapshot() to analyze slide")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should succeed"
    assert result["result"]["snapshot_success"] == True, "Snapshot should succeed"
    assert result["result"]["object_count"] > 0, "Should have objects from previous test"
    print("\nPASS: Test 2 - skills.snapshot() works")
    input("\nPress ENTER to continue to next test...")


def test_slide_duplication():
    """Test 3: Use skills.manage_slide() to duplicate."""
    print("\n" + "="*60)
    print("TEST 3: Using skills.manage_slide('duplicate')")
    print("="*60)

    code = """
# Duplicate the current slide
dup_result = skills.manage_slide("duplicate", 1)

# Verify duplication
total_slides_after = presentation.Slides.Count

result = {
    "duplication_success": dup_result.get("success", False),
    "total_slides": total_slides_after,
    "duplicate_result": dup_result
}
"""

    result = powerpoint_evaluate(code=code, description="Use skills.manage_slide to duplicate")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should succeed"
    assert result["result"]["duplication_success"] == True, "Duplication should succeed"
    assert result["result"]["total_slides"] == 2, "Should have 2 slides now"
    print("\nPASS: Test 3 - skills.manage_slide('duplicate') works")
    input("\nPress ENTER to continue to next test...")


def test_speaker_notes():
    """Test 4: Use skills.add_speaker_notes()."""
    print("\n" + "="*60)
    print("TEST 4: Using skills.add_speaker_notes()")
    print("="*60)

    code = """
# Add speaker notes to both slides
notes_text_1 = "This is the original slide with circular layout"
notes_text_2 = "This is the duplicated slide"

notes_result_1 = skills.add_speaker_notes(1, notes_text_1)
notes_result_2 = skills.add_speaker_notes(2, notes_text_2)

result = {
    "notes_1_success": notes_result_1.get("success", False),
    "notes_2_success": notes_result_2.get("success", False),
    "both_added": notes_result_1.get("success", False) and notes_result_2.get("success", False)
}
"""

    result = powerpoint_evaluate(code=code, description="Use skills.add_speaker_notes()")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should succeed"
    assert result["result"]["both_added"] == True, "Both notes should be added"
    print("\nPASS: Test 4 - skills.add_speaker_notes() works")
    input("\nPress ENTER to continue to next test...")


def test_combined_operations():
    """Test 5: Complex workflow combining multiple skills."""
    print("\n" + "="*60)
    print("TEST 5: Combined Skills Operations")
    print("="*60)

    code = """
# 1. Switch to slide 2
switch_result = skills.switch_slide(2)

# 2. Get snapshot to see what's on the slide
snapshot = skills.snapshot(include_screenshot=False)

# 3. Modify shapes using populate_placeholder
modifications_made = []
if snapshot.get("success") and "objects" in snapshot:
    # Find circle shapes and update their text
    for obj in snapshot["objects"]:
        if obj.get("name", "").startswith("Circle_"):
            shape_name = obj["name"]
            # Update with different colored text
            new_text = f"<red><b>COPY</b></red>"
            mod_result = skills.populate_placeholder(shape_name, new_text)
            modifications_made.append({
                "shape": shape_name,
                "success": mod_result.get("success", False)
            })

# 4. Add a note about the modifications
note_text = f"Modified {len(modifications_made)} shapes on this duplicated slide"
notes_result = skills.add_speaker_notes(2, note_text)

result = {
    "switch_success": switch_result.get("success", False),
    "snapshot_success": snapshot.get("success", False),
    "modifications_count": len(modifications_made),
    "all_modifications_success": all(m["success"] for m in modifications_made),
    "notes_added": notes_result.get("success", False)
}
"""

    result = powerpoint_evaluate(code=code, description="Complex workflow with multiple skills")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should succeed"
    assert result["result"]["switch_success"] == True, "Should switch to slide 2"
    assert result["result"]["snapshot_success"] == True, "Should get snapshot"
    assert result["result"]["modifications_count"] > 0, "Should modify shapes"
    assert result["result"]["all_modifications_success"] == True, "All modifications should succeed"
    print("\nPASS: Test 5 - Combined skills operations work")
    input("\nPress ENTER to continue to cleanup...")


def cleanup():
    """Cleanup: Close the test presentation."""
    print("\n" + "="*60)
    print("CLEANUP")
    print("="*60)
    manage_presentation("close", presentation_name="test_evaluate_skills.pptx")
    print("INFO: Test presentation closed")


if __name__ == "__main__":
    try:
        setup_test_presentation()
        test_circular_layout_with_skills()
        test_snapshot_integration()
        test_slide_duplication()
        test_speaker_notes()
        test_combined_operations()

        print("\n" + "="*60)
        print("ALL SKILLS INTEGRATION TESTS PASSED!")
        print("="*60)

    except AssertionError as e:
        print(f"\nFAIL: TEST FAILED: {e}")
    except Exception as e:
        print(f"\nFAIL: UNEXPECTED ERROR: {type(e).__name__}: {e}")
    finally:
        cleanup()
