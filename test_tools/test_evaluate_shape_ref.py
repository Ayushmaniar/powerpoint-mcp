"""
Test shape_ref parameter functionality with powerpoint_evaluate tool.
Tests: targeting specific shapes, manipulating shapes by reference.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'python'))

from powerpoint_mcp.tools.evaluate import powerpoint_evaluate, generate_mcp_response
from powerpoint_mcp.tools.presentation import manage_presentation
import win32com.client


def setup_test_presentation():
    """Setup: Create presentation with named shapes."""
    print("\n" + "="*60)
    print("SETUP: Creating Test Presentation")
    print("="*60)

    manage_presentation("create", file_path=r"C:\Users\Rajat\test_evaluate_shape_ref.pptx")
    
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

    # Add test shapes with specific names
    setup_code = """
# Get slide dimensions
slide_width = slide.Master.Width
slide_height = slide.Master.Height

# Create 3 shapes with specific names
shapes_to_create = [
    {"name": "Target1", "x": 0.2, "y": 0.2, "color": 0xFF0000},  # Red
    {"name": "Target2", "x": 0.5, "y": 0.3, "color": 0x00FF00},  # Green
    {"name": "Target3", "x": 0.3, "y": 0.6, "color": 0x0000FF},  # Blue
]

for spec in shapes_to_create:
    shape = slide.Shapes.AddShape(
        Type=1,  # msoShapeRectangle
        Left=slide_width * spec["x"],
        Top=slide_height * spec["y"],
        Width=slide_width * 0.15,
        Height=slide_height * 0.10
    )
    shape.Name = spec["name"]
    shape.Fill.ForeColor.RGB = spec["color"]
    shape.TextFrame.TextRange.Text = spec["name"]
"""

    result = powerpoint_evaluate(code=setup_code, description="Setup named shapes")
    print(generate_mcp_response(result))
    print("INFO: Test shapes created")
    input("\nPress ENTER to continue to next test...")


def test_access_shape_by_name():
    """Test 1: Access specific shape using shape_ref parameter."""
    print("\n" + "="*60)
    print("TEST 1: Access Shape by Name")
    print("="*60)

    code = """
# The 'shape' variable is automatically set to the target shape
result = {
    "shape_name": shape.Name,
    "shape_id": shape.Id,
    "shape_type": shape.Type,
    "position": {
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2)
    },
    "dimensions": {
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2)
    },
    "has_text": shape.HasTextFrame
}
"""

    result = powerpoint_evaluate(
        code=code,
        shape_ref="Target1",
        description="Get info about Target1"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should access shape by name"
    assert result["result"]["shape_name"] == "Target1", "Should be Target1"
    print("\nPASS: Test 1 - Shape access by name works")
    input("\nPress ENTER to continue to next test...")


def test_modify_shape_text():
    """Test 2: Modify specific shape's text using shape_ref and skills."""
    print("\n" + "="*60)
    print("TEST 2: Modify Shape Text with Skills")
    print("="*60)

    code = """
# Get old text
old_text = shape.TextFrame.TextRange.Text

# Use skills to populate with formatted text
populate_result = skills.populate_placeholder(shape.Name, "<blue><b>Modified with Skills!</b></blue>")

# Get new text
new_text = shape.TextFrame.TextRange.Text

result = {
    "shape_name": shape.Name,
    "old_text": old_text,
    "new_text": new_text,
    "text_changed": old_text != new_text,
    "skills_success": populate_result.get("success", False)
}
"""

    result = powerpoint_evaluate(
        code=code,
        shape_ref="Target2",
        description="Modify Target2 text with skills"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should modify text"
    assert result["result"]["text_changed"] == True, "Text should change"
    assert result["result"]["skills_success"] == True, "Skills should succeed"
    print("\nPASS: Test 2 - Shape text modification with skills works")
    input("\nPress ENTER to continue to next test...")


def test_move_shape():
    """Test 3: Move specific shape using shape_ref."""
    print("\n" + "="*60)
    print("TEST 3: Move Shape")
    print("="*60)

    code = """
# Get current position
old_left = shape.Left
old_top = shape.Top

# Move shape by offset
offset_x = 50
offset_y = 30
shape.Left = old_left + offset_x
shape.Top = old_top + offset_y

result = {
    "shape_name": shape.Name,
    "old_position": {
        "left": round(old_left, 2),
        "top": round(old_top, 2)
    },
    "new_position": {
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2)
    },
    "offset_applied": {
        "x": offset_x,
        "y": offset_y
    }
}
"""

    result = powerpoint_evaluate(
        code=code,
        shape_ref="Target3",
        description="Move Target3 by offset"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should move shape"
    print("\nPASS: Test 3 - Shape movement works")
    input("\nPress ENTER to continue to next test...")


def test_resize_shape():
    """Test 4: Resize specific shape using shape_ref."""
    print("\n" + "="*60)
    print("TEST 4: Resize Shape")
    print("="*60)

    code = """
# Get current dimensions
old_width = shape.Width
old_height = shape.Height

# Resize (double the size)
shape.Width = old_width * 1.5
shape.Height = old_height * 1.5

result = {
    "shape_name": shape.Name,
    "old_dimensions": {
        "width": round(old_width, 2),
        "height": round(old_height, 2)
    },
    "new_dimensions": {
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2)
    },
    "scale_factor": 1.5
}
"""

    result = powerpoint_evaluate(
        code=code,
        shape_ref="Target1",
        description="Resize Target1"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should resize shape"
    print("\nPASS: Test 4 - Shape resizing works")
    input("\nPress ENTER to continue to next test...")


def test_change_shape_color():
    """Test 5: Change shape fill color using shape_ref."""
    print("\n" + "="*60)
    print("TEST 5: Change Shape Color")
    print("="*60)

    code = """
# Get old color
old_rgb = shape.Fill.ForeColor.RGB

# Change to purple
new_color = 0x800080  # Purple
shape.Fill.ForeColor.RGB = new_color

# Verify change
new_rgb = shape.Fill.ForeColor.RGB

result = {
    "shape_name": shape.Name,
    "old_color_hex": f"0x{old_rgb:06X}",
    "new_color_hex": f"0x{new_rgb:06X}",
    "color_changed": old_rgb != new_rgb
}
"""

    result = powerpoint_evaluate(
        code=code,
        shape_ref="Target2",
        description="Change Target2 color to purple"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should change color"
    assert result["result"]["color_changed"] == True, "Color should change"
    print("\nPASS: Test 5 - Shape color change works")
    input("\nPress ENTER to continue to next test...")


def test_rotate_shape():
    """Test 6: Rotate shape using shape_ref."""
    print("\n" + "="*60)
    print("TEST 6: Rotate Shape")
    print("="*60)

    code = """
# Get old rotation
old_rotation = shape.Rotation

# Rotate by 45 degrees
shape.Rotation = old_rotation + 45

result = {
    "shape_name": shape.Name,
    "old_rotation": round(old_rotation, 2),
    "new_rotation": round(shape.Rotation, 2),
    "rotation_applied": 45
}
"""

    result = powerpoint_evaluate(
        code=code,
        shape_ref="Target3",
        description="Rotate Target3 by 45 degrees"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should rotate shape"
    print("\nPASS: Test 6 - Shape rotation works")
    input("\nPress ENTER to continue to next test...")


def test_shape_ref_by_id():
    """Test 7: Access shape using ID instead of name."""
    print("\n" + "="*60)
    print("TEST 7: Access Shape by ID")
    print("="*60)

    # First, get the ID of Target1
    get_id_code = """
target_id = None
for s in slide.Shapes:
    if s.Name == "Target1":
        target_id = str(s.Id)
        break
result = {"target_id": target_id}
"""

    id_result = powerpoint_evaluate(code=get_id_code, description="Get Target1 ID")
    target_id = id_result["result"]["target_id"]

    # Now access by ID
    code = """
result = {
    "accessed_by": "ID",
    "shape_name": shape.Name,
    "shape_id": shape.Id,
    "verification": "Accessed by ID successfully"
}
"""

    result = powerpoint_evaluate(
        code=code,
        shape_ref=target_id,
        description="Access shape by ID"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should access shape by ID"
    assert result["result"]["shape_name"] == "Target1", "Should be Target1"
    print("\nPASS: Test 7 - Shape access by ID works")
    input("\nPress ENTER to continue to next test...")


def test_shape_ref_not_found():
    """Test 8: Handle shape_ref that doesn't exist."""
    print("\n" + "="*60)
    print("TEST 8: Handle Non-Existent Shape Reference")
    print("="*60)

    code = """
result = {"should_not_execute": True}
"""

    result = powerpoint_evaluate(
        code=code,
        shape_ref="NonExistentShape",
        description="Try to access non-existent shape"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("error") is not None, "Should return error"
    assert "not found" in result["error"].lower(), "Should indicate shape not found"
    print("\nPASS: Test 8 - Non-existent shape handled correctly")
    input("\nPress ENTER to continue to next test...")


def test_combined_shape_and_slide_operations():
    """Test 9: Combine shape_ref with slide-level operations and skills."""
    print("\n" + "="*60)
    print("TEST 9: Combined Shape, Slide Operations and Skills")
    print("="*60)

    code = """
# Work with specific shape and also access slide
shape_center_x = shape.Left + shape.Width / 2
shape_center_y = shape.Top + shape.Height / 2

slide_center_x = slide.Master.Width / 2
slide_center_y = slide.Master.Height / 2

# Calculate distance from shape to slide center
distance = math.sqrt(
    (shape_center_x - slide_center_x)**2 +
    (shape_center_y - slide_center_y)**2
)

# Use skills to get snapshot data
snapshot = skills.snapshot(include_screenshot=False)
total_shapes = snapshot.get("object_count", 0)

# Add speaker notes using skills
notes_text = f"Shape '{shape.Name}' is {round(distance, 2)} points from center"
notes_result = skills.add_speaker_notes(1, notes_text)

result = {
    "target_shape": shape.Name,
    "shape_center": {
        "x": round(shape_center_x, 2),
        "y": round(shape_center_y, 2)
    },
    "slide_center": {
        "x": round(slide_center_x, 2),
        "y": round(slide_center_y, 2)
    },
    "distance_from_center": round(distance, 2),
    "total_shapes_on_slide": total_shapes,
    "snapshot_success": snapshot.get("success", False),
    "notes_added": notes_result.get("success", False)
}
"""

    result = powerpoint_evaluate(
        code=code,
        shape_ref="Target2",
        description="Analyze Target2 with skills integration"
    )
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should combine operations"
    assert result["result"]["target_shape"] == "Target2", "Should target correct shape"
    assert result["result"]["snapshot_success"] == True, "Snapshot should succeed"
    assert result["result"]["notes_added"] == True, "Notes should be added"
    print("\nPASS: Test 9 - Combined operations with skills work")
    input("\nPress ENTER to continue to cleanup...")


def cleanup():
    """Cleanup: Close the test presentation."""
    print("\n" + "="*60)
    print("CLEANUP")
    print("="*60)
    manage_presentation("close", presentation_name="test_evaluate_shape_ref.pptx")
    print("INFO: Test presentation closed")


if __name__ == "__main__":
    try:
        setup_test_presentation()
        test_access_shape_by_name()
        test_modify_shape_text()
        test_move_shape()
        test_resize_shape()
        test_change_shape_color()
        test_rotate_shape()
        test_shape_ref_by_id()
        test_shape_ref_not_found()
        test_combined_shape_and_slide_operations()

        print("\n" + "="*60)
        print("ALL SHAPE_REF TESTS PASSED!")
        print("="*60)

    except AssertionError as e:
        print(f"\nFAIL: TEST FAILED: {e}")
    except Exception as e:
        print(f"\nFAIL: UNEXPECTED ERROR: {type(e).__name__}: {e}")
    finally:
        cleanup()