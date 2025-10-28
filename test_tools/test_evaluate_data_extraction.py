"""
Test data extraction capabilities with powerpoint_evaluate tool.
Tests: extracting shape positions, text content, formatting, colors, etc.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'python'))

from powerpoint_mcp.tools.evaluate import powerpoint_evaluate, generate_mcp_response
from powerpoint_mcp.tools.presentation import manage_presentation
import win32com.client


def setup_test_presentation():
    """Setup: Create presentation with various shapes for testing."""
    print("\n" + "="*60)
    print("SETUP: Creating Test Presentation")
    print("="*60)

    manage_presentation("create", file_path=r"C:\Users\Rajat\test_evaluate_data.pptx")
    
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

    # Add some test shapes
    setup_code = """
# Get slide dimensions
slide_width = slide.Master.Width
slide_height = slide.Master.Height

# Add various shapes for testing
# Shape 1: Textbox with formatted text
textbox = slide.Shapes.AddTextbox(
    Orientation=1,
    Left=slide_width * 0.1,
    Top=slide_height * 0.1,
    Width=slide_width * 0.3,
    Height=slide_height * 0.15
)
textbox.TextFrame.TextRange.Text = "Bold Blue Text"
textbox.TextFrame.TextRange.Font.Bold = True
textbox.TextFrame.TextRange.Font.Color.RGB = 0x0000FF
textbox.Name = "FormattedText"

# Shape 2: Rectangle with fill
rect = slide.Shapes.AddShape(
    Type=1,  # msoShapeRectangle
    Left=slide_width * 0.5,
    Top=slide_height * 0.1,
    Width=slide_width * 0.2,
    Height=slide_height * 0.2
)
rect.Fill.ForeColor.RGB = 0xFF0000  # Red
rect.Name = "RedRectangle"

# Shape 3: Circle
circle = slide.Shapes.AddShape(
    Type=9,  # msoShapeOval
    Left=slide_width * 0.1,
    Top=slide_height * 0.4,
    Width=slide_width * 0.15,
    Height=slide_height * 0.15
)
circle.Fill.ForeColor.RGB = 0x00FF00  # Green
circle.Name = "GreenCircle"

# Shape 4: Textbox with multi-line
multi_text = slide.Shapes.AddTextbox(
    Orientation=1,
    Left=slide_width * 0.5,
    Top=slide_height * 0.5,
    Width=slide_width * 0.4,
    Height=slide_height * 0.3
)
multi_text.TextFrame.TextRange.Text = "Line 1\\nLine 2\\nLine 3"
multi_text.Name = "MultiLineText"
"""

    result = powerpoint_evaluate(code=setup_code, description="Setup test shapes")
    print(generate_mcp_response(result))
    print("INFO: Test shapes created")
    input("\nPress ENTER to continue to next test...")


def test_extract_all_shapes():
    """Test 1: Extract all shape information using skills.snapshot()."""
    print("\n" + "="*60)
    print("TEST 1: Extract All Shape Information with Skills")
    print("="*60)

    code = """
# Use skills.snapshot() to get all shape data efficiently
snapshot = skills.snapshot(include_screenshot=False)

# Also manually extract for comparison
shapes_info_manual = []
for shape in slide.Shapes:
    info = {
        "name": shape.Name,
        "id": shape.Id,
        "type": shape.Type,
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2)
    }
    shapes_info_manual.append(info)

result = {
    "total_shapes_manual": len(shapes_info_manual),
    "total_shapes_snapshot": snapshot.get("object_count", 0),
    "snapshot_success": snapshot.get("success", False),
    "shapes_manual": shapes_info_manual,
    "shapes_snapshot": snapshot.get("objects", [])[:3]  # First 3 for brevity
}
"""

    result = powerpoint_evaluate(code=code, description="Extract shape data with skills.snapshot()")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should extract shape data"
    assert result["result"]["total_shapes_manual"] > 0, "Should have shapes"
    assert result["result"]["snapshot_success"] == True, "Snapshot should succeed"
    assert result["result"]["total_shapes_manual"] == result["result"]["total_shapes_snapshot"], "Counts should match"
    print("\nPASS: Test 1 - Shape extraction with skills works")
    input("\nPress ENTER to continue to next test...")


def test_extract_text_content():
    """Test 2: Extract text content from all shapes."""
    print("\n" + "="*60)
    print("TEST 2: Extract Text Content")
    print("="*60)

    code = """
text_shapes = []

for shape in slide.Shapes:
    if shape.HasTextFrame:
        if shape.TextFrame.HasText:
            text_info = {
                "name": shape.Name,
                "text": shape.TextFrame.TextRange.Text,
                "character_count": len(shape.TextFrame.TextRange.Text),
                "has_bold": shape.TextFrame.TextRange.Font.Bold == -1
            }
            text_shapes.append(text_info)

result = {
    "text_shapes_found": len(text_shapes),
    "text_content": text_shapes
}
"""

    result = powerpoint_evaluate(code=code, description="Extract text from shapes")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should extract text"
    assert result["result"]["text_shapes_found"] > 0, "Should find text shapes"
    print("\nPASS: Test 2 - Text extraction works")
    input("\nPress ENTER to continue to next test...")


def test_extract_colors():
    """Test 3: Extract fill colors from shapes."""
    print("\n" + "="*60)
    print("TEST 3: Extract Fill Colors")
    print("="*60)

    code = """
colored_shapes = []

for shape in slide.Shapes:
    if shape.Fill.Visible:
        try:
            rgb = shape.Fill.ForeColor.RGB
            # Convert RGB integer to hex string
            r = rgb & 0xFF
            g = (rgb >> 8) & 0xFF
            b = (rgb >> 16) & 0xFF
            hex_color = f"#{r:02X}{g:02X}{b:02X}"

            colored_shapes.append({
                "name": shape.Name,
                "rgb_int": rgb,
                "rgb_hex": hex_color,
                "rgb_components": {"r": r, "g": g, "b": b}
            })
        except:
            pass

result = {
    "colored_shapes": len(colored_shapes),
    "colors": colored_shapes
}
"""

    result = powerpoint_evaluate(code=code, description="Extract shape colors")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should extract colors"
    print("\nPASS: Test 3 - Color extraction works")
    input("\nPress ENTER to continue to next test...")


def test_find_specific_shapes():
    """Test 4: Find shapes by criteria."""
    print("\n" + "="*60)
    print("TEST 4: Find Shapes by Criteria")
    print("="*60)

    code = """
# Find shapes by type
rectangles = []
circles = []
textboxes = []

for shape in slide.Shapes:
    if shape.Type == 1:  # msoShapeRectangle
        rectangles.append(shape.Name)
    elif shape.Type == 9:  # msoShapeOval
        circles.append(shape.Name)
    elif shape.Type == 17:  # msoTextBox
        textboxes.append(shape.Name)

result = {
    "rectangles": rectangles,
    "circles": circles,
    "textboxes": textboxes
}
"""

    result = powerpoint_evaluate(code=code, description="Find shapes by type")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should find shapes by type"
    print("\nPASS: Test 4 - Shape filtering works")
    input("\nPress ENTER to continue to next test...")


def test_calculate_layout_metrics():
    """Test 5: Calculate layout metrics (bounding box, center, etc.)."""
    print("\n" + "="*60)
    print("TEST 5: Calculate Layout Metrics")
    print("="*60)

    code = """
# Calculate bounding box of all shapes
if slide.Shapes.Count > 0:
    min_x = min(shape.Left for shape in slide.Shapes)
    min_y = min(shape.Top for shape in slide.Shapes)
    max_x = max(shape.Left + shape.Width for shape in slide.Shapes)
    max_y = max(shape.Top + shape.Height for shape in slide.Shapes)

    # Calculate centers
    centers = []
    for shape in slide.Shapes:
        centers.append({
            "name": shape.Name,
            "center_x": round(shape.Left + shape.Width/2, 2),
            "center_y": round(shape.Top + shape.Height/2, 2)
        })

    # Calculate center of mass
    avg_x = sum(c["center_x"] for c in centers) / len(centers)
    avg_y = sum(c["center_y"] for c in centers) / len(centers)

    result = {
        "bounding_box": {
            "left": round(min_x, 2),
            "top": round(min_y, 2),
            "right": round(max_x, 2),
            "bottom": round(max_y, 2),
            "width": round(max_x - min_x, 2),
            "height": round(max_y - min_y, 2)
        },
        "center_of_mass": {
            "x": round(avg_x, 2),
            "y": round(avg_y, 2)
        },
        "shape_centers": centers
    }
else:
    result = {"error": "No shapes on slide"}
"""

    result = powerpoint_evaluate(code=code, description="Calculate layout metrics")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should calculate metrics"
    assert "bounding_box" in result["result"], "Should have bounding box"
    assert "center_of_mass" in result["result"], "Should have center of mass"
    print("\nPASS: Test 5 - Layout metrics calculated")
    input("\nPress ENTER to continue to next test...")


def test_find_overlapping_shapes():
    """Test 6: Detect overlapping shapes."""
    print("\n" + "="*60)
    print("TEST 6: Detect Overlapping Shapes")
    print("="*60)

    code = """
def rectangles_overlap(shape1, shape2):
    # Check if two rectangles overlap
    x1_left = shape1.Left
    x1_right = shape1.Left + shape1.Width
    y1_top = shape1.Top
    y1_bottom = shape1.Top + shape1.Height

    x2_left = shape2.Left
    x2_right = shape2.Left + shape2.Width
    y2_top = shape2.Top
    y2_bottom = shape2.Top + shape2.Height

    # No overlap if one is completely to the left/right/above/below the other
    if x1_right < x2_left or x2_right < x1_left:
        return False
    if y1_bottom < y2_top or y2_bottom < y1_top:
        return False

    return True

overlaps = []
shapes_list = list(slide.Shapes)

for i, shape1 in enumerate(shapes_list):
    for shape2 in shapes_list[i+1:]:
        if rectangles_overlap(shape1, shape2):
            overlaps.append({
                "shape1": shape1.Name,
                "shape2": shape2.Name
            })

result = {
    "overlapping_pairs": len(overlaps),
    "overlaps": overlaps
}
"""

    result = powerpoint_evaluate(code=code, description="Find overlapping shapes")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should check for overlaps"
    print("\nPASS: Test 6 - Overlap detection works")
    input("\nPress ENTER to continue to next test...")


def test_extract_by_name():
    """Test 7: Extract specific shape by name."""
    print("\n" + "="*60)
    print("TEST 7: Extract Shape by Name")
    print("="*60)

    code = """
# Find shape named "RedRectangle"
target_name = "RedRectangle"
found_shape = None

for shape in slide.Shapes:
    if shape.Name == target_name:
        found_shape = {
            "name": shape.Name,
            "id": shape.Id,
            "type": shape.Type,
            "position": {
                "left": round(shape.Left, 2),
                "top": round(shape.Top, 2)
            },
            "dimensions": {
                "width": round(shape.Width, 2),
                "height": round(shape.Height, 2)
            }
        }
        break

result = {
    "search_name": target_name,
    "found": found_shape is not None,
    "shape_data": found_shape
}
"""

    result = powerpoint_evaluate(code=code, description="Find shape by name")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Should find shape by name"
    print("\nPASS: Test 7 - Name-based extraction works")
    input("\nPress ENTER to continue to cleanup...")


def cleanup():
    """Cleanup: Close the test presentation."""
    print("\n" + "="*60)
    print("CLEANUP")
    print("="*60)
    manage_presentation("close", presentation_name="test_evaluate_data.pptx")
    print("INFO: Test presentation closed")


if __name__ == "__main__":
    try:
        setup_test_presentation()
        test_extract_all_shapes()
        test_extract_text_content()
        test_extract_colors()
        test_find_specific_shapes()
        test_calculate_layout_metrics()
        test_find_overlapping_shapes()
        test_extract_by_name()

        print("\n" + "="*60)
        print("ALL DATA EXTRACTION TESTS PASSED!")
        print("="*60)

    except AssertionError as e:
        print(f"\nFAIL: TEST FAILED: {e}")
    except Exception as e:
        print(f"\nFAIL: UNEXPECTED ERROR: {type(e).__name__}: {e}")
    finally:
        cleanup()