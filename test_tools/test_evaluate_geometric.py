"""
Test geometric calculations with powerpoint_evaluate tool.
Tests: circular layouts, grid layouts, relative positioning.
Uses slide dimensions programmatically (not hardcoded values).
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'python'))

from powerpoint_mcp.tools.evaluate import powerpoint_evaluate, generate_mcp_response
from powerpoint_mcp.tools.presentation import manage_presentation
import win32com.client


def test_circular_layout():
    """Test 1: Create circular flowchart with 5 circles."""
    print("\n" + "="*60)
    print("TEST 1: Circular Flowchart Layout")
    print("="*60)

    # Create test presentation
    manage_presentation("create", file_path=r"C:\Users\Rajat\test_evaluate_geometric.pptx")
    
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
# Get slide dimensions
slide_width = slide.Master.Width
slide_height = slide.Master.Height

# Calculate center of slide
center_x = slide_width / 2
center_y = slide_height / 2

# Circular layout parameters
n_circles = 5
radius = min(slide_width, slide_height) * 0.25  # 25% of smaller dimension
circle_size = min(slide_width, slide_height) * 0.10  # 10% of smaller dimension

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

    # Set shape name and color
    circle.Name = f"Circle_{i+1}"
    circle.Fill.ForeColor.RGB = 0x4472C4  # Blue
    
    # Use skills to populate with formatted text
    skills.populate_placeholder(f"Circle_{i+1}", f"<b>Step {i+1}</b>")

    circles_created.append({
        "index": i,
        "x": round(x, 2),
        "y": round(y, 2),
        "theta_degrees": round(math.degrees(theta), 2)
    })

result = {
    "circles_created": len(circles_created),
    "circle_positions": circles_created,
    "slide_dimensions": {
        "width": slide_width,
        "height": slide_height
    }
}
"""

    result = powerpoint_evaluate(code=code, description="Create circular flowchart with 5 steps")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Circular layout should succeed"
    assert result["result"]["circles_created"] == 5, "Should create 5 circles"
    print("\nPASS: Test 1 - Circular layout created")
    input("\nPress ENTER to continue to next test...")


def test_grid_layout():
    """Test 2: Create grid layout of rectangles."""
    print("\n" + "="*60)
    print("TEST 2: Grid Layout")
    print("="*60)

    code = """
# Get slide dimensions
slide_width = slide.Master.Width
slide_height = slide.Master.Height

# Grid parameters
rows = 3
cols = 4
total_items = rows * cols

# Calculate grid dimensions with margins
margin_x = slide_width * 0.05  # 5% margin
margin_y = slide_height * 0.10  # 10% margin
grid_width = slide_width - 2 * margin_x
grid_height = slide_height - 2 * margin_y

# Calculate item size and spacing
spacing_ratio = 0.1  # 10% spacing between items
item_width = (grid_width - (cols - 1) * grid_width * spacing_ratio) / cols
item_height = (grid_height - (rows - 1) * grid_height * spacing_ratio) / rows
spacing_x = grid_width * spacing_ratio
spacing_y = grid_height * spacing_ratio

# Create grid
grid_items = []
for i in range(total_items):
    row, col = divmod(i, cols)

    # Calculate position
    x = margin_x + col * (item_width + spacing_x)
    y = margin_y + row * (item_height + spacing_y)

    # Create rectangle
    rect = slide.Shapes.AddShape(
        Type=1,  # msoShapeRectangle
        Left=x,
        Top=y,
        Width=item_width,
        Height=item_height
    )

    # Set shape name and color
    rect.Name = f"GridItem_{i+1}"
    rect.Fill.ForeColor.RGB = 0x70AD47  # Green
    
    # Use skills to populate with formatted text
    skills.populate_placeholder(f"GridItem_{i+1}", f"<green><b>Item {i+1}</b></green>")

    grid_items.append({
        "index": i,
        "row": row,
        "col": col,
        "x": round(x, 2),
        "y": round(y, 2)
    })

result = {
    "items_created": len(grid_items),
    "grid_dimensions": {"rows": rows, "cols": cols},
    "item_dimensions": {
        "width": round(item_width, 2),
        "height": round(item_height, 2)
    },
    "positions": grid_items[:4]  # First 4 items for verification
}
"""

    result = powerpoint_evaluate(code=code, slide_number=1, description="Create 3x4 grid layout")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Grid layout should succeed"
    assert result["result"]["items_created"] == 12, "Should create 12 items (3x4)"
    print("\nPASS: Test 2 - Grid layout created")
    input("\nPress ENTER to continue to next test...")


def test_spiral_pattern():
    """Test 3: Create spiral pattern."""
    print("\n" + "="*60)
    print("TEST 3: Spiral Pattern")
    print("="*60)

    code = """
# Get slide dimensions
slide_width = slide.Master.Width
slide_height = slide.Master.Height

# Calculate center
center_x = slide_width / 2
center_y = slide_height / 2

# Spiral parameters
n_points = 8
max_radius = min(slide_width, slide_height) * 0.35
circle_size = min(slide_width, slide_height) * 0.06

spiral_points = []

for i in range(n_points):
    # Spiral: radius increases with angle
    # Create a spiral that does ~2.5 full rotations
    theta = 5 * math.pi * (i / n_points)  # 5π = 2.5 rotations
    radius = max_radius * (i / n_points)

    # Calculate position
    x = center_x + radius * math.cos(theta) - circle_size/2
    y = center_y + radius * math.sin(theta) - circle_size/2

    # Create circle
    circle = slide.Shapes.AddShape(
        Type=9,  # msoShapeOval
        Left=x,
        Top=y,
        Width=circle_size,
        Height=circle_size
    )

    # Color gradient from red to blue
    hue = i / n_points
    if hue < 0.5:
        # Red to yellow
        rgb = int(0xFF0000 + 0x00FF00 * (hue * 2))
    else:
        # Yellow to blue
        rgb = int(0xFFFF00 - 0xFF0000 * ((hue - 0.5) * 2))

    circle.Fill.ForeColor.RGB = rgb

    spiral_points.append({
        "index": i,
        "radius": round(radius, 2),
        "angle_degrees": round(math.degrees(theta), 2)
    })

result = {
    "points_created": len(spiral_points),
    "pattern": "spiral",
    "positions": spiral_points
}
"""

    result = powerpoint_evaluate(code=code, description="Create spiral pattern")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Spiral pattern should succeed"
    assert result["result"]["points_created"] == 8, "Should create 8 points"
    print("\nPASS: Test 3 - Spiral pattern created")
    input("\nPress ENTER to continue to next test...")


def test_relative_positioning():
    """Test 4: Position shapes relative to slide corners."""
    print("\n" + "="*60)
    print("TEST 4: Relative Positioning (Corners & Edges)")
    print("="*60)

    code = """
# Get slide dimensions
slide_width = slide.Master.Width
slide_height = slide.Master.Height

# Shape size
box_size = min(slide_width, slide_height) * 0.08
margin = min(slide_width, slide_height) * 0.05

# Position boxes at corners and edges
positions = {
    "top_left": (margin, margin),
    "top_right": (slide_width - box_size - margin, margin),
    "bottom_left": (margin, slide_height - box_size - margin),
    "bottom_right": (slide_width - box_size - margin, slide_height - box_size - margin),
    "center": ((slide_width - box_size)/2, (slide_height - box_size)/2),
    "top_center": ((slide_width - box_size)/2, margin),
    "bottom_center": ((slide_width - box_size)/2, slide_height - box_size - margin),
    "left_center": (margin, (slide_height - box_size)/2),
    "right_center": (slide_width - box_size - margin, (slide_height - box_size)/2)
}

created_shapes = []

for name, (x, y) in positions.items():
    shape = slide.Shapes.AddShape(
        Type=5,  # msoShapeDiamond
        Left=x,
        Top=y,
        Width=box_size,
        Height=box_size
    )
    shape.TextFrame.TextRange.Text = name.replace("_", "\\n")
    shape.TextFrame.TextRange.Font.Size = 8
    shape.Fill.ForeColor.RGB = 0xC55A11  # Orange

    created_shapes.append({
        "name": name,
        "x": round(x, 2),
        "y": round(y, 2)
    })

result = {
    "shapes_created": len(created_shapes),
    "positions": created_shapes
}
"""

    result = powerpoint_evaluate(code=code, description="Position shapes at corners and edges")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Relative positioning should succeed"
    assert result["result"]["shapes_created"] == 9, "Should create 9 shapes"
    print("\nPASS: Test 4 - Relative positioning works")
    input("\nPress ENTER to continue to next test...")


def test_connecting_shapes():
    """Test 5: Create shapes with connecting lines."""
    print("\n" + "="*60)
    print("TEST 5: Connected Flowchart")
    print("="*60)

    code = """
# Get slide dimensions
slide_width = slide.Master.Width
slide_height = slide.Master.Height

# Create vertical flowchart
n_steps = 4
margin_top = slide_height * 0.15
margin_bottom = slide_height * 0.15
available_height = slide_height - margin_top - margin_bottom

box_width = slide_width * 0.30
box_height = slide_height * 0.12
spacing = (available_height - n_steps * box_height) / (n_steps - 1)

center_x = (slide_width - box_width) / 2
flowchart_steps = []

# Create boxes
for i in range(n_steps):
    y = margin_top + i * (box_height + spacing)

    box = slide.Shapes.AddShape(
        Type=1,  # msoShapeRectangle
        Left=center_x,
        Top=y,
        Width=box_width,
        Height=box_height
    )
    box.Name = f"Process_{i+1}"
    box.Fill.ForeColor.RGB = 0x5B9BD5  # Blue
    
    # Use skills to populate with formatted text
    skills.populate_placeholder(f"Process_{i+1}", f"<b>Process {i+1}</b>")

    flowchart_steps.append({
        "index": i,
        "center_x": center_x + box_width/2,
        "center_y": y + box_height/2
    })

# Create connecting arrows
arrows_created = 0
for i in range(n_steps - 1):
    # Arrow from bottom of box i to top of box i+1
    start_x = flowchart_steps[i]["center_x"]
    start_y = flowchart_steps[i]["center_y"] + box_height/2
    end_x = flowchart_steps[i+1]["center_x"]
    end_y = flowchart_steps[i+1]["center_y"] - box_height/2

    arrow = slide.Shapes.AddConnector(
        Type=1,  # msoConnectorStraight
        BeginX=start_x,
        BeginY=start_y,
        EndX=end_x,
        EndY=end_y
    )
    arrow.Line.EndArrowheadStyle = 3  # msoArrowheadTriangle
    arrow.Line.Weight = 2.5
    arrows_created += 1

result = {
    "boxes_created": len(flowchart_steps),
    "arrows_created": arrows_created,
    "flowchart_type": "vertical"
}
"""

    result = powerpoint_evaluate(code=code, description="Create vertical flowchart with connectors")
    print("\nResult:")
    print(generate_mcp_response(result))

    assert result.get("success") == True, "Connected flowchart should succeed"
    assert result["result"]["boxes_created"] == 4, "Should create 4 boxes"
    assert result["result"]["arrows_created"] == 3, "Should create 3 arrows"
    print("\nPASS: Test 5 - Connected flowchart created")
    input("\nPress ENTER to continue to cleanup...")


def cleanup():
    """Cleanup: Close the test presentation."""
    print("\n" + "="*60)
    print("CLEANUP")
    print("="*60)
    manage_presentation("close", presentation_name="test_evaluate_geometric.pptx")
    print("INFO: Test presentation closed")


if __name__ == "__main__":
    try:
        test_circular_layout()
        test_grid_layout()
        test_spiral_pattern()
        test_relative_positioning()
        test_connecting_shapes()

        print("\n" + "="*60)
        print("ALL GEOMETRIC TESTS PASSED!")
        print("="*60)

    except AssertionError as e:
        print(f"\nFAIL: TEST FAILED: {e}")
    except Exception as e:
        print(f"\nFAIL: UNEXPECTED ERROR: {type(e).__name__}: {e}")
    finally:
        cleanup()