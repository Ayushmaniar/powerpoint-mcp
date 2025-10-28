"""
Test script for the manage_slide tool.
Tests duplicate, delete, and move operations.
"""

import sys
import os

# Add the parent directory to Python path to import our tools
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'python'))

from powerpoint_mcp.tools.manage_slide import powerpoint_manage_slide, generate_mcp_response
from powerpoint_mcp.tools.presentation import manage_presentation
from powerpoint_mcp.tools.add_slide_with_layout import powerpoint_add_slide_with_layout


def test_manage_slide():
    """Test the manage_slide tool with all operations."""
    print("Testing manage_slide tool...")
    print("=" * 50)

    # First, ensure we have a presentation with multiple slides
    print("\n1. Creating test presentation...")
    create_result = manage_presentation("create")
    print(f"Create result: {create_result}")

    # Add a few slides for testing
    print("\n2. Adding test slides...")
    for i in range(3):
        # We'll add blank slides by using a basic layout - check what's available
        slide_result = powerpoint_add_slide_with_layout("current", "Title and Content", i)
        print(f"Added slide {i+2}: {slide_result}")

    print("\n3. Testing DUPLICATE operation...")
    # Test basic duplication
    result = powerpoint_manage_slide("duplicate", 2)
    print("Duplicate slide 2 (default position):")
    print(generate_mcp_response(result))
    print()

    # Test duplication with specific target position
    result = powerpoint_manage_slide("duplicate", 1, 5)
    print("Duplicate slide 1 to position 5:")
    print(generate_mcp_response(result))
    print()

    print("\n4. Testing MOVE operation...")
    # Test moving a slide
    result = powerpoint_manage_slide("move", 3, 1)
    print("Move slide 3 to position 1:")
    print(generate_mcp_response(result))
    print()

    # Test moving to same position (should be no-op)
    result = powerpoint_manage_slide("move", 2, 2)
    print("Move slide 2 to position 2 (no-op):")
    print(generate_mcp_response(result))
    print()

    print("\n5. Testing DELETE operation...")
    # Test deleting a slide
    result = powerpoint_manage_slide("delete", 4)
    print("Delete slide 4:")
    print(generate_mcp_response(result))
    print()

    print("\n6. Testing ERROR conditions...")
    # Test invalid operation
    result = powerpoint_manage_slide("invalid", 1)
    print("Invalid operation:")
    print(generate_mcp_response(result))
    print()

    # Test invalid slide number
    result = powerpoint_manage_slide("duplicate", 999)
    print("Invalid slide number:")
    print(generate_mcp_response(result))
    print()

    # Test move without target_position
    result = powerpoint_manage_slide("move", 1, None)
    print("Move without target_position:")
    print(generate_mcp_response(result))
    print()

    print("\n7. Final presentation state:")
    # You can manually check the presentation to verify the operations worked
    print("Check your PowerPoint presentation to verify the slide management operations.")
    print("The tool should have created, duplicated, moved, and deleted slides as requested.")


if __name__ == "__main__":
    try:
        test_manage_slide()
        print("\n✅ Test completed successfully!")
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()