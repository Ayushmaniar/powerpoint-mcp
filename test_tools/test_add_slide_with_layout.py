"""
Test file for add_slide_with_layout functionality.
Creates slides with specific template layouts at specified positions.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'python'))

import win32com.client
from datetime import datetime


def get_template_directories():
    """Get common PowerPoint template directories. (Reused from list_templates)"""
    from pathlib import Path

    directories = []
    username = os.environ.get('USERNAME', '')

    # Personal templates directory (Office 365/2019+)
    personal_templates = Path(f"C:/Users/{username}/Documents/Custom Office Templates")
    if personal_templates.exists():
        directories.append(str(personal_templates))

    # User templates directory (AppData)
    user_templates = Path(f"C:/Users/{username}/AppData/Roaming/Microsoft/Templates")
    if user_templates.exists():
        directories.append(str(user_templates))

    # System templates - multiple possible locations
    system_locations = [
        "C:/Program Files/Microsoft Office/Templates",
        "C:/Program Files/Microsoft Office/root/Templates",
        "C:/Program Files (x86)/Microsoft Office/Templates",
        "C:/Program Files (x86)/Microsoft Office/root/Templates"
    ]

    for location in system_locations:
        if Path(location).exists():
            directories.append(location)

    return directories


def find_template_by_name(template_name):
    """Find a template file by name in standard template directories. (Reused from analyze_template)"""
    from pathlib import Path

    template_extensions = {'.potx', '.potm', '.pot'}

    # Search all template directories
    for directory in get_template_directories():
        directory_path = Path(directory)

        # Search recursively for the template
        for file_path in directory_path.rglob('*'):
            if (file_path.is_file() and
                file_path.suffix.lower() in template_extensions and
                file_path.stem.lower() == template_name.lower()):
                return str(file_path)

    return None


def resolve_template_path(template_name):
    """
    Resolve template name to actual file path.

    Args:
        template_name: Name of the template

    Returns:
        Template file path or None if not found
    """
    try:
        # Search for template by name
        template_path = find_template_by_name(template_name)
        if template_path:
            print(f"✅ Found template: {template_name}")
            print(f"   Path: {template_path}")
            return template_path
        else:
            print(f"❌ Template not found: {template_name}")
            return None

    except Exception as e:
        print(f"❌ Error resolving template: {e}")
        return None


def find_layout_in_template(template_path, layout_name):
    """
    Find a specific layout by name within a template.

    Args:
        template_path: Path to the template file
        layout_name: Name of the layout to find

    Returns:
        Layout object and temp presentation, or (None, None) if not found
    """
    try:
        # Get PowerPoint application
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")

        # Create hidden temporary presentation to access template layouts
        print(f"🔍 Searching for layout '{layout_name}' in template...")
        temp_presentation = ppt_app.Presentations.Add(WithWindow=False)
        temp_presentation.ApplyTemplate(template_path)

        # Search for layout by name
        for i in range(1, temp_presentation.SlideMaster.CustomLayouts.Count + 1):
            layout = temp_presentation.SlideMaster.CustomLayouts(i)
            print(f"   Checking layout {i}: '{layout.Name}'")

            if layout.Name.lower() == layout_name.lower():
                print(f"✅ Found layout: '{layout.Name}' (index {i})")
                return layout, temp_presentation

        # Layout not found
        print(f"❌ Layout '{layout_name}' not found in template")
        print(f"   Available layouts:")
        for i in range(1, temp_presentation.SlideMaster.CustomLayouts.Count + 1):
            layout = temp_presentation.SlideMaster.CustomLayouts(i)
            print(f"     {i}. '{layout.Name}'")

        # Clean up temp presentation
        temp_presentation.Close()
        return None, None

    except Exception as e:
        print(f"❌ Error finding layout: {e}")
        return None, None


def test_add_slide_with_layout(template_name, layout_name, after_slide):
    """
    Test function for adding a slide with a specific layout.

    Args:
        template_name: Name of the template (e.g., "Pitchbook")
        layout_name: Name of the layout (e.g., "Title", "Agenda")
        after_slide: Insert slide after this position

    Returns:
        Dictionary with results
    """
    temp_presentation = None

    try:
        print("🚀 Testing Add Slide with Layout")
        print("=" * 50)
        print(f"Template: {template_name}")
        print(f"Layout: {layout_name}")
        print(f"Insert after slide: {after_slide}")
        print()

        # 1. Check if PowerPoint is available and has active presentation
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            active_presentation = ppt_app.ActivePresentation
            original_slide_count = active_presentation.Slides.Count
            print(f"✅ Active presentation: {active_presentation.Name}")
            print(f"   Current slides: {original_slide_count}")
        except:
            return {"error": "No active PowerPoint presentation found. Please open a presentation first."}

        # 2. Validate after_slide parameter
        if after_slide < 0 or after_slide > original_slide_count:
            return {"error": f"Invalid after_slide position. Must be between 0 and {original_slide_count}"}

        # 3. Resolve template name to file path
        template_path = resolve_template_path(template_name)
        if not template_path:
            return {"error": f"Template '{template_name}' not found. Use list_templates() to see available templates."}

        # 4. Find layout in template
        target_layout, temp_presentation = find_layout_in_template(template_path, layout_name)
        if not target_layout:
            return {"error": f"Layout '{layout_name}' not found in template '{template_name}'. Use analyze_template() to see available layouts."}

        # Store layout name before closing temp presentation (to avoid object reference issues)
        actual_layout_name = target_layout.Name

        # 5. Create slide from template without affecting existing slides
        print(f"📄 Creating slide from template layout without changing existing slides...")

        # Create a temporary slide in the temp presentation with the target layout
        temp_slide = temp_presentation.Slides.AddSlide(1, target_layout)

        # Copy the slide from temp presentation to active presentation
        new_slide_position = after_slide + 1
        print(f"📄 Copying slide to position {new_slide_position}...")

        # Copy the slide to the active presentation
        temp_slide.Copy()

        # Paste it at the correct position in the active presentation
        if new_slide_position <= active_presentation.Slides.Count:
            # Insert at specific position
            active_presentation.Slides.Paste(new_slide_position)
        else:
            # Add at the end
            active_presentation.Slides.Paste()

        new_slide_count = active_presentation.Slides.Count

        print(f"✅ Slide added successfully without changing existing slides!")
        print(f"   New slide position: {new_slide_position}")
        print(f"   Total slides: {new_slide_count} (was {original_slide_count})")
        print(f"   Original slides remain unchanged! ✅")

        # 6. Clean up temp presentation
        if temp_presentation:
            temp_presentation.Close()
            temp_presentation = None

        # 7. Return success result
        result = {
            "success": True,
            "new_slide_number": new_slide_position,
            "layout_name": actual_layout_name,
            "template_name": template_name,
            "original_slide_count": original_slide_count,
            "new_slide_count": new_slide_count,
            "message": f"Added slide {new_slide_position} using '{actual_layout_name}' layout from '{template_name}' template (existing slides unchanged)"
        }

        print(f"🎉 Success! {result['message']}")
        return result

    except Exception as e:
        # Always clean up temp presentation on error
        if temp_presentation:
            try:
                temp_presentation.Close()
            except:
                pass

        error_message = f"Failed to add slide: {str(e)}"
        print(f"❌ {error_message}")
        return {"error": error_message}


def print_test_results(result):
    """Print formatted test results."""
    print(f"\n{'='*60}")
    print(f"📊 ADD SLIDE WITH LAYOUT TEST RESULTS")
    print(f"{'='*60}")

    if result.get('success'):
        print(f"✅ SUCCESS!")
        print(f"   New slide number: {result['new_slide_number']}")
        print(f"   Layout used: '{result['layout_name']}'")
        print(f"   Template: '{result['template_name']}'")
        print(f"   Slide count: {result['original_slide_count']} → {result['new_slide_count']}")
        print(f"   Message: {result['message']}")
    else:
        print(f"❌ FAILED!")
        print(f"   Error: {result['error']}")

    print(f"{'='*60}")


if __name__ == "__main__":
    print("🚀 Testing Add Slide with Layout Functionality")
    print("=" * 80)

    # Test configuration - modify these to test different scenarios
    test_template = "Training"      # Change to any template from list_templates()
    test_layout = "Title Slide"     # Change to any layout from analyze_template()
    test_after_slide = 1           # Insert after this slide number

    print(f"📋 Test Configuration:")
    print(f"   Template: {test_template}")
    print(f"   Layout: {test_layout}")
    print(f"   Insert after slide: {test_after_slide}")
    print()

    print("💡 Instructions:")
    print("   1. Open PowerPoint and have an active presentation")
    print("   2. Modify test_template and test_layout variables above")
    print("   3. Run this test to add a slide with the specified layout")
    print()

    # Run the test
    result = test_add_slide_with_layout(test_template, test_layout, test_after_slide)

    # Print results
    print_test_results(result)

    print(f"\n📋 Next Steps:")
    if result.get('success'):
        print(f"   • The new slide is ready at position {result['new_slide_number']}")
        print(f"   • Check PowerPoint to see the new slide with '{result['layout_name']}' layout")
    else:
        print(f"   • Fix the error above and try again")
        print(f"   • Use list_templates() to see available templates")
        print(f"   • Use analyze_template() to see available layouts")