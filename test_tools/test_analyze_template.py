"""
Minimal test for analyze_template functionality using hidden temporary presentation approach.
Tests template analysis without interfering with user's active presentation.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'python'))

import win32com.client
from datetime import datetime
from pathlib import Path
import tempfile
from PIL import Image, ImageDraw, ImageFont
import shutil


def analyze_slide_placeholders(slide):
    """Analyze placeholders in a slide using our proven Microsoft approach."""
    placeholders = []

    try:
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            try:
                # Check if shape is a placeholder (msoPlaceholder = 14)
                if hasattr(shape, 'Type') and shape.Type == 14:
                    # Get placeholder type using our Microsoft-documented approach
                    placeholder_type = shape.PlaceholderFormat.Type

                    placeholder_info = {
                        'index': i,
                        'type_value': placeholder_type,
                        'type_name': get_placeholder_type_name(placeholder_type),
                        'name': shape.Name,
                        'position': f"({round(shape.Left, 1)}, {round(shape.Top, 1)})",
                        'size': f"{round(shape.Width, 1)} x {round(shape.Height, 1)}"
                    }
                    placeholders.append(placeholder_info)
            except:
                continue

    except Exception as e:
        print(f"Error analyzing placeholders: {e}")

    return placeholders


def get_placeholder_type_name(type_value):
    """Convert placeholder type constants to readable names."""
    type_names = {
        1: "ppPlaceholderTitle",
        2: "ppPlaceholderBody",
        3: "ppPlaceholderCenterTitle",
        4: "ppPlaceholderSubtitle",
        7: "ppPlaceholderObject",
        8: "ppPlaceholderChart",
        12: "ppPlaceholderTable",
        13: "ppPlaceholderSlideNumber",
        14: "ppPlaceholderHeader",
        15: "ppPlaceholderFooter",
        16: "ppPlaceholderDate"
    }
    return type_names.get(type_value, f"Unknown_{type_value}")


def create_analysis_folder(template_name):
    """Create output folder in test_tools directory without timestamp."""
    test_tools_dir = Path(__file__).parent
    folder_name = f"{template_name}_analysis"
    output_path = test_tools_dir / folder_name

    # Remove existing folder if it exists (for clean results)
    if output_path.exists():
        import shutil
        shutil.rmtree(output_path)

    output_path.mkdir(exist_ok=True)
    return str(output_path)


def populate_placeholder_defaults(slide):
    """Populate placeholders with default text to make them visible in screenshots."""
    try:
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            try:
                # Check if shape is a placeholder
                if hasattr(shape, 'Type') and shape.Type == 14:
                    if hasattr(shape, 'TextFrame') and shape.TextFrame and hasattr(shape, 'PlaceholderFormat'):
                        placeholder_type = shape.PlaceholderFormat.Type

                        # Add default text based on placeholder type
                        default_texts = {
                            1: "Click to edit Master title style",  # ppPlaceholderTitle
                            2: "Click to edit Master text styles\n• Second level\n  • Third level\n    • Fourth level\n      • Fifth level",  # ppPlaceholderBody
                            3: "Click to edit Master title style",  # ppPlaceholderCenterTitle
                            4: "Click to edit Master subtitle style",  # ppPlaceholderSubtitle
                            7: "Click to add content",  # ppPlaceholderObject
                            8: "Click to add chart",  # ppPlaceholderChart
                            12: "Click to add table",  # ppPlaceholderTable
                        }

                        if placeholder_type in default_texts:
                            shape.TextFrame.TextRange.Text = default_texts[placeholder_type]

            except Exception as e:
                # Skip problematic shapes
                continue

    except Exception as e:
        print(f"Warning: Could not populate placeholder defaults: {e}")


def add_bounding_box_overlays(image_path, slide_data, presentation):
    """Add bounding box overlays like slide_snapshot tool with correct dimensions."""
    try:
        # Load the image
        image = Image.open(image_path)
        draw = ImageDraw.Draw(image)

        # Calculate scaling factors using ACTUAL slide dimensions
        img_width, img_height = image.size

        # Get actual slide dimensions from presentation
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight

        print(f"    📐 Image size: {img_width}x{img_height}")
        print(f"    📐 Slide size: {slide_width}x{slide_height} points")

        scale_x = img_width / slide_width
        scale_y = img_height / slide_height

        print(f"    📐 Scale factors: X={scale_x:.2f}, Y={scale_y:.2f}")

        # Try to load a font
        try:
            font_size = max(12, int(img_width / 100))
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.load_default()

        # Colors (RGB for PIL)
        box_color = (0, 255, 0)      # Green
        bg_color = (255, 255, 0)     # Yellow background
        text_color = (0, 0, 0)       # Black text

        # Draw bounding box and ID for each placeholder
        for placeholder in slide_data:
            try:
                # Parse position and size
                pos_str = placeholder['position'].strip('()')
                x_pos, y_pos = map(float, pos_str.split(', '))

                size_str = placeholder['size']
                width, height = map(float, size_str.split(' x '))

                print(f"      📍 {placeholder['type_name']}: PPT({x_pos}, {y_pos}) Size({width}x{height})")

                # Convert PowerPoint coordinates to image coordinates
                x = int(x_pos * scale_x)
                y = int(y_pos * scale_y)
                w = int(width * scale_x)
                h = int(height * scale_y)

                print(f"      🖼️  Image coords: ({x}, {y}) Size({w}x{h})")

                # Ensure coordinates are within image bounds
                x = max(0, min(x, img_width))
                y = max(0, min(y, img_height))
                w = max(1, min(w, img_width - x))
                h = max(1, min(h, img_height - y))

                # Draw bounding box
                draw.rectangle([x, y, x + w, y + h], outline=box_color, width=3)

                # Draw compact ID label
                id_text = f"ID:{placeholder['index']}"

                # Get text size
                try:
                    bbox = draw.textbbox((0, 0), id_text, font=font)
                    text_w = bbox[2] - bbox[0]
                    text_h = bbox[3] - bbox[1]
                except AttributeError:
                    # Fallback for older PIL versions
                    text_w, text_h = draw.textsize(id_text, font=font)

                # Position label (above box, or below if no space)
                label_x = x
                label_y = y - text_h - 5
                if label_y < 0:
                    label_y = y + h + 5

                # Ensure label stays within image bounds
                label_x = max(0, min(label_x, img_width - text_w - 4))
                label_y = max(0, min(label_y, img_height - text_h - 2))

                # Draw label background and text
                draw.rectangle([label_x, label_y, label_x + text_w + 4, label_y + text_h + 2],
                             fill=bg_color)
                draw.text((label_x + 2, label_y + 1), id_text, fill=text_color, font=font)

            except Exception as e:
                # Skip problematic placeholders
                continue

        # Save the annotated image
        image.save(image_path, "PNG")
        return True

    except Exception as e:
        print(f"Warning: Could not add bounding boxes: {e}")
        return False


def get_template_directories():
    """Get common PowerPoint template directories."""
    import os

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
    """Find a template file by name in standard template directories."""
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


def resolve_template_source(source):
    """
    Resolve template source to actual template information.

    Args:
        source: Can be "current", template name, or full path

    Returns:
        Dictionary with template_path, template_name, and source_type
    """
    try:
        # Case 1: "current" - use active presentation
        if source == "current":
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            active_presentation = ppt_app.ActivePresentation
            return {
                'template_path': active_presentation.FullName,
                'template_name': active_presentation.Name.replace('.pptx', '').replace('.potx', ''),
                'source_type': 'current_presentation',
                'ppt_app': ppt_app
            }

        # Case 2: Full path provided
        elif source.endswith(('.potx', '.potm', '.pot', '.pptx', '.pptm', '.ppt')):
            template_path = Path(source)
            if template_path.exists():
                return {
                    'template_path': str(template_path),
                    'template_name': template_path.stem,
                    'source_type': 'file_path',
                    'ppt_app': win32com.client.GetActiveObject("PowerPoint.Application")
                }
            else:
                return {'error': f"Template file not found: {source}"}

        # Case 3: Template name - search in template directories
        else:
            found_path = find_template_by_name(source)
            if found_path:
                return {
                    'template_path': found_path,
                    'template_name': source,
                    'source_type': 'template_name',
                    'ppt_app': win32com.client.GetActiveObject("PowerPoint.Application")
                }
            else:
                return {'error': f"Template not found: '{source}'. Use list_templates() to see available templates."}

    except Exception as e:
        return {'error': f"Failed to resolve template source: {str(e)}"}


def powerpoint_analyze_template_test(source="current", output_folder_name=None):
    """
    Test version of analyze_template using hidden temporary presentation.

    Args:
        source: "current" for active presentation, template name, or full path
        output_folder_name: Custom folder name (defaults to template name)

    Returns:
        Dictionary with template analysis results
    """
    temp_presentation = None

    try:
        print("🔍 Starting template analysis...")

        # 1. Resolve template source (current, name, or path)
        print(f"🔎 Resolving template source: '{source}'")
        template_info = resolve_template_source(source)

        if 'error' in template_info:
            print(f"❌ {template_info['error']}")
            return {"error": template_info['error']}

        ppt_app = template_info['ppt_app']
        template_path = template_info['template_path']
        template_name = template_info['template_name']
        source_type = template_info['source_type']

        print(f"✅ Resolved template: {template_name} ({source_type})")
        print(f"📄 Template path: {template_path}")

        # 2. Create HIDDEN temporary presentation 🎯
        print("🔒 Creating hidden temporary presentation...")
        temp_presentation = ppt_app.Presentations.Add(WithWindow=False)
        print("✅ Hidden presentation created (user won't see this!)")

        # 3. Apply template to hidden presentation
        print("🎨 Applying template to hidden presentation...")
        temp_presentation.ApplyTemplate(template_path)
        print("✅ Template applied to hidden presentation")

        # 4. Create output folder
        folder_name = output_folder_name or template_name
        output_path = create_analysis_folder(folder_name)
        print(f"📁 Output folder: {output_path}")

        # 5. Analyze each layout in the HIDDEN presentation
        print(f"🔎 Analyzing {temp_presentation.SlideMaster.CustomLayouts.Count} layouts...")
        layouts_data = []

        for i in range(1, temp_presentation.SlideMaster.CustomLayouts.Count + 1):
            layout = temp_presentation.SlideMaster.CustomLayouts(i)
            layout_name = layout.Name
            print(f"  📋 Layout {i}: {layout_name}")

            # Create temporary slide in HIDDEN presentation (user never sees this!)
            temp_slide = temp_presentation.Slides.AddSlide(1, layout)

            # 🎯 KEY FIX: Populate placeholders with default text to make them visible
            print(f"    ✏️  Populating placeholder defaults...")
            populate_placeholder_defaults(temp_slide)

            # Analyze placeholders using our proven approach
            placeholder_data = analyze_slide_placeholders(temp_slide)
            print(f"    🎯 Found {len(placeholder_data)} placeholders")

            # Take screenshot AFTER populating text
            safe_name = layout_name.replace(' ', '-').replace('/', '-').lower()
            screenshot_filename = f"layout-{i}-{safe_name}.png"
            screenshot_path = os.path.join(output_path, screenshot_filename)
            temp_slide.Export(screenshot_path, "PNG")
            print(f"    📸 Screenshot: {screenshot_filename}")

            # 🎯 KEY FIX: Add bounding boxes like slide_snapshot tool
            print(f"    🔲 Adding bounding boxes...")
            add_bounding_box_overlays(screenshot_path, placeholder_data, temp_presentation)

            # Store layout info
            layout_info = {
                "index": i,
                "name": layout_name,
                "screenshot_file": screenshot_filename,
                "placeholders": placeholder_data,
                "placeholder_count": len(placeholder_data)
            }
            layouts_data.append(layout_info)

        # 6. Clean up - close hidden presentation
        print("🧹 Cleaning up hidden presentation...")
        temp_presentation.Close()
        temp_presentation = None
        print("✅ Hidden presentation closed (clean exit)")

        # 7. Return comprehensive results
        result = {
            "success": True,
            "source": source,
            "source_type": source_type,
            "template_name": template_name,
            "template_path": template_path,
            "output_folder": output_path,
            "total_layouts": len(layouts_data),
            "layouts": layouts_data,
            "timestamp": datetime.now().isoformat()
        }

        print("🎉 Template analysis completed successfully!")
        return result

    except Exception as e:
        print(f"❌ Error during template analysis: {str(e)}")

        # Always clean up temp presentation even on error
        if temp_presentation:
            try:
                print("🧹 Emergency cleanup of hidden presentation...")
                temp_presentation.Close()
                print("✅ Emergency cleanup successful")
            except:
                print("⚠️  Could not clean up hidden presentation")

        return {"error": f"Template analysis failed: {str(e)}"}


def print_analysis_results(result):
    """Print detailed analysis results to terminal."""
    if not result.get('success'):
        print(f"\n❌ ANALYSIS FAILED: {result.get('error')}")
        return

    print(f"\n{'='*60}")
    print(f"📊 TEMPLATE ANALYSIS RESULTS")
    print(f"{'='*60}")
    print(f"Source: {result['source']} ({result['source_type']})")
    print(f"Template Name: {result['template_name']}")
    print(f"Template Path: {result['template_path']}")
    print(f"Output Folder: {result['output_folder']}")
    print(f"Total Layouts: {result['total_layouts']}")
    print(f"Analysis Time: {result['timestamp']}")

    print(f"\n📋 LAYOUT DETAILS:")
    print(f"{'-'*60}")

    for layout in result['layouts']:
        print(f"\n{layout['index']}. {layout['name']}")
        print(f"   Screenshot: {layout['screenshot_file']}")
        print(f"   Placeholders: {layout['placeholder_count']}")

        if layout['placeholders']:
            for ph in layout['placeholders']:
                print(f"     • ID:{ph['index']} {ph['type_name']} - {ph['name']}")
                print(f"       Position: {ph['position']}, Size: {ph['size']}")
        else:
            print(f"     • No placeholders found")


if __name__ == "__main__":
    print("🚀 Testing PowerPoint Template Analysis (Enhanced with Template Discovery)")
    print("=" * 80)

    # Example 1: Test with current active presentation
    print("\n" + "="*50)
    print("📋 TEST 1: Current Active Presentation")
    print("="*50)
    result1 = powerpoint_analyze_template_test(source="current")
    print_analysis_results(result1)

    if result1.get('success'):
        print(f"\n✅ SUCCESS! Check the output folder for layout screenshots:")
        print(f"   {result1['output_folder']}")
    else:
        print(f"\n❌ FAILED! Error: {result1.get('error')}")

    # Example 2: Test with template name (user can modify this)
    print("\n" + "="*50)
    print("📋 TEST 2: Template by Name")
    print("="*50)
    print("💡 To test template discovery, replace 'Template_Name_Here' with an actual template name")
    print("   Use the list_templates tool to see available templates")

    # Uncomment and modify the line below to test with an actual template name:
    result2 = powerpoint_analyze_template_test(source="Training")
    print_analysis_results(result2)

    print("\n📋 Usage Examples:")
    print("  analyze_template(source='current')                    # Current presentation")
    print("  analyze_template(source='Business Template')          # Template by name")
    print("  analyze_template(source='C:/path/to/template.potx')   # Full path")
    print("\n💡 Use list_templates() first to discover available template names")