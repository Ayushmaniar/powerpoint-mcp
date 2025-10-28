"""
PowerPoint template analysis tool for MCP server.
Analyzes template layouts and generates screenshots with placeholder analysis.
"""

import os
import win32com.client
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from typing import Optional


def get_template_directories():
    """Get common PowerPoint template directories."""
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

        # Case 2: Full path provided (don't need active PowerPoint yet)
        elif source.endswith(('.potx', '.potm', '.pot', '.pptx', '.pptm', '.ppt')):
            template_path = Path(source)
            if template_path.exists():
                return {
                    'template_path': str(template_path),
                    'template_name': template_path.stem,
                    'source_type': 'file_path',
                    'ppt_app': None  # Will get/create PowerPoint app later
                }
            else:
                return {'error': f"Template file not found: {source}"}

        # Case 3: Template name - search in template directories (don't need active PowerPoint yet)
        else:
            found_path = find_template_by_name(source)
            if found_path:
                return {
                    'template_path': found_path,
                    'template_name': source,
                    'source_type': 'template_name',
                    'ppt_app': None  # Will get/create PowerPoint app later
                }
            else:
                return {'error': f"Template not found: '{source}'. Use list_templates() to see available templates."}

    except Exception as e:
        return {'error': f"Failed to resolve template source: {str(e)}"}


def get_output_file(template_name: str, filename: Optional[str] = None) -> str:
    """
    Create an organized output file path in template-specific folder.
    Uses the same .powerpoint-mcp directory as slide_snapshot tool with template subfolder.

    Args:
        template_name: Name of the template for subfolder organization
        filename: Optional custom filename

    Returns:
        Full path to the output file in template-specific subfolder
    """
    if filename is None:
        timestamp = datetime.now().isoformat().replace(':', '-').replace('.', '-')
        filename = f"template-{timestamp}.png"

    # Clean template name for folder usage
    safe_template_name = template_name.replace(' ', '-').replace('/', '-').replace('\\', '-')

    # Follow Playwright's pattern: try user directory first, fallback to temp
    try:
        # Try user's home directory first (like Playwright does with rootPath)
        user_home = Path.home()
        template_dir = user_home / ".powerpoint-mcp" / safe_template_name
        template_dir.mkdir(parents=True, exist_ok=True)
        return str(template_dir / filename)
    except (PermissionError, OSError):
        # Fallback to system temp directory (like Playwright's fallback)
        import tempfile
        temp_dir = Path(tempfile.gettempdir()) / "powerpoint-mcp-output" / safe_template_name
        temp_dir.mkdir(parents=True, exist_ok=True)
        return str(temp_dir / filename)


def analyze_slide_placeholders(slide):
    """Analyze placeholders in a slide using Microsoft-documented approach."""
    placeholders = []

    try:
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            try:
                # Check if shape is a placeholder (msoPlaceholder = 14)
                if hasattr(shape, 'Type') and shape.Type == 14:
                    # Get placeholder type using Microsoft-documented approach
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

    except Exception:
        pass

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

            except Exception:
                # Skip problematic shapes
                continue

    except Exception:
        pass


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

        scale_x = img_width / slide_width
        scale_y = img_height / slide_height

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

                # Convert PowerPoint coordinates to image coordinates
                x = int(x_pos * scale_x)
                y = int(y_pos * scale_y)
                w = int(width * scale_x)
                h = int(height * scale_y)

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

            except Exception:
                # Skip problematic placeholders
                continue

        # Save the annotated image
        image.save(image_path, "PNG")
        return True

    except Exception:
        return False


def powerpoint_analyze_template(source="current"):
    """
    Analyze PowerPoint template layouts using hidden temporary presentation.

    Args:
        source: "current" for active presentation, template name, or full path

    Returns:
        Dictionary with template analysis results
    """
    temp_presentation = None

    try:
        # 1. Resolve template source (current, name, or path)
        template_info = resolve_template_source(source)

        if 'error' in template_info:
            return {"error": template_info['error']}

        ppt_app = template_info['ppt_app']
        template_path = template_info['template_path']
        template_name = template_info['template_name']
        source_type = template_info['source_type']

        # 2. Get or create PowerPoint application if needed
        if ppt_app is None:
            try:
                # Try to get existing PowerPoint application
                ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                # If no PowerPoint running, create a new one
                ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                ppt_app.Visible = True  # Make sure it's visible for COM operations

        # 3. Create HIDDEN temporary presentation
        temp_presentation = ppt_app.Presentations.Add(WithWindow=False)

        # 4. Apply template to hidden presentation
        temp_presentation.ApplyTemplate(template_path)

        # 5. Analyze each layout in the HIDDEN presentation
        layouts_data = []
        screenshot_info = {}

        for i in range(1, temp_presentation.SlideMaster.CustomLayouts.Count + 1):
            layout = temp_presentation.SlideMaster.CustomLayouts(i)
            layout_name = layout.Name

            # Create temporary slide in HIDDEN presentation
            temp_slide = temp_presentation.Slides.AddSlide(1, layout)

            # Populate placeholders with default text to make them visible
            populate_placeholder_defaults(temp_slide)

            # Analyze placeholders
            placeholder_data = analyze_slide_placeholders(temp_slide)

            # Take screenshot
            safe_name = layout_name.replace(' ', '-').replace('/', '-').lower()
            screenshot_filename = f"layout-{i}-{safe_name}.png"
            screenshot_path = get_output_file(template_name, screenshot_filename)
            temp_slide.Export(screenshot_path, "PNG")

            # Add bounding boxes
            add_bounding_box_overlays(screenshot_path, placeholder_data, temp_presentation)

            # Store layout info
            layout_info = {
                "index": i,
                "name": layout_name,
                "screenshot_file": screenshot_filename,
                "screenshot_path": screenshot_path,
                "placeholders": placeholder_data,
                "placeholder_count": len(placeholder_data)
            }
            layouts_data.append(layout_info)

            # Track screenshot info for MCP response
            screenshot_info[screenshot_filename] = screenshot_path

        # 6. Clean up - close hidden presentation
        temp_presentation.Close()
        temp_presentation = None

        # 7. Return comprehensive results
        safe_template_name = template_name.replace(' ', '-').replace('/', '-').replace('\\', '-')
        template_screenshot_dir = str(Path.home() / ".powerpoint-mcp" / safe_template_name)

        result = {
            "success": True,
            "source": source,
            "source_type": source_type,
            "template_name": template_name,
            "template_path": template_path,
            "total_layouts": len(layouts_data),
            "layouts": layouts_data,
            "screenshot_directory": template_screenshot_dir,
            "base_screenshot_directory": str(Path.home() / ".powerpoint-mcp"),
            "screenshots": screenshot_info,
            "timestamp": datetime.now().isoformat()
        }

        return result

    except Exception as e:
        # Always clean up temp presentation even on error
        if temp_presentation:
            try:
                temp_presentation.Close()
            except:
                pass

        return {"error": f"Template analysis failed: {str(e)}"}


def generate_mcp_response(result, detailed=False):
    """Generate the MCP tool response for the LLM.

    Args:
        result: Analysis result dictionary
        detailed: If True, include position and size info for placeholders
    """
    if not result.get('success'):
        return f"Template analysis failed: {result.get('error')}"

    # Build response for LLM
    response_lines = [
        f"Template Analysis: {result['template_name']} ({result['source_type']})",
        f"Found {result['total_layouts']} layouts with screenshots and placeholder analysis",
        f"Screenshots saved to: {result['screenshot_directory']}",
        ""
    ]

    # Add layout details
    for layout in result['layouts']:
        response_lines.append(f"Layout {layout['index']}: \"{layout['name']}\"")
        response_lines.append(f"  Screenshot: {layout['screenshot_file']}")
        response_lines.append(f"  Placeholders: {layout['placeholder_count']}")

        if layout['placeholders']:
            for ph in layout['placeholders']:
                response_lines.append(f"    • ID:{ph['index']} {ph['type_name']} - {ph['name']}")
                if detailed:
                    response_lines.append(f"      Position: {ph['position']}, Size: {ph['size']}")
        else:
            response_lines.append(f"    • No placeholders found")

        response_lines.append("")

    # Add usage instructions
    response_lines.extend([
        "Screenshot Usage:",
        f"• Screenshots are saved in template-specific folder: {result['screenshot_directory']}",
        f"• Use Read tool to view screenshots: Read(file_path=\"{result['screenshot_directory']}/layout-1-title-slide.png\")",
        f"• Each screenshot shows green bounding boxes with yellow ID labels for placeholders",
        f"• Template folder structure: ~/.powerpoint-mcp/{result['template_name'].replace(' ', '-')}/",
        ""
    ])

    return "\n".join(response_lines)