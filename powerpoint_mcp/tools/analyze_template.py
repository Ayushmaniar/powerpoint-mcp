"""
PowerPoint template analysis tool for MCP server.
Analyzes template layouts and generates screenshots with placeholder analysis.
"""

import os
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from typing import Optional

from ..backends import get_backend
from ..backends.types import UnsupportedFeatureError


def get_template_directories():
    """Get template directories from the backend."""
    backend = get_backend()
    template_dirs = backend.get_template_directories()
    return [td.path for td in template_dirs]


def find_template_by_name(template_name):
    """Find a template file by name in standard template directories."""
    template_extensions = {'.potx', '.potm', '.pot'}

    for directory in get_template_directories():
        directory_path = Path(directory)

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
        if source == "current":
            backend = get_backend()
            backend.connect()
            info = backend.get_active_presentation_info()
            return {
                'template_path': info.full_path,
                'template_name': info.name.replace('.pptx', '').replace('.potx', ''),
                'source_type': 'current_presentation',
            }

        elif source.endswith(('.potx', '.potm', '.pot', '.pptx', '.pptm', '.ppt')):
            template_path = Path(source)
            if template_path.exists():
                return {
                    'template_path': str(template_path),
                    'template_name': template_path.stem,
                    'source_type': 'file_path',
                }
            else:
                return {'error': f"Template file not found: {source}"}

        else:
            found_path = find_template_by_name(source)
            if found_path:
                return {
                    'template_path': found_path,
                    'template_name': source,
                    'source_type': 'template_name',
                }
            else:
                return {'error': f"Template not found: '{source}'. Use list_templates() to see available templates."}

    except Exception as e:
        return {'error': f"Failed to resolve template source: {str(e)}"}


def get_output_file(template_name: str, filename: Optional[str] = None) -> str:
    """
    Create an organized output file path in template-specific folder.
    """
    if filename is None:
        timestamp = datetime.now().isoformat().replace(':', '-').replace('.', '-')
        filename = f"template-{timestamp}.png"

    safe_template_name = template_name.replace(' ', '-').replace('/', '-').replace('\\', '-')

    try:
        user_home = Path.home()
        template_dir = user_home / ".powerpoint-mcp" / safe_template_name
        template_dir.mkdir(parents=True, exist_ok=True)
        return str(template_dir / filename)
    except (PermissionError, OSError):
        import tempfile
        temp_dir = Path(tempfile.gettempdir()) / "powerpoint-mcp-output" / safe_template_name
        temp_dir.mkdir(parents=True, exist_ok=True)
        return str(temp_dir / filename)


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


def _get_font():
    """Get appropriate font for bounding box labels."""
    import sys
    font_candidates = ["arial.ttf"]
    if sys.platform == "darwin":
        font_candidates = [
            "/System/Library/Fonts/Helvetica.ttc",
            "/System/Library/Fonts/SFNSText.ttf",
            "/Library/Fonts/Arial.ttf",
            "arial.ttf",
        ]
    for font_path in font_candidates:
        try:
            return font_path
        except:
            continue
    return None


def add_bounding_box_overlays(image_path, slide_data, slide_width, slide_height):
    """Add bounding box overlays with correct dimensions."""
    try:
        image = Image.open(image_path)
        draw = ImageDraw.Draw(image)

        img_width, img_height = image.size
        scale_x = img_width / slide_width
        scale_y = img_height / slide_height

        font_path = _get_font()
        try:
            font_size = max(12, int(img_width / 100))
            font = ImageFont.truetype(font_path, font_size)
        except:
            font = ImageFont.load_default()

        box_color = (0, 255, 0)
        bg_color = (255, 255, 0)
        text_color = (0, 0, 0)

        for placeholder in slide_data:
            try:
                pos_str = placeholder.position.strip('()')
                x_pos, y_pos = map(float, pos_str.split(', '))

                size_str = placeholder.size
                width, height = map(float, size_str.split(' x '))

                x = int(x_pos * scale_x)
                y = int(y_pos * scale_y)
                w = int(width * scale_x)
                h = int(height * scale_y)

                x = max(0, min(x, img_width))
                y = max(0, min(y, img_height))
                w = max(1, min(w, img_width - x))
                h = max(1, min(h, img_height - y))

                draw.rectangle([x, y, x + w, y + h], outline=box_color, width=3)

                id_text = f"ID:{placeholder.index}"

                try:
                    bbox = draw.textbbox((0, 0), id_text, font=font)
                    text_w = bbox[2] - bbox[0]
                    text_h = bbox[3] - bbox[1]
                except AttributeError:
                    text_w, text_h = draw.textsize(id_text, font=font)

                label_x = x
                label_y = y - text_h - 5
                if label_y < 0:
                    label_y = y + h + 5

                label_x = max(0, min(label_x, img_width - text_w - 4))
                label_y = max(0, min(label_y, img_height - text_h - 2))

                draw.rectangle([label_x, label_y, label_x + text_w + 4, label_y + text_h + 2],
                             fill=bg_color)
                draw.text((label_x + 2, label_y + 1), id_text, fill=text_color, font=font)

            except Exception:
                continue

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
    try:
        template_info = resolve_template_source(source)

        if 'error' in template_info:
            return {"error": template_info['error']}

        template_path = template_info['template_path']
        template_name = template_info['template_name']
        source_type = template_info['source_type']

        backend = get_backend()
        backend.connect()

        layouts_data = []
        screenshot_info = {}

        with backend.hidden_presentation(template_path) as hidden:
            layouts = hidden.get_layouts()

            for layout in layouts:
                hidden.add_slide(layout.index)
                hidden.populate_placeholder_defaults(1)

                placeholder_data = hidden.get_placeholders(1)

                safe_name = layout.name.replace(' ', '-').replace('/', '-').lower()
                screenshot_filename = f"layout-{layout.index}-{safe_name}.png"
                screenshot_path = get_output_file(template_name, screenshot_filename)
                hidden.export_slide(1, screenshot_path)

                slide_w, slide_h = hidden.get_dimensions()
                add_bounding_box_overlays(screenshot_path, placeholder_data, slide_w, slide_h)

                layout_info = {
                    "index": layout.index,
                    "name": layout.name,
                    "screenshot_file": screenshot_filename,
                    "screenshot_path": screenshot_path,
                    "placeholders": [
                        {
                            'index': ph.index,
                            'type_value': ph.type_value,
                            'type_name': ph.type_name,
                            'name': ph.name,
                            'position': ph.position,
                            'size': ph.size,
                        }
                        for ph in placeholder_data
                    ],
                    "placeholder_count": len(placeholder_data)
                }
                layouts_data.append(layout_info)
                screenshot_info[screenshot_filename] = screenshot_path

                # Delete the temp slide to prepare for next layout
                # (The hidden presentation always adds at position 1)

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

    except UnsupportedFeatureError as e:
        return {"error": str(e)}
    except Exception as e:
        return {"error": f"Template analysis failed: {str(e)}"}


def generate_mcp_response(result, detailed=False):
    """Generate the MCP tool response for the LLM."""
    if not result.get('success'):
        return f"Template analysis failed: {result.get('error')}"

    response_lines = [
        f"Template Analysis: {result['template_name']} ({result['source_type']})",
        f"Found {result['total_layouts']} layouts with screenshots and placeholder analysis",
        f"Screenshots saved to: {result['screenshot_directory']}",
        ""
    ]

    for layout in result['layouts']:
        response_lines.append(f"Layout {layout['index']}: \"{layout['name']}\"")
        response_lines.append(f"  Screenshot: {layout['screenshot_file']}")
        response_lines.append(f"  Placeholders: {layout['placeholder_count']}")

        if layout['placeholders']:
            for ph in layout['placeholders']:
                response_lines.append(f"    ID:{ph['index']} {ph['type_name']} - {ph['name']}")
                if detailed:
                    response_lines.append(f"      Position: {ph['position']}, Size: {ph['size']}")
        else:
            response_lines.append(f"    No placeholders found")

        response_lines.append("")

    response_lines.extend([
        "Screenshot Usage:",
        f"  Screenshots are saved in template-specific folder: {result['screenshot_directory']}",
        f"  Use Read tool to view screenshots: Read(file_path=\"{result['screenshot_directory']}/layout-1-title-slide.png\")",
        f"  Each screenshot shows green bounding boxes with yellow ID labels for placeholders",
        f"  Template folder structure: ~/.powerpoint-mcp/{result['template_name'].replace(' ', '-')}/",
        ""
    ])

    return "\n".join(response_lines)
