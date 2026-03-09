"""
PowerPoint slide context snapshot tool.

Provides comprehensive slide content analysis similar to Playwright's browser_snapshot,
but returns detailed slide context including all shapes, text, tables, and formatting.
Now includes screenshot functionality with object bounding box overlays.
"""

import os
import sys
import tempfile
from typing import Optional
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

from ..backends import get_backend


def generate_markdown_table(table_cells_html):
    """Generate a markdown table with HTML formatted cell content."""
    if not table_cells_html or len(table_cells_html) == 0:
        return "Empty table"

    try:
        col_count = len(table_cells_html[0]) if table_cells_html else 0
        if col_count == 0:
            return "Empty table"

        markdown_lines = []

        if len(table_cells_html) > 0:
            header_row = "| " + " | ".join(table_cells_html[0]) + " |"
            markdown_lines.append(header_row)

            separator = "| " + " | ".join(["---"] * col_count) + " |"
            markdown_lines.append(separator)

            for row in table_cells_html[1:]:
                padded_row = row + [""] * (col_count - len(row))
                data_row = "| " + " | ".join(padded_row[:col_count]) + " |"
                markdown_lines.append(data_row)

        return "\n".join(markdown_lines)

    except Exception as e:
        return f"Error generating markdown table: {e}"


def get_output_file(filename: Optional[str] = None) -> str:
    """Create an organized output file path following Playwright's pattern."""
    if filename is None:
        timestamp = datetime.now().isoformat().replace(':', '-').replace('.', '-')
        filename = f"slide-{timestamp}.png"

    try:
        user_home = Path.home()
        output_dir = user_home / ".powerpoint-mcp"
        output_dir.mkdir(exist_ok=True)
        return str(output_dir / filename)
    except (PermissionError, OSError):
        temp_dir = Path(tempfile.gettempdir()) / "powerpoint-mcp-output"
        temp_dir.mkdir(exist_ok=True)
        return str(temp_dir / filename)


def _get_font():
    """Get appropriate font for bounding box labels."""
    if sys.platform == "darwin":
        candidates = [
            "/System/Library/Fonts/Helvetica.ttc",
            "/System/Library/Fonts/SFNSText.ttf",
            "/Library/Fonts/Arial.ttf",
        ]
        for path in candidates:
            if os.path.exists(path):
                return path
    return "arial.ttf"


def add_bounding_box_overlays(image: Image.Image, shapes, slide_width, slide_height) -> Image.Image:
    """Add bounding box overlays to the slide image."""
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

    for shape in shapes:
        try:
            x = int(shape.left * scale_x)
            y = int(shape.top * scale_y)
            w = int(shape.width * scale_x)
            h = int(shape.height * scale_y)

            draw.rectangle([x, y, x + w, y + h], outline=box_color, width=2)

            id_text = f"ID:{shape.id}"

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

            draw.rectangle([label_x, label_y, label_x + text_w + 4, label_y + text_h + 2],
                         fill=bg_color)
            draw.text((label_x + 2, label_y + 1), id_text, fill=text_color, font=font)

        except Exception:
            continue

    return image


def format_slide_context(slide_data):
    """Format slide data into a readable context string."""
    context_parts = [
        "=== POWERPOINT SLIDE CONTEXT ===",
        f"Slide: {slide_data['slide_number']} of {slide_data['total_slides']}",
        f"Layout: {slide_data['layout']}",
        f"Objects: {slide_data['object_count']}",
        "",
        "=== SLIDE CONTENT ==="
    ]

    for i, shape in enumerate(slide_data['shapes'], 1):
        context_parts.append(f"\n--- Object {i}: {shape['name']} ---")
        context_parts.append(f"Type: {shape['type']}")
        context_parts.append(f"ID: {shape['id']}")
        context_parts.append(f"Position: {shape.get('position', 'Unknown')}")
        context_parts.append(f"Size: {shape.get('size', 'Unknown')}")

        if 'html_text' in shape and shape['html_text']:
            context_parts.append(f"Text: {shape['html_text']}")
            if 'font' in shape:
                context_parts.append(f"Font: {shape['font']}")
        elif 'text' in shape and shape['text']:
            context_parts.append(f"Text: {shape['text']}")
            if 'font' in shape:
                context_parts.append(f"Font: {shape['font']}")

        if 'is_table' in shape and shape['is_table']:
            context_parts.append(f"Table: {shape['table_info']}")
            if 'container_type' in shape:
                context_parts.append(f"Container: {shape['container_type']}")

            if 'table_markdown' in shape:
                context_parts.append("Table content (Markdown with HTML formatting):")
                context_parts.append(shape['table_markdown'])
            elif 'table_content_html' in shape:
                context_parts.append("Table content (HTML formatted):")
                for row_idx, row_data in enumerate(shape['table_content_html']):
                    row_str = " | ".join(row_data)
                    context_parts.append(f"  Row {row_idx + 1}: {row_str}")
            elif 'table_content' in shape:
                context_parts.append("Table content (plain text):")
                for row_idx, row_data in enumerate(shape['table_content']):
                    row_str = " | ".join(row_data)
                    context_parts.append(f"  Row {row_idx + 1}: {row_str}")

            if 'table_hyperlinks' in shape and shape['table_hyperlinks']:
                context_parts.append("Table Hyperlinks:")
                for link in shape['table_hyperlinks']:
                    if isinstance(link, dict):
                        cell_pos = link.get('cell_position', 'Unknown position')
                        address = link.get('address', 'Unknown URL')
                        text = link.get('text', 'No text')
                        context_parts.append(f"  -> {address} (Text: {text}, Cell: {cell_pos})")

            if 'table_error' in shape:
                context_parts.append(f"Table Error: {shape['table_error']}")

        if 'chart_info' in shape:
            context_parts.append(f"Chart: {shape['chart_info']}")
            if 'chart_title' in shape:
                context_parts.append(f"Title: {shape['chart_title']}")

            if 'chart_data' in shape and isinstance(shape['chart_data'], dict):
                chart_data = shape['chart_data']

                if 'axes' in chart_data and isinstance(chart_data['axes'], dict):
                    axes = chart_data['axes']
                    if 'category_axis' in axes and isinstance(axes['category_axis'], dict):
                        context_parts.append(f"X-Axis: {axes['category_axis'].get('title', 'No title')}")
                    if 'value_axis' in axes and isinstance(axes['value_axis'], dict):
                        context_parts.append(f"Y-Axis: {axes['value_axis'].get('title', 'No title')}")

                if 'categories' in chart_data and isinstance(chart_data['categories'], list) and chart_data['categories']:
                    if len(chart_data['categories']) <= 10:
                        context_parts.append(f"Categories: {chart_data['categories']}")
                    else:
                        context_parts.append(f"Categories: {len(chart_data['categories'])} items ({chart_data['categories'][:5]}...)")

                if 'series' in chart_data and isinstance(chart_data['series'], list):
                    context_parts.append("Chart Data Series:")
                    for si, series in enumerate(chart_data['series'], 1):
                        if isinstance(series, dict):
                            series_name = series.get('name', f'Series {si}')
                            values = series.get('values', [])
                            context_parts.append(f"  Series {si}: {series_name}")
                            if values and isinstance(values, list) and len(values) <= 10:
                                context_parts.append(f"    Values: {values}")
                            elif values and isinstance(values, list):
                                context_parts.append(f"    Values: {len(values)} data points")

            if 'chart_error' in shape:
                context_parts.append(f"Chart Error: {shape['chart_error']}")

        if 'hyperlinks' in shape and shape['hyperlinks']:
            context_parts.append("Hyperlinks:")
            for link in shape['hyperlinks']:
                if isinstance(link, dict):
                    context_parts.append(f"  -> {link.get('address', 'Unknown URL')} (Text: {link.get('text', 'No text')})")

    # Notes section
    if slide_data.get('notes'):
        context_parts.extend(["", "=== SLIDE NOTES (HTML formatted) ===", slide_data['notes']])

    # Comments section
    if slide_data.get('comments'):
        context_parts.extend(["", "=== SLIDE COMMENTS ==="])
        for i, comment in enumerate(slide_data['comments'], 1):
            if isinstance(comment, dict):
                context_parts.append(f"Comment {i}:")
                context_parts.append(f"  Author: {comment.get('author', 'Unknown')}")
                context_parts.append(f"  Date: {comment.get('date', 'Unknown')}")
                context_parts.append(f"  Position: {comment.get('position', 'Unknown')}")

                if 'associated_object' in comment:
                    obj = comment['associated_object']
                    context_parts.append(f"  Associated Object: {obj.get('name', 'Unknown')} (ID: {obj.get('id', 'Unknown')}, Type: {obj.get('type', 'Unknown')})")

                context_parts.append(f"  Text: {comment.get('text', 'No text')}")
            else:
                context_parts.append(f"Comment {i}: {comment}")

    context_parts.append("\n=== END CONTEXT ===")
    return "\n".join(context_parts)


def _shape_info_to_dict(shape_info):
    """Convert a ShapeInfo dataclass to the dict format expected by format_slide_context."""
    d = {
        'name': shape_info.name,
        'type': shape_info.type_name,
        'id': shape_info.id,
        'position': f"({shape_info.left}, {shape_info.top})",
        'size': f"{shape_info.width} x {shape_info.height}",
    }

    if shape_info.html_text:
        d['html_text'] = shape_info.html_text
    if shape_info.text:
        d['text'] = shape_info.text
    if shape_info.font_info:
        d['font'] = shape_info.font_info
    if shape_info.hyperlinks:
        d['hyperlinks'] = shape_info.hyperlinks

    if shape_info.is_table:
        d['is_table'] = True
        d['table_info'] = shape_info.table_info
        if shape_info.container_type:
            d['container_type'] = shape_info.container_type
        if shape_info.table_content:
            d['table_content'] = shape_info.table_content
        if shape_info.table_content_html:
            d['table_content_html'] = shape_info.table_content_html
            d['table_markdown'] = generate_markdown_table(shape_info.table_content_html)
        if shape_info.table_hyperlinks:
            d['table_hyperlinks'] = shape_info.table_hyperlinks
        if shape_info.table_error:
            d['table_error'] = shape_info.table_error

    if shape_info.chart_info:
        d['chart_info'] = shape_info.chart_info
    if shape_info.chart_title:
        d['chart_title'] = shape_info.chart_title
    if shape_info.chart_data:
        d['chart_data'] = shape_info.chart_data
    if shape_info.chart_error:
        d['chart_error'] = shape_info.chart_error

    return d


def powerpoint_snapshot(slide_number: Optional[int] = None,
                       include_screenshot: bool = True,
                       screenshot_filename: Optional[str] = None) -> dict:
    """
    Capture comprehensive context of a PowerPoint slide with optional screenshot.

    Args:
        slide_number: Slide number to capture (1-based). If None, uses current slide.
        include_screenshot: Whether to save a screenshot with bounding boxes. Default True.
        screenshot_filename: Optional custom filename for screenshot.

    Returns:
        Dictionary with slide context data, screenshot info (if enabled), or error information
    """
    try:
        backend = get_backend()
        backend.connect()

        if not backend.get_presentation_count():
            return {"error": "No PowerPoint presentation is open"}

        # Determine slide to analyze
        if slide_number is None:
            slide_number = backend.get_current_slide_index()
            if slide_number is None:
                slide_number = 1

        slide_count = backend.get_slide_count()
        if slide_number < 1 or slide_number > slide_count:
            return {"error": f"Invalid slide number {slide_number}. Presentation has {slide_count} slides."}

        # Get slide info
        slide_info = backend.get_slide_info(slide_number)

        # Get all shapes
        shapes = backend.get_shapes(slide_number)

        # Convert ShapeInfo objects to dicts for format_slide_context
        shape_dicts = [_shape_info_to_dict(s) for s in shapes]

        # Build slide data dict
        slide_data = {
            'slide_number': slide_number,
            'total_slides': slide_info.total_slides,
            'slide_name': slide_info.name,
            'layout': slide_info.layout_name,
            'object_count': slide_info.shape_count,
            'timestamp': datetime.now().isoformat(),
            'shapes': shape_dicts
        }

        # Get speaker notes
        notes_html = backend.get_speaker_notes(slide_number)
        if notes_html:
            slide_data['notes'] = notes_html
            notes_plain = backend.get_speaker_notes_plain(slide_number)
            if notes_plain:
                slide_data['notes_plain'] = notes_plain

        # Get comments
        comments = backend.get_comments(slide_number)
        if comments:
            slide_data['comments'] = [
                {
                    'text': c.text,
                    'author': c.author,
                    'date': c.date,
                    'position': c.position,
                    **(({'associated_object': c.associated_object} if c.associated_object else {}))
                }
                for c in comments
            ]

        # Format context
        formatted_context = format_slide_context(slide_data)

        # Screenshot functionality
        screenshot_info = {}
        if include_screenshot:
            try:
                slide_width, slide_height = backend.get_slide_dimensions()

                # Export slide to temp file
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                    temp_path = temp_file.name

                backend.export_slide_image(slide_number, temp_path)

                image = Image.open(temp_path)

                # Add bounding box overlays using ShapeInfo objects directly
                annotated_image = add_bounding_box_overlays(image, shapes, slide_width, slide_height)

                output_path = get_output_file(screenshot_filename)
                annotated_image.save(output_path, "PNG")

                os.unlink(temp_path)

                screenshot_info = {
                    "screenshot_saved": True,
                    "screenshot_path": output_path,
                    "image_size": f"{annotated_image.size[0]}x{annotated_image.size[1]}",
                    "objects_annotated": len(shapes),
                    "screenshot_message": f"Screenshot saved with {len(shapes)} object annotations"
                }
            except Exception as screenshot_error:
                screenshot_info = {
                    "screenshot_saved": False,
                    "screenshot_error": f"Failed to take screenshot: {str(screenshot_error)}"
                }

        result = {
            "success": True,
            "slide_number": slide_number,
            "total_slides": slide_info.total_slides,
            "object_count": slide_info.shape_count,
            "context": formatted_context,
            "slide_data": slide_data
        }

        if include_screenshot:
            result.update(screenshot_info)

        return result

    except Exception as e:
        return {"error": f"Failed to capture slide context: {str(e)}"}
