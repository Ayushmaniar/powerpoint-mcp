"""
PowerPoint placeholder population tool for MCP server.
Populates placeholders with text content (with basic HTML formatting) or images.
"""

import os
import re
import tempfile
from typing import Optional

# Initialize matplotlib at module level to avoid first-call delays
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend (must be set before importing pyplot)
import matplotlib.pyplot as plt
import numpy as np

from ..backends import get_backend
from ..backends.types import UnsupportedFeatureError


def detect_content_type(content: str) -> str:
    """Auto-detect if content is an image file or text."""
    image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp', '.svg'}

    if any(content.lower().endswith(ext) for ext in image_extensions):
        return "image"
    else:
        return "text"


def get_powerpoint_char_length(text: str) -> int:
    """Calculate character length as PowerPoint COM would see it (UTF-16 encoding)."""
    return len(text.encode('utf-16-le')) // 2


def get_powerpoint_char_position(full_text: str, python_position: int) -> int:
    """Convert Python string position to PowerPoint COM position (accounting for emojis)."""
    substring = full_text[:python_position]
    return get_powerpoint_char_length(substring) + 1


def process_simple_html(html_text: str):
    """
    Process simplified HTML tags and return plain text with formatting segments.

    Supports:
    - <b>bold</b>, <i>italic</i>, <u>underline</u>
    - <red>text</red>, <blue>text</blue>, <green>text</green>, etc.
    - <ul><li>item</li></ul>, <ol><li>item</li></ol>
    - <latex>equation</latex> for LaTeX equations
    - <para>text</para> for animation grouping

    Returns:
        tuple: (plain_text, format_segments, latex_segments, para_count)
    """
    # Extract <para> segments for counting
    para_segments = []
    para_pattern = r'<para>(.*?)</para>'
    para_matches = list(re.finditer(para_pattern, html_text, re.IGNORECASE | re.DOTALL))
    for match in para_matches:
        para_segments.append({'content': match.group(1)})

    # Process <para> tags
    text = re.sub(r'<para>', '', html_text, flags=re.IGNORECASE)
    text = re.sub(r'</para>', '\r', text, flags=re.IGNORECASE)

    # Process numbered lists FIRST
    ol_pattern = r'<ol>(.*?)</ol>'
    def replace_ol(match):
        ol_content = match.group(1)
        items = re.findall(r'<li>\s*(.*?)\s*</li>', ol_content, re.DOTALL)
        numbered_items = []
        for i, item in enumerate(items, 1):
            formatted_item = item.strip()
            numbered_items.append(f"{i}. {formatted_item}")
        return '\n' + '\n'.join(numbered_items) if numbered_items else ''

    text = re.sub(ol_pattern, replace_ol, text, flags=re.DOTALL)

    # Then process unordered lists
    text = re.sub(r'<ul>\s*', '\n', text)
    text = re.sub(r'</ul>\s*', '', text)
    text = re.sub(r'<li>\s*', '• ', text)
    text = re.sub(r'</li>\s*', '\n', text)

    # Handle basic line breaks
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)

    # Extract LaTeX segments before processing other tags
    latex_segments = []
    latex_pattern = r'<latex>(.*?)</latex>'
    latex_matches = list(re.finditer(latex_pattern, text, re.IGNORECASE | re.DOTALL))

    for match in latex_matches:
        latex_content = match.group(1).strip()
        temp_plain = re.sub(r'<[^>]+>', '', text[:match.start()])
        start_pos = get_powerpoint_char_length(temp_plain) + 1
        length = get_powerpoint_char_length(latex_content)

        latex_segments.append({
            'start': start_pos,
            'length': length,
            'latex': latex_content
        })

    # Define supported tags and their formatting
    format_tags = {
        'b': {'bold': True},
        'i': {'italic': True},
        'u': {'underline': True},
        'red': {'color': 'red'},
        'blue': {'color': 'blue'},
        'green': {'color': 'green'},
        'orange': {'color': 'orange'},
        'purple': {'color': 'purple'},
        'yellow': {'color': 'yellow'},
        'black': {'color': 'black'},
        'white': {'color': 'white'}
    }

    format_segments = []
    plain_text = text

    for tag_name, formatting in format_tags.items():
        tag_pattern = f'<{tag_name}>(.*?)</{tag_name}>'
        matches = list(re.finditer(tag_pattern, plain_text, re.IGNORECASE | re.DOTALL))

        for match in matches:
            tag_content_with_tags = match.group(1)
            tag_content_plain = re.sub(r'<[^>]+>', '', tag_content_with_tags)
            temp_plain = re.sub(r'<[^>]+>', '', plain_text[:match.start()])

            start_pos = get_powerpoint_char_length(temp_plain) + 1
            content_length = get_powerpoint_char_length(tag_content_plain)

            if tag_content_plain:
                format_segments.append({
                    'start': start_pos,
                    'length': content_length,
                    'formatting': formatting
                })

    plain_text = re.sub(r'<[^>]+>', '', plain_text)
    plain_text = plain_text.strip('\n\r')
    para_count = len(para_segments)

    return plain_text, format_segments, latex_segments, para_count


def adjust_formatting_positions_after_latex(format_segments: list, latex_segments: list) -> list:
    """Adjust formatting segment positions after LaTeX equations have been converted."""
    if not latex_segments or not format_segments:
        return format_segments

    latex_shifts = []
    for latex_seg in latex_segments:
        old_start = latex_seg['start']
        old_length = latex_seg['length']
        old_end = old_start + old_length - 1
        new_length = latex_seg.get('actual_new_length', old_length)
        shift = new_length - old_length
        latex_shifts.append({
            'original_start': old_start,
            'original_end': old_end,
            'shift': shift
        })

    adjusted_formats = []
    for fmt_seg in format_segments:
        fmt_start = fmt_seg['start']
        fmt_length = fmt_seg['length']

        total_shift = 0
        for latex_shift in latex_shifts:
            if latex_shift['original_end'] < fmt_start:
                total_shift += latex_shift['shift']

        new_start = fmt_start + total_shift
        adjusted_formats.append({
            'start': new_start,
            'length': fmt_length,
            'formatting': fmt_seg['formatting']
        })

    return adjusted_formats


def render_matplotlib_plot(matplotlib_code: str) -> str:
    """Execute matplotlib code and return the path to the generated image."""
    try:
        temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        temp_path = temp_file.name
        temp_file.close()

        cleaned_code = re.sub(r'plt\.savefig\s*\([^)]*\)', '', matplotlib_code)
        cleaned_code = re.sub(r'plt\.close\s*\([^)]*\)', '', cleaned_code)

        exec_namespace = {
            'plt': plt,
            'matplotlib': matplotlib,
            'np': np,
            'numpy': np,
            '__builtins__': __builtins__
        }

        exec(cleaned_code, exec_namespace)

        plt.savefig(temp_path, dpi=300, bbox_inches='tight')
        plt.close('all')

        return temp_path

    except Exception as e:
        raise Exception(f"Failed to render matplotlib plot: {str(e)}")


def populate_text_placeholder(backend, slide_number: int, shape_name: str, content: str):
    """Populate a placeholder with text content, HTML formatting, and LaTeX equations."""
    has_html = bool(re.search(r'<[^>]+>', content))

    if has_html:
        plain_text, format_segments, latex_segments, para_count = process_simple_html(content)

        # Set the text first
        backend.set_text(slide_number, shape_name, plain_text)

        # Apply LaTeX equation conversion FIRST (before other formatting)
        latex_warning = None
        if latex_segments:
            try:
                backend.convert_latex_to_equation(slide_number, shape_name, latex_segments)
            except UnsupportedFeatureError as e:
                latex_warning = str(e)

        # Clear bullets after LaTeX conversion
        backend.clear_bullets(slide_number, shape_name)

        # Adjust formatting positions after LaTeX conversion
        if latex_segments and format_segments:
            format_segments = adjust_formatting_positions_after_latex(format_segments, latex_segments)

        # Apply other formatting (bold, italic, colors)
        if format_segments:
            backend.apply_character_formatting(slide_number, shape_name, format_segments)

        result = {
            "success": True,
            "content_type": "formatted_text",
            "html_input": content,
            "plain_text": plain_text,
            "format_segments_applied": len(format_segments),
            "latex_equations_applied": len(latex_segments),
            "para_segments_detected": para_count
        }

        if latex_warning:
            result["latex_warning"] = latex_warning

        return result
    else:
        # Simple plain text
        backend.clear_bullets(slide_number, shape_name)
        backend.set_text(slide_number, shape_name, content)
        backend.clear_bullets(slide_number, shape_name)

        return {
            "success": True,
            "content_type": "plain_text",
            "text_set": content
        }


def powerpoint_populate_placeholder(
    placeholder_name: str,
    content: str,
    content_type: str = "auto",
    slide_number: Optional[int] = None
) -> dict:
    """
    Populate a PowerPoint placeholder with content.

    Args:
        placeholder_name: Name of the placeholder (e.g., "Title 1", "Subtitle 2")
        content: Text content (with optional HTML), image file path, or matplotlib code
        content_type: "text", "image", "plot", or "auto" (auto-detect based on content)
        slide_number: Target slide number (1-based). If None, uses current active slide

    Returns:
        Dictionary with success status and operation details
    """
    try:
        backend = get_backend()
        backend.connect()

        if not backend.get_presentation_count():
            return {"error": "No PowerPoint presentation is open. Please open a presentation first."}

        # Determine target slide
        if slide_number is None:
            slide_number = backend.get_current_slide_index()
            if slide_number is None:
                slide_number = 1

        total_slides = backend.get_slide_count()
        if slide_number < 1 or slide_number > total_slides:
            return {"error": f"Invalid slide number {slide_number}. Must be between 1 and {total_slides}."}

        # Verify shape exists by getting shapes
        shapes = backend.get_shapes(slide_number)
        shape_found = any(s.name.lower() == placeholder_name.lower() for s in shapes)
        if not shape_found:
            available_names = [s.name for s in shapes]
            return {
                "error": f"Placeholder '{placeholder_name}' not found on slide {slide_number}",
                "available_placeholders": available_names
            }

        # Auto-detect content type if needed
        if content_type == "auto":
            content_type = detect_content_type(content)

        # Handle matplotlib plot rendering
        temp_plot_path = None
        matplotlib_code_for_alt_text = None
        if content_type == "plot":
            try:
                temp_plot_path = render_matplotlib_plot(content)
                actual_content = temp_plot_path
                actual_content_type = "image"
                matplotlib_code_for_alt_text = content
            except Exception as e:
                return {"error": f"Failed to render matplotlib plot: {str(e)}"}
        else:
            actual_content = content
            actual_content_type = content_type

        # Populate based on content type
        if actual_content_type == "text":
            result = populate_text_placeholder(backend, slide_number, placeholder_name, actual_content)
        elif actual_content_type == "image":
            if not os.path.exists(actual_content):
                return {"error": f"Image file not found: {actual_content}"}
            result = backend.insert_image(slide_number, placeholder_name, actual_content, matplotlib_code_for_alt_text)
        else:
            return {"error": f"Unsupported content type '{content_type}'. Use 'text', 'image', 'plot', or 'auto'."}

        # Clean up temporary plot file if created
        if temp_plot_path and os.path.exists(temp_plot_path):
            try:
                os.unlink(temp_plot_path)
            except:
                pass

        # Add common success information
        if result.get("success"):
            result.update({
                "placeholder_name": placeholder_name,
                "slide_number": slide_number,
                "total_slides": total_slides,
                "detected_content_type": content_type,
                "was_matplotlib_plot": content_type == "plot"
            })

        return result

    except Exception as e:
        return {"error": f"Failed to populate placeholder '{placeholder_name}': {str(e)}"}


def generate_mcp_response(result):
    """Generate the MCP tool response for the LLM."""
    if not result.get('success'):
        return f"Failed to populate placeholder: {result.get('error')}"

    response_lines = [
        f"Populated placeholder '{result['placeholder_name']}' on slide {result['slide_number']}"
    ]

    if result['content_type'] == 'formatted_text':
        response_lines.append(f"Content: HTML-formatted text with {result['format_segments_applied']} formatting segments")
        response_lines.append(f"Plain text: '{result['plain_text']}'")
        if result.get('latex_warning'):
            response_lines.append(f"Warning: {result['latex_warning']}")
    elif result['content_type'] == 'plain_text':
        response_lines.append(f"Content: Plain text '{result['text_set']}'")
    elif result['content_type'] == 'image':
        if result.get('was_matplotlib_plot'):
            response_lines.append(f"Content: Matplotlib plot (rendered to image)")
        else:
            response_lines.append(f"Content: Image from '{result['image_path']}'")
        response_lines.append(f"Dimensions: {result['dimensions']}")
        if result.get('placeholder_renamed_from'):
            response_lines.append(
                f"Note: placeholder '{result['placeholder_renamed_from']}' is now '{result['new_shape_name']}'"
            )

    response_lines.append(f"Content type: {result['detected_content_type']}")

    return "\n".join(response_lines)
