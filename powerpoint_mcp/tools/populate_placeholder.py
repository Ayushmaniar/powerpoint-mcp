"""
PowerPoint placeholder population tool for MCP server.
Populates placeholders with text content (with basic HTML formatting) or images.
"""

import os
import re
import tempfile
import win32com.client
from typing import Optional


def detect_content_type(content: str) -> str:
    """Auto-detect if content is an image file or text."""
    image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp', '.svg'}

    if any(content.lower().endswith(ext) for ext in image_extensions):
        return "image"
    else:
        return "text"


def find_shape_by_name(slide, shape_name: str):
    """Find a shape on the slide by its name (case-insensitive)."""
    for shape in slide.Shapes:
        if shape.Name.lower() == shape_name.lower():
            return shape
    return None


def get_powerpoint_char_length(text: str) -> int:
    """Calculate character length as PowerPoint COM would see it (UTF-16 encoding)."""
    return len(text.encode('utf-16-le')) // 2


def get_powerpoint_char_position(full_text: str, python_position: int) -> int:
    """Convert Python string position to PowerPoint COM position (accounting for emojis)."""
    # Get the substring up to the Python position
    substring = full_text[:python_position]
    # Return PowerPoint character count + 1 (PowerPoint uses 1-based indexing)
    return get_powerpoint_char_length(substring) + 1


def process_simple_html(html_text: str):
    """
    Process simplified HTML tags and return plain text with formatting segments.

    Supports:
    - <b>bold</b>, <i>italic</i>, <u>underline</u>
    - <red>text</red>, <blue>text</blue>, <green>text</green>, etc.
    - <ul><li>item</li></ul>, <ol><li>item</li></ol>
    - <latex>equation</latex> for LaTeX equations

    Returns:
        tuple: (plain_text, format_segments, latex_segments)
    """
    # Handle lists first (convert to plain text with bullets/numbers)
    text = html_text

    # IMPORTANT: Process numbered lists FIRST before bullet lists
    # Convert ordered lists to numbered points
    ol_pattern = r'<ol>(.*?)</ol>'
    def replace_ol(match):
        ol_content = match.group(1)
        items = re.findall(r'<li>\s*(.*?)\s*</li>', ol_content, re.DOTALL)

        numbered_items = []
        for i, item in enumerate(items, 1):
            # Keep formatting tags intact, just add the number prefix
            formatted_item = item.strip()
            numbered_items.append(f"{i}. {formatted_item}")

        return '\n' + '\n'.join(numbered_items) + '\n' if numbered_items else ''

    text = re.sub(ol_pattern, replace_ol, text, flags=re.DOTALL)

    # THEN process unordered lists (bullet points)
    text = re.sub(r'<ul>\s*', '\n', text)  # Add newline before starting list
    text = re.sub(r'</ul>\s*', '\n', text)
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

        # Calculate position in text without any HTML tags
        temp_plain = re.sub(r'<[^>]+>', '', text[:match.start()])
        start_pos = get_powerpoint_char_length(temp_plain) + 1  # 1-based for PowerPoint
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

    # Better approach: handle nested tags properly
    format_segments = []
    plain_text = text

    # Process each format tag, allowing nested content
    for tag_name, formatting in format_tags.items():
        # Updated pattern to handle nested tags: match everything until closing tag
        tag_pattern = f'<{tag_name}>(.*?)</{tag_name}>'
        matches = list(re.finditer(tag_pattern, plain_text, re.IGNORECASE | re.DOTALL))

        for match in matches:
            tag_content_with_tags = match.group(1)  # May contain nested tags

            # Extract plain text from the tagged content
            tag_content_plain = re.sub(r'<[^>]+>', '', tag_content_with_tags)

            # Find where this content will be in the final plain text
            temp_plain = re.sub(r'<[^>]+>', '', plain_text[:match.start()])

            # Use PowerPoint-compatible character counting
            start_pos = get_powerpoint_char_length(temp_plain) + 1  # 1-based for PowerPoint
            content_length = get_powerpoint_char_length(tag_content_plain)

            if tag_content_plain:  # Only add if there's actual text content
                format_segments.append({
                    'start': start_pos,
                    'length': content_length,
                    'formatting': formatting
                })

    # Remove all HTML tags to get final plain text
    plain_text = re.sub(r'<[^>]+>', '', plain_text)

    return plain_text, format_segments, latex_segments


def clear_placeholder_bullets(text_range):
    """Remove default bullet formatting from a placeholder's TextRange."""
    try:
        paragraph_format = text_range.ParagraphFormat
        bullet_format = paragraph_format.Bullet
        bullet_format.Visible = 0
        bullet_format.Type = 0
    except Exception:
        pass

    try:
        paragraphs = text_range.Paragraphs()
        count = paragraphs.Count
        for idx in range(1, count + 1):
            para_range = paragraphs(idx)
            try:
                para_bullet = para_range.ParagraphFormat.Bullet
                para_bullet.Visible = 0
                para_bullet.Type = 0
            except Exception:
                continue
    except Exception:
        pass


def apply_formatting_segments(text_range, format_segments: list):
    """
    Apply formatting segments to a PowerPoint TextRange WITHOUT resetting the text.
    This is used after LaTeX equations have been converted.
    """
    # Apply each formatting segment
    for segment in format_segments:
        try:
            start_pos = segment['start']
            length = segment['length']
            formatting = segment['formatting']

            # Get the character range for this segment
            char_range = text_range.Characters(start_pos, length)

            # Apply formatting
            if formatting.get('bold'):
                char_range.Font.Bold = True
            if formatting.get('italic'):
                char_range.Font.Italic = True
            if formatting.get('underline'):
                char_range.Font.Underline = True

            # Apply color formatting
            if formatting.get('color'):
                color_name = formatting['color'].lower()
                color_map = {
                    'red': 255,                    # RGB(255, 0, 0)
                    'blue': 16711680,             # RGB(0, 0, 255)
                    'green': 65280,               # RGB(0, 255, 0)
                    'orange': 33023,              # RGB(255, 127, 0)
                    'purple': 8388736,            # RGB(128, 0, 128)
                    'yellow': 65535,              # RGB(255, 255, 0)
                    'black': 0,                   # RGB(0, 0, 0)
                    'white': 16777215             # RGB(255, 255, 255)
                }

                if color_name in color_map:
                    char_range.Font.Color.RGB = color_map[color_name]

        except Exception:
            # If formatting fails for this segment, continue with others
            pass


def adjust_formatting_positions_after_latex(format_segments: list, latex_segments: list, text_range) -> list:
    """
    Adjust formatting segment positions after LaTeX equations have been converted.
    
    When LaTeX is converted to equation objects, the character positions shift because
    equation objects have different lengths than the original LaTeX text.
    
    Args:
        format_segments: List of formatting segments with 'start' and 'length' keys
        latex_segments: List of LaTeX segments that were converted (with 'start' and 'length' keys)
        text_range: The TextRange after LaTeX conversion (to measure new lengths)
    
    Returns:
        List of adjusted formatting segments
    """
    if not latex_segments or not format_segments:
        return format_segments
    
    # Calculate the shift for each latex segment (new_length - old_length)
    # Note: LaTeX segments are processed in reverse order during conversion,
    # but we need to calculate shifts based on original positions
    latex_shifts = []
    
    for latex_seg in latex_segments:
        old_start = latex_seg['start']
        old_length = latex_seg['length']
        old_end = old_start + old_length - 1  # Last character position of the LaTeX segment
        
        # Get the actual new length measured during conversion
        # If not measured, default to assuming it stayed the same
        new_length = latex_seg.get('actual_new_length', old_length)
        shift = new_length - old_length
        
        latex_shifts.append({
            'original_start': old_start,
            'original_end': old_end,
            'shift': shift
        })
    
    # Adjust each formatting segment based on latex shifts
    adjusted_formats = []
    for fmt_seg in format_segments:
        fmt_start = fmt_seg['start']
        fmt_length = fmt_seg['length']
        
        # Calculate cumulative shift up to this format segment's start
        total_shift = 0
        for latex_shift in latex_shifts:
            # If the latex segment ENDS before this format segment starts, apply the shift
            if latex_shift['original_end'] < fmt_start:
                total_shift += latex_shift['shift']
        
        new_start = fmt_start + total_shift
        
        adjusted_formats.append({
            'start': new_start,
            'length': fmt_length,
            'formatting': fmt_seg['formatting']
        })
    
    return adjusted_formats


def apply_latex_equations(ppt_app, text_range, latex_segments: list):
    """
    Convert LaTeX segments in a TextRange to PowerPoint equations.

    Uses PowerPoint's built-in LaTeX equation conversion via ExecuteMso.
    Based on the VBA pattern: set text, select equation portion, ExecuteMso.

    Args:
        ppt_app: PowerPoint Application COM object
        text_range: TextRange containing the text with LaTeX content
        latex_segments: List of dicts with 'start', 'length', and 'latex' keys
    """
    if not latex_segments:
        return

    try:
        import time

        # Ensure PowerPoint window is active and the containing slide is in view
        try:
            ppt_app.Activate()
            
            # Get the slide that contains this text range
            slide = text_range.Parent.Parent.Parent  # TextFrame -> Shape -> Slide
            slide_index = slide.SlideIndex
            
            # Switch to the slide containing this shape
            if hasattr(ppt_app.ActiveWindow, 'View'):
                view = ppt_app.ActiveWindow.View
                # Switch to Normal view first
                ppt_app.ActiveWindow.ViewType = 1  # ppViewNormal
                time.sleep(0.1)
                # Navigate to the slide
                if hasattr(view, 'GotoSlide'):
                    view.GotoSlide(slide_index)
                    time.sleep(0.1)
            
            ppt_app.ActiveWindow.Activate()
        except:
            pass

        # Process LaTeX segments in reverse order to avoid position shifting
        for segment in reversed(latex_segments):
            try:
                start_pos = segment['start']
                length = segment['length']

                # Get the LaTeX text range (it's already set as plain text)
                char_range = text_range.Characters(start_pos, length)

                # Get the latex content and measure length before conversion
                latex_content = segment['latex']
                old_eq_length = length
                text_length_before = get_powerpoint_char_length(text_range.Text)

                # Convert LaTeX to equation using ExecuteMso commands
                try:
                    # Select the character range containing the LaTeX
                    char_range.Select()
                    time.sleep(0.05)

                    # Execute PowerPoint's equation commands
                    ppt_app.CommandBars.ExecuteMso("InsertBuildingBlocksEquationsGallery")
                    time.sleep(0.1)
                    ppt_app.CommandBars.ExecuteMso("EquationLaTeXToMath")
                    time.sleep(0.05)

                    # Measure the actual new length by comparing total text length (using PowerPoint's UTF-16 counting)
                    try:
                        time.sleep(0.05)  # Give time for the conversion to complete
                        text_length_after = get_powerpoint_char_length(text_range.Text)

                        # Calculate new equation length: new_total = old_total - old_eq_length + new_eq_length
                        new_eq_length = text_length_after - text_length_before + old_eq_length

                        # Store the actual new length for position adjustment
                        segment['actual_new_length'] = new_eq_length
                    except:
                        # If measurement fails, assume no change
                        segment['actual_new_length'] = old_eq_length

                except:
                    # If conversion fails, leave the LaTeX text as-is
                    segment['actual_new_length'] = old_eq_length
                    continue

            except Exception:
                # Continue with other segments even if one fails
                pass

        # After processing ALL LaTeX segments, switch back to normal view
        # This prevents PowerPoint from staying in the Equation tab
        try:
            ppt_app.ActiveWindow.ViewType = 9  # ppViewNormal = 9 (official Microsoft constant)
        except:
            pass

    except Exception:
        # Silently fail if LaTeX processing encounters an error
        pass


def populate_text_placeholder(ppt_app, shape, content: str):
    """Populate a placeholder with text content, HTML formatting, and LaTeX equations."""
    if not hasattr(shape, 'TextFrame'):
        return {"error": f"Shape '{shape.Name}' cannot hold text (no TextFrame)"}

    # Check if content has HTML tags (including LaTeX)
    has_html = bool(re.search(r'<[^>]+>', content))

    if has_html:
        # Process HTML formatting and LaTeX tags
        plain_text, format_segments, latex_segments = process_simple_html(content)
        
        # Set the text first WITHOUT clearing bullets yet
        text_range = shape.TextFrame.TextRange
        text_range.Text = plain_text
        
        # Apply LaTeX equation conversion FIRST (before other formatting or bullet clearing)
        # This is critical because PowerPoint equation commands don't work on pre-formatted text
        if latex_segments:
            apply_latex_equations(ppt_app, shape.TextFrame.TextRange, latex_segments)
        
        # NOW clear bullets after LaTeX conversion
        clear_placeholder_bullets(shape.TextFrame.TextRange)
        
        # Adjust formatting positions after LaTeX conversion
        # (equations change character positions)
        if latex_segments and format_segments:
            format_segments = adjust_formatting_positions_after_latex(
                format_segments, latex_segments, shape.TextFrame.TextRange
            )
        
        # Then apply other formatting (bold, italic, colors) to non-equation text
        if format_segments:
            apply_formatting_segments(shape.TextFrame.TextRange, format_segments)

        return {
            "success": True,
            "content_type": "formatted_text",
            "html_input": content,
            "plain_text": plain_text,
            "format_segments_applied": len(format_segments),
            "latex_equations_applied": len(latex_segments)
        }
    else:
        # Simple plain text
        text_range = shape.TextFrame.TextRange
        clear_placeholder_bullets(text_range)
        text_range.Text = content
        clear_placeholder_bullets(text_range)

        return {
            "success": True,
            "content_type": "plain_text",
            "text_set": content
        }


def render_matplotlib_plot(matplotlib_code: str) -> str:
    """
    Execute matplotlib code and return the path to the generated image.

    Args:
        matplotlib_code: Python code that generates a matplotlib plot

    Returns:
        Path to the generated PNG image file
    """
    try:
        import matplotlib
        matplotlib.use('Agg')  # Use non-interactive backend
        import matplotlib.pyplot as plt

        # Create a temporary file for the plot
        temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        temp_path = temp_file.name
        temp_file.close()

        # Create a clean namespace for code execution
        exec_namespace = {
            'plt': plt,
            'matplotlib': matplotlib,
            '__builtins__': __builtins__
        }

        # Execute the matplotlib code
        exec(matplotlib_code, exec_namespace)

        # Save the figure
        plt.savefig(temp_path, dpi=300, bbox_inches='tight')
        plt.close('all')

        return temp_path

    except Exception as e:
        raise Exception(f"Failed to render matplotlib plot: {str(e)}")


def populate_image_placeholder(shape, image_path: str):
    """Populate a placeholder with an image."""
    if not os.path.exists(image_path):
        return {"error": f"Image file not found: {image_path}"}

    try:
        # Get the slide that contains this shape
        slide = shape.Parent

        # Get shape position and size for replacement
        placeholder_left = shape.Left
        placeholder_top = shape.Top
        placeholder_width = shape.Width
        placeholder_height = shape.Height

        placeholder_name = getattr(shape, "Name", None)

        # Delete the placeholder shape
        shape.Delete()

        # Add the image first with default size to get its natural dimensions
        temp_shape = slide.Shapes.AddPicture(
            FileName=image_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=-1,  # Use original width
            Height=-1  # Use original height
        )

        # Get the natural image dimensions
        image_width = temp_shape.Width
        image_height = temp_shape.Height

        # Calculate aspect ratios
        placeholder_aspect = placeholder_width / placeholder_height
        image_aspect = image_width / image_height

        # Determine scaled dimensions while maintaining aspect ratio
        if image_aspect > placeholder_aspect:
            # Image is wider - width will hit the limit first
            final_width = placeholder_width
            final_height = placeholder_width / image_aspect
        else:
            # Image is taller - height will hit the limit first
            final_height = placeholder_height
            final_width = placeholder_height * image_aspect

        # Center the image within the placeholder bounds
        final_left = placeholder_left + (placeholder_width - final_width) / 2
        final_top = placeholder_top + (placeholder_height - final_height) / 2

        # Update the temporary shape with final dimensions and position
        temp_shape.Left = final_left
        temp_shape.Top = final_top
        temp_shape.Width = final_width
        temp_shape.Height = final_height

        new_shape = temp_shape

        # Attempt to preserve original placeholder name for easier follow-up actions
        original_name = placeholder_name if isinstance(placeholder_name, str) else None
        if original_name:
            try:
                new_shape.Name = original_name
            except Exception:
                pass

        result = {
            "success": True,
            "content_type": "image",
            "image_path": image_path,
            "new_shape_id": new_shape.Id,
            "new_shape_name": new_shape.Name,
            "dimensions": f"{final_width} x {final_height}"
        }

        if (
            original_name
            and isinstance(new_shape.Name, str)
            and original_name.lower() != new_shape.Name.lower()
        ):
            result["placeholder_renamed_from"] = original_name

        return result

    except Exception as e:
        return {"error": f"Failed to insert image: {str(e)}"}


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

    Content Types:
        - "text": Plain text or HTML-formatted text
        - "image": Path to an image file
        - "plot": Matplotlib code that generates a plot (rendered to image)
        - "auto": Auto-detect based on content

    Returns:
        Dictionary with success status and operation details
    """
    try:
        # Connect to PowerPoint (reusing pattern from existing tools)
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        if not ppt_app.Presentations.Count:
            return {"error": "No PowerPoint presentation is open. Please open a presentation first."}

        active_presentation = ppt_app.ActivePresentation

        # Determine target slide
        if slide_number is not None:
            if slide_number < 1 or slide_number > active_presentation.Slides.Count:
                return {"error": f"Invalid slide number {slide_number}. Must be between 1 and {active_presentation.Slides.Count}."}
            target_slide = active_presentation.Slides(slide_number)
        else:
            # Use current active slide (reusing pattern from snapshot.py)
            try:
                active_window = ppt_app.ActiveWindow
                if hasattr(active_window, 'View') and hasattr(active_window.View, 'Slide'):
                    target_slide = active_window.View.Slide
                else:
                    target_slide = active_presentation.Slides(1)  # Fallback to first slide
            except:
                target_slide = active_presentation.Slides(1)

        # Find the placeholder by name
        target_shape = find_shape_by_name(target_slide, placeholder_name)
        if not target_shape:
            # Provide helpful error with available placeholder names
            available_names = [shape.Name for shape in target_slide.Shapes]
            return {
                "error": f"Placeholder '{placeholder_name}' not found on slide {target_slide.SlideIndex}",
                "available_placeholders": available_names
            }

        # Auto-detect content type if needed
        if content_type == "auto":
            content_type = detect_content_type(content)

        # Handle matplotlib plot rendering
        temp_plot_path = None
        if content_type == "plot":
            try:
                temp_plot_path = render_matplotlib_plot(content)
                # Treat the rendered plot as an image
                actual_content = temp_plot_path
                actual_content_type = "image"
            except Exception as e:
                return {"error": f"Failed to render matplotlib plot: {str(e)}"}
        else:
            actual_content = content
            actual_content_type = content_type

        # Populate based on content type
        if actual_content_type == "text":
            result = populate_text_placeholder(ppt_app, target_shape, actual_content)
        elif actual_content_type == "image":
            result = populate_image_placeholder(target_shape, actual_content)
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
                "slide_number": target_slide.SlideIndex,
                "total_slides": active_presentation.Slides.Count,
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

    # Create clean response for LLM
    response_lines = [
        f"✅ Populated placeholder '{result['placeholder_name']}' on slide {result['slide_number']}"
    ]

    if result['content_type'] == 'formatted_text':
        response_lines.append(f"Content: HTML-formatted text with {result['format_segments_applied']} formatting segments")
        response_lines.append(f"Plain text: '{result['plain_text']}'")
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