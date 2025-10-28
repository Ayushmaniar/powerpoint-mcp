"""
PowerPoint slide context snapshot tool.

Provides comprehensive slide content analysis similar to Playwright's browser_snapshot,
but returns detailed slide context including all shapes, text, tables, and formatting.
Now includes screenshot functionality with object bounding box overlays.
"""

import os
import tempfile
import win32com.client
from typing import Optional
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


def get_current_slide_index(ppt_app):
    """Get the index of the currently selected/active slide."""
    try:
        if not ppt_app:
            return None

        active_window = ppt_app.ActiveWindow

        # Method 1: Try to get from the current view (most reliable)
        try:
            if hasattr(active_window, 'View') and hasattr(active_window.View, 'Slide'):
                slide_index = active_window.View.Slide.SlideIndex
                if slide_index > 0:
                    return slide_index
        except:
            pass

        # Method 2: Try to get from selection (works in slide sorter view)
        try:
            if (hasattr(active_window, 'Selection') and
                hasattr(active_window.Selection, 'SlideRange') and
                active_window.Selection.SlideRange.Count > 0):
                return active_window.Selection.SlideRange[0].SlideIndex
        except:
            pass

        # Method 3: Try SlideShowWindow if in slideshow mode
        try:
            if hasattr(ppt_app, 'SlideShowWindows') and ppt_app.SlideShowWindows.Count > 0:
                slide_show = ppt_app.SlideShowWindows(1)
                if hasattr(slide_show, 'View') and hasattr(slide_show.View, 'CurrentShowPosition'):
                    return slide_show.View.CurrentShowPosition
        except:
            pass

        # Fallback: return 1 if presentation exists
        presentation = ppt_app.ActivePresentation
        if presentation and presentation.Slides.Count > 0:
            return 1

        return None

    except Exception:
        return 1  # Safe fallback


def convert_text_to_html(text_range):
    """Convert PowerPoint text formatting to HTML using the lightning-fast runs approach."""
    try:
        if not hasattr(text_range, 'Runs') or not text_range.Text:
            return text_range.Text if hasattr(text_range, 'Text') else ""

        html_parts = []
        runs = text_range.Runs()
        if not runs:
            return text_range.Text

        for run in runs:
            run_font = run.Font
            run_text = run.Text

            if not run_text.strip():
                html_parts.append(run_text)
                continue

            open_tags = []
            close_tags = []

            # Check formatting properties
            if run_font.Bold:
                open_tags.append('<b>')
                close_tags.insert(0, '</b>')
            if run_font.Italic:
                open_tags.append('<i>')
                close_tags.insert(0, '</i>')
            if run_font.Underline:
                open_tags.append('<u>')
                close_tags.insert(0, '</u>')

            try:
                if run_font.Strikethrough:
                    open_tags.append('<s>')
                    close_tags.insert(0, '</s>')
            except:
                pass

            # Handle color
            try:
                color_bgr = run_font.Color.RGB
                r = color_bgr & 0xFF
                g = (color_bgr >> 8) & 0xFF
                b = (color_bgr >> 16) & 0xFF
                hex_color = f"#{r:02x}{g:02x}{b:02x}"

                if hex_color.lower() != "#000000":
                    open_tags.append(f'<span style="color: {hex_color}">')
                    close_tags.insert(0, '</span>')
            except:
                pass

            # Handle line breaks and escape HTML
            escaped_text = run_text.replace('\r\n', '<br>').replace('\r', '<br>').replace('\n', '<br>')
            escaped_text = escaped_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            escaped_text = escaped_text.replace('&lt;br&gt;', '<br>')

            formatted_text = ''.join(open_tags) + escaped_text + ''.join(close_tags)
            html_parts.append(formatted_text)

        return ''.join(html_parts)

    except Exception:
        return text_range.Text if hasattr(text_range, 'Text') else ""


def get_shape_type_name(shape_type):
    """Convert shape type number to readable name."""
    shape_types = {
        1: "AutoShape", 2: "Callout", 3: "Chart", 4: "Comment", 5: "Freeform",
        6: "Group", 7: "Embedded OLE Object", 8: "Line", 9: "Linked OLE Object",
        10: "Linked Picture", 11: "Media", 12: "OLE Control", 13: "Picture",
        14: "Placeholder", 15: "Text Effect", 16: "Title", 17: "Picture",
        18: "Script Anchor", 19: "Table", 20: "Canvas", 21: "Diagram",
        22: "Ink", 23: "Ink Comment", 24: "Smart Art", 25: "Web Video"
    }
    return shape_types.get(shape_type, f"Unknown Type ({shape_type})")


def generate_markdown_table(table_cells_html):
    """Generate a markdown table with HTML formatted cell content."""
    if not table_cells_html or len(table_cells_html) == 0:
        return "Empty table"

    try:
        # Calculate column widths for better formatting
        col_count = len(table_cells_html[0]) if table_cells_html else 0
        if col_count == 0:
            return "Empty table"

        # Build markdown table
        markdown_lines = []

        # Header row (first row)
        if len(table_cells_html) > 0:
            header_row = "| " + " | ".join(table_cells_html[0]) + " |"
            markdown_lines.append(header_row)

            # Separator row
            separator = "| " + " | ".join(["---"] * col_count) + " |"
            markdown_lines.append(separator)

            # Data rows (remaining rows)
            for row in table_cells_html[1:]:
                # Ensure row has the right number of columns
                padded_row = row + [""] * (col_count - len(row))
                data_row = "| " + " | ".join(padded_row[:col_count]) + " |"
                markdown_lines.append(data_row)

        return "\n".join(markdown_lines)

    except Exception as e:
        return f"Error generating markdown table: {e}"


def extract_chart_data(chart):
    """Extract comprehensive chart data including series, labels, and data points."""
    try:
        chart_data = {
            'chart_type': chart.ChartType,
            'has_title': chart.HasTitle,
            'title': chart.ChartTitle.Text if chart.HasTitle else "No title"
        }

        # Extract data series
        try:
            series_data = []
            for i in range(1, chart.SeriesCollection().Count + 1):
                series = chart.SeriesCollection(i)
                series_info = {
                    'name': series.Name,
                    'values': [],
                    'categories': []
                }

                # Get series name more reliably
                try:
                    if hasattr(series, 'Name') and series.Name:
                        series_info['name'] = str(series.Name)
                    elif hasattr(series, 'SeriesTitle') and series.SeriesTitle:
                        series_info['name'] = str(series.SeriesTitle)
                    else:
                        series_info['name'] = f"Series {i}"
                except:
                    series_info['name'] = f"Series {i}"

                # Get data values
                try:
                    values = series.Values
                    if values:
                        # Convert COM array to Python list
                        series_info['values'] = [float(v) if v is not None else 0 for v in values]
                except:
                    series_info['values'] = ["Error reading values"]

                # Get category labels using multiple methods
                if i == 1:  # Only get categories once from first series
                    categories = []

                    # Method 1: Try chart data source
                    try:
                        chart_data_source = chart.ChartData
                        if hasattr(chart_data_source, 'Workbook'):
                            workbook = chart_data_source.Workbook
                            worksheet = workbook.Worksheets(1)
                            # Try to read category labels from first column
                            for row in range(2, 10):  # Check first few rows
                                try:
                                    cell_value = worksheet.Cells(row, 1).Value
                                    if cell_value and str(cell_value).strip():
                                        categories.append(str(cell_value))
                                except:
                                    break
                    except:
                        pass

                    # Method 2: Try series XValues if Method 1 failed
                    if not categories:
                        try:
                            if hasattr(series, 'XValues') and series.XValues:
                                categories = [str(c) if c is not None else "" for c in series.XValues]
                        except:
                            pass

                    # Method 3: Try category axis properties
                    if not categories:
                        try:
                            if hasattr(chart, 'Axes'):
                                category_axis = chart.Axes(1)
                                if hasattr(category_axis, 'CategoryNames'):
                                    categories = [str(c) for c in category_axis.CategoryNames]
                        except:
                            pass

                    # Store categories in chart_data (not per series)
                    if categories:
                        chart_data['categories'] = categories

                series_data.append(series_info)

            chart_data['series'] = series_data
        except:
            chart_data['series'] = "Error reading series data"

        # Extract axis information
        try:
            axes_info = {}

            # Category axis (X-axis)
            if hasattr(chart, 'Axes'):
                try:
                    category_axis = chart.Axes(1)  # xlCategory = 1
                    axes_info['category_axis'] = {
                        'title': category_axis.AxisTitle.Text if category_axis.HasTitle else "No title",
                        'has_title': category_axis.HasTitle
                    }
                except:
                    axes_info['category_axis'] = "Error reading category axis"

                # Value axis (Y-axis)
                try:
                    value_axis = chart.Axes(2)  # xlValue = 2
                    axes_info['value_axis'] = {
                        'title': value_axis.AxisTitle.Text if value_axis.HasTitle else "No title",
                        'has_title': value_axis.HasTitle,
                        'minimum': getattr(value_axis, 'MinimumScale', 'Auto'),
                        'maximum': getattr(value_axis, 'MaximumScale', 'Auto')
                    }
                except:
                    axes_info['value_axis'] = "Error reading value axis"

            chart_data['axes'] = axes_info
        except:
            chart_data['axes'] = "Error reading axes"

        # Extract legend information
        try:
            if chart.HasLegend:
                legend = chart.Legend
                chart_data['legend'] = {
                    'has_legend': True,
                    'position': getattr(legend, 'Position', 'Unknown')
                }
            else:
                chart_data['legend'] = {'has_legend': False}
        except:
            chart_data['legend'] = "Error reading legend"

        return chart_data

    except Exception as e:
        return f"Error extracting chart data: {str(e)}"


def extract_hyperlinks(text_range):
    """Extract hyperlinks from text range using multiple methods."""
    try:
        hyperlinks = []

        # Method 1: Check ActionSettings (for shape-level hyperlinks)
        if hasattr(text_range, 'ActionSettings'):
            try:
                click_action = text_range.ActionSettings(1)  # ppMouseClick = 1
                if hasattr(click_action, 'Hyperlink') and click_action.Hyperlink.Address:
                    hyperlinks.append({
                        'address': click_action.Hyperlink.Address,
                        'text': text_range.Text,
                        'type': 'shape_click'
                    })
            except:
                pass

        # Method 2: Check for hyperlinks in text runs (character-level hyperlinks)
        try:
            if hasattr(text_range, 'Runs'):
                runs = text_range.Runs()
                for run in runs:
                    try:
                        # Check if this run has a hyperlink
                        if hasattr(run, 'ActionSettings'):
                            run_action = run.ActionSettings(1)
                            if hasattr(run_action, 'Hyperlink') and run_action.Hyperlink.Address:
                                hyperlinks.append({
                                    'address': run_action.Hyperlink.Address,
                                    'text': run.Text,
                                    'type': 'text_run'
                                })
                    except:
                        continue
        except:
            pass

        # Method 3: Try to access hyperlinks collection from parent slide
        try:
            # This requires getting the parent slide, which is more complex
            # For now, we'll implement a basic text-based detection as fallback
            text = text_range.Text if hasattr(text_range, 'Text') else ""
            if "http://" in text or "https://" in text:
                # Basic URL pattern detection as fallback
                import re
                url_pattern = r'https?://[^\s<>"{}|\\^`\[\]]+'
                urls = re.findall(url_pattern, text)
                for url in urls:
                    hyperlinks.append({
                        'address': url,
                        'text': url,
                        'type': 'detected_url'
                    })
        except:
            pass

        return hyperlinks

    except:
        return []


def extract_slide_comments(slide):
    """Extract all comments from a slide with author, timestamp, and associated objects."""
    try:
        comments = []

        if hasattr(slide, 'Comments') and slide.Comments.Count > 0:
            for i in range(1, slide.Comments.Count + 1):
                comment = slide.Comments(i)
                comment_info = {
                    'text': comment.Text,
                    'author': comment.Author,
                    'date': str(comment.DateTime) if hasattr(comment, 'DateTime') else "Unknown date",
                    'position': f"({round(comment.Left, 1)}, {round(comment.Top, 1)})" if hasattr(comment, 'Left') else "Unknown position"
                }

                # Try to find the associated object using the slide's hyperlinks collection approach
                try:
                    # Method 1: Check if comment has direct parent shape reference
                    if hasattr(comment, 'Parent'):
                        parent = comment.Parent
                        # Comments might be linked to shapes in a hierarchy
                        if hasattr(parent, 'Parent') and hasattr(parent.Parent, 'ID'):
                            shape = parent.Parent
                            comment_info['associated_object'] = {
                                'name': shape.Name,
                                'id': shape.ID,
                                'type': get_shape_type_name(shape.Type) if hasattr(shape, 'Type') else 'Unknown'
                            }
                        elif hasattr(parent, 'ID') and hasattr(parent, 'Name'):
                            comment_info['associated_object'] = {
                                'name': parent.Name,
                                'id': parent.ID,
                                'type': get_shape_type_name(parent.Type) if hasattr(parent, 'Type') else 'Unknown'
                            }

                    # Method 2: Try to find via slide hyperlinks collection (research-based approach)
                    if 'associated_object' not in comment_info:
                        try:
                            # Get the slide and check hyperlinks for comment associations
                            if hasattr(slide, 'Hyperlinks'):
                                for j in range(1, slide.Hyperlinks.Count + 1):
                                    hyperlink = slide.Hyperlinks(j)
                                    # Check if this hyperlink is associated with the comment
                                    if hasattr(hyperlink, 'Parent'):
                                        # Navigate parent hierarchy based on hyperlink type
                                        hl_parent = hyperlink.Parent
                                        if hasattr(hl_parent, 'Parent'):
                                            # Try different parent levels for different hyperlink types
                                            potential_shape = None
                                            if hasattr(hl_parent.Parent, 'ID'):
                                                potential_shape = hl_parent.Parent
                                            elif hasattr(hl_parent.Parent, 'Parent') and hasattr(hl_parent.Parent.Parent, 'ID'):
                                                potential_shape = hl_parent.Parent.Parent

                                            if potential_shape and hasattr(potential_shape, 'Name'):
                                                comment_info['associated_object'] = {
                                                    'name': potential_shape.Name,
                                                    'id': potential_shape.ID,
                                                    'type': get_shape_type_name(potential_shape.Type) if hasattr(potential_shape, 'Type') else 'Unknown'
                                                }
                                                break
                        except:
                            pass

                    # Method 3: Alternative COM properties exploration
                    if 'associated_object' not in comment_info:
                        # Try other potential properties that might exist
                        for prop_name in ['Scope', 'Target', 'Anchor', 'Shape']:
                            try:
                                if hasattr(comment, prop_name):
                                    obj = getattr(comment, prop_name)
                                    if hasattr(obj, 'ID') and hasattr(obj, 'Name'):
                                        comment_info['associated_object'] = {
                                            'name': obj.Name,
                                            'id': obj.ID,
                                            'type': get_shape_type_name(obj.Type) if hasattr(obj, 'Type') else 'Unknown'
                                        }
                                        break
                            except:
                                continue

                except Exception as e:
                    # If we can't find the associated object, that's okay
                    pass

                comments.append(comment_info)

        return comments

    except Exception as e:
        return [f"Error reading comments: {str(e)}"]


def get_output_file(filename: Optional[str] = None) -> str:
    """
    Create an organized output file path following Playwright's pattern.

    Args:
        filename: Optional custom filename, defaults to slide-{timestamp}.png

    Returns:
        Full path to the output file
    """
    if filename is None:
        timestamp = datetime.now().isoformat().replace(':', '-').replace('.', '-')
        filename = f"slide-{timestamp}.png"

    # Follow Playwright's pattern: try user directory first, fallback to temp
    try:
        # Try user's home directory first (like Playwright does with rootPath)
        user_home = Path.home()
        output_dir = user_home / ".powerpoint-mcp"
        output_dir.mkdir(exist_ok=True)
        return str(output_dir / filename)
    except (PermissionError, OSError):
        # Fallback to system temp directory (like Playwright's fallback)
        temp_dir = Path(tempfile.gettempdir()) / "powerpoint-mcp-output"
        temp_dir.mkdir(exist_ok=True)
        return str(temp_dir / filename)


def add_bounding_box_overlays(image: Image.Image, slide_data: dict) -> Image.Image:
    """
    Add bounding box overlays to the slide image.
    Clean, simple implementation focused on core functionality.

    Args:
        image: PIL Image of the slide
        slide_data: Slide context data with shapes

    Returns:
        PIL Image with bounding boxes and ID labels overlaid
    """
    # Create a drawing context
    draw = ImageDraw.Draw(image)

    # Calculate scaling factors (image pixels to PowerPoint points)
    img_width, img_height = image.size
    slide_width = slide_data.get('slide_width', 720)  # Default if missing
    slide_height = slide_data.get('slide_height', 540)

    scale_x = img_width / slide_width
    scale_y = img_height / slide_height

    # Try to load a reasonable font, fallback to default
    try:
        font_size = max(12, int(img_width / 100))  # Dynamic font size
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        font = ImageFont.load_default()

    # Colors (RGB for PIL)
    box_color = (0, 255, 0)      # Green
    bg_color = (255, 255, 0)     # Yellow background
    text_color = (0, 0, 0)       # Black text

    # Draw bounding box and ID for each shape
    for shape in slide_data.get('shapes', []):
        try:
            # Convert PowerPoint coordinates to image coordinates
            x = int(shape['left'] * scale_x)
            y = int(shape['top'] * scale_y)
            w = int(shape['width'] * scale_x)
            h = int(shape['height'] * scale_y)

            # Draw bounding box
            draw.rectangle([x, y, x + w, y + h], outline=box_color, width=2)

            # Draw ID label
            id_text = f"ID:{shape['id']}"

            # Get text size (handle different PIL versions)
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

            # Draw label background and text
            draw.rectangle([label_x, label_y, label_x + text_w + 4, label_y + text_h + 2],
                         fill=bg_color)
            draw.text((label_x + 2, label_y + 1), id_text, fill=text_color, font=font)

        except Exception:
            # Skip problematic shapes
            continue

    return image


def get_slide_context_data_for_screenshot(presentation, slide_index: int) -> dict:
    """
    Get slide context data including object positions for screenshots.

    Args:
        presentation: PowerPoint presentation object
        slide_index: 1-based slide index

    Returns:
        Dictionary with slide data including shapes list with position info
    """
    try:
        slide = presentation.Slides(slide_index)
        shapes = []

        for shape in slide.Shapes:
            try:
                shape_data = {
                    'id': shape.Id,
                    'left': shape.Left,
                    'top': shape.Top,
                    'width': shape.Width,
                    'height': shape.Height,
                    'name': shape.Name
                }
                shapes.append(shape_data)
            except Exception:
                # Skip shapes that can't be accessed
                continue

        return {
            'slide_index': slide_index,
            'shapes': shapes,
            'slide_width': presentation.PageSetup.SlideWidth,
            'slide_height': presentation.PageSetup.SlideHeight
        }
    except Exception as e:
        return {'error': str(e), 'shapes': []}


def analyze_shape(shape):
    """Analyze a single shape and extract essential properties."""
    try:
        shape_info = {
            'name': shape.Name,
            'type': get_shape_type_name(shape.Type),
            'id': shape.ID,
            'position': f"({round(shape.Left, 1)}, {round(shape.Top, 1)})",
            'size': f"{round(shape.Width, 1)} x {round(shape.Height, 1)}"
        }

        # Text content with hyperlink detection
        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
            try:
                text_range = shape.TextFrame.TextRange
                raw_text = text_range.Text
                html_text = convert_text_to_html(text_range)

                shape_info['text'] = raw_text
                shape_info['html_text'] = html_text
                shape_info['font'] = f"{text_range.Font.Name}, {text_range.Font.Size}pt"

                # Extract hyperlinks from text
                hyperlinks = extract_hyperlinks(text_range)
                if hyperlinks:
                    shape_info['hyperlinks'] = hyperlinks

            except:
                shape_info['text'] = "Could not read text"

        # Enhanced table detection using Microsoft-recommended HasTable property
        table_found = False

        # Method 1: Use Microsoft's recommended HasTable property (most reliable)
        try:
            if hasattr(shape, 'HasTable') and shape.HasTable:
                table_found = True
                table = shape.Table
        except:
            # Method 2: Fallback to type checking for compatibility
            if shape.Type == 19 and hasattr(shape, 'Table'):
                table_found = True
                table = shape.Table

        # Check for tables inside groups (if not already found)
        if not table_found and shape.Type == 6:  # Group only
            try:
                # Check if it actually has GroupItems before accessing
                if hasattr(shape, 'GroupItems') and shape.GroupItems.Count > 0:
                    # Search through grouped items for tables
                    for i in range(1, shape.GroupItems.Count + 1):
                        item = shape.GroupItems(i)
                        if item.Type == 19 and hasattr(item, 'Table'):
                            table_found = True
                            table = item.Table
                            shape_info['container_type'] = 'Group'
                            break
            except:
                pass

        # For placeholders, check if they contain a table directly
        if not table_found and shape.Type == 14:  # Placeholder
            try:
                # First, try Microsoft's recommended HasTable property
                if hasattr(shape, 'HasTable') and shape.HasTable:
                    table_found = True
                    table = shape.Table
                    shape_info['container_type'] = 'Placeholder'
                # Fallback: Check PlaceholderFormat for table-specific placeholders
                elif (hasattr(shape, 'PlaceholderFormat') and
                      hasattr(shape.PlaceholderFormat, 'Type') and
                      shape.PlaceholderFormat.Type == 12 and  # ppPlaceholderTable = 12
                      hasattr(shape, 'Table') and shape.Table):
                    table_found = True
                    table = shape.Table
                    shape_info['container_type'] = 'Placeholder (Table)'
                # Final fallback: Direct table property check
                elif hasattr(shape, 'Table') and shape.Table:
                    table_found = True
                    table = shape.Table
                    shape_info['container_type'] = 'Placeholder'
            except:
                pass

        # Process table if found
        if table_found:
            try:
                shape_info['table_info'] = f"{table.Rows.Count} rows x {table.Columns.Count} columns"
                shape_info['is_table'] = True

                # Read ALL cell content with HTML formatting and hyperlinks
                table_cells_html = []
                table_cells_plain = []
                table_hyperlinks = []

                for row in range(table.Rows.Count):
                    row_cells_html = []
                    row_cells_plain = []
                    row_hyperlinks = []
                    for col in range(table.Columns.Count):
                        try:
                            cell_shape = table.Cell(row + 1, col + 1).Shape
                            cell_text = cell_shape.TextFrame.TextRange.Text.strip()
                            cell_html = convert_text_to_html(cell_shape.TextFrame.TextRange)

                            row_cells_plain.append(cell_text if cell_text else "[Empty]")
                            row_cells_html.append(cell_html if cell_html else "[Empty]")

                            # Extract hyperlinks from this cell
                            cell_hyperlinks = extract_hyperlinks(cell_shape.TextFrame.TextRange)
                            if cell_hyperlinks:
                                for link in cell_hyperlinks:
                                    link['cell_position'] = f"Row {row + 1}, Col {col + 1}"
                                row_hyperlinks.extend(cell_hyperlinks)
                        except:
                            row_cells_plain.append("[Error]")
                            row_cells_html.append("[Error]")

                    table_cells_plain.append(row_cells_plain)
                    table_cells_html.append(row_cells_html)
                    if row_hyperlinks:
                        table_hyperlinks.extend(row_hyperlinks)

                shape_info['table_content'] = table_cells_plain
                shape_info['table_content_html'] = table_cells_html

                # Store table hyperlinks if any were found
                if table_hyperlinks:
                    shape_info['table_hyperlinks'] = table_hyperlinks

                # Generate markdown table with HTML formatting
                shape_info['table_markdown'] = generate_markdown_table(table_cells_html)

            except Exception as table_error:
                shape_info['table_error'] = f"Error reading table: {table_error}"

        # Enhanced chart content extraction using Microsoft-recommended HasChart property
        chart_detected = False

        # Method 1: Use Microsoft's recommended HasChart property (most reliable)
        try:
            if hasattr(shape, 'HasChart') and shape.HasChart:
                chart_detected = True
        except:
            # Method 2: Fallback to type checking for compatibility
            if shape.Type == 3 and hasattr(shape, 'Chart'):
                chart_detected = True

        if chart_detected:
            try:
                chart = shape.Chart
                # Basic chart info (keeping for compatibility)
                shape_info['chart_info'] = f"Type: {chart.ChartType}"
                if chart.HasTitle:
                    shape_info['chart_title'] = chart.ChartTitle.Text

                # Detailed chart data extraction
                detailed_chart_data = extract_chart_data(chart)
                shape_info['chart_data'] = detailed_chart_data

            except Exception as chart_error:
                shape_info['chart_error'] = f"Error reading chart: {chart_error}"

        return shape_info

    except Exception as e:
        # Try to get basic shape info even when there's an error
        try:
            position = f"({round(shape.Left, 1)}, {round(shape.Top, 1)})"
            size = f"{round(shape.Width, 1)} x {round(shape.Height, 1)}"
        except:
            position = "Unknown"
            size = "Unknown"

        return {
            'name': f"Shape analysis error: {str(e)}",
            'type': 'Unknown',
            'id': getattr(shape, 'ID', 'Unknown'),
            'position': position,
            'size': size
        }


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
            # Fallback to plain text only if HTML conversion failed
            context_parts.append(f"Text: {shape['text']}")
            if 'font' in shape:
                context_parts.append(f"Font: {shape['font']}")

        if 'is_table' in shape and shape['is_table']:
            context_parts.append(f"Table: {shape['table_info']}")
            if 'container_type' in shape:
                context_parts.append(f"Container: {shape['container_type']}")

            # Show markdown formatted table with HTML content
            if 'table_markdown' in shape:
                context_parts.append("Table content (Markdown with HTML formatting):")
                context_parts.append(shape['table_markdown'])

            # Fallback to simple format if markdown failed
            elif 'table_content_html' in shape:
                context_parts.append("Table content (HTML formatted):")
                for row_idx, row_data in enumerate(shape['table_content_html']):
                    row_str = " | ".join(row_data)
                    context_parts.append(f"  Row {row_idx + 1}: {row_str}")

            # Final fallback to plain text
            elif 'table_content' in shape:
                context_parts.append("Table content (plain text):")
                for row_idx, row_data in enumerate(shape['table_content']):
                    row_str = " | ".join(row_data)
                    context_parts.append(f"  Row {row_idx + 1}: {row_str}")

            # Display table hyperlinks
            if 'table_hyperlinks' in shape and shape['table_hyperlinks']:
                context_parts.append("Table Hyperlinks:")
                for link in shape['table_hyperlinks']:
                    if isinstance(link, dict):
                        cell_pos = link.get('cell_position', 'Unknown position')
                        address = link.get('address', 'Unknown URL')
                        text = link.get('text', 'No text')
                        context_parts.append(f"  → {address} (Text: {text}, Cell: {cell_pos})")

            if 'table_error' in shape:
                context_parts.append(f"Table Error: {shape['table_error']}")

        # Enhanced chart information display
        if 'chart_info' in shape:
            context_parts.append(f"Chart: {shape['chart_info']}")
            if 'chart_title' in shape:
                context_parts.append(f"Title: {shape['chart_title']}")

            # Detailed chart data
            if 'chart_data' in shape and isinstance(shape['chart_data'], dict):
                chart_data = shape['chart_data']

                # Display axes information
                if 'axes' in chart_data and isinstance(chart_data['axes'], dict):
                    axes = chart_data['axes']
                    if 'category_axis' in axes and isinstance(axes['category_axis'], dict):
                        context_parts.append(f"X-Axis: {axes['category_axis'].get('title', 'No title')}")
                    if 'value_axis' in axes and isinstance(axes['value_axis'], dict):
                        value_axis = axes['value_axis']
                        context_parts.append(f"Y-Axis: {value_axis.get('title', 'No title')}")

                # Display category labels (X-axis labels)
                if 'categories' in chart_data and isinstance(chart_data['categories'], list) and chart_data['categories']:
                    if len(chart_data['categories']) <= 10:
                        context_parts.append(f"Categories: {chart_data['categories']}")
                    else:
                        context_parts.append(f"Categories: {len(chart_data['categories'])} items ({chart_data['categories'][:5]}...)")

                # Display series data
                if 'series' in chart_data and isinstance(chart_data['series'], list):
                    context_parts.append("Chart Data Series:")
                    for i, series in enumerate(chart_data['series'], 1):
                        if isinstance(series, dict):
                            series_name = series.get('name', f'Series {i}')
                            values = series.get('values', [])

                            context_parts.append(f"  Series {i}: {series_name}")
                            if values and isinstance(values, list) and len(values) <= 10:
                                # Show values if not too many
                                context_parts.append(f"    Values: {values}")
                            elif values and isinstance(values, list):
                                # Show summary for large datasets
                                context_parts.append(f"    Values: {len(values)} data points")

            if 'chart_error' in shape:
                context_parts.append(f"Chart Error: {shape['chart_error']}")

        # Display hyperlinks
        if 'hyperlinks' in shape and shape['hyperlinks']:
            context_parts.append("Hyperlinks:")
            for link in shape['hyperlinks']:
                if isinstance(link, dict):
                    context_parts.append(f"  → {link.get('address', 'Unknown URL')} (Text: {link.get('text', 'No text')})")

    # Enhanced notes section with HTML formatting
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

                # Show associated object if available
                if 'associated_object' in comment:
                    obj = comment['associated_object']
                    context_parts.append(f"  Associated Object: {obj.get('name', 'Unknown')} (ID: {obj.get('id', 'Unknown')}, Type: {obj.get('type', 'Unknown')})")

                context_parts.append(f"  Text: {comment.get('text', 'No text')}")
            else:
                context_parts.append(f"Comment {i}: {comment}")

    context_parts.append("\n=== END CONTEXT ===")
    return "\n".join(context_parts)


def powerpoint_snapshot(slide_number: Optional[int] = None,
                       include_screenshot: bool = True,
                       screenshot_filename: Optional[str] = None) -> dict:
    """
    Capture comprehensive context of a PowerPoint slide with optional screenshot.

    Similar to browser_snapshot in Playwright, this tool provides detailed information
    about the current (or specified) slide including all objects, text content with
    HTML formatting, tables, charts, and layout details.

    Now includes screenshot functionality with object bounding box overlays.

    Args:
        slide_number: Slide number to capture (1-based). If None, uses current slide.
        include_screenshot: Whether to save a screenshot with bounding boxes. Default True.
        screenshot_filename: Optional custom filename for screenshot. If None, generates slide-{timestamp}.png

    Returns:
        Dictionary with slide context data, screenshot info (if enabled), or error information
    """
    try:
        # Connect to PowerPoint
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        if not ppt_app.Presentations.Count:
            return {"error": "No PowerPoint presentation is open"}

        presentation = ppt_app.ActivePresentation

        # Determine slide to analyze
        if slide_number is None:
            slide_number = get_current_slide_index(ppt_app)
            if slide_number is None:
                slide_number = 1

        if slide_number < 1 or slide_number > presentation.Slides.Count:
            return {"error": f"Invalid slide number {slide_number}. Presentation has {presentation.Slides.Count} slides."}

        slide = presentation.Slides(slide_number)

        # Collect slide data
        slide_data = {
            'slide_number': slide_number,
            'total_slides': presentation.Slides.Count,
            'slide_name': slide.Name,
            'layout': getattr(slide.Layout, 'Name', 'Unknown Layout'),
            'object_count': slide.Shapes.Count,
            'timestamp': datetime.now().isoformat(),
            'shapes': []
        }

        # Analyze all shapes
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            shape_info = analyze_shape(shape)
            slide_data['shapes'].append(shape_info)

        # Enhanced slide notes extraction with HTML formatting using Microsoft-documented approach
        try:
            notes_page = slide.NotesPage
            if notes_page.Shapes.Count > 0:
                # Find the correct notes placeholder using the same approach as add_speaker_notes
                notes_shape = None
                for i in range(1, notes_page.Shapes.Count + 1):
                    shape = notes_page.Shapes(i)
                    try:
                        # Check if shape is a placeholder (msoPlaceholder = 14)
                        if hasattr(shape, 'Type') and shape.Type == 14:
                            # Check if it's the body placeholder (ppPlaceholderBody = 2)
                            if (hasattr(shape, 'PlaceholderFormat') and
                                hasattr(shape.PlaceholderFormat, 'Type') and
                                shape.PlaceholderFormat.Type == 2):
                                notes_shape = shape
                                break
                    except:
                        continue

                # Extract notes if we found the correct placeholder
                if notes_shape and hasattr(notes_shape, 'TextFrame') and notes_shape.TextFrame.HasText:
                    notes_text_range = notes_shape.TextFrame.TextRange
                    slide_data['notes'] = convert_text_to_html(notes_text_range)
                    slide_data['notes_plain'] = notes_text_range.Text
        except:
            pass

        # Extract slide comments with author and timestamp
        try:
            slide_comments = extract_slide_comments(slide)
            if slide_comments:
                slide_data['comments'] = slide_comments
        except:
            pass

        # Format context
        formatted_context = format_slide_context(slide_data)

        # Screenshot functionality (optional)
        screenshot_info = {}
        if include_screenshot:
            try:
                # Get slide context data for screenshot with position info
                screenshot_slide_data = get_slide_context_data_for_screenshot(presentation, slide_number)
                if 'error' not in screenshot_slide_data:
                    # Export slide to temporary file using PowerPoint's default resolution
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                        temp_path = temp_file.name

                    # Export at default resolution (no ScaleWidth/ScaleHeight = 96 DPI default)
                    slide.Export(temp_path, "PNG")

                    # Load image with PIL
                    image = Image.open(temp_path)

                    # Add bounding box overlays
                    annotated_image = add_bounding_box_overlays(image, screenshot_slide_data)

                    # Save to final output location
                    output_path = get_output_file(screenshot_filename)
                    annotated_image.save(output_path, "PNG")

                    # Clean up temp file
                    os.unlink(temp_path)

                    screenshot_info = {
                        "screenshot_saved": True,
                        "screenshot_path": output_path,
                        "image_size": f"{annotated_image.size[0]}x{annotated_image.size[1]}",
                        "objects_annotated": len(screenshot_slide_data.get('shapes', [])),
                        "screenshot_message": f"Screenshot saved with {len(screenshot_slide_data.get('shapes', []))} object annotations"
                    }
                else:
                    screenshot_info = {
                        "screenshot_saved": False,
                        "screenshot_error": f"Failed to get slide data for screenshot: {screenshot_slide_data.get('error', 'Unknown error')}"
                    }
            except Exception as screenshot_error:
                screenshot_info = {
                    "screenshot_saved": False,
                    "screenshot_error": f"Failed to take screenshot: {str(screenshot_error)}"
                }

        result = {
            "success": True,
            "slide_number": slide_number,
            "total_slides": presentation.Slides.Count,
            "object_count": slide_data['object_count'],
            "context": formatted_context,
            "slide_data": slide_data  # Raw data for potential future use
        }

        # Add screenshot info if enabled
        if include_screenshot:
            result.update(screenshot_info)

        return result

    except Exception as e:
        return {"error": f"Failed to capture slide context: {str(e)}"}