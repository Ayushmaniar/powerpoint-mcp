"""
Windows COM backend for PowerPoint automation via pywin32.

Extracts all win32com.client calls from the tool files into a single backend.
"""

import os
import time
from contextlib import contextmanager
from typing import Optional

from .base import PowerPointBackend, HiddenPresentation
from .types import (
    PresentationInfo, SlideInfo, ShapeInfo, TextRun, CommentInfo,
    LayoutInfo, PlaceholderInfo, TemplateDir, FeatureSupport,
)


def _get_shape_type_name(shape_type):
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


def _convert_text_to_html(text_range):
    """Convert PowerPoint text formatting to HTML using the runs approach."""
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

            escaped_text = run_text.replace('\r\n', '<br>').replace('\r', '<br>').replace('\n', '<br>')
            escaped_text = escaped_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            escaped_text = escaped_text.replace('&lt;br&gt;', '<br>')

            formatted_text = ''.join(open_tags) + escaped_text + ''.join(close_tags)
            html_parts.append(formatted_text)

        return ''.join(html_parts)

    except Exception:
        return text_range.Text if hasattr(text_range, 'Text') else ""


def _extract_hyperlinks(text_range):
    """Extract hyperlinks from text range."""
    try:
        hyperlinks = []

        if hasattr(text_range, 'ActionSettings'):
            try:
                click_action = text_range.ActionSettings(1)
                if hasattr(click_action, 'Hyperlink') and click_action.Hyperlink.Address:
                    hyperlinks.append({
                        'address': click_action.Hyperlink.Address,
                        'text': text_range.Text,
                        'type': 'shape_click'
                    })
            except:
                pass

        try:
            if hasattr(text_range, 'Runs'):
                runs = text_range.Runs()
                for run in runs:
                    try:
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

        try:
            text = text_range.Text if hasattr(text_range, 'Text') else ""
            if "http://" in text or "https://" in text:
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


def _extract_chart_data(chart):
    """Extract comprehensive chart data."""
    try:
        chart_data = {
            'chart_type': chart.ChartType,
            'has_title': chart.HasTitle,
            'title': chart.ChartTitle.Text if chart.HasTitle else "No title"
        }

        try:
            series_data = []
            for i in range(1, chart.SeriesCollection().Count + 1):
                series = chart.SeriesCollection(i)
                series_info = {'name': f"Series {i}", 'values': [], 'categories': []}

                try:
                    if hasattr(series, 'Name') and series.Name:
                        series_info['name'] = str(series.Name)
                except:
                    pass

                try:
                    values = series.Values
                    if values:
                        series_info['values'] = [float(v) if v is not None else 0 for v in values]
                except:
                    series_info['values'] = ["Error reading values"]

                if i == 1:
                    categories = []
                    try:
                        chart_data_source = chart.ChartData
                        if hasattr(chart_data_source, 'Workbook'):
                            workbook = chart_data_source.Workbook
                            worksheet = workbook.Worksheets(1)
                            for row in range(2, 10):
                                try:
                                    cell_value = worksheet.Cells(row, 1).Value
                                    if cell_value and str(cell_value).strip():
                                        categories.append(str(cell_value))
                                except:
                                    break
                    except:
                        pass

                    if not categories:
                        try:
                            if hasattr(series, 'XValues') and series.XValues:
                                categories = [str(c) if c is not None else "" for c in series.XValues]
                        except:
                            pass

                    if not categories:
                        try:
                            if hasattr(chart, 'Axes'):
                                category_axis = chart.Axes(1)
                                if hasattr(category_axis, 'CategoryNames'):
                                    categories = [str(c) for c in category_axis.CategoryNames]
                        except:
                            pass

                    if categories:
                        chart_data['categories'] = categories

                series_data.append(series_info)

            chart_data['series'] = series_data
        except:
            chart_data['series'] = "Error reading series data"

        try:
            axes_info = {}
            if hasattr(chart, 'Axes'):
                try:
                    category_axis = chart.Axes(1)
                    axes_info['category_axis'] = {
                        'title': category_axis.AxisTitle.Text if category_axis.HasTitle else "No title",
                        'has_title': category_axis.HasTitle
                    }
                except:
                    axes_info['category_axis'] = "Error reading category axis"

                try:
                    value_axis = chart.Axes(2)
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

        try:
            if chart.HasLegend:
                chart_data['legend'] = {
                    'has_legend': True,
                    'position': getattr(chart.Legend, 'Position', 'Unknown')
                }
            else:
                chart_data['legend'] = {'has_legend': False}
        except:
            chart_data['legend'] = "Error reading legend"

        return chart_data

    except Exception as e:
        return f"Error extracting chart data: {str(e)}"


def _analyze_shape(shape) -> ShapeInfo:
    """Analyze a single COM shape and return a ShapeInfo dataclass."""
    try:
        info = ShapeInfo(
            name=shape.Name,
            id=shape.ID,
            type_name=_get_shape_type_name(shape.Type),
            type_value=shape.Type,
            left=round(shape.Left, 1),
            top=round(shape.Top, 1),
            width=round(shape.Width, 1),
            height=round(shape.Height, 1),
        )

        # Text content
        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
            try:
                text_range = shape.TextFrame.TextRange
                info.text = text_range.Text
                info.html_text = _convert_text_to_html(text_range)
                info.font_info = f"{text_range.Font.Name}, {text_range.Font.Size}pt"
                hyperlinks = _extract_hyperlinks(text_range)
                if hyperlinks:
                    info.hyperlinks = hyperlinks
            except:
                info.text = "Could not read text"

        # Table detection
        table = None
        table_found = False

        try:
            if hasattr(shape, 'HasTable') and shape.HasTable:
                table_found = True
                table = shape.Table
        except:
            if shape.Type == 19 and hasattr(shape, 'Table'):
                table_found = True
                table = shape.Table

        if not table_found and shape.Type == 6:
            try:
                if hasattr(shape, 'GroupItems') and shape.GroupItems.Count > 0:
                    for i in range(1, shape.GroupItems.Count + 1):
                        item = shape.GroupItems(i)
                        if item.Type == 19 and hasattr(item, 'Table'):
                            table_found = True
                            table = item.Table
                            info.container_type = 'Group'
                            break
            except:
                pass

        if not table_found and shape.Type == 14:
            try:
                if hasattr(shape, 'HasTable') and shape.HasTable:
                    table_found = True
                    table = shape.Table
                    info.container_type = 'Placeholder'
                elif (hasattr(shape, 'PlaceholderFormat') and
                      hasattr(shape.PlaceholderFormat, 'Type') and
                      shape.PlaceholderFormat.Type == 12 and
                      hasattr(shape, 'Table') and shape.Table):
                    table_found = True
                    table = shape.Table
                    info.container_type = 'Placeholder (Table)'
                elif hasattr(shape, 'Table') and shape.Table:
                    table_found = True
                    table = shape.Table
                    info.container_type = 'Placeholder'
            except:
                pass

        if table_found and table:
            try:
                info.is_table = True
                info.table_info = f"{table.Rows.Count} rows x {table.Columns.Count} columns"

                cells_html = []
                cells_plain = []
                table_hyperlinks_list = []

                for row in range(table.Rows.Count):
                    row_html = []
                    row_plain = []
                    for col in range(table.Columns.Count):
                        try:
                            cell_shape = table.Cell(row + 1, col + 1).Shape
                            cell_text = cell_shape.TextFrame.TextRange.Text.strip()
                            cell_html = _convert_text_to_html(cell_shape.TextFrame.TextRange)
                            row_plain.append(cell_text if cell_text else "[Empty]")
                            row_html.append(cell_html if cell_html else "[Empty]")
                            cell_links = _extract_hyperlinks(cell_shape.TextFrame.TextRange)
                            if cell_links:
                                for link in cell_links:
                                    link['cell_position'] = f"Row {row + 1}, Col {col + 1}"
                                table_hyperlinks_list.extend(cell_links)
                        except:
                            row_plain.append("[Error]")
                            row_html.append("[Error]")
                    cells_plain.append(row_plain)
                    cells_html.append(row_html)

                info.table_content = cells_plain
                info.table_content_html = cells_html
                if table_hyperlinks_list:
                    info.table_hyperlinks = table_hyperlinks_list
            except Exception as e:
                info.table_error = f"Error reading table: {e}"

        # Chart detection
        chart_detected = False
        try:
            if hasattr(shape, 'HasChart') and shape.HasChart:
                chart_detected = True
        except:
            if shape.Type == 3 and hasattr(shape, 'Chart'):
                chart_detected = True

        if chart_detected:
            try:
                chart = shape.Chart
                info.chart_info = f"Type: {chart.ChartType}"
                if chart.HasTitle:
                    info.chart_title = chart.ChartTitle.Text
                info.chart_data = _extract_chart_data(chart)
            except Exception as e:
                info.chart_error = f"Error reading chart: {e}"

        return info

    except Exception as e:
        try:
            left = round(shape.Left, 1)
            top = round(shape.Top, 1)
            width = round(shape.Width, 1)
            height = round(shape.Height, 1)
        except:
            left = top = width = height = 0.0

        return ShapeInfo(
            name=f"Shape analysis error: {str(e)}",
            id=getattr(shape, 'ID', 0),
            type_name='Unknown',
            type_value=0,
            left=left, top=top, width=width, height=height,
        )


def _get_powerpoint_char_length(text: str) -> int:
    """Calculate character length as PowerPoint COM sees it (UTF-16)."""
    return len(text.encode('utf-16-le')) // 2


class WindowsBackend(PowerPointBackend):
    """Windows COM backend using pywin32."""

    def __init__(self):
        self._ppt_app = None

    def _get_app(self):
        """Get or create PowerPoint COM application object."""
        import win32com.client
        if self._ppt_app is None:
            try:
                self._ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                self._ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                self._ppt_app.Visible = True
        return self._ppt_app

    def _ensure_presentation(self):
        """Ensure at least one presentation is open."""
        app = self._get_app()
        if not app.Presentations.Count:
            raise RuntimeError("No PowerPoint presentation is open")
        return app

    def _get_slide(self, slide_number: int):
        """Get a slide object by number, with validation."""
        app = self._ensure_presentation()
        presentation = app.ActivePresentation
        if slide_number < 1 or slide_number > presentation.Slides.Count:
            raise ValueError(f"Invalid slide number {slide_number}. Presentation has {presentation.Slides.Count} slides.")
        return presentation.Slides(slide_number)

    def _find_shape(self, slide, shape_name: str):
        """Find a shape on a slide by name (case-insensitive)."""
        for shape in slide.Shapes:
            if shape.Name.lower() == shape_name.lower():
                return shape
        available = [s.Name for s in slide.Shapes]
        raise ValueError(f"Shape '{shape_name}' not found. Available shapes: {available}")

    def _goto_slide_view(self, slide_number: int):
        """Switch the PowerPoint window to show the given slide."""
        app = self._get_app()
        try:
            if hasattr(app, 'ActiveWindow') and app.ActiveWindow:
                view = app.ActiveWindow.View
                if hasattr(view, 'GotoSlide'):
                    view.GotoSlide(slide_number)
                elif hasattr(view, 'Slide'):
                    view.Slide = app.ActivePresentation.Slides(slide_number)
        except Exception:
            pass

    # --- Application lifecycle ---

    def connect(self):
        self._get_app()

    def get_presentation_count(self) -> int:
        return self._get_app().Presentations.Count

    def get_active_presentation_info(self) -> PresentationInfo:
        app = self._ensure_presentation()
        p = app.ActivePresentation
        return PresentationInfo(name=p.Name, full_path=p.FullName, slide_count=p.Slides.Count)

    # --- Presentation management ---

    def open_presentation(self, file_path: str) -> PresentationInfo:
        app = self._get_app()
        app.Visible = True
        abs_path = os.path.abspath(file_path)

        # Check if already open
        for presentation in app.Presentations:
            try:
                if os.path.samefile(presentation.FullName, abs_path):
                    return PresentationInfo(
                        name=presentation.Name,
                        full_path=presentation.FullName,
                        slide_count=presentation.Slides.Count
                    )
            except (OSError, AttributeError):
                continue

        presentation = app.Presentations.Open(abs_path)
        return PresentationInfo(
            name=presentation.Name,
            full_path=presentation.FullName,
            slide_count=presentation.Slides.Count
        )

    def close_presentation(self, presentation_name: Optional[str] = None) -> str:
        app = self._ensure_presentation()
        if presentation_name:
            for presentation in app.Presentations:
                if presentation.Name == presentation_name:
                    presentation.Close()
                    return f"Successfully closed presentation '{presentation_name}'"
            raise ValueError(f"Presentation '{presentation_name}' not found")
        else:
            name = app.ActivePresentation.Name
            app.ActivePresentation.Close()
            return f"Successfully closed active presentation '{name}'"

    def create_presentation(self, file_path: Optional[str] = None, template_path: Optional[str] = None) -> PresentationInfo:
        app = self._get_app()
        app.Visible = True

        if template_path:
            abs_template = os.path.abspath(template_path)
            presentation = app.Presentations.Open(abs_template)
        else:
            presentation = app.Presentations.Add()

        if file_path:
            abs_path = os.path.abspath(file_path)
            os.makedirs(os.path.dirname(abs_path), exist_ok=True)
            presentation.SaveAs(abs_path)

        return PresentationInfo(
            name=presentation.Name,
            full_path=presentation.FullName if file_path else "",
            slide_count=presentation.Slides.Count
        )

    def save_presentation(self) -> str:
        app = self._ensure_presentation()
        presentation = app.ActivePresentation
        if hasattr(presentation, 'FullName') and presentation.FullName:
            presentation.Save()
            return f"Successfully saved '{presentation.Name}' to {presentation.FullName}"
        raise RuntimeError(f"Presentation '{presentation.Name}' has never been saved. Use save_as.")

    def save_presentation_as(self, save_path: str) -> str:
        app = self._ensure_presentation()
        presentation = app.ActivePresentation
        abs_path = os.path.abspath(save_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        presentation.SaveAs(abs_path)
        return f"Successfully saved '{presentation.Name}' to {abs_path}"

    # --- Slide navigation ---

    def get_current_slide_index(self) -> Optional[int]:
        try:
            app = self._get_app()
            if not app.Presentations.Count:
                return None

            try:
                active_window = app.ActiveWindow
                try:
                    if hasattr(active_window, 'View') and hasattr(active_window.View, 'Slide'):
                        idx = active_window.View.Slide.SlideIndex
                        if idx > 0:
                            return idx
                except:
                    pass

                try:
                    if (hasattr(active_window, 'Selection') and
                        hasattr(active_window.Selection, 'SlideRange') and
                        active_window.Selection.SlideRange.Count > 0):
                        return active_window.Selection.SlideRange[0].SlideIndex
                except:
                    pass
            except:
                pass

            try:
                if hasattr(app, 'SlideShowWindows') and app.SlideShowWindows.Count > 0:
                    return app.SlideShowWindows(1).View.CurrentShowPosition
            except:
                pass

            if app.ActivePresentation.Slides.Count > 0:
                return 1
            return None

        except Exception:
            return 1

    def get_slide_count(self) -> int:
        return self._ensure_presentation().ActivePresentation.Slides.Count

    def goto_slide(self, slide_number: int) -> SlideInfo:
        slide = self._get_slide(slide_number)
        self._goto_slide_view(slide_number)
        app = self._get_app()
        presentation = app.ActivePresentation
        return SlideInfo(
            slide_number=slide_number,
            total_slides=presentation.Slides.Count,
            name=getattr(slide, 'Name', f"Slide {slide_number}"),
            layout_name=getattr(slide.Layout, 'Name', 'Unknown Layout') if hasattr(slide, 'Layout') else 'Unknown',
            shape_count=slide.Shapes.Count
        )

    # --- Slide management ---

    def duplicate_slide(self, slide_number: int, target_position: Optional[int] = None) -> dict:
        app = self._ensure_presentation()
        presentation = app.ActivePresentation
        original_count = presentation.Slides.Count

        source_slide = presentation.Slides(slide_number)
        source_slide.Duplicate()

        if target_position is not None:
            current_position = slide_number + 1
            new_count = presentation.Slides.Count
            if target_position < 1 or target_position > new_count:
                raise ValueError(f"Invalid target_position {target_position}. Must be between 1 and {new_count}.")
            if target_position != current_position:
                presentation.Slides(current_position).MoveTo(target_position)
            final_position = target_position
        else:
            final_position = slide_number + 1

        self._goto_slide_view(final_position)

        return {
            "success": True,
            "operation": "duplicate",
            "original_slide": slide_number,
            "new_slide": final_position,
            "original_slide_count": original_count,
            "new_slide_count": presentation.Slides.Count,
            "message": f"Duplicated slide {slide_number} to position {final_position}"
        }

    def delete_slide(self, slide_number: int) -> dict:
        app = self._ensure_presentation()
        presentation = app.ActivePresentation
        original_count = presentation.Slides.Count

        if original_count == 1:
            raise RuntimeError("Cannot delete the last remaining slide in the presentation.")

        presentation.Slides(slide_number).Delete()
        new_count = presentation.Slides.Count
        next_slide = min(slide_number, new_count)
        self._goto_slide_view(next_slide)

        return {
            "success": True,
            "operation": "delete",
            "deleted_slide": slide_number,
            "current_slide": next_slide,
            "original_slide_count": original_count,
            "new_slide_count": new_count,
            "message": f"Deleted slide {slide_number}. Now viewing slide {next_slide}."
        }

    def move_slide(self, slide_number: int, target_position: int) -> dict:
        app = self._ensure_presentation()
        presentation = app.ActivePresentation
        slide_count = presentation.Slides.Count

        if target_position < 1 or target_position > slide_count:
            raise ValueError(f"Invalid target_position {target_position}. Must be between 1 and {slide_count}.")

        if slide_number == target_position:
            return {
                "success": True,
                "operation": "move",
                "slide_number": slide_number,
                "target_position": target_position,
                "message": f"Slide {slide_number} is already at position {target_position}. No move needed."
            }

        presentation.Slides(slide_number).MoveTo(target_position)
        self._goto_slide_view(target_position)

        return {
            "success": True,
            "operation": "move",
            "slide_number": slide_number,
            "original_position": slide_number,
            "new_position": target_position,
            "total_slides": slide_count,
            "message": f"Moved slide from position {slide_number} to position {target_position}"
        }

    # --- Shape/content reading ---

    def get_slide_info(self, slide_number: int) -> SlideInfo:
        slide = self._get_slide(slide_number)
        app = self._get_app()
        presentation = app.ActivePresentation
        return SlideInfo(
            slide_number=slide_number,
            total_slides=presentation.Slides.Count,
            name=slide.Name,
            layout_name=getattr(slide.Layout, 'Name', 'Unknown Layout') if hasattr(slide, 'Layout') else 'Unknown',
            shape_count=slide.Shapes.Count
        )

    def get_shapes(self, slide_number: int) -> list[ShapeInfo]:
        slide = self._get_slide(slide_number)
        shapes = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            shapes.append(_analyze_shape(shape))
        return shapes

    def get_speaker_notes(self, slide_number: int) -> Optional[str]:
        slide = self._get_slide(slide_number)
        try:
            notes_page = slide.NotesPage
            for i in range(1, notes_page.Shapes.Count + 1):
                shape = notes_page.Shapes(i)
                try:
                    if hasattr(shape, 'Type') and shape.Type == 14:
                        if (hasattr(shape, 'PlaceholderFormat') and
                            hasattr(shape.PlaceholderFormat, 'Type') and
                            shape.PlaceholderFormat.Type == 2):
                            if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                                return _convert_text_to_html(shape.TextFrame.TextRange)
                except:
                    continue
        except:
            pass
        return None

    def get_speaker_notes_plain(self, slide_number: int) -> Optional[str]:
        slide = self._get_slide(slide_number)
        try:
            notes_page = slide.NotesPage
            for i in range(1, notes_page.Shapes.Count + 1):
                shape = notes_page.Shapes(i)
                try:
                    if hasattr(shape, 'Type') and shape.Type == 14:
                        if (hasattr(shape, 'PlaceholderFormat') and
                            hasattr(shape.PlaceholderFormat, 'Type') and
                            shape.PlaceholderFormat.Type == 2):
                            if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                                return shape.TextFrame.TextRange.Text
                except:
                    continue
        except:
            pass
        return None

    def get_comments(self, slide_number: int) -> list[CommentInfo]:
        slide = self._get_slide(slide_number)
        comments = []
        try:
            if hasattr(slide, 'Comments') and slide.Comments.Count > 0:
                for i in range(1, slide.Comments.Count + 1):
                    comment = slide.Comments(i)
                    assoc = None

                    try:
                        if hasattr(comment, 'Parent'):
                            parent = comment.Parent
                            if hasattr(parent, 'Parent') and hasattr(parent.Parent, 'ID'):
                                s = parent.Parent
                                assoc = {'name': s.Name, 'id': s.ID,
                                         'type': _get_shape_type_name(s.Type) if hasattr(s, 'Type') else 'Unknown'}
                            elif hasattr(parent, 'ID') and hasattr(parent, 'Name'):
                                assoc = {'name': parent.Name, 'id': parent.ID,
                                         'type': _get_shape_type_name(parent.Type) if hasattr(parent, 'Type') else 'Unknown'}
                    except:
                        pass

                    if assoc is None:
                        for prop_name in ['Scope', 'Target', 'Anchor', 'Shape']:
                            try:
                                if hasattr(comment, prop_name):
                                    obj = getattr(comment, prop_name)
                                    if hasattr(obj, 'ID') and hasattr(obj, 'Name'):
                                        assoc = {'name': obj.Name, 'id': obj.ID,
                                                 'type': _get_shape_type_name(obj.Type) if hasattr(obj, 'Type') else 'Unknown'}
                                        break
                            except:
                                continue

                    comments.append(CommentInfo(
                        text=comment.Text,
                        author=comment.Author,
                        date=str(comment.DateTime) if hasattr(comment, 'DateTime') else "Unknown date",
                        position=f"({round(comment.Left, 1)}, {round(comment.Top, 1)})" if hasattr(comment, 'Left') else "Unknown position",
                        associated_object=assoc
                    ))
        except:
            pass
        return comments

    def export_slide_image(self, slide_number: int, output_path: str):
        slide = self._get_slide(slide_number)
        slide.Export(output_path, "PNG")

    def get_slide_dimensions(self) -> tuple[float, float]:
        app = self._ensure_presentation()
        p = app.ActivePresentation
        return (p.PageSetup.SlideWidth, p.PageSetup.SlideHeight)

    # --- Content writing ---

    def set_text(self, slide_number: int, shape_name: str, text: str):
        slide = self._get_slide(slide_number)
        shape = self._find_shape(slide, shape_name)
        if not hasattr(shape, 'TextFrame'):
            raise ValueError(f"Shape '{shape_name}' cannot hold text (no TextFrame)")
        shape.TextFrame.TextRange.Text = text

    def apply_character_formatting(self, slide_number: int, shape_name: str, segments: list[dict]):
        slide = self._get_slide(slide_number)
        shape = self._find_shape(slide, shape_name)
        text_range = shape.TextFrame.TextRange

        color_map = {
            'red': 255, 'blue': 16711680, 'green': 65280,
            'orange': 33023, 'purple': 8388736, 'yellow': 65535,
            'black': 0, 'white': 16777215
        }

        for segment in segments:
            try:
                start = segment['start']
                length = segment['length']
                fmt = segment['formatting']
                char_range = text_range.Characters(start, length)

                if fmt.get('bold'):
                    char_range.Font.Bold = True
                if fmt.get('italic'):
                    char_range.Font.Italic = True
                if fmt.get('underline'):
                    char_range.Font.Underline = True

                color_name = fmt.get('color', '').lower()
                if color_name in color_map:
                    char_range.Font.Color.RGB = color_map[color_name]
            except Exception:
                pass

    def clear_bullets(self, slide_number: int, shape_name: str):
        slide = self._get_slide(slide_number)
        shape = self._find_shape(slide, shape_name)
        text_range = shape.TextFrame.TextRange

        try:
            pf = text_range.ParagraphFormat
            pf.Bullet.Visible = 0
            pf.Bullet.Type = 0
        except:
            pass

        try:
            paragraphs = text_range.Paragraphs()
            for idx in range(1, paragraphs.Count + 1):
                try:
                    para_bullet = paragraphs(idx).ParagraphFormat.Bullet
                    para_bullet.Visible = 0
                    para_bullet.Type = 0
                except:
                    continue
        except:
            pass

    def insert_image(self, slide_number: int, shape_name: str, image_path: str,
                     matplotlib_code: Optional[str] = None) -> dict:
        slide = self._get_slide(slide_number)
        shape = self._find_shape(slide, shape_name)

        placeholder_left = shape.Left
        placeholder_top = shape.Top
        placeholder_width = shape.Width
        placeholder_height = shape.Height
        original_name = shape.Name

        shape.Delete()

        temp_shape = slide.Shapes.AddPicture(
            FileName=image_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0, Top=0, Width=-1, Height=-1
        )

        image_width = temp_shape.Width
        image_height = temp_shape.Height
        placeholder_aspect = placeholder_width / placeholder_height
        image_aspect = image_width / image_height

        if image_aspect > placeholder_aspect:
            final_width = placeholder_width
            final_height = placeholder_width / image_aspect
        else:
            final_height = placeholder_height
            final_width = placeholder_height * image_aspect

        final_left = placeholder_left + (placeholder_width - final_width) / 2
        final_top = placeholder_top + (placeholder_height - final_height) / 2

        temp_shape.Left = final_left
        temp_shape.Top = final_top
        temp_shape.Width = final_width
        temp_shape.Height = final_height

        try:
            temp_shape.Name = original_name
        except:
            pass

        if matplotlib_code:
            try:
                temp_shape.AlternativeText = f"Code used to generate this image:\n\n{matplotlib_code}"
            except:
                pass

        result = {
            "success": True,
            "content_type": "image",
            "image_path": image_path,
            "new_shape_id": temp_shape.Id,
            "new_shape_name": temp_shape.Name,
            "dimensions": f"{final_width} x {final_height}",
            "alt_text_added": matplotlib_code is not None
        }

        if original_name and original_name.lower() != temp_shape.Name.lower():
            result["placeholder_renamed_from"] = original_name

        return result

    def set_speaker_notes(self, slide_number: int, notes_text: str):
        slide = self._get_slide(slide_number)
        notes_page = slide.NotesPage
        notes_shape = None

        for i in range(1, notes_page.Shapes.Count + 1):
            shape = notes_page.Shapes(i)
            try:
                if hasattr(shape, 'Type') and shape.Type == 14:
                    if (hasattr(shape, 'PlaceholderFormat') and
                        hasattr(shape.PlaceholderFormat, 'Type') and
                        shape.PlaceholderFormat.Type == 2):
                        notes_shape = shape
                        break
            except:
                continue

        if notes_shape and hasattr(notes_shape, 'TextFrame') and notes_shape.TextFrame:
            notes_shape.TextFrame.TextRange.Text = notes_text
        else:
            raise RuntimeError(f"Could not find notes placeholder for slide {slide_number}")

    # --- LaTeX ---

    def convert_latex_to_equation(self, slide_number: int, shape_name: str, latex_segments: list[dict]):
        slide = self._get_slide(slide_number)
        shape = self._find_shape(slide, shape_name)
        text_range = shape.TextFrame.TextRange
        app = self._get_app()

        try:
            app.Activate()
            if hasattr(app.ActiveWindow, 'View'):
                app.ActiveWindow.ViewType = 1  # ppViewNormal
                time.sleep(0.1)
                view = app.ActiveWindow.View
                if hasattr(view, 'GotoSlide'):
                    view.GotoSlide(slide_number)
                    time.sleep(0.1)
            app.ActiveWindow.Activate()
        except:
            pass

        for segment in reversed(latex_segments):
            try:
                start_pos = segment['start']
                length = segment['length']
                old_eq_length = length
                text_length_before = _get_powerpoint_char_length(text_range.Text)

                char_range = text_range.Characters(start_pos, length)
                char_range.Select()
                time.sleep(0.05)

                app.CommandBars.ExecuteMso("InsertBuildingBlocksEquationsGallery")
                time.sleep(0.1)
                app.CommandBars.ExecuteMso("EquationLaTeXToMath")
                time.sleep(0.05)

                try:
                    time.sleep(0.05)
                    text_length_after = _get_powerpoint_char_length(text_range.Text)
                    segment['actual_new_length'] = text_length_after - text_length_before + old_eq_length
                except:
                    segment['actual_new_length'] = old_eq_length

            except Exception:
                segment['actual_new_length'] = segment.get('length', 0)
                continue

        try:
            app.ActiveWindow.ViewType = 9  # ppViewNormal
        except:
            pass

    # --- Animation ---

    def add_animation_effect(self, slide_number: int, shape_name: str, effect_id: int,
                             level: int = 0, trigger: int = 1, duration: float = 0.5) -> int:
        slide = self._get_slide(slide_number)
        shape = self._find_shape(slide, shape_name)
        main_sequence = slide.TimeLine.MainSequence

        new_effect = main_sequence.AddEffect(
            Shape=shape,
            effectId=effect_id,
            Level=level,
            trigger=trigger,
            Index=-1
        )
        new_effect.Timing.Duration = duration
        new_effect.Timing.TriggerDelayTime = 0.0

        return main_sequence.Count

    def remove_shape_animations(self, slide_number: int, shape_name: str) -> int:
        slide = self._get_slide(slide_number)
        shape = self._find_shape(slide, shape_name)
        main_sequence = slide.TimeLine.MainSequence

        effects_to_remove = []
        for i in range(1, main_sequence.Count + 1):
            if main_sequence.Item(i).Shape.Name == shape.Name:
                effects_to_remove.append(i)

        for idx in reversed(effects_to_remove):
            main_sequence.Item(idx).Delete()

        return len(effects_to_remove)

    def get_paragraph_count(self, slide_number: int, shape_name: str) -> int:
        slide = self._get_slide(slide_number)
        shape = self._find_shape(slide, shape_name)
        if hasattr(shape, 'TextFrame'):
            try:
                return shape.TextFrame.TextRange.Paragraphs().Count
            except:
                return 0
        return 0

    # --- Templates ---

    def get_template_directories(self) -> list[TemplateDir]:
        from pathlib import Path
        dirs = []
        username = os.environ.get('USERNAME', '')

        personal = Path(f"C:/Users/{username}/Documents/Custom Office Templates")
        if personal.exists():
            dirs.append(TemplateDir(path=str(personal), dir_type='personal'))

        user = Path(f"C:/Users/{username}/AppData/Roaming/Microsoft/Templates")
        if user.exists():
            dirs.append(TemplateDir(path=str(user), dir_type='user'))

        system_locations = [
            "C:/Program Files/Microsoft Office/Templates",
            "C:/Program Files/Microsoft Office/root/Templates",
            "C:/Program Files (x86)/Microsoft Office/Templates",
            "C:/Program Files (x86)/Microsoft Office/root/Templates"
        ]
        for loc in system_locations:
            if Path(loc).exists():
                dirs.append(TemplateDir(path=loc, dir_type='system'))

        return dirs

    def get_layouts(self) -> list[LayoutInfo]:
        app = self._ensure_presentation()
        presentation = app.ActivePresentation
        layouts = []
        master = presentation.SlideMaster
        for i in range(1, master.CustomLayouts.Count + 1):
            layouts.append(LayoutInfo(index=i, name=master.CustomLayouts(i).Name))
        return layouts

    def add_slide_with_layout(self, template_path: str, layout_name: str, after_slide: int) -> dict:
        app = self._ensure_presentation()
        presentation = app.ActivePresentation
        original_count = presentation.Slides.Count

        design = presentation.Designs.Load(template_path)
        slide_master = design.SlideMaster

        target_layout = None
        for i in range(1, slide_master.CustomLayouts.Count + 1):
            layout = slide_master.CustomLayouts(i)
            if layout.Name.lower() == layout_name.lower():
                target_layout = layout
                break

        if not target_layout:
            raise ValueError(f"Layout '{layout_name}' not found in template")

        new_pos = after_slide + 1
        presentation.Slides.AddSlide(new_pos, target_layout)
        self._goto_slide_view(new_pos)

        return {
            "success": True,
            "new_slide_number": new_pos,
            "layout_name": target_layout.Name,
            "original_slide_count": original_count,
            "new_slide_count": presentation.Slides.Count,
        }

    # --- Hidden presentations ---

    @contextmanager
    def hidden_presentation(self, template_path: str):
        app = self._get_app()
        if app is None:
            import win32com.client
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                app = win32com.client.Dispatch("PowerPoint.Application")
                app.Visible = True

        temp_pres = app.Presentations.Add(WithWindow=False)
        temp_pres.ApplyTemplate(template_path)

        try:
            yield WindowsHiddenPresentation(temp_pres)
        finally:
            try:
                temp_pres.Close()
            except:
                pass

    # --- Feature support ---

    def get_feature_support(self) -> FeatureSupport:
        return FeatureSupport(
            latex_equations=True,
            animations=True,
            animation_by_paragraph=True,
            raw_evaluate=True,
            hidden_presentations=True,
            character_formatting=True,
        )

    def get_raw_context(self, slide_number: Optional[int] = None, shape_ref: Optional[str] = None) -> dict:
        import math
        app = self._get_app()

        if not app.Presentations.Count:
            return {}

        presentation = app.ActivePresentation

        if slide_number is not None:
            if slide_number < 1 or slide_number > presentation.Slides.Count:
                return {}
            slide = presentation.Slides(slide_number)
        else:
            try:
                slide = app.ActiveWindow.View.Slide
            except:
                if presentation.Slides.Count > 0:
                    slide = presentation.Slides(1)
                else:
                    return {}

        shape = None
        if shape_ref:
            for s in slide.Shapes:
                if s.Name == shape_ref or str(s.Id) == shape_ref:
                    shape = s
                    break

        try:
            import numpy as np
        except ImportError:
            np = None

        return {
            'ppt': app,
            'presentation': presentation,
            'slide': slide,
            'shape': shape,
            'math': math,
            'np': np,
            'has_numpy': np is not None,
        }


class WindowsHiddenPresentation(HiddenPresentation):
    """Windows COM implementation of hidden presentation for template analysis."""

    def __init__(self, com_presentation):
        self._pres = com_presentation

    def get_layouts(self) -> list[LayoutInfo]:
        layouts = []
        master = self._pres.SlideMaster
        for i in range(1, master.CustomLayouts.Count + 1):
            layouts.append(LayoutInfo(index=i, name=master.CustomLayouts(i).Name))
        return layouts

    def add_slide(self, layout_index: int):
        layout = self._pres.SlideMaster.CustomLayouts(layout_index)
        self._pres.Slides.AddSlide(1, layout)

    def get_placeholders(self, slide_index: int) -> list[PlaceholderInfo]:
        slide = self._pres.Slides(slide_index)
        placeholders = []

        type_names = {
            1: "ppPlaceholderTitle", 2: "ppPlaceholderBody",
            3: "ppPlaceholderCenterTitle", 4: "ppPlaceholderSubtitle",
            7: "ppPlaceholderObject", 8: "ppPlaceholderChart",
            12: "ppPlaceholderTable", 13: "ppPlaceholderSlideNumber",
            14: "ppPlaceholderHeader", 15: "ppPlaceholderFooter",
            16: "ppPlaceholderDate"
        }

        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            try:
                if hasattr(shape, 'Type') and shape.Type == 14:
                    pt = shape.PlaceholderFormat.Type
                    placeholders.append(PlaceholderInfo(
                        index=i,
                        type_value=pt,
                        type_name=type_names.get(pt, f"Unknown_{pt}"),
                        name=shape.Name,
                        position=f"({round(shape.Left, 1)}, {round(shape.Top, 1)})",
                        size=f"{round(shape.Width, 1)} x {round(shape.Height, 1)}"
                    ))
            except:
                continue
        return placeholders

    def populate_placeholder_defaults(self, slide_index: int):
        slide = self._pres.Slides(slide_index)
        default_texts = {
            1: "Click to edit Master title style",
            2: "Click to edit Master text styles\n• Second level\n  • Third level\n    • Fourth level\n      • Fifth level",
            3: "Click to edit Master title style",
            4: "Click to edit Master subtitle style",
            7: "Click to add content",
            8: "Click to add chart",
            12: "Click to add table",
        }

        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            try:
                if hasattr(shape, 'Type') and shape.Type == 14:
                    if hasattr(shape, 'TextFrame') and shape.TextFrame and hasattr(shape, 'PlaceholderFormat'):
                        pt = shape.PlaceholderFormat.Type
                        if pt in default_texts:
                            shape.TextFrame.TextRange.Text = default_texts[pt]
            except:
                continue

    def export_slide(self, slide_index: int, output_path: str):
        self._pres.Slides(slide_index).Export(output_path, "PNG")

    def get_dimensions(self) -> tuple[float, float]:
        return (self._pres.PageSetup.SlideWidth, self._pres.PageSetup.SlideHeight)
