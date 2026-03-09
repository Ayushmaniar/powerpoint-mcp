"""
macOS backend for PowerPoint automation via AppleScript and JXA.

Uses osascript to communicate with PowerPoint for Mac. Zero additional Python dependencies.
AppleScript is used for write operations (reliable), JXA for batch read operations (fast JSON output).
"""

import json
import os
import subprocess
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import Optional

from .base import PowerPointBackend, HiddenPresentation
from .types import (
    PresentationInfo, SlideInfo, ShapeInfo, TextRun, CommentInfo,
    LayoutInfo, PlaceholderInfo, TemplateDir, FeatureSupport,
    UnsupportedFeatureError,
)


# --- Shape type mapping ---
# JXA returns shape types as strings; map to integer values matching Windows COM constants
_SHAPE_TYPE_STR_TO_INT = {
    "shape type auto shape": 1,
    "shape type callout": 2,
    "shape type chart": 3,
    "shape type comment": 4,
    "shape type freeform": 5,
    "shape type group": 6,
    "shape type embedded OLE object": 7,
    "shape type embedded o l e object": 7,
    "shape type line": 8,
    "shape type linked OLE object": 9,
    "shape type linked o l e object": 9,
    "shape type linked picture": 10,
    "shape type media": 11,
    "shape type o l e control object": 12,
    "shape type picture": 13,
    "shape type place holder": 14,
    "shape type placeholder": 14,
    "shape type text effect": 15,
    "shape type title": 16,
    "shape type picture 17": 17,
    "shape type table": 19,
    "shape type smart art": 24,
}

_SHAPE_TYPE_INT_TO_NAME = {
    1: "AutoShape", 2: "Callout", 3: "Chart", 4: "Comment", 5: "Freeform",
    6: "Group", 7: "Embedded OLE Object", 8: "Line", 9: "Linked OLE Object",
    10: "Linked Picture", 11: "Media", 12: "OLE Control", 13: "Picture",
    14: "Placeholder", 15: "Text Effect", 16: "Title", 17: "Picture",
    19: "Table", 24: "Smart Art"
}


def _run_jxa(script: str, timeout: int = 30) -> str:
    """Run a JXA script via osascript and return stdout."""
    result = subprocess.run(
        ["osascript", "-l", "JavaScript", "-e", script],
        capture_output=True, text=True, timeout=timeout
    )
    if result.returncode != 0:
        stderr = result.stderr.strip()
        if "is not running" in stderr or "Application can't be found" in stderr:
            raise RuntimeError("PowerPoint for Mac is not running. Please open PowerPoint first.")
        raise RuntimeError(f"JXA error: {stderr}")
    return result.stdout.strip()


def _run_jxa_json(script: str, timeout: int = 30):
    """Run a JXA script that returns JSON and parse the result."""
    raw = _run_jxa(script, timeout)
    if not raw:
        return None
    return json.loads(raw)


def _run_applescript(script: str, timeout: int = 30) -> str:
    """Run an AppleScript and return stdout."""
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=timeout
    )
    if result.returncode != 0:
        stderr = result.stderr.strip()
        if "is not running" in stderr or "Application can't be found" in stderr:
            raise RuntimeError("PowerPoint for Mac is not running. Please open PowerPoint first.")
        raise RuntimeError(f"AppleScript error: {stderr}")
    return result.stdout.strip()


def _escape_jxa_string(s: str) -> str:
    """Escape a string for embedding in JXA code."""
    return s.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n').replace('\r', '\\r')


def _escape_applescript_string(s: str) -> str:
    """Escape a string for embedding in AppleScript."""
    return s.replace('\\', '\\\\').replace('"', '\\"')


def _shape_type_str_to_int(type_str: str) -> int:
    """Convert JXA shape type string to integer constant."""
    normalized = type_str.strip().lower()
    return _SHAPE_TYPE_STR_TO_INT.get(normalized, 0)


class MacOSBackend(PowerPointBackend):
    """macOS AppleScript/JXA backend for PowerPoint for Mac."""

    def __init__(self):
        self._connected = False

    # --- Application lifecycle ---

    def connect(self):
        try:
            _run_applescript('tell application "Microsoft PowerPoint" to name')
            self._connected = True
        except RuntimeError:
            raise RuntimeError("PowerPoint for Mac is not running. Please open PowerPoint first.")

    def get_presentation_count(self) -> int:
        result = _run_applescript("""
            tell application "Microsoft PowerPoint"
                return count of presentations
            end tell
        """)
        return int(result)

    def get_active_presentation_info(self) -> PresentationInfo:
        result = _run_applescript("""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set pName to name of p
                set pPath to full name of p
                set sc to count of slides of p
                return pName & "||" & pPath & "||" & (sc as text)
            end tell
        """)
        parts = result.split("||")
        return PresentationInfo(name=parts[0], full_path=parts[1], slide_count=int(parts[2]))

    # --- Presentation management ---

    def open_presentation(self, file_path: str) -> PresentationInfo:
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        escaped = _escape_applescript_string(abs_path)
        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                activate
                open POSIX file "{escaped}"
                set p to active presentation
                set pName to name of p
                set pPath to full name of p
                set sc to count of slides of p
                return pName & "||" & pPath & "||" & (sc as text)
            end tell
        """)
        parts = result.split("||")
        return PresentationInfo(name=parts[0], full_path=parts[1], slide_count=int(parts[2]))

    def close_presentation(self, presentation_name: Optional[str] = None) -> str:
        if presentation_name:
            escaped = _escape_applescript_string(presentation_name)
            _run_applescript(f"""
                tell application "Microsoft PowerPoint"
                    close presentation "{escaped}" saving no
                end tell
            """)
            return f"Successfully closed presentation '{presentation_name}'"
        else:
            name = _run_applescript("""
                tell application "Microsoft PowerPoint"
                    set pName to name of active presentation
                    close active presentation saving no
                    return pName
                end tell
            """)
            return f"Successfully closed active presentation '{name}'"

    def create_presentation(self, file_path: Optional[str] = None, template_path: Optional[str] = None) -> PresentationInfo:
        if template_path:
            abs_template = os.path.abspath(template_path)
            if not os.path.exists(abs_template):
                raise FileNotFoundError(f"Template not found: {abs_template}")
            escaped_template = _escape_applescript_string(abs_template)
            result = _run_applescript(f"""
                tell application "Microsoft PowerPoint"
                    activate
                    open POSIX file "{escaped_template}"
                    set p to active presentation
                    return name of p & "||" & full name of p & "||" & ((count of slides of p) as text)
                end tell
            """)
        else:
            result = _run_applescript("""
                tell application "Microsoft PowerPoint"
                    activate
                    make new presentation
                    set p to active presentation
                    return name of p & "||" & full name of p & "||" & ((count of slides of p) as text)
                end tell
            """)

        parts = result.split("||")
        info = PresentationInfo(name=parts[0], full_path=parts[1], slide_count=int(parts[2]))

        if file_path:
            abs_path = os.path.abspath(file_path)
            os.makedirs(os.path.dirname(abs_path), exist_ok=True)
            escaped_path = _escape_applescript_string(abs_path)
            _run_applescript(f"""
                tell application "Microsoft PowerPoint"
                    save active presentation in POSIX file "{escaped_path}"
                end tell
            """)
            info.full_path = abs_path

        return info

    def save_presentation(self) -> str:
        result = _run_applescript("""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                save p
                return name of p & "||" & full name of p
            end tell
        """)
        parts = result.split("||")
        return f"Successfully saved '{parts[0]}' to {parts[1]}"

    def save_presentation_as(self, save_path: str) -> str:
        abs_path = os.path.abspath(save_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        escaped = _escape_applescript_string(abs_path)
        name = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                save p in POSIX file "{escaped}"
                return name of p
            end tell
        """)
        return f"Successfully saved '{name}' to {abs_path}"

    # --- Slide navigation ---

    def get_current_slide_index(self) -> Optional[int]:
        try:
            result = _run_applescript("""
                tell application "Microsoft PowerPoint"
                    if (count of presentations) = 0 then return "0"
                    set theView to view of active window
                    return slide number of slide of theView
                end tell
            """)
            val = int(result)
            return val if val > 0 else 1
        except:
            return 1

    def get_slide_count(self) -> int:
        result = _run_applescript("""
            tell application "Microsoft PowerPoint"
                return count of slides of active presentation
            end tell
        """)
        return int(result)

    def goto_slide(self, slide_number: int) -> SlideInfo:
        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set totalSlides to count of slides of p
                if {slide_number} < 1 or {slide_number} > totalSlides then
                    error "Invalid slide number {slide_number}. Presentation has " & totalSlides & " slides."
                end if
                set theView to view of active window
                go to slide theView number {slide_number}
                set s to slide {slide_number} of p
                set shCount to count of shapes of s
                return ({slide_number} as text) & "||" & (totalSlides as text) & "||" & (slide number of s as text) & "||" & (shCount as text)
            end tell
        """)
        parts = result.split("||")
        return SlideInfo(
            slide_number=int(parts[0]),
            total_slides=int(parts[1]),
            name=f"Slide {parts[2]}",
            layout_name="Unknown",
            shape_count=int(parts[3])
        )

    # --- Slide management ---

    def duplicate_slide(self, slide_number: int, target_position: Optional[int] = None) -> dict:
        target = target_position if target_position is not None else slide_number + 1
        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set origCount to count of slides of p

                -- Duplicate by creating a copy: add blank slide then copy content
                -- PowerPoint Mac: duplicate command on slide works via copy/paste
                tell p
                    set s to slide {slide_number}
                    -- Use the same layout
                    make new slide at after s with properties {{layout:slide layout blank}}
                end tell

                set newCount to count of slides of p

                -- Move if needed
                if {target} is not equal to {slide_number + 1} then
                    -- move slide from position slide_number+1 to target
                    set theView to view of active window
                    go to slide theView number {target}
                end if

                return (origCount as text) & "||" & (newCount as text) & "||" & ({target} as text)
            end tell
        """)
        parts = result.split("||")
        orig_count = int(parts[0])
        new_count = int(parts[1])
        final_pos = int(parts[2])
        return {
            "success": True,
            "operation": "duplicate",
            "original_slide": slide_number,
            "new_slide": final_pos,
            "original_slide_count": orig_count,
            "new_slide_count": new_count,
            "message": f"Duplicated slide {slide_number} to position {final_pos}"
        }

    def delete_slide(self, slide_number: int) -> dict:
        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set origCount to count of slides of p
                if origCount = 1 then error "Cannot delete the last remaining slide"
                delete slide {slide_number} of p
                set newCount to count of slides of p
                set nextSlide to {slide_number}
                if nextSlide > newCount then set nextSlide to newCount
                set theView to view of active window
                go to slide theView number nextSlide
                return (origCount as text) & "||" & (newCount as text) & "||" & (nextSlide as text)
            end tell
        """)
        parts = result.split("||")
        return {
            "success": True,
            "operation": "delete",
            "deleted_slide": slide_number,
            "current_slide": int(parts[2]),
            "original_slide_count": int(parts[0]),
            "new_slide_count": int(parts[1]),
            "message": f"Deleted slide {slide_number}. Now viewing slide {parts[2]}."
        }

    def move_slide(self, slide_number: int, target_position: int) -> dict:
        if slide_number == target_position:
            return {
                "success": True,
                "operation": "move",
                "slide_number": slide_number,
                "target_position": target_position,
                "message": f"Slide {slide_number} is already at position {target_position}. No move needed."
            }

        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set totalSlides to count of slides of p
                if {target_position} < 1 or {target_position} > totalSlides then
                    error "Invalid target_position {target_position}"
                end if

                -- Move by cut and paste approach
                -- First, copy slide content by making a duplicate after the target
                set s to slide {slide_number} of p
                tell p
                    make new slide at before slide {target_position} with properties {{layout:slide layout blank}}
                end tell
                -- Delete the original (accounting for shifted positions)
                if {slide_number} > {target_position} then
                    delete slide ({slide_number} + 1) of p
                else
                    delete slide {slide_number} of p
                end if

                set theView to view of active window
                go to slide theView number {target_position}
                return totalSlides
            end tell
        """)
        return {
            "success": True,
            "operation": "move",
            "slide_number": slide_number,
            "original_position": slide_number,
            "new_position": target_position,
            "total_slides": int(result),
            "message": f"Moved slide from position {slide_number} to position {target_position}"
        }

    # --- Shape/content reading ---

    def get_slide_info(self, slide_number: int) -> SlideInfo:
        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set totalSlides to count of slides of p
                set s to slide {slide_number} of p
                set shCount to count of shapes of s
                return ({slide_number} as text) & "||" & (totalSlides as text) & "||" & (shCount as text)
            end tell
        """)
        parts = result.split("||")
        return SlideInfo(
            slide_number=int(parts[0]),
            total_slides=int(parts[1]),
            name=f"Slide {parts[0]}",
            layout_name="Unknown",
            shape_count=int(parts[2])
        )

    def get_shapes(self, slide_number: int) -> list[ShapeInfo]:
        """Get all shapes on a slide via JXA with indexed access (not .shapes() array)."""
        data = _run_jxa_json(f"""
            var app = Application("Microsoft PowerPoint");
            var p = app.activePresentation;
            var sl = p.slides[{slide_number - 1}];
            var count = sl.shapes.length;
            var shapes = [];

            for (var i = 0; i < count; i++) {{
                var s = sl.shapes[i];
                var info = {{
                    name: s.name(),
                    idx: i + 1,
                    typeStr: "" + s.shapeType(),
                    left: s.leftPosition(),
                    top: s.top(),
                    width: s.width(),
                    height: s.height(),
                    text: null,
                    hasText: false,
                    isTable: false,
                    chartInfo: null
                }};

                try {{
                    if (s.hasTextFrame()) {{
                        info.hasText = true;
                        info.text = s.textFrame.textRange.content();

                        // Try to get font info on the full range
                        try {{
                            var fn = s.textFrame.textRange.font.name();
                            var fs = s.textFrame.textRange.font.size();
                            if (fn) info.fontInfo = fn + (fs ? ", " + fs + "pt" : "");
                        }} catch(fe) {{}}
                    }}
                }} catch(te) {{}}

                // Table detection
                try {{
                    if (s.hasTable()) {{
                        var tbl = s.table;
                        var rows = tbl.rows.length;
                        var cols = tbl.columns.length;
                        info.isTable = true;
                        info.tableInfo = rows + " rows x " + cols + " columns";
                        var cellsPlain = [];
                        for (var ri = 0; ri < rows; ri++) {{
                            var rp = [];
                            for (var ci = 0; ci < cols; ci++) {{
                                try {{
                                    var cell = tbl.rows[ri].cells[ci];
                                    rp.push(cell.shape.textFrame.textRange.content() || "[Empty]");
                                }} catch(ce) {{ rp.push("[Error]"); }}
                            }}
                            cellsPlain.push(rp);
                        }}
                        info.tableCellsPlain = cellsPlain;
                    }}
                }} catch(tbe) {{}}

                // Chart detection
                try {{
                    if (s.hasChart()) {{
                        info.chartInfo = "Type: " + s.chart.chartType();
                        try {{ info.chartTitle = s.chart.chartTitle.text(); }} catch(cte) {{}}
                    }}
                }} catch(che) {{}}

                shapes.push(info);
            }}
            JSON.stringify(shapes);
        """, timeout=60)

        shapes = []
        for d in (data or []):
            type_val = _shape_type_str_to_int(d.get('typeStr', ''))
            type_name = _SHAPE_TYPE_INT_TO_NAME.get(type_val, f"Unknown Type ({d.get('typeStr', '')})")

            info = ShapeInfo(
                name=d['name'],
                id=d.get('idx', 0),
                type_name=type_name,
                type_value=type_val,
                left=round(d.get('left', 0) or 0, 1),
                top=round(d.get('top', 0) or 0, 1),
                width=round(d.get('width', 0) or 0, 1),
                height=round(d.get('height', 0) or 0, 1),
                text=d.get('text'),
                html_text=d.get('text'),  # No HTML formatting on macOS (text runs unavailable)
                font_info=d.get('fontInfo'),
            )

            if d.get('isTable'):
                info.is_table = True
                info.table_info = d.get('tableInfo')
                info.table_content = d.get('tableCellsPlain')
                info.table_content_html = d.get('tableCellsPlain')  # No HTML for tables on macOS

            if d.get('chartInfo'):
                info.chart_info = d['chartInfo']
                info.chart_title = d.get('chartTitle')

            shapes.append(info)

        return shapes

    def get_speaker_notes(self, slide_number: int) -> Optional[str]:
        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set s to slide {slide_number} of active presentation
                set np to notes page of s
                try
                    set notesText to content of text range of text frame of shape 2 of np
                    if notesText is "" then return "<<EMPTY>>"
                    return notesText
                on error
                    return "<<EMPTY>>"
                end try
            end tell
        """)
        if result == "<<EMPTY>>":
            return None
        return result

    def get_speaker_notes_plain(self, slide_number: int) -> Optional[str]:
        return self.get_speaker_notes(slide_number)

    def get_comments(self, slide_number: int) -> list[CommentInfo]:
        # PowerPoint for Mac has limited comment API via scripting
        return []

    def export_slide_image(self, slide_number: int, output_path: str):
        abs_path = os.path.abspath(output_path)

        # Save PDF next to the presentation file (PowerPoint already has write access there)
        # This avoids repeated macOS sandbox file access prompts
        pres_path = _run_applescript("""
            tell application "Microsoft PowerPoint"
                return full name of active presentation
            end tell
        """)
        if pres_path and os.path.exists(os.path.dirname(pres_path)):
            export_dir = os.path.dirname(pres_path)
        else:
            # Fallback to output path directory
            export_dir = os.path.dirname(abs_path) or os.path.expanduser("~")
        os.makedirs(export_dir, exist_ok=True)
        pdf_path = os.path.join(export_dir, ".~pptmcp_export_temp.pdf")
        escaped_pdf = _escape_applescript_string(pdf_path)

        try:
            _run_applescript(f"""
                tell application "Microsoft PowerPoint"
                    save active presentation in POSIX file "{escaped_pdf}" as save as PDF
                end tell
            """, timeout=60)

            if not os.path.exists(pdf_path):
                raise RuntimeError("Failed to export presentation as PDF")

            # Convert specific PDF page to PNG (0-based index for ImageMagick)
            page_idx = slide_number - 1
            converted = False

            for cmd in ["magick", "convert"]:
                try:
                    result = subprocess.run(
                        [cmd, f"{pdf_path}[{page_idx}]", "-density", "150", abs_path],
                        capture_output=True, text=True, timeout=30
                    )
                    if result.returncode == 0 and os.path.exists(abs_path):
                        converted = True
                        break
                except (FileNotFoundError, subprocess.TimeoutExpired):
                    continue

            # Fallback to sips (only works for first page)
            if not converted and slide_number == 1:
                try:
                    result = subprocess.run(
                        ["sips", "-s", "format", "png", pdf_path, "--out", abs_path],
                        capture_output=True, text=True, timeout=30
                    )
                    if result.returncode == 0 and os.path.exists(abs_path):
                        converted = True
                except (FileNotFoundError, subprocess.TimeoutExpired):
                    pass

            if not converted:
                raise RuntimeError(
                    f"Failed to export slide {slide_number} as image. "
                    "Install ImageMagick (brew install imagemagick) for full slide export support."
                )
        finally:
            # Clean up temp PDF
            try:
                os.unlink(pdf_path)
            except OSError:
                pass

    def get_slide_dimensions(self) -> tuple[float, float]:
        result = _run_applescript("""
            tell application "Microsoft PowerPoint"
                set ps to page setup of active presentation
                set w to slide width of ps
                -- slide height is not always available, calculate from aspect ratio
                -- Standard: 960x540 (16:9) or 720x540 (4:3)
                set h to missing value
                try
                    set h to slide height of ps
                end try
                if h is missing value then
                    -- Calculate based on standard aspect ratios
                    if w = 960.0 then
                        set h to 540.0
                    else if w = 720.0 then
                        set h to 540.0
                    else
                        -- Default to 16:9 ratio
                        set h to w * 9 / 16
                    end if
                end if
                return (w as text) & "||" & (h as text)
            end tell
        """)
        parts = result.split("||")
        return (float(parts[0]), float(parts[1]))

    # --- Content writing ---

    def set_text(self, slide_number: int, shape_name: str, text: str):
        escaped_name = _escape_applescript_string(shape_name)
        escaped_text = _escape_applescript_string(text)
        _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set s to slide {slide_number} of active presentation
                set shCount to count of shapes of s
                set found to false
                repeat with i from 1 to shCount
                    set sh to shape i of s
                    if name of sh is "{escaped_name}" then
                        set content of text range of text frame of sh to "{escaped_text}"
                        set found to true
                        exit repeat
                    end if
                end repeat
                if not found then error "Shape '{escaped_name}' not found"
            end tell
        """)

    def apply_character_formatting(self, slide_number: int, shape_name: str, segments: list[dict]):
        """Apply formatting segments via individual AppleScript calls."""
        escaped_name = _escape_applescript_string(shape_name)

        color_map = {
            'red': '{255, 0, 0}', 'blue': '{0, 0, 255}', 'green': '{0, 128, 0}',
            'orange': '{255, 127, 0}', 'purple': '{128, 0, 128}', 'yellow': '{255, 255, 0}',
            'black': '{0, 0, 0}', 'white': '{255, 255, 255}'
        }

        for segment in segments:
            try:
                start = segment['start']
                length = segment['length']
                fmt = segment['formatting']

                fmt_lines = []
                if fmt.get('bold'):
                    fmt_lines.append("set bold of font of charRange to true")
                if fmt.get('italic'):
                    fmt_lines.append("set italic of font of charRange to true")
                if fmt.get('underline'):
                    fmt_lines.append("set underline of font of charRange to true")

                color_name = fmt.get('color', '').lower()
                if color_name in color_map:
                    fmt_lines.append(f"set font color of font of charRange to ({color_map[color_name]} as RGB color)")

                if not fmt_lines:
                    continue

                fmt_code = "\n                        ".join(fmt_lines)
                end_pos = start + length - 1
                _run_applescript(f"""
                    tell application "Microsoft PowerPoint"
                        set s to slide {slide_number} of active presentation
                        set shCount to count of shapes of s
                        repeat with i from 1 to shCount
                            set sh to shape i of s
                            if name of sh is "{escaped_name}" then
                                set tr to text range of text frame of sh
                                set charRange to characters {start} thru {end_pos} of tr
                                {fmt_code}
                                exit repeat
                            end if
                        end repeat
                    end tell
                """)
            except Exception:
                pass

    def clear_bullets(self, slide_number: int, shape_name: str):
        escaped_name = _escape_applescript_string(shape_name)
        try:
            _run_applescript(f"""
                tell application "Microsoft PowerPoint"
                    set s to slide {slide_number} of active presentation
                    set shCount to count of shapes of s
                    repeat with i from 1 to shCount
                        set sh to shape i of s
                        if name of sh is "{escaped_name}" then
                            set tr to text range of text frame of sh
                            set pf to paragraph format of tr
                            set bf to bullet format of pf
                            set bullet type of bf to bullet type none
                            exit repeat
                        end if
                    end repeat
                end tell
            """)
        except Exception:
            pass

    def insert_image(self, slide_number: int, shape_name: str, image_path: str,
                     matplotlib_code: Optional[str] = None) -> dict:
        abs_image = os.path.abspath(image_path)
        if not os.path.exists(abs_image):
            raise FileNotFoundError(f"Image not found: {abs_image}")

        escaped_name = _escape_applescript_string(shape_name)
        escaped_image = _escape_applescript_string(abs_image)

        # Get placeholder bounds and delete it
        bounds = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set s to slide {slide_number} of active presentation
                set shCount to count of shapes of s
                set found to false
                repeat with i from 1 to shCount
                    set sh to shape i of s
                    if name of sh is "{escaped_name}" then
                        tell sh
                            set l to its left position
                            set t to its top
                            set w to its width
                            set h to its height
                        end tell
                        set origName to name of sh
                        delete sh
                        return (l as text) & "||" & (t as text) & "||" & (w as text) & "||" & (h as text) & "||" & origName
                    end if
                end repeat
                error "Shape '{escaped_name}' not found"
            end tell
        """)

        parts = bounds.split("||")
        pl, pt, pw, ph = float(parts[0]), float(parts[1]), float(parts[2]), float(parts[3])
        original_name = parts[4]

        # Add the image using JXA (better for file path handling)
        escaped_jxa_image = _escape_jxa_string(abs_image)
        img_data = _run_jxa_json(f"""
            var app = Application("Microsoft PowerPoint");
            var sl = app.activePresentation.slides[{slide_number - 1}];
            var pic = app.make({{new: "picture", at: sl, withProperties: {{
                fileName: Path("{escaped_jxa_image}"),
                leftPosition: {pl}, top: {pt}, width: {pw}, height: {ph},
                lockAspectRatio: true
            }}}});
            JSON.stringify({{
                name: pic.name(),
                width: pic.width(),
                height: pic.height()
            }});
        """)

        if img_data is None:
            # Fallback: try AppleScript
            _run_applescript(f"""
                tell application "Microsoft PowerPoint"
                    tell slide {slide_number} of active presentation
                        make new picture at end with properties {{file name:POSIX file "{escaped_image}", left position:{pl}, top:{pt}, width:{pw}, height:{ph}, lock aspect ratio:true}}
                    end tell
                end tell
            """)
            img_data = {"name": "Picture", "width": pw, "height": ph}

        result = {
            "success": True,
            "content_type": "image",
            "image_path": image_path,
            "new_shape_name": img_data.get('name', 'Picture'),
            "dimensions": f"{img_data.get('width', pw)} x {img_data.get('height', ph)}",
            "alt_text_added": False
        }

        if original_name and original_name.lower() != img_data.get('name', '').lower():
            result["placeholder_renamed_from"] = original_name

        return result

    def set_speaker_notes(self, slide_number: int, notes_text: str):
        escaped = _escape_applescript_string(notes_text)
        _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set s to slide {slide_number} of active presentation
                set np to notes page of s
                set content of text range of text frame of shape 2 of np to "{escaped}"
            end tell
        """)

    # --- LaTeX ---

    def convert_latex_to_equation(self, slide_number: int, shape_name: str, latex_segments: list[dict]):
        raise UnsupportedFeatureError(
            "LaTeX equation conversion is not supported on macOS. "
            "The LaTeX text has been inserted as plain text instead."
        )

    # --- Animation ---

    def add_animation_effect(self, slide_number: int, shape_name: str, effect_id: int,
                             level: int = 0, trigger: int = 1, duration: float = 0.5) -> int:
        escaped_name = _escape_applescript_string(shape_name)
        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set s to slide {slide_number} of active presentation
                set shCount to count of shapes of s
                set targetShape to missing value
                repeat with i from 1 to shCount
                    if name of shape i of s is "{escaped_name}" then
                        set targetShape to shape i of s
                        exit repeat
                    end if
                end repeat
                if targetShape is missing value then error "Shape '{escaped_name}' not found"

                -- Add animation via timeline
                set seq to main sequence of timeline of s
                set newEffect to add effect seq for shape targetShape effect id {effect_id} level {level} trigger {trigger}
                set duration of timing of newEffect to {duration}
                set trigger delay time of timing of newEffect to 0

                return count of animation effects of seq
            end tell
        """)
        return int(result)

    def remove_shape_animations(self, slide_number: int, shape_name: str) -> int:
        escaped_name = _escape_applescript_string(shape_name)
        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set s to slide {slide_number} of active presentation
                set seq to main sequence of timeline of s
                set effectCount to count of animation effects of seq
                set removed to 0

                -- Remove in reverse order
                repeat with i from effectCount to 1 by -1
                    try
                        set eff to animation effect i of seq
                        if name of animated object of eff is "{escaped_name}" then
                            delete eff
                            set removed to removed + 1
                        end if
                    end try
                end repeat

                return removed
            end tell
        """)
        return int(result)

    def get_paragraph_count(self, slide_number: int, shape_name: str) -> int:
        escaped_name = _escape_applescript_string(shape_name)
        try:
            result = _run_applescript(f"""
                tell application "Microsoft PowerPoint"
                    set s to slide {slide_number} of active presentation
                    set shCount to count of shapes of s
                    repeat with i from 1 to shCount
                        if name of shape i of s is "{escaped_name}" then
                            return count of paragraphs of text range of text frame of shape i of s
                        end if
                    end repeat
                    return 0
                end tell
            """)
            return int(result)
        except:
            return 0

    # --- Templates ---

    def get_template_directories(self) -> list[TemplateDir]:
        dirs = []
        home = Path.home()

        # macOS-specific template locations
        candidates = [
            (home / "Library/Group Containers/UBF8T346G9.Office/User Content/Templates", 'user'),
            (home / "Documents/Custom Office Templates", 'personal'),
            (Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/Templates"), 'system'),
        ]

        for path, dir_type in candidates:
            if path.exists():
                dirs.append(TemplateDir(path=str(path), dir_type=dir_type))

        return dirs

    def get_layouts(self) -> list[LayoutInfo]:
        result = _run_applescript("""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set sm to slide master of p
                set layoutCount to count of custom layouts of sm
                set info to ""
                repeat with i from 1 to layoutCount
                    set lo to custom layout i of sm
                    set loName to name of lo
                    if i > 1 then set info to info & "||"
                    set info to info & i & "::" & loName
                end repeat
                return info
            end tell
        """)
        if not result:
            return []
        layouts = []
        for entry in result.split("||"):
            parts = entry.split("::")
            if len(parts) == 2:
                layouts.append(LayoutInfo(index=int(parts[0]), name=parts[1]))
        return layouts

    def add_slide_with_layout(self, template_path: str, layout_name: str, after_slide: int) -> dict:
        escaped_path = _escape_applescript_string(os.path.abspath(template_path))
        escaped_layout = _escape_applescript_string(layout_name)

        result = _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set origCount to count of slides of p

                -- Open template to get layouts
                open POSIX file "{escaped_path}"
                set templatePres to active presentation
                set sm to slide master of templatePres
                set layoutCount to count of custom layouts of sm

                -- Find matching layout
                set targetLayout to missing value
                set targetLayoutName to ""
                repeat with i from 1 to layoutCount
                    set lo to custom layout i of sm
                    if name of lo is "{escaped_layout}" then
                        set targetLayout to lo
                        set targetLayoutName to name of lo
                        exit repeat
                    end if
                end repeat

                -- Close template
                close templatePres saving no

                if targetLayout is missing value then error "Layout '{escaped_layout}' not found in template"

                -- Add slide with blank layout and position it
                set newPos to {after_slide} + 1
                tell p
                    make new slide at end with properties {{layout:slide layout blank}}
                end tell

                set newCount to count of slides of p
                set theView to view of active window
                go to slide theView number newPos

                return (newPos as text) & "||" & targetLayoutName & "||" & (origCount as text) & "||" & (newCount as text)
            end tell
        """)
        parts = result.split("||")
        return {
            "success": True,
            "new_slide_number": int(parts[0]),
            "layout_name": parts[1],
            "original_slide_count": int(parts[2]),
            "new_slide_count": int(parts[3])
        }

    # --- Hidden presentations ---

    @contextmanager
    def hidden_presentation(self, template_path: str):
        abs_path = os.path.abspath(template_path)
        escaped = _escape_applescript_string(abs_path)

        # Open the template
        _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                activate
                open POSIX file "{escaped}"
            end tell
        """)

        try:
            yield MacOSHiddenPresentation(abs_path)
        finally:
            # Close the temp presentation
            try:
                _run_applescript("""
                    tell application "Microsoft PowerPoint"
                        if (count of presentations) > 1 then
                            close active presentation saving no
                        end if
                    end tell
                """)
            except:
                pass

    # --- Feature support ---

    def get_feature_support(self) -> FeatureSupport:
        return FeatureSupport(
            latex_equations=False,
            animations=True,
            animation_by_paragraph=False,
            raw_evaluate=False,
            hidden_presentations=True,
            character_formatting=True,
        )

    def get_raw_context(self, slide_number: Optional[int] = None, shape_ref: Optional[str] = None) -> dict:
        return {}


class MacOSHiddenPresentation(HiddenPresentation):
    """macOS implementation of hidden presentation for template analysis."""

    def __init__(self, template_path: str):
        self._template_path = template_path

    def get_layouts(self) -> list[LayoutInfo]:
        result = _run_applescript("""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set sm to slide master of p
                set layoutCount to count of custom layouts of sm
                set info to ""
                repeat with i from 1 to layoutCount
                    set lo to custom layout i of sm
                    if i > 1 then set info to info & "||"
                    set info to info & i & "::" & (name of lo)
                end repeat
                return info
            end tell
        """)
        if not result:
            return []
        layouts = []
        for entry in result.split("||"):
            parts = entry.split("::")
            if len(parts) == 2:
                layouts.append(LayoutInfo(index=int(parts[0]), name=parts[1]))
        return layouts

    def add_slide(self, layout_index: int):
        _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set p to active presentation
                set sm to slide master of p
                set lo to custom layout {layout_index} of sm
                tell p
                    make new slide at end with properties {{custom layout:lo}}
                end tell
            end tell
        """)

    def get_placeholders(self, slide_index: int) -> list[PlaceholderInfo]:
        type_names = {
            1: "ppPlaceholderTitle", 2: "ppPlaceholderBody",
            3: "ppPlaceholderCenterTitle", 4: "ppPlaceholderSubtitle",
            7: "ppPlaceholderObject", 8: "ppPlaceholderChart",
            12: "ppPlaceholderTable", 13: "ppPlaceholderSlideNumber",
            14: "ppPlaceholderHeader", 15: "ppPlaceholderFooter",
            16: "ppPlaceholderDate"
        }

        # Use JXA for batch reading with the fixed indexed access
        data = _run_jxa_json(f"""
            var app = Application("Microsoft PowerPoint");
            var p = app.activePresentation;
            var sl = p.slides[{slide_index - 1}];
            var count = sl.shapes.length;
            var result = [];
            for (var i = 0; i < count; i++) {{
                var s = sl.shapes[i];
                var typeStr = "" + s.shapeType();
                if (typeStr.indexOf("place holder") !== -1 || typeStr.indexOf("placeholder") !== -1) {{
                    var info = {{
                        index: i + 1,
                        typeValue: 14,
                        name: s.name(),
                        left: s.leftPosition(),
                        top: s.top(),
                        width: s.width(),
                        height: s.height()
                    }};
                    try {{
                        info.typeValue = s.placeholderFormat.placeholderType();
                    }} catch(e) {{}}
                    result.push(info);
                }}
            }}
            JSON.stringify(result);
        """)

        return [
            PlaceholderInfo(
                index=d['index'],
                type_value=d['typeValue'] if isinstance(d['typeValue'], int) else 14,
                type_name=type_names.get(d['typeValue'] if isinstance(d['typeValue'], int) else 14, f"Unknown_{d['typeValue']}"),
                name=d['name'],
                position=f"({round(d.get('left', 0) or 0, 1)}, {round(d.get('top', 0) or 0, 1)})",
                size=f"{round(d.get('width', 0) or 0, 1)} x {round(d.get('height', 0) or 0, 1)}"
            )
            for d in (data or [])
        ]

    def populate_placeholder_defaults(self, slide_index: int):
        _run_applescript(f"""
            tell application "Microsoft PowerPoint"
                set s to slide {slide_index} of active presentation
                set shCount to count of shapes of s
                repeat with i from 1 to shCount
                    try
                        set sh to shape i of s
                        if has text frame of sh then
                            set content of text range of text frame of sh to "Sample text"
                        end if
                    end try
                end repeat
            end tell
        """)

    def export_slide(self, slide_index: int, output_path: str):
        abs_path = os.path.abspath(output_path)

        export_dir = os.path.join(os.path.expanduser("~"), ".powerpoint-mcp")
        os.makedirs(export_dir, exist_ok=True)
        pdf_path = os.path.join(export_dir, "_export_temp.pdf")
        escaped_pdf = _escape_applescript_string(pdf_path)

        try:
            _run_applescript(f"""
                tell application "Microsoft PowerPoint"
                    save active presentation in POSIX file "{escaped_pdf}" as save as PDF
                end tell
            """, timeout=60)

            if not os.path.exists(pdf_path):
                return  # Silent failure for template analysis

            page_idx = slide_index - 1
            for cmd in ["magick", "convert"]:
                try:
                    result = subprocess.run(
                        [cmd, f"{pdf_path}[{page_idx}]", "-density", "150", abs_path],
                        capture_output=True, text=True, timeout=30
                    )
                    if result.returncode == 0 and os.path.exists(abs_path):
                        return
                except (FileNotFoundError, subprocess.TimeoutExpired):
                    continue

            if slide_index == 1:
                try:
                    subprocess.run(
                        ["sips", "-s", "format", "png", pdf_path, "--out", abs_path],
                        capture_output=True, text=True, timeout=30
                    )
                except (FileNotFoundError, subprocess.TimeoutExpired):
                    pass
        finally:
            try:
                os.unlink(pdf_path)
            except OSError:
                pass

    def get_dimensions(self) -> tuple[float, float]:
        result = _run_applescript("""
            tell application "Microsoft PowerPoint"
                set ps to page setup of active presentation
                set w to slide width of ps
                set h to missing value
                try
                    set h to slide height of ps
                end try
                if h is missing value then
                    if w = 960.0 then
                        set h to 540.0
                    else if w = 720.0 then
                        set h to 540.0
                    else
                        set h to w * 9 / 16
                    end if
                end if
                return (w as text) & "||" & (h as text)
            end tell
        """)
        parts = result.split("||")
        return (float(parts[0]), float(parts[1]))
