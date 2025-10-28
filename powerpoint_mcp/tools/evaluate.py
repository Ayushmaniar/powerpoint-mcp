"""
PowerPoint code evaluation tool for MCP server.
Execute arbitrary Python code in PowerPoint automation context.
"""

import math
import win32com.client
from typing import Optional, Any
import json
from .skills import skills


def powerpoint_evaluate(
    code: str,
    slide_number: Optional[int] = None,
    shape_ref: Optional[str] = None,
    description: Optional[str] = None
) -> dict:
    """
    Execute arbitrary Python code in PowerPoint automation context.

    Enables free-form PowerPoint automation beyond predefined tools, including:
    - Complex geometric calculations (circular flowcharts, grid layouts)
    - Advanced shape manipulations (rotations, custom positioning)
    - Data extraction using custom logic
    - Multi-shape operations with mathematical relationships
    - Access to all PowerPoint MCP tools via the 'skills' object

    Args:
        code: Python code to execute. Has access to:
              - ppt: PowerPoint Application object
              - presentation: Active presentation
              - slide: Current or specified slide
              - shape: Target shape (if shape_ref provided)
              - math: Python math module for calculations
              - skills: Access to all PowerPoint MCP tools
                  - skills.populate_placeholder(name, content)
                  - skills.snapshot()
                  - skills.switch_slide(n)
                  - skills.add_speaker_notes(n, text)
                  - skills.manage_slide(operation, n, target)
                  - skills.manage_presentation(action, ...)
                  - And more... (see skills.py for full list)

        slide_number: Target slide (1-based). If None, uses current slide
        shape_ref: Optional shape ID/Name to operate on (from slide_snapshot)
        description: Human-readable description of operation intent

    Returns:
        Dictionary with success/error status and optional result data
    """

    try:
        # Get PowerPoint objects
        ppt = win32com.client.Dispatch("PowerPoint.Application")

        if not ppt.Presentations.Count:
            return {"error": "No presentation is currently open"}

        presentation = ppt.ActivePresentation

        # Get target slide
        if slide_number is not None:
            if slide_number < 1 or slide_number > presentation.Slides.Count:
                return {
                    "error": f"Slide {slide_number} out of range (1-{presentation.Slides.Count})"
                }
            slide = presentation.Slides(slide_number)
        else:
            # Use current active slide
            try:
                slide = ppt.ActiveWindow.View.Slide
            except:
                # Fallback to first slide if no active window
                if presentation.Slides.Count > 0:
                    slide = presentation.Slides(1)
                else:
                    return {"error": "No slides in presentation"}

        # Get target shape if specified
        shape = None
        if shape_ref:
            for s in slide.Shapes:
                if s.Name == shape_ref or str(s.Id) == shape_ref:
                    shape = s
                    break
            if not shape:
                return {"error": f"Shape '{shape_ref}' not found on slide {slide.SlideNumber}"}

        # Try to import numpy, but don't fail if unavailable
        try:
            import numpy as np
            has_numpy = True
        except ImportError:
            np = None
            has_numpy = False

        # Create execution context with PowerPoint objects and math utilities
        context = {
            'ppt': ppt,
            'presentation': presentation,
            'slide': slide,
            'shape': shape,
            'math': math,
            'np': np,
            'has_numpy': has_numpy,
            'skills': skills,  # Access to all PowerPoint MCP tools
            # Python builtins for general programming
            'range': range,
            'len': len,
            'str': str,
            'int': int,
            'float': float,
            'bool': bool,
            'list': list,
            'dict': dict,
            'tuple': tuple,
            'set': set,
            'enumerate': enumerate,
            'zip': zip,
            'round': round,
            'min': min,
            'max': max,
            'sum': sum,
            'sorted': sorted,
            'reversed': reversed,
            'abs': abs,
            'divmod': divmod,
            'pow': pow,
            'print': print,
        }

        # Execute the code
        exec(code, context)

        # Check if code set a return value (via variable assignment)
        result = context.get('result', None)

        # Prepare response
        if result is not None:
            # Try to serialize result to ensure it's JSON-compatible
            try:
                json.dumps(result)
                return {
                    "success": True,
                    "result": result,
                    "description": description or "Code executed successfully with return value",
                    "slide_number": slide.SlideNumber,
                    "total_slides": presentation.Slides.Count
                }
            except (TypeError, ValueError) as e:
                return {
                    "success": True,
                    "result": str(result),
                    "description": description or "Code executed successfully (result converted to string)",
                    "slide_number": slide.SlideNumber,
                    "total_slides": presentation.Slides.Count,
                    "warning": f"Result was not JSON-serializable: {str(e)}"
                }
        else:
            return {
                "success": True,
                "message": "Code executed successfully (no return value)",
                "description": description,
                "slide_number": slide.SlideNumber,
                "total_slides": presentation.Slides.Count
            }

    except Exception as e:
        return {
            "error": f"Execution error: {type(e).__name__}: {str(e)}",
            "description": description
        }


def generate_mcp_response(result: dict) -> str:
    """Generate formatted MCP response from evaluation result."""

    if result.get("error"):
        return f"Error: {result['error']}"

    response_parts = []

    if result.get("description"):
        response_parts.append(f"SUCCESS: {result['description']}")
    else:
        response_parts.append("SUCCESS: Code executed successfully")

    response_parts.append(f"Slide: {result['slide_number']} of {result['total_slides']}")

    if result.get("result") is not None:
        response_parts.append("\nReturned data:")
        response_parts.append(json.dumps(result['result'], indent=2))

    if result.get("warning"):
        response_parts.append(f"\nWARNING: {result['warning']}")

    return "\n".join(response_parts)