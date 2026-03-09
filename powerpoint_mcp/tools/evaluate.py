"""
PowerPoint code evaluation tool for MCP server.
Execute arbitrary Python code in PowerPoint automation context.
"""

import math
import json
from typing import Optional, Any

from ..backends import get_backend
from .skills import skills


def powerpoint_evaluate(
    code: str,
    slide_number: Optional[int] = None,
    shape_ref: Optional[str] = None,
    description: Optional[str] = None
) -> dict:
    """
    Execute arbitrary Python code in PowerPoint automation context.

    Args:
        code: Python code to execute.
        slide_number: Target slide (1-based). If None, uses current slide
        shape_ref: Optional shape ID/Name to operate on
        description: Human-readable description of operation intent

    Returns:
        Dictionary with success/error status and optional result data
    """
    try:
        backend = get_backend()
        backend.connect()

        feature_support = backend.get_feature_support()

        if not feature_support.raw_evaluate:
            return {
                "error": "The evaluate tool is not supported on this platform. "
                         "Use the dedicated MCP tools (populate_placeholder, add_animation, etc.) instead."
            }

        if not backend.get_presentation_count():
            return {"error": "No presentation is currently open"}

        # Get raw COM/platform context
        raw_context = backend.get_raw_context(slide_number, shape_ref)
        if not raw_context:
            return {"error": "Failed to get PowerPoint context"}

        ppt = raw_context.get('ppt')
        presentation = raw_context.get('presentation')
        slide = raw_context.get('slide')
        shape = raw_context.get('shape')
        np = raw_context.get('np')

        if slide_number is not None and slide is None:
            return {"error": f"Slide {slide_number} out of range"}

        if shape_ref and shape is None:
            return {"error": f"Shape '{shape_ref}' not found"}

        # Create execution context
        context = {
            'ppt': ppt,
            'presentation': presentation,
            'slide': slide,
            'shape': shape,
            'math': math,
            'np': np,
            'has_numpy': np is not None,
            'skills': skills,
            # Python builtins
            'range': range, 'len': len, 'str': str, 'int': int,
            'float': float, 'bool': bool, 'list': list, 'dict': dict,
            'tuple': tuple, 'set': set, 'enumerate': enumerate, 'zip': zip,
            'round': round, 'min': min, 'max': max, 'sum': sum,
            'sorted': sorted, 'reversed': reversed, 'abs': abs,
            'divmod': divmod, 'pow': pow, 'print': print,
        }

        # Execute the code
        exec(code, context)

        # Check for return value
        result = context.get('result', None)

        if result is not None:
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
