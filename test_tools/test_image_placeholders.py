"""
Utility to exercise the image branch of the populate_placeholder tool.

Run this while a presentation is open in PowerPoint. The script targets the
current slide, finds any placeholders that can reasonably host an image, and
invokes powerpoint_populate_placeholder with a sample PNG that ships with the
repo. Results (and any scaling behaviour) are printed to stdout so you can
inspect the inserted shapes.

⚠️ The tool replaces the placeholder with an actual picture shape. Use a
scratch slide if you want to keep the original placeholders intact.
"""

import os
import sys
from pathlib import Path
from typing import Iterable, Tuple, Optional

import win32com.client


sys.path.append(os.path.join(os.path.dirname(__file__), "..", "python"))

from powerpoint_mcp.tools.populate_placeholder import powerpoint_populate_placeholder, find_shape_by_name


PLACEHOLDER_TYPES = {
    1: "Title",
    2: "Body/Content",
    3: "Center Title",
    4: "Subtitle",
    5: "Vertical Title",
    6: "Vertical Body",
    7: "Table",
    8: "Chart",
    9: "SmartArt/Diagram",
    10: "Media",
    11: "Org Chart",
    14: "Slide Number",
    18: "Picture",
}

PLACEHOLDER_CONTAINED_TYPES = {
    1: "Title",
    2: "Text",
    3: "Table",
    4: "Chart",
    5: "SmartArt",
    6: "Media",
    7: "ClipArt",
    8: "Diagram",
    9: "OrgChart",
    10: "MediaPlaceholder",
    11: "VerticalText",
    12: "Picture",
}

IMAGE_CAPABLE_TYPES = {2, 10, 12, 18}  # Content, Media, Picture


def describe_placeholder(value: Optional[int], mapping: dict) -> str:
    if value is None:
        return "None"
    return mapping.get(int(value), f"Unknown ({value})")


def connect_to_powerpoint():
    try:
        return win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        return win32com.client.Dispatch("PowerPoint.Application")


def get_active_slide(ppt_app):
    presentation = ppt_app.ActivePresentation
    if not presentation:
        raise RuntimeError("No presentation is currently open in PowerPoint")

    try:
        active_window = ppt_app.ActiveWindow
        if hasattr(active_window, "View") and hasattr(active_window.View, "Slide"):
            return active_window.View.Slide
    except Exception:
        pass

    # Fallback to the first slide
    return presentation.Slides(1)


def list_image_placeholders(slide) -> Iterable[Tuple[int, str, Optional[int], Optional[int], float, float]]:
    for shape in slide.Shapes:
        placeholder_type = None
        contained_type = None
        if hasattr(shape, "PlaceholderFormat"):
            try:
                placeholder_format = shape.PlaceholderFormat
                placeholder_type = getattr(placeholder_format, "Type", None)
                contained_type = getattr(placeholder_format, "ContainedType", None)
            except Exception:
                placeholder_type = None

        numeric_placeholder = int(placeholder_type) if placeholder_type is not None else None
        numeric_contained = int(contained_type) if contained_type is not None else None

        name_lower = shape.Name.lower()
        is_image_capable = False

        if numeric_placeholder in IMAGE_CAPABLE_TYPES or numeric_contained in IMAGE_CAPABLE_TYPES:
            is_image_capable = True
        elif "content placeholder" in name_lower:
            is_image_capable = True
        elif numeric_placeholder in {3, 7}:  # tables, charts still allow image insertion via AddPicture
            is_image_capable = any(keyword in name_lower for keyword in ["content", "picture"])

        if is_image_capable:
            yield shape.Id, shape.Name, numeric_placeholder, numeric_contained, shape.Width, shape.Height


def snapshot_placeholders(slide, label: str):
    print(f"\n📋 Placeholder snapshot: {label}")
    rows = []
    for shape in slide.Shapes:
        try:
            placeholder_format = getattr(shape, "PlaceholderFormat", None)
            placeholder_type = getattr(placeholder_format, "Type", None) if placeholder_format else None
            contained_type = getattr(placeholder_format, "ContainedType", None) if placeholder_format else None
        except Exception:
            placeholder_type = None
            contained_type = None

        rows.append(
            (
                shape.Id,
                shape.Name,
                describe_placeholder(placeholder_type, PLACEHOLDER_TYPES),
                describe_placeholder(contained_type, PLACEHOLDER_CONTAINED_TYPES),
                shape.Width,
                shape.Height,
            )
        )

    if not rows:
        print("   (no shapes on slide)")
        return

    for shape_id, name, p_type, c_type, width, height in rows:
        print(
            f"   • Id {shape_id}: '{name}'"
            f" → Type: {p_type}; Contained: {c_type}; Size: {width:.1f} x {height:.1f}"
        )


def run_image_tests():
    ppt_app = connect_to_powerpoint()

    if not ppt_app.Presentations.Count:
        print("❌ Open a presentation in PowerPoint before running this script")
        return

    slide = get_active_slide(ppt_app)

    snapshot_placeholders(slide, "initial state")

    placeholders = list(list_image_placeholders(slide))

    if not placeholders:
        # Check if specific targets were provided via env var (helps debug)
        target_name = os.environ.get("IMAGE_PLACEHOLDER_NAME")
        if target_name:
            print(f"🔍 Target placeholder hint: {target_name}")
            alt = powerpoint_populate_placeholder(
                placeholder_name=target_name,
                content=str(Path(__file__).parent / "OpenAI-Logo-2022.png"),
                content_type="image",
                slide_number=slide.SlideIndex,
            )
            print(f"   → Direct populate attempt: {alt}")

    if not placeholders:
        print("⚠️ No image-capable placeholders found on the active slide.")
        snapshot_placeholders(slide, "no image-capable placeholders")

        shapes = []
        for shape in slide.Shapes:
            placeholder_type = None
            contained_type = None
            if hasattr(shape, "PlaceholderFormat"):
                try:
                    placeholder_format = shape.PlaceholderFormat
                    placeholder_type = getattr(placeholder_format, "Type", None)
                    contained_type = getattr(placeholder_format, "ContainedType", None)
                except Exception:
                    pass

            shapes.append(
                (
                    shape.Name,
                    describe_placeholder(placeholder_type, PLACEHOLDER_TYPES),
                    describe_placeholder(contained_type, PLACEHOLDER_CONTAINED_TYPES),
                )
            )

        if shapes:
            print("   Available placeholders on this slide:")
            for name, p_type, c_type in shapes:
                print(f"     • {name} → Type: {p_type}; Contained: {c_type}")
        else:
            print("   No shapes detected on this slide.")

        print("   Try selecting a slide with a content or picture placeholder.")
        return

    print(f"🎯 Testing slide {slide.SlideIndex} with {len(placeholders)} candidate placeholders:")
    for shape_id, name, ph_type, contained_type, width, height in placeholders:
        type_label = describe_placeholder(ph_type, PLACEHOLDER_TYPES)
        contained_label = describe_placeholder(contained_type, PLACEHOLDER_CONTAINED_TYPES)
        print(f"   • {name} (Id {shape_id}) → Type: {type_label}; Contained: {contained_label} ({width:.1f} x {height:.1f})")

    sample_image = Path(__file__).parent / "OpenAI-Logo-2022.png"
    if not sample_image.exists():
        print(f"❌ Sample image missing: {sample_image}")
        return

    for shape_id, original_name, ph_type, contained_type, width, height in placeholders:
        try:
            current_shape = slide.Shapes.FindById(shape_id)
            placeholder_name = current_shape.Name
        except Exception:
            current_shape = None
            placeholder_name = original_name

        if placeholder_name != original_name:
            print(f"\n🧪 Populating '{original_name}' (currently '{placeholder_name}') with sample image…")
        else:
            print(f"\n🧪 Populating '{placeholder_name}' with sample image…")

        result = powerpoint_populate_placeholder(
            placeholder_name=placeholder_name,
            content=str(sample_image),
            content_type="image",
            slide_number=slide.SlideIndex,
        )

        if not result.get("success"):
            print(f"   ❌ Failed: {result.get('error')}")
            snapshot_placeholders(slide, f"after failed attempt on '{placeholder_name}'")
            continue

        new_shape_id = result.get("new_shape_id")
        try:
            new_shape = slide.Shapes.FindById(new_shape_id)
        except Exception:
            new_shape = None

        initial_name = result.get("placeholder_renamed_from", placeholder_name)
        final_name = result.get("new_shape_name", placeholder_name)

        if initial_name and final_name:
            print(f"   🛈 Placeholder mapping: '{initial_name}' → '{final_name}'")

        if new_shape:
            new_width = new_shape.Width
            new_height = new_shape.Height
            aspect = new_width / new_height if new_height else float("inf")
            print(
                f"   ✅ Inserted image shape {new_shape_id}"
                f" → size {new_width:.1f} x {new_height:.1f}"
            )
            print(f"   ℹ️ Aspect ratio now {aspect:.3f}; placeholder was {width/height:.3f} (before).")
        else:
            # Placeholder Delete could have renumbered the second placeholder; re-scan by name
            try:
                replacement = find_shape_by_name(slide, name)
            except Exception:
                replacement = None

            if replacement:
                print(f"   ❇️ Found shape named '{name}' after insertion → {replacement.Width:.1f} x {replacement.Height:.1f}")
            else:
                print("   ⚠️ Inserted image, but placeholder name now resolves differently (likely deleted).")

        snapshot_placeholders(slide, f"after populating '{placeholder_name}'")


if __name__ == "__main__":
    run_image_tests()

