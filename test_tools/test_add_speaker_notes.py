"""
Minimal test for add_speaker_notes functionality.
Tests adding "Hello World" to the current active slide's speaker notes.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'python'))

from powerpoint_mcp.tools.add_speaker_notes import powerpoint_add_speaker_notes


def test_add_hello_world():
    """Test adding Hello World to current active slide's speaker notes."""

    print("Testing with current active slide (auto-detect)...")
    result = powerpoint_add_speaker_notes(slide_number=None, notes_text="Hello World")

    print("Test Result:")
    print(f"Success: {result.get('success', False)}")

    if result.get('success'):
        print(f"Added to slide: {result.get('slide_number')}")
        print(f"Notes length: {result.get('notes_length')}")
        print(f"Message: {result.get('message')}")
    else:
        print(f"Error: {result.get('error')}")

    return result


def test_slide_1_specifically():
    """Test adding notes specifically to slide 1 (the problematic slide)."""

    print("Testing specifically with slide 1...")
    result = powerpoint_add_speaker_notes(slide_number=1, notes_text="Hello World - Slide 1 Test")

    print("Slide 1 Test Result:")
    print(f"Success: {result.get('success', False)}")

    if result.get('success'):
        print(f"Added to slide: {result.get('slide_number')}")
        print(f"Notes length: {result.get('notes_length')}")
        print(f"Message: {result.get('message')}")
    else:
        print(f"Error: {result.get('error')}")

    return result


if __name__ == "__main__":
    print("Testing add_speaker_notes functionality...\n")

    # Test 1: Auto-detect current slide
    result1 = test_add_hello_world()
    print()

    # Test 2: Specifically test slide 1 (the problematic one)
    result2 = test_slide_1_specifically()
    print()

    # Summary
    if result1.get('success') and result2.get('success'):
        print("✅ ALL TESTS PASSED - Speaker notes functionality working!")
    elif result1.get('success') or result2.get('success'):
        print("⚠️  PARTIAL SUCCESS - Some tests passed")
        if not result1.get('success'):
            print(f"   Auto-detect failed: {result1.get('error')}")
        if not result2.get('success'):
            print(f"   Slide 1 failed: {result2.get('error')}")
    else:
        print("❌ ALL TESTS FAILED")
        print(f"   Auto-detect: {result1.get('error')}")
        print(f"   Slide 1: {result2.get('error')}")