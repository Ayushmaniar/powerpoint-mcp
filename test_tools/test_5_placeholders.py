"""
Comprehensive test optimized for exactly 5 placeholders.
Each placeholder gets a different category of formatting tests.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'python'))

import win32com.client
from powerpoint_mcp.tools.populate_placeholder import powerpoint_populate_placeholder


def get_current_slide_placeholders():
    """Get all text placeholders on the current slide."""
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        active_presentation = ppt_app.ActivePresentation

        # Get current slide
        try:
            active_window = ppt_app.ActiveWindow
            if hasattr(active_window, 'View') and hasattr(active_window.View, 'Slide'):
                current_slide = active_window.View.Slide
            else:
                current_slide = active_presentation.Slides(1)
        except:
            current_slide = active_presentation.Slides(1)

        # Find text placeholders
        text_placeholders = []
        for shape in current_slide.Shapes:
            if hasattr(shape, 'TextFrame'):
                try:
                    text_range = shape.TextFrame.TextRange
                    text_placeholders.append(shape.Name)
                except:
                    pass

        return text_placeholders, current_slide.SlideIndex

    except Exception as e:
        print(f"❌ Failed to get placeholders: {e}")
        return [], None


def test_5_placeholder_scenarios():
    """Test comprehensive formatting across exactly 5 placeholders."""

    # Get available placeholders
    text_placeholders, slide_number = get_current_slide_placeholders()

    if len(text_placeholders) < 5:
        print(f"❌ Need at least 5 text placeholders, found only {len(text_placeholders)}")
        print(f"Available placeholders: {text_placeholders}")
        return

    print(f"🎯 Using 5 placeholders on slide {slide_number}:")
    for i, placeholder in enumerate(text_placeholders[:5], 1):
        print(f"   {i}. '{placeholder}'")

    # 5 comprehensive test scenarios - each tests different aspects
    test_scenarios = [
        {
            "placeholder": text_placeholders[0],
            "name": "BASIC FORMATTING TESTS",
            "content": """<b>Bold Text Test</b> | <i>Italic Text Test</i> | <u>Underlined Text Test</u>

Mixed formatting: <b>Bold</b>, <i>Italic</i>, <u>Underline</u> in one sentence.

Special combo: <b><i>Bold AND Italic</i></b> together.""",
            "expect": "Bold, italic, underline formatting. Bold+italic combination."
        },
        {
            "placeholder": text_placeholders[1],
            "name": "COLOR FORMATTING TESTS",
            "content": """Single colors: <red>Red Text</red> | <blue>Blue Text</blue> | <green>Green Text</green>
More colors: <orange>Orange</orange> | <purple>Purple</purple> | <yellow>Yellow</yellow>
CRITICAL TEST - Nested formatting:
<red><b>Bold Red Text</b></red> | <blue><i>Italic Blue Text</i></blue> | <green><u>Underlined Green</u></green>
Triple combo: <red><b><u>Bold Red Underlined</u></b></red>""",
            "expect": "All colors work. Nested formatting: red+bold, blue+italic, green+underline."
        },
        {
            "placeholder": text_placeholders[2],
            "name": "LIST FORMATTING TESTS",
            "content": """Simple bullet list:
<ul>
<li>First bullet point</li>
<li>Second bullet point</li>
<li>Third bullet point</li>
</ul>

Formatted bullet list:
<ul>
<li><b>Bold bullet</b></li>
<li><i>Italic bullet</i></li>
<li><red>Red bullet</red></li>
<li><blue><b>Blue bold bullet</b></blue></li>
</ul>

Simple numbered list:
<ol>
<li>First numbered item</li>
<li>Second numbered item</li>
<li>Third numbered item</li>
</ol>

Formatted numbered list:
<ol>
<li><green>Green numbered item</green></li>
<li><purple><i>Purple italic item</i></purple></li>
</ol>""",
            "expect": "Bullets (•) and numbers (1,2,3). Formatted list items work."
        },
        {
            "placeholder": text_placeholders[3],
            "name": "LINE BREAKS & MIXED CONTENT",
            "content": """Line break test:
First line with <b>bold</b>
Second line with <red>red text</red>

Empty line above this paragraph.

Complex mixed content:
Introduction with <green>green text</green>.

Key points:
<ul>
<li>Point with <b>bold emphasis</b></li>
<li>Point with <red>red highlighting</red></li>
</ul>

Steps to follow:
<ol>
<li>First do <i>this italic step</i></li>
<li>Then do <blue><b>this blue bold step</b></blue></li>
</ol>

Conclusion: <purple><b>Purple bold summary</b></purple>.""",
            "expect": "Line breaks work. Mixed lists with text. Paragraph structure."
        },
        {
            "placeholder": text_placeholders[4],
            "name": "REAL-WORLD BUSINESS SCENARIO",
            "content": """<b>Q4 2024 Business Review</b>

<green><b>✅ ACHIEVEMENTS</b></green>
<ul>
<li>Revenue: <green><b>$2.5M (+25% YoY)</b></green></li>
<li>Customers: <blue><b>1,200 new clients</b></blue></li>
<li>Profit: <red><b>22% margin (↑3%)</b></red></li>
</ul>

<orange><b>🎯 Q1 2025 GOALS</b></orange>
<ol>
<li>Expand to <b>European markets</b></li>
<li>Launch <i>AI-powered features</i></li>
<li>Achieve <u>$3M revenue target</u></li>
<li>Hire <red><b>50 new employees</b></red></li>
</ol>

<purple><b>STATUS: ON TRACK FOR SUCCESS! 🚀</b></purple>

Critical metrics to watch:
• <green>Customer satisfaction: 98%</green>
• <blue>Team productivity: +40%</blue>
• <red>Market share: Leading position</red>""",
            "expect": "Business presentation with emojis, nested formatting, mixed content."
        }
    ]

    print(f"\n🚀 Starting 5-placeholder comprehensive test...")
    print(f"Each placeholder tests different formatting aspects.")

    results = []

    for i, scenario in enumerate(test_scenarios):
        print(f"\n{'='*80}")
        print(f"📝 PLACEHOLDER {i+1}: {scenario['name']}")
        print(f"🎯 Target: '{scenario['placeholder']}'")
        print(f"📄 Expected: {scenario['expect']}")
        print(f"📊 Content length: {len(scenario['content'])} characters")

        # Show content preview
        lines = scenario['content'].split('\n')
        print(f"📋 Content preview ({len(lines)} lines):")
        for j, line in enumerate(lines[:5]):  # Show first 5 lines
            print(f"   {j+1}. {line[:60]}{'...' if len(line) > 60 else ''}")
        if len(lines) > 5:
            print(f"   ... and {len(lines) - 5} more lines")

        # Call the populate function
        result = powerpoint_populate_placeholder(
            placeholder_name=scenario['placeholder'],
            content=scenario['content'],
            content_type="auto"
        )

        success = result.get('success', False)
        print(f"\n🔄 RESULT: {'✅ SUCCESS' if success else '❌ FAILED'}")

        if success:
            print(f"📊 Content Type: {result.get('content_type', 'unknown')}")
            if result.get('content_type') == 'formatted_text':
                segments = result.get('format_segments_applied', 0)
                print(f"🎨 Format Segments Applied: {segments}")

                plain_text = result.get('plain_text', '')
                print(f"📝 Plain Text Length: {len(plain_text)} characters")
                print(f"📑 Lines: {plain_text.count('\\n') + 1}")

                # Show plain text preview
                preview = plain_text.replace('\n', '\\n')[:100]
                print(f"📝 Plain Text Preview: '{preview}{'...' if len(plain_text) > 100 else ''}'")

        else:
            print(f"❌ Error: {result.get('error', 'Unknown')}")

        results.append({
            'scenario': scenario['name'],
            'placeholder': scenario['placeholder'],
            'success': success,
            'result': result,
            'expected': scenario['expect']
        })

        # Detailed visual check instructions
        print(f"\n👀 VISUAL CHECK FOR PLACEHOLDER {i+1}:")
        print(f"   🎯 Look at '{scenario['placeholder']}' in PowerPoint")
        print(f"   📋 Should see: {scenario['expect']}")
        if success and result.get('format_segments_applied', 0) > 0:
            print(f"   🎨 Should have {result['format_segments_applied']} different formatting styles")
        else:
            print(f"   ⚠️  No formatting applied - check for issues")

        input(f"Press Enter after checking placeholder {i+1}...")

    # Comprehensive summary
    print(f"\n{'='*80}")
    print("📊 FINAL COMPREHENSIVE TEST SUMMARY")
    print(f"{'='*80}")

    total_tests = len(results)
    passed_tests = len([r for r in results if r['success']])
    total_chars = sum(len(s['content']) for s in test_scenarios)
    total_segments = sum(r['result'].get('format_segments_applied', 0) for r in results if r['success'])

    print(f"🎯 Placeholders tested: {total_tests}")
    print(f"✅ Successful: {passed_tests}")
    print(f"📊 Success rate: {(passed_tests/total_tests*100):.1f}%")
    print(f"📝 Total characters: {total_chars:,}")
    print(f"🎨 Total format segments: {total_segments}")
    print()

    print("📋 DETAILED RESULTS:")
    for i, result in enumerate(results, 1):
        status = "✅ PASSED" if result['success'] else "❌ FAILED"
        segments = result['result'].get('format_segments_applied', 0) if result['success'] else 0
        print(f"   {status} Placeholder {i}: {result['scenario']} ({segments} segments)")
        print(f"      Expected: {result['expected']}")

    print(f"\n💡 CRITICAL TESTS TO VERIFY VISUALLY:")
    print(f"   🔴 Placeholder 2: <red><b>Bold Red Text</b></red> should be BOTH red AND bold")
    print(f"   🔵 Placeholder 2: <blue><i>Italic Blue Text</i></blue> should be BOTH blue AND italic")
    print(f"   🟢 Placeholder 2: <green><u>Underlined Green</u></green> should be BOTH green AND underlined")
    print(f"   🟣 Placeholder 2: <red><b><u>Bold Red Underlined</u></b></red> should be ALL THREE")
    print(f"   📋 Placeholder 3: Lists should have bullets (•) and numbers (1,2,3)")
    print(f"   📄 Placeholder 4: Line breaks should create new lines")

    if passed_tests == total_tests:
        print(f"\n🎉 PERFECT! All formatting works correctly!")
        print(f"   🚀 The populate_placeholder tool is production ready!")
    elif passed_tests >= 4:
        print(f"\n✨ EXCELLENT! {passed_tests}/{total_tests} placeholders work")
        print(f"   Minor issues with {total_tests - passed_tests} placeholder(s)")
    elif passed_tests >= 3:
        print(f"\n⚠️  GOOD: {passed_tests}/{total_tests} placeholders work")
        print(f"   Need to fix {total_tests - passed_tests} placeholder(s)")
    else:
        print(f"\n❌ NEEDS WORK: Only {passed_tests}/{total_tests} placeholders work")
        print(f"   Significant formatting issues need fixing")

    print(f"\n🔍 If nested formatting (red+bold) doesn't work, the HTML parser needs more fixes.")


if __name__ == "__main__":
    print("🚀 COMPREHENSIVE 5-PLACEHOLDER TEST")
    print("=" * 60)
    print("💡 This test covers ALL formatting scenarios across 5 placeholders:")
    print("   1️⃣  Basic formatting (bold, italic, underline)")
    print("   2️⃣  Colors + nested formatting (RED+BOLD combinations)")
    print("   3️⃣  Lists (bullets, numbers, formatted items)")
    print("   4️⃣  Line breaks + mixed content")
    print("   5️⃣  Real business presentation content")
    print()
    print("🎯 Critical test: Nested formatting like <red><b>Bold Red</b></red>")
    print()

    input("Press Enter when you're on a slide with 5+ text placeholders...")
    test_5_placeholder_scenarios()