"""
Test script for matplotlib plot rendering in populate_placeholder tool.

This script tests the new "plot" content_type functionality that allows LLM-written
matplotlib code to be rendered and inserted into PowerPoint placeholders.
"""

import sys
import os

# Add parent directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'python'))

from powerpoint_mcp.tools.populate_placeholder import powerpoint_populate_placeholder
from powerpoint_mcp.tools.presentation import manage_presentation
from powerpoint_mcp.tools.add_slide_with_layout import powerpoint_add_slide_with_layout
import win32com.client


def test_educational_quadratic():
    """Test an educational quadratic equation plot with roots and LaTeX annotations."""
    print("\n=== Test 1: Educational Quadratic Equation (with roots and LaTeX) ===")

    matplotlib_code = """
import numpy as np

# Define the quadratic equation: y = x^2 - 4x + 3
# This factors to: y = (x-1)(x-3), so roots are at x=1 and x=3
a, b, c = 1, -4, 3

# Create x values
x = np.linspace(-1, 5, 300)
y = a * x**2 + b * x + c

# Calculate roots using quadratic formula
discriminant = b**2 - 4*a*c
root1 = (-b + np.sqrt(discriminant)) / (2*a)
root2 = (-b - np.sqrt(discriminant)) / (2*a)

# Calculate vertex (minimum point)
vertex_x = -b / (2*a)
vertex_y = a * vertex_x**2 + b * vertex_x + c

# Create the plot
fig, ax = plt.subplots(figsize=(12, 9))

# Plot the parabola with a gradient color
ax.plot(x, y, 'b-', linewidth=3, label=r'$f(x) = x^2 - 4x + 3$', zorder=3)

# Mark the roots with red points
ax.plot([root1, root2], [0, 0], 'ro', markersize=15, label='Roots', zorder=5)

# Mark the vertex with a green point
ax.plot(vertex_x, vertex_y, 'go', markersize=12, label=f'Vertex ({vertex_x:.1f}, {vertex_y:.1f})', zorder=5)

# Add grid
ax.grid(True, alpha=0.4, linestyle='--', linewidth=1)
ax.axhline(y=0, color='k', linewidth=1.5, alpha=0.7)
ax.axvline(x=0, color='k', linewidth=1.5, alpha=0.7)

# Add LaTeX annotations for roots
ax.annotate(f'Root 1\\n$x_1 = {root2:.0f}$',
            xy=(root2, 0), xytext=(root2 - 0.8, -1.5),
            fontsize=14, color='red', weight='bold',
            bbox=dict(boxstyle='round,pad=0.5', facecolor='yellow', alpha=0.7),
            arrowprops=dict(arrowstyle='->', color='red', lw=2))

ax.annotate(f'Root 2\\n$x_2 = {root1:.0f}$',
            xy=(root1, 0), xytext=(root1 + 0.5, -1.5),
            fontsize=14, color='red', weight='bold',
            bbox=dict(boxstyle='round,pad=0.5', facecolor='yellow', alpha=0.7),
            arrowprops=dict(arrowstyle='->', color='red', lw=2))

# Add LaTeX annotation for the vertex
ax.annotate(f'Vertex (Minimum)\\n$({vertex_x:.0f}, {vertex_y:.0f})$',
            xy=(vertex_x, vertex_y), xytext=(vertex_x - 1.5, vertex_y - 2),
            fontsize=14, color='green', weight='bold',
            bbox=dict(boxstyle='round,pad=0.5', facecolor='lightgreen', alpha=0.7),
            arrowprops=dict(arrowstyle='->', color='green', lw=2))

# Add the quadratic formula as a text box
formula_text = r'$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$'
ax.text(0.02, 0.98, f'Quadratic Formula:\\n{formula_text}\\n\\nWhere: $a={a}$, $b={b}$, $c={c}$',
        transform=ax.transAxes, fontsize=13,
        verticalalignment='top',
        bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8, pad=1))

# Add discriminant information
discriminant_text = f'Discriminant: $\\Delta = b^2 - 4ac = {discriminant:.0f}$\\n$\\Delta > 0$ → Two real roots'
ax.text(0.98, 0.98, discriminant_text,
        transform=ax.transAxes, fontsize=12,
        verticalalignment='top', horizontalalignment='right',
        bbox=dict(boxstyle='round', facecolor='lightblue', alpha=0.8, pad=1))

# Labels and title
ax.set_xlabel('x', fontsize=16, weight='bold')
ax.set_ylabel('f(x)', fontsize=16, weight='bold')
ax.set_title('Solving Quadratic Equations: $f(x) = x^2 - 4x + 3$',
             fontsize=18, weight='bold', pad=20)

# Set axis limits for better visualization
ax.set_xlim(-1, 5)
ax.set_ylim(-3, 8)

# Add legend
ax.legend(loc='upper right', fontsize=12, framealpha=0.9)

# Add minor gridlines for better precision
ax.minorticks_on()
ax.grid(which='minor', alpha=0.2, linestyle=':', linewidth=0.5)

plt.tight_layout()
"""

    result = powerpoint_populate_placeholder(
        placeholder_name="Text Placeholder 2",
        content=matplotlib_code,
        content_type="plot",
        slide_number=1
    )

    print(f"Result: {result}")
    return result


def test_bar_chart():
    """Test a bar chart."""
    print("\n=== Test 2: Bar Chart ===")

    matplotlib_code = """
import numpy as np

categories = ['Category A', 'Category B', 'Category C', 'Category D', 'Category E']
values = [23, 45, 56, 78, 32]

plt.figure(figsize=(10, 6))
bars = plt.bar(categories, values, color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#FFA07A', '#98D8C8'])
plt.xlabel('Categories', fontsize=12)
plt.ylabel('Values', fontsize=12)
plt.title('Sales by Category', fontsize=14, fontweight='bold')
plt.grid(axis='y', alpha=0.3)

# Add value labels on bars
for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2., height,
             f'{height}',
             ha='center', va='bottom', fontsize=10)
"""

    result = powerpoint_populate_placeholder(
        placeholder_name="Text Placeholder 2",
        content=matplotlib_code,
        content_type="plot",
        slide_number=2
    )

    print(f"Result: {result}")
    return result


def test_scatter_plot():
    """Test a scatter plot with multiple datasets."""
    print("\n=== Test 3: Scatter Plot ===")

    matplotlib_code = """
import numpy as np

np.random.seed(42)
x1 = np.random.randn(50)
y1 = np.random.randn(50)
x2 = np.random.randn(50) + 2
y2 = np.random.randn(50) + 2

plt.figure(figsize=(10, 6))
plt.scatter(x1, y1, alpha=0.6, s=100, c='blue', label='Dataset 1')
plt.scatter(x2, y2, alpha=0.6, s=100, c='red', label='Dataset 2')
plt.xlabel('X Values', fontsize=12)
plt.ylabel('Y Values', fontsize=12)
plt.title('Scatter Plot Comparison', fontsize=14, fontweight='bold')
plt.legend()
plt.grid(True, alpha=0.3)
"""

    result = powerpoint_populate_placeholder(
        placeholder_name="Text Placeholder 2",
        content=matplotlib_code,
        content_type="plot",
        slide_number=3
    )

    print(f"Result: {result}")
    return result


def test_pie_chart():
    """Test a pie chart."""
    print("\n=== Test 4: Pie Chart ===")

    matplotlib_code = """
sizes = [30, 25, 20, 15, 10]
labels = ['Product A', 'Product B', 'Product C', 'Product D', 'Product E']
colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99', '#ff99cc']
explode = (0.1, 0, 0, 0, 0)  # explode first slice

plt.figure(figsize=(10, 8))
plt.pie(sizes, explode=explode, labels=labels, colors=colors,
        autopct='%1.1f%%', shadow=True, startangle=90)
plt.title('Market Share Distribution', fontsize=14, fontweight='bold')
plt.axis('equal')
"""

    result = powerpoint_populate_placeholder(
        placeholder_name="Text Placeholder 2",
        content=matplotlib_code,
        content_type="plot",
        slide_number=4
    )

    print(f"Result: {result}")
    return result


def test_complex_subplot():
    """Test a complex plot with subplots."""
    print("\n=== Test 5: Complex Subplot ===")

    matplotlib_code = """
import numpy as np

x = np.linspace(0, 10, 100)

fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 10))

# Subplot 1: Sine wave
ax1.plot(x, np.sin(x), 'b-', linewidth=2)
ax1.set_title('Sine Wave', fontweight='bold')
ax1.grid(True, alpha=0.3)

# Subplot 2: Cosine wave
ax2.plot(x, np.cos(x), 'r-', linewidth=2)
ax2.set_title('Cosine Wave', fontweight='bold')
ax2.grid(True, alpha=0.3)

# Subplot 3: Exponential
ax3.plot(x, np.exp(-x/3), 'g-', linewidth=2)
ax3.set_title('Exponential Decay', fontweight='bold')
ax3.grid(True, alpha=0.3)

# Subplot 4: Combination
ax4.plot(x, np.sin(x) * np.exp(-x/5), 'm-', linewidth=2)
ax4.set_title('Damped Oscillation', fontweight='bold')
ax4.grid(True, alpha=0.3)

plt.tight_layout()
"""

    result = powerpoint_populate_placeholder(
        placeholder_name="Text Placeholder 2",
        content=matplotlib_code,
        content_type="plot",
        slide_number=5
    )

    print(f"Result: {result}")
    return result


def setup_test_presentation():
    """Setup: Create a new PowerPoint presentation with slides for testing."""
    print("\n" + "=" * 80)
    print("SETUP: Creating Test Presentation")
    print("=" * 80)

    # Create new presentation
    print("\nCreating new presentation at: C:\\Users\\Rajat\\test_matplotlib_plots.pptx")
    manage_presentation("create", file_path=r"C:\Users\Rajat\test_matplotlib_plots.pptx")

    # Activate PowerPoint window
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        ppt.Activate()
        presentation = ppt.ActivePresentation

        # Add 5 slides with "Title and Content" layout (has Picture Placeholder 2)
        print("Adding 5 slides with placeholders...")
        for i in range(5):
            # Layout 2 is typically "Title and Content" which has a content placeholder
            # that can accept both text and pictures
            presentation.Slides.Add(Index=i+1, Layout=2)  # ppLayoutText = 2

        # Activate the first slide
        if presentation.Windows.Count > 0:
            presentation.Windows(1).Activate()

        print(f"Created {presentation.Slides.Count} slides")
        print("\nSlides created with 'Title and Content' layout")
        print("Each slide has placeholders: 'Title 1' and 'Text Placeholder 2'")

    except Exception as e:
        print(f"Error setting up presentation: {e}")
        raise

    input("\nPress ENTER to start running tests...")


def cleanup():
    """Cleanup: Save and close the test presentation."""
    print("\n" + "=" * 80)
    print("CLEANUP")
    print("=" * 80)

    try:
        # Save the presentation
        print("Saving presentation...")
        manage_presentation("save")

        print("Test presentation saved at: C:\\Users\\Rajat\\test_matplotlib_plots.pptx")
        print("You can open it to view all the generated plots!")

        # Ask user if they want to close
        close_choice = input("\nDo you want to close the presentation? (y/n): ").lower()
        if close_choice == 'y':
            manage_presentation("close", presentation_name="test_matplotlib_plots.pptx")
            print("Presentation closed")
        else:
            print("Presentation left open for inspection")

    except Exception as e:
        print(f"Error during cleanup: {e}")


def main():
    """Run all tests."""
    print("=" * 80)
    print("Testing Matplotlib Plot Rendering in populate_placeholder")
    print("=" * 80)

    tests = [
        ("Educational Quadratic Equation", test_educational_quadratic, "Text Placeholder 2", 1),
        ("Bar Chart", test_bar_chart, "Text Placeholder 2", 2),
        ("Scatter Plot", test_scatter_plot, "Text Placeholder 2", 3),
        ("Pie Chart", test_pie_chart, "Text Placeholder 2", 4),
        ("Complex Subplot", test_complex_subplot, "Text Placeholder 2", 5)
    ]

    try:
        setup_test_presentation()

        results = []
        for test_name, test_func, placeholder, slide_num in tests:
            print("\n" + "=" * 80)
            print(f"Running: {test_name} (Slide {slide_num})")
            print("=" * 80)

            try:
                # Update the test function call to use correct placeholder and slide
                result = test_func()
                results.append((test_name, result))

                if result.get("success"):
                    print(f"✅ SUCCESS: {test_name} completed")
                else:
                    print(f"❌ FAILED: {test_name}")
                    print(f"   Error: {result.get('error', 'Unknown error')}")

                # Wait for user to see the result before continuing
                input(f"\nPlot generated on slide {slide_num}. Press ENTER to continue to next test...")

            except Exception as e:
                print(f"❌ ERROR in {test_name}: {e}")
                results.append((test_name, {"error": str(e)}))
                input("\nPress ENTER to continue despite error...")

        # Summary
        print("\n" + "=" * 80)
        print("Test Summary")
        print("=" * 80)

        for test_name, result in results:
            status = "✅ PASSED" if result.get("success") else "❌ FAILED"
            print(f"{status} - {test_name}")
            if not result.get("success"):
                print(f"  Error: {result.get('error', 'Unknown error')}")

        success_count = sum(1 for _, r in results if r.get("success"))
        print(f"\nTotal: {success_count}/{len(results)} tests passed")

        if success_count == len(results):
            print("\n🎉 ALL TESTS PASSED!")
        else:
            print(f"\n⚠️  {len(results) - success_count} test(s) failed")

    except Exception as e:
        print(f"\n❌ FATAL ERROR: {e}")
        import traceback
        traceback.print_exc()

    finally:
        cleanup()


if __name__ == "__main__":
    main()
