"""
PowerPoint template discovery and enumeration tool.
Discovers available PowerPoint templates from common locations and validates them via COM.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'python'))

import win32com.client
from pathlib import Path
import json
from datetime import datetime


def get_template_directories():
    """Get common PowerPoint template directories."""
    import os

    directories = []
    username = os.environ.get('USERNAME', '')

    # Personal templates directory (Office 365/2019+)
    personal_templates = Path(f"C:/Users/{username}/Documents/Custom Office Templates")
    if personal_templates.exists():
        directories.append(str(personal_templates))

    # User templates directory (AppData)
    user_templates = Path(f"C:/Users/{username}/AppData/Roaming/Microsoft/Templates")
    if user_templates.exists():
        directories.append(str(user_templates))

    # System templates - multiple possible locations
    system_locations = [
        "C:/Program Files/Microsoft Office/Templates",
        "C:/Program Files/Microsoft Office/root/Templates",
        "C:/Program Files (x86)/Microsoft Office/Templates",
        "C:/Program Files (x86)/Microsoft Office/root/Templates"
    ]

    for location in system_locations:
        if Path(location).exists():
            directories.append(location)

    return directories


def scan_template_directory(directory_path):
    """Scan a directory for PowerPoint template files."""
    template_extensions = {'.potx', '.potm', '.pot'}
    templates = []

    try:
        directory = Path(directory_path)
        print(f"  📁 Scanning: {directory}")

        # Recursively search for template files
        for file_path in directory.rglob('*'):
            if file_path.is_file() and file_path.suffix.lower() in template_extensions:
                try:
                    template_info = {
                        'name': file_path.stem,
                        'filename': file_path.name,
                        'path': str(file_path),
                        'extension': file_path.suffix,
                        'size_mb': round(file_path.stat().st_size / (1024 * 1024), 2),
                        'modified': datetime.fromtimestamp(file_path.stat().st_mtime).isoformat(),
                        'directory_type': get_directory_type(directory_path)
                    }
                    templates.append(template_info)
                    print(f"    ✅ Found: {file_path.name}")
                except Exception as e:
                    print(f"    ⚠️  Error reading {file_path.name}: {e}")
                    continue

        print(f"  📊 Found {len(templates)} templates in {directory}")

    except Exception as e:
        print(f"  ❌ Error scanning directory {directory_path}: {e}")

    return templates


def get_directory_type(directory_path):
    """Classify template directory type."""
    path_lower = directory_path.lower()

    if 'custom office templates' in path_lower:
        return 'personal'
    elif 'appdata' in path_lower and 'templates' in path_lower:
        return 'user'
    elif 'program files' in path_lower:
        return 'system'
    else:
        return 'other'



def powerpoint_list_templates(max_templates=50):
    """
    Discover and list available PowerPoint templates.

    Args:
        max_templates: Maximum number of templates to return (to avoid overwhelming results)

    Returns:
        Dictionary with template discovery results
    """
    try:
        print("🔍 Discovering PowerPoint templates...")

        # 1. Get template directories
        template_dirs = get_template_directories()
        print(f"📁 Found {len(template_dirs)} template directories to scan")

        # 2. Scan all directories for templates
        all_templates = []
        directory_stats = {}

        for directory in template_dirs:
            templates = scan_template_directory(directory)
            all_templates.extend(templates)
            directory_stats[directory] = len(templates)

        print(f"\n📊 Discovery Summary:")
        print(f"  Total templates found: {len(all_templates)}")
        for directory, count in directory_stats.items():
            print(f"  {directory}: {count} templates")

        # 3. Sort templates by type and name
        all_templates.sort(key=lambda t: (t['directory_type'], t['name'].lower()))

        # 4. Limit results if too many
        if len(all_templates) > max_templates:
            print(f"⚠️  Limiting results to first {max_templates} templates")
            all_templates = all_templates[:max_templates]

        # 5. Prepare results
        result = {
            'success': True,
            'total_found': len(all_templates),
            'directories_scanned': template_dirs,
            'directory_stats': directory_stats,
            'templates': all_templates,
            'timestamp': datetime.now().isoformat()
        }

        print(f"\n🎉 Template discovery completed!")
        print(f"  Found: {result['total_found']} templates")

        return result

    except Exception as e:
        print(f"❌ Error during template discovery: {str(e)}")
        return {
            'success': False,
            'error': f"Template discovery failed: {str(e)}",
            'templates': [],
            'timestamp': datetime.now().isoformat()
        }


def print_template_results(result):
    """Print formatted template discovery results."""
    if not result.get('success'):
        print(f"\n❌ TEMPLATE DISCOVERY FAILED: {result.get('error')}")
        return

    print(f"\n{'='*70}")
    print(f"📋 POWERPOINT TEMPLATE DISCOVERY RESULTS")
    print(f"{'='*70}")
    print(f"Total Templates Found: {result['total_found']}")
    print(f"Discovery Time: {result['timestamp']}")

    print(f"\n📁 DIRECTORIES SCANNED:")
    print(f"{'-'*70}")
    for directory, count in result['directory_stats'].items():
        print(f"  {directory} ({count} templates)")

    if result['templates']:
        print(f"\n📋 AVAILABLE TEMPLATES:")
        print(f"{'-'*70}")

        # Group by directory type
        by_type = {}
        for template in result['templates']:
            dir_type = template['directory_type']
            if dir_type not in by_type:
                by_type[dir_type] = []
            by_type[dir_type].append(template)

        for dir_type in ['personal', 'user', 'system', 'other']:
            if dir_type in by_type:
                print(f"\n  📂 {dir_type.upper()} TEMPLATES:")
                for template in by_type[dir_type]:
                    print(f"    • {template['name']}")
                    print(f"       Path: {template['path']}")
                    print(f"       Size: {template['size_mb']} MB, Modified: {template['modified'][:10]}")


def generate_mcp_response(result):
    """Generate the actual MCP tool response that would be returned to the LLM."""
    if not result.get('success'):
        return f"Template discovery failed: {result.get('error')}"

    # Group templates by type for organized presentation
    templates_by_type = {}
    for template in result['templates']:
        dir_type = template['directory_type'].title()
        if dir_type not in templates_by_type:
            templates_by_type[dir_type] = []
        templates_by_type[dir_type].append(template)

    summary_lines = []

    # Add template listings
    for dir_type in ['Personal', 'User', 'System', 'Other']:
        if dir_type in templates_by_type:
            templates = templates_by_type[dir_type]
            summary_lines.append(f"{dir_type} Templates:")

            for template in templates:
                summary_lines.append(f"  \"{template['name']}\"")

    # Add usage instructions
    summary_lines.extend([
        "",
        "Usage: analyze_template(source=\"template_name\") for any template listed above"
    ])

    return "\n".join(summary_lines)


if __name__ == "__main__":
    print("🚀 Testing PowerPoint Template Discovery")
    print("=" * 80)

    # Test template discovery
    result = powerpoint_list_templates(max_templates=20)

    # Print detailed results (for developer/debugging)
    print_template_results(result)

    # Also save results to JSON file for inspection
    output_file = Path(__file__).parent / "template_discovery_results.json"
    try:
        with open(output_file, 'w') as f:
            json.dump(result, f, indent=2)
        print(f"\n💾 Results saved to: {output_file}")
    except Exception as e:
        print(f"\n⚠️  Could not save results: {e}")

    # Show the actual MCP tool response that would be returned to LLM
    print(f"\n{'='*80}")
    print("🤖 MCP TOOL RESPONSE (What the LLM would receive):")
    print(f"{'='*80}")
    mcp_response = generate_mcp_response(result)
    print(mcp_response)
    print(f"{'='*80}")

    # Calculate response length for context awareness
    response_length = len(mcp_response)
    print(f"📊 MCP Response Stats: {response_length} characters, ~{response_length//4} tokens")