"""
PowerPoint template discovery tool for MCP server.
Discovers available PowerPoint templates from common locations.
"""

import os
from pathlib import Path
from datetime import datetime


def get_template_directories():
    """Get common PowerPoint template directories."""
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
                except Exception:
                    # Skip files that can't be read
                    continue

    except Exception:
        # Skip directories that can't be accessed
        pass

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


def generate_mcp_response(result):
    """Generate the MCP tool response for the LLM."""
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
                summary_lines.append(f"  \"{template['name']}\" - {template['path']}")

    # Add usage instructions
    summary_lines.extend([
        "",
        "Usage: analyze_template(source=\"template_name\") for any template listed above"
    ])

    return "\n".join(summary_lines)


def powerpoint_list_templates():
    """
    Discover and list available PowerPoint templates.

    Returns:
        Dictionary with template discovery results
    """
    try:
        # Get template directories
        template_dirs = get_template_directories()

        # Scan all directories for templates
        all_templates = []
        directory_stats = {}

        for directory in template_dirs:
            templates = scan_template_directory(directory)
            all_templates.extend(templates)
            directory_stats[directory] = len(templates)

        # Sort templates by type and name
        all_templates.sort(key=lambda t: (t['directory_type'], t['name'].lower()))

        # Prepare results
        result = {
            'success': True,
            'total_found': len(all_templates),
            'directories_scanned': template_dirs,
            'directory_stats': directory_stats,
            'templates': all_templates,
            'timestamp': datetime.now().isoformat()
        }

        return result

    except Exception as e:
        return {
            'success': False,
            'error': f"Template discovery failed: {str(e)}",
            'templates': [],
            'timestamp': datetime.now().isoformat()
        }