"""
PowerPoint backend factory with auto-detection.

Provides get_backend() which returns the appropriate backend for the current platform.
"""

import sys
from typing import Optional

from .base import PowerPointBackend
from .types import UnsupportedFeatureError

_backend_instance: Optional[PowerPointBackend] = None


def get_backend() -> PowerPointBackend:
    """Get the PowerPoint backend for the current platform.

    Uses singleton pattern. Override with POWERPOINT_MCP_BACKEND env var.
    """
    global _backend_instance
    if _backend_instance is not None:
        return _backend_instance

    import os
    override = os.environ.get("POWERPOINT_MCP_BACKEND", "").lower()

    if override == "windows":
        from .windows import WindowsBackend
        _backend_instance = WindowsBackend()
    elif override == "macos":
        from .macos import MacOSBackend
        _backend_instance = MacOSBackend()
    elif override:
        raise RuntimeError(f"Unknown backend override: '{override}'. Supported: 'windows', 'macos'")
    elif sys.platform == "win32":
        from .windows import WindowsBackend
        _backend_instance = WindowsBackend()
    elif sys.platform == "darwin":
        from .macos import MacOSBackend
        _backend_instance = MacOSBackend()
    else:
        raise RuntimeError(
            f"Unsupported platform: {sys.platform}. "
            "PowerPoint MCP supports Windows and macOS."
        )

    return _backend_instance


def reset_backend():
    """Reset the singleton backend (for testing)."""
    global _backend_instance
    _backend_instance = None
