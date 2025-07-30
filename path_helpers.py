import os
import sys


def get_base_path():
    """Return the base path for bundled or normal execution."""
    if getattr(sys, 'frozen', False):
         # When bundled by PyInstaller, resources live next to the executable
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))