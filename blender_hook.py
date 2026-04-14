"""Compatibility wrapper for blender_listener.py.

Example:
    blender -b project.blend --python blender_hook.py -- results\blender_results.csv blender_001
"""

from __future__ import annotations

import sys
from pathlib import Path


def ensure_script_dir_on_path() -> None:
    if "__file__" in globals():
        script_dir = Path(__file__).resolve().parent
    else:
        script_dir = Path.cwd()
    script_dir_text = str(script_dir)
    if script_dir_text not in sys.path:
        sys.path.insert(0, script_dir_text)


ensure_script_dir_on_path()

from blender_listener import register, unregister  # noqa: E402


if __name__ == "__main__":
    register()
