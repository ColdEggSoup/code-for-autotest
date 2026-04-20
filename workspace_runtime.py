from __future__ import annotations

import sys
from pathlib import Path


def configure_workspace_runtime() -> None:
    try:
        repo_root = Path(__file__).resolve().parent
        pycache_root = repo_root / "trash" / "__pycache__"
        pycache_root.mkdir(parents=True, exist_ok=True)
        if sys.pycache_prefix is None:
            sys.pycache_prefix = str(pycache_root)
    except Exception:
        pass
