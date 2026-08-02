from __future__ import annotations

import os
from pathlib import Path


def render_disk_root() -> Path | None:
    root = Path("/var/data")
    if not os.environ.get("RENDER") or not root.is_dir():
        return None
    return root


def data_path(
    persistent_relative: str,
    env_key: str | None = None,
    local_default: str | None = None,
) -> Path:
    if env_key:
        configured = os.environ.get(env_key)
        if configured:
            return Path(configured)
    root = render_disk_root()
    if root is not None:
        return root / persistent_relative
    return Path(local_default or persistent_relative)
