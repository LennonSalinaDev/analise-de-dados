from __future__ import annotations

import os
from pathlib import Path


def data_path(default_relative: str, env_key: str | None = None) -> Path:
    if env_key:
        configured = os.environ.get(env_key)
        if configured:
            return Path(configured)
    if os.environ.get("RENDER"):
        return Path("/var/data") / default_relative
    return Path(default_relative)
