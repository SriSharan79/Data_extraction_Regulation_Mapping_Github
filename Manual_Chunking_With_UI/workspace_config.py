"""Shared workspace configuration for the Docling chunking tools.

Single source of truth for the workspace registry path so the extraction
launcher and the cache launcher always agree on where the file-to-storage
registry lives.

The two launchers historically hardcoded two different paths that pointed at
the *same* share via different mounts (a Windows drive letter vs. a POSIX
mount). The default below is therefore OS-aware, and can be overridden entirely
with the ``DOCLING_WORKSPACE_REGISTRY`` environment variable.
"""

import os

_DEFAULT_WINDOWS = r"U:\ALR DATA\00_Container\docling_workspace_registry.json"
_DEFAULT_POSIX = (
    "/remotedata/U/DLR+kata_du/ALR DATA/00_Container/docling_workspace_registry.json"
)

REGISTRY_FILE = os.environ.get(
    "DOCLING_WORKSPACE_REGISTRY",
    _DEFAULT_WINDOWS if os.name == "nt" else _DEFAULT_POSIX,
)
