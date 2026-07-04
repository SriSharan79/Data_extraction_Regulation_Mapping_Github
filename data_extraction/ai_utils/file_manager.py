"""Base workspace folder used by the AI utilities for LLM logs and the API-key
config file.

Defaults to ``~/.data_extraction`` and can be overridden with the
``ALR_MAIN_FOLDER`` environment variable (e.g. to point at an existing ALR
workspace that already holds ``API_keys_config.json``).
"""

import os
from pathlib import Path

ALR_main_folder = Path(os.environ.get("ALR_MAIN_FOLDER", Path.home() / ".data_extraction"))

try:
    ALR_main_folder.mkdir(parents=True, exist_ok=True)
except OSError:
    # Read-only home / sandbox — the LLM helpers create subfolders on demand.
    pass
