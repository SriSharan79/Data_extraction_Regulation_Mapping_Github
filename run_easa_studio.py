#!/usr/bin/env python3
"""Entry point: launch the EASA Studio (extraction + JSON review).

    python run_easa_studio.py
"""

from data_extraction.studio.easa import main

if __name__ == "__main__":
    main()
