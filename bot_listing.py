#!/usr/bin/env python3
"""
bot_listing.py
==============
Discord Bot khusus untuk /listing (Generate & Upload Outlet Info ke Google Drive).
Menjalankan layanan bot mandiri untuk channel #listing.
"""
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from discord_bot import main

if __name__ == "__main__":
    main()

