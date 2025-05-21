#!/usr/bin/env python3
"""
pyktok_dl.py URL

Download TikTok video using yt_dlp and save to ~/Pictures/Screenshots
"""

import sys, os, logging, pathlib
from yt_dlp import YoutubeDL

# Setup
SAVE_DIR = os.path.expanduser("~/Pictures/Screenshots")
os.makedirs(SAVE_DIR, exist_ok=True)

LOG_FILE = pathlib.Path.home() / ".cache" / "pyktok_dl.log"
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

logger = logging.getLogger("pyktok_dl")

def main():
    logger.info("=== TikTok download started ===")

    if len(sys.argv) != 2:
        logger.error("Incorrect usage: expected 1 argument (URL)")
        print("Usage: pyktok_dl.py <tiktok_url>")
        sys.exit(1)

    url = sys.argv[1].strip()
    logger.info(f"Received TikTok URL: {url}")
    logger.info(f"Saving to directory: {SAVE_DIR}")

    ydl_opts = {
        "outtmpl": os.path.join(SAVE_DIR, "%(title).50s.%(ext)s"),
        "format": "mp4",
        "quiet": True
    }

    try:
        with YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            filename = ydl.prepare_filename(info)
            logger.info(f"Download successful: {filename}")
            print("✓ Saved to", filename)
    except Exception as e:
        logger.exception(f"Download failed: {e}")
        sys.exit(1)

    logger.info("=== TikTok download complete ===")

if __name__ == "__main__":
    main()
