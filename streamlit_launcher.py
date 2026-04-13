from __future__ import annotations

import os
import sys
from pathlib import Path

from streamlit.web import cli as streamlit_cli


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
else:
    BASE_DIR = Path(__file__).resolve().parent
    BUNDLE_DIR = BASE_DIR

APP_FILE = BUNDLE_DIR / "streamlit_app.py"


def main() -> None:
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "false")

    sys.argv = [
        "streamlit",
        "run",
        str(APP_FILE),
        "--global.developmentMode=false",
        "--server.address=127.0.0.1",
        "--server.port=8501",
        "--browser.gatherUsageStats=false",
    ]
    streamlit_cli.main()


if __name__ == "__main__":
    main()
