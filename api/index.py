"""Vercel entrypoint. The runtime wants a BaseHTTPRequestHandler named `handler`."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import Handler as handler  # noqa: E402
