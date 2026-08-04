"""Vercel entrypoint. The runtime scans this file for a top-level `handler`;
an import alias doesn't register, so subclass it."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import Handler  # noqa: E402


class handler(Handler):  # noqa: N801 - the name is the runtime's contract
    pass
