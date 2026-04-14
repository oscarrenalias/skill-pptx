#!/usr/bin/env python3
import os
import subprocess
import sys
from pathlib import Path

_SKILL_DIR = Path(__file__).parent
_VENV = _SKILL_DIR / ".venv"
_DEPS = ["click", "defusedxml", "python-pptx"]


def _bootstrap() -> None:
    if str(_VENV) in sys.prefix:
        return  # already inside the managed venv

    venv_py = _VENV / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")

    if not venv_py.exists():
        print("pypptx: first run, installing dependencies...", file=sys.stderr)
        subprocess.check_call([sys.executable, "-m", "venv", str(_VENV)])
        subprocess.check_call([str(venv_py), "-m", "pip", "install", "--quiet", *_DEPS])

    os.execv(str(venv_py), [str(venv_py)] + sys.argv)


_bootstrap()

sys.path.insert(0, str(_SKILL_DIR))

from pypptx.cli import cli  # noqa: E402

if __name__ == "__main__":
    cli()
