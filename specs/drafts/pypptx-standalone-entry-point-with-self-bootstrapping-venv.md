---
name: pypptx standalone entry point with self-bootstrapping venv
id: spec-5636a620
description: Adds a self-contained pypptx.py entry script at the repo root that manages its own virtualenv and dependencies without requiring uv, pip, or any external tooling.
dependencies: spec-b1c47a0f
priority: high
complexity: low
status: draft
tags:
- cli
- packaging
scope:
  in: pypptx.py entry script at repo root; .venv management
  out: changes to the pypptx package itself
feature_root_id:
---
# pypptx standalone entry point with self-bootstrapping venv

## Objective

Create a `pypptx.py` script at the repo root that can be invoked directly with `python3 pypptx.py <command>` — no `uv`, no `pip install`, no package manager required. On first run it creates a `.venv` next to itself, installs the required dependencies, and re-executes itself inside that environment. Subsequent runs skip straight to re-execution.

## Background

The current entry point is `uv run pypptx`, which requires `uv` to be installed and the package to be set up. Agents using this as a skill should be able to invoke it with only a system Python 3 available.

The pattern is taken from `memory.py` in the `skill-memory` project: a `_bootstrap()` function runs before any imports, creates a `.venv` next to the script using the stdlib `venv` module, installs deps via the venv's own `pip`, then calls `os.execv` to re-execute the process inside the venv. When re-executed, `str(_VENV) in sys.prefix` is true so `_bootstrap()` returns immediately.

## Implementation

`pypptx.py` at the repo root:

```python
#!/usr/bin/env python3
import sys, os, subprocess
from pathlib import Path

_SKILL_DIR = Path(__file__).parent
_VENV = _SKILL_DIR / ".venv"
_DEPS = ["click", "defusedxml", "python-pptx"]

def _bootstrap():
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
from pypptx.cli import cli
cli()
```

Key points:
- `_SKILL_DIR` is the repo root (where `pypptx/` package directory lives), so `sys.path.insert(0, str(_SKILL_DIR))` makes `from pypptx.cli import cli` work.
- `_DEPS` does **not** include Pillow — that is the optional thumbnails dependency and is out of scope for this spec.
- The script is chmod +x and has a `#!/usr/bin/env python3` shebang so it can also be invoked as `./pypptx.py`.
- `.venv/` is added to `.gitignore` if not already present.

## Files to Modify

| File | Change |
|---|---|
| `pypptx.py` | New — standalone entry script at repo root |
| `.gitignore` | Ensure `.venv/` is ignored |

## Acceptance Criteria

- `python3 pypptx.py --version` works on a system with only Python 3 stdlib (no uv, no pre-installed packages).
- First invocation prints `pypptx: first run, installing dependencies...` to stderr, creates `.venv/`, and completes successfully.
- Second invocation produces no bootstrap output and runs at normal speed.
- All existing commands work identically to `uv run pypptx`: `python3 pypptx.py unpack`, `slide list`, `slide add`, `extract-text`, etc.
- `.venv/` is present in `.gitignore`.
- The `pypptx/` package directory and `pyproject.toml` are unchanged.
