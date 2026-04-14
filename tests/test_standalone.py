"""Tests for pypptx.py standalone entry script.

Covers:
- Bootstrap guard: returns early when _VENV is already in sys.prefix.
- Venv creation: subprocess creates venv when venv_py is absent.
- _DEPS list: exactly the three required packages.
- os.execv: called with the correct venv python and argv.
- Version invocation: subprocess call exits 0 with "pypptx" in output.
- First-run bootstrap: .venv/ is created when using a fresh Python.
- Second-run silence: no stderr output once already bootstrapped.
"""
import importlib.util
import shutil
import subprocess
import sys
from pathlib import Path
from unittest.mock import call, patch

import pytest

_REPO_ROOT = Path(__file__).parent.parent
_SCRIPT = _REPO_ROOT / ".apm" / "skills" / "pypptx" / "pypptx.py"


# ---------------------------------------------------------------------------
# Load the standalone script as a module
# (safe because uv run pytest already activates .venv → _bootstrap() no-ops)
# ---------------------------------------------------------------------------


def _load_standalone():
    key = "pypptx_standalone_script"
    if key in sys.modules:
        return sys.modules[key]
    spec = importlib.util.spec_from_file_location(key, _SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[key] = mod
    spec.loader.exec_module(mod)
    return mod


_mod = _load_standalone()


# ---------------------------------------------------------------------------
# Unit tests — _bootstrap() internals
# ---------------------------------------------------------------------------


class TestBootstrapGuard:
    """_bootstrap() must skip all side-effects when the venv is already active."""

    def test_returns_early_when_venv_in_prefix(self):
        """When sys.prefix contains str(_VENV), no subprocess call is made."""
        with patch.object(sys, "prefix", str(_mod._VENV)), \
             patch("subprocess.check_call") as mock_sub:
            _mod._bootstrap()
        mock_sub.assert_not_called()

    def test_deps_list_contains_exactly_three_packages(self):
        """_DEPS must be exactly ["click", "defusedxml", "python-pptx"]."""
        assert _mod._DEPS == ["click", "defusedxml", "python-pptx"]

    def test_venv_created_when_python_absent(self):
        """When venv_py does not exist, subprocess.check_call creates the venv."""
        with patch.object(sys, "prefix", "/other/prefix"), \
             patch.object(sys, "argv", ["pypptx.py"]), \
             patch("subprocess.check_call") as mock_sub, \
             patch("pathlib.Path.exists", return_value=False), \
             patch("os.execv"):
            _mod._bootstrap()
        first_call = mock_sub.call_args_list[0]
        assert first_call == call([sys.executable, "-m", "venv", str(_mod._VENV)])

    def test_first_run_prints_stderr_message(self):
        """On first run (venv absent), a message is printed to stderr."""
        with patch.object(sys, "prefix", "/other/prefix"), \
             patch.object(sys, "argv", ["pypptx.py"]), \
             patch("subprocess.check_call"), \
             patch("pathlib.Path.exists", return_value=False), \
             patch("os.execv"), \
             patch("builtins.print") as mock_print:
            _mod._bootstrap()
        mock_print.assert_called_once_with(
            "pypptx: first run, installing dependencies...", file=sys.stderr
        )

    def test_pip_install_runs_only_on_first_run(self):
        """pip install is called when venv is absent; NOT called when venv already exists."""
        # Second-run scenario: venv_py exists, bootstrap just calls execv.
        with patch.object(sys, "prefix", "/other/prefix"), \
             patch.object(sys, "argv", ["pypptx.py"]), \
             patch("subprocess.check_call") as mock_sub, \
             patch("pathlib.Path.exists", return_value=True), \
             patch("os.execv"):
            _mod._bootstrap()
        pip_calls = [c for c in mock_sub.call_args_list if "pip" in str(c)]
        assert pip_calls == [], "pip install must not run when venv already exists"

    def test_execv_called_with_correct_args(self):
        """os.execv receives venv python path and the original sys.argv."""
        fake_argv = ["pypptx.py", "--version"]
        expected_py = str(_mod._VENV / "bin" / "python")
        with patch.object(sys, "prefix", "/other/prefix"), \
             patch.object(sys, "argv", fake_argv), \
             patch("subprocess.check_call"), \
             patch("pathlib.Path.exists", return_value=True), \
             patch("os.execv") as mock_execv:
            _mod._bootstrap()
        mock_execv.assert_called_once_with(expected_py, [expected_py] + fake_argv)


# ---------------------------------------------------------------------------
# Integration test — version invocation (venv already active)
# ---------------------------------------------------------------------------


class TestVersionInvocation:
    """Run pypptx.py --version as a subprocess using the active venv python."""

    def test_version_exits_zero(self):
        result = subprocess.run(
            [sys.executable, str(_SCRIPT), "--version"],
            capture_output=True,
            cwd=str(_REPO_ROOT),
        )
        assert result.returncode == 0

    def test_version_output_mentions_pypptx(self):
        result = subprocess.run(
            [sys.executable, str(_SCRIPT), "--version"],
            capture_output=True,
            text=True,
            cwd=str(_REPO_ROOT),
        )
        assert result.returncode == 0
        assert "pypptx" in result.stdout


# ---------------------------------------------------------------------------
# Integration tests — first-run and second-run bootstrap (needs system Python)
# ---------------------------------------------------------------------------


def _find_base_python() -> str | None:
    """Return a Python interpreter outside the current venv, or None."""
    base = getattr(sys, "_base_executable", None)
    if base and base != sys.executable:
        return base
    venv_bin = str(Path(sys.prefix) / "bin")
    for name in ("python3", "python"):
        found = shutil.which(name)
        if found and not found.startswith(venv_bin):
            return found
    return None


_BASE_PYTHON = _find_base_python()

_needs_base_python = pytest.mark.skipif(
    _BASE_PYTHON is None,
    reason="No Python interpreter found outside current venv; skipping bootstrap integration tests",
)


@pytest.fixture
def isolated_repo(tmp_path):
    """Temp directory containing pypptx.py and the pypptx package."""
    shutil.copy2(_SCRIPT, tmp_path / "pypptx.py")
    shutil.copytree(_REPO_ROOT / ".apm" / "skills" / "pypptx" / "pypptx", tmp_path / "pypptx")
    return tmp_path


@_needs_base_python
class TestFirstRunBootstrap:
    """End-to-end bootstrap tests that create a real venv from scratch."""

    def test_first_run_creates_venv(self, isolated_repo):
        """Running with a non-venv Python creates .venv/ in the script directory."""
        subprocess.run(
            [_BASE_PYTHON, str(isolated_repo / "pypptx.py"), "--version"],
            capture_output=True,
            cwd=str(isolated_repo),
            timeout=180,
        )
        assert (isolated_repo / ".venv").exists()

    def test_second_run_produces_no_stderr(self, isolated_repo):
        """After bootstrap, a second run from within the venv emits nothing to stderr."""
        # First run — bootstrap happens (may take a while)
        subprocess.run(
            [_BASE_PYTHON, str(isolated_repo / "pypptx.py"), "--version"],
            capture_output=True,
            cwd=str(isolated_repo),
            timeout=180,
        )
        venv_py = isolated_repo / ".venv" / "bin" / "python"
        if not venv_py.exists():
            pytest.skip(".venv not created by first run; cannot test second-run behaviour")
        # Second run — already in venv, bootstrap must be skipped
        result = subprocess.run(
            [str(venv_py), str(isolated_repo / "pypptx.py"), "--version"],
            capture_output=True,
            text=True,
            cwd=str(isolated_repo),
            timeout=30,
        )
        assert result.returncode == 0
        assert result.stderr == ""
