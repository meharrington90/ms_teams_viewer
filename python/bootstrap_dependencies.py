#!/usr/bin/env python3

import importlib.util
import subprocess
import sys
from pathlib import Path


CCL_REQUIREMENT = "ccl_chromium_reader @ https://codeload.github.com/cclgroupltd/ccl_chromium_reader/zip/refs/heads/master"
REPO_ROOT = Path(__file__).resolve().parents[1]


def has_ccl() -> bool:
    return importlib.util.find_spec("ccl_chromium_reader") is not None


def pip_install(*args: str) -> None:
    subprocess.run([sys.executable, "-m", "pip", "install", *args], check=True, cwd=REPO_ROOT)


def main() -> int:
    if has_ccl():
        return 0

    try:
        pip_install("-r", str(REPO_ROOT / "requirements.txt"))
        pip_install("--no-deps", CCL_REQUIREMENT)
    except subprocess.CalledProcessError:
        return 1

    return 0 if has_ccl() else 1


if __name__ == "__main__":
    raise SystemExit(main())
