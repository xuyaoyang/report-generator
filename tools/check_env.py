"""Check the local report-generator runtime environment."""
from __future__ import annotations

import importlib.util
import shutil
import sys


REQUIRED_MODULES = [
    "docx",
    "docxcompose",
    "openpyxl",
    "PIL",
    "PySide6",
    "win32com",
    "fitz",
]


def module_status(name: str) -> str:
    return "OK" if importlib.util.find_spec(name) else "MISSING"


def check_word() -> str:
    try:
        import win32com.client

        word = win32com.client.Dispatch("Word.Application")
        version = word.Version
        word.Quit()
        return f"OK ({version})"
    except Exception as exc:  # pragma: no cover - depends on local Windows apps.
        return f"MISSING ({exc})"


def main() -> int:
    print(f"Python: {sys.executable}")
    print(f"Version: {sys.version.split()[0]}")
    print()
    print("Python modules:")
    missing = []
    for module in REQUIRED_MODULES:
        status = module_status(module)
        print(f"  {module}: {status}")
        if status != "OK":
            missing.append(module)

    print()
    print("External tools:")
    print(f"  Microsoft Word COM: {check_word()}")
    print(f"  soffice: {shutil.which('soffice') or 'not found'}")
    print(f"  pdftoppm: {shutil.which('pdftoppm') or 'not found'}")

    if missing:
        print()
        print("Missing required Python modules:", ", ".join(missing))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
