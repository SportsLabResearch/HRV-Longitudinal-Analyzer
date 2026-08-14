"""
=========================================================
HRV-Longitudinal-Analyzer
Utility functions
=========================================================
"""

from datetime import datetime
from pathlib import Path

from config import LINE


def print_header(title: str) -> None:
    """Print a formatted console header."""
    print("\n" + LINE)
    print(title.upper())
    print(LINE)


def print_step(message: str) -> None:
    """Print a formatted process step."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


def ensure_directory(path: Path) -> None:
    """Create a directory if it does not exist."""
    path.mkdir(parents=True, exist_ok=True)


def ensure_project_directories(directories: list[Path]) -> None:
    """Create all required project directories."""
    for directory in directories:
        ensure_directory(directory)


def file_exists(path: Path) -> bool:
    """Check whether a file exists."""
    return path.exists() and path.is_file()


def directory_exists(path: Path) -> bool:
    """Check whether a directory exists."""
    return path.exists() and path.is_dir()


def count_files(path: Path, extensions: tuple[str, ...] = None) -> int:
    """Count files in a directory, optionally filtering by extension."""
    if not directory_exists(path):
        return 0

    files = [f for f in path.rglob("*") if f.is_file()]

    if extensions is not None:
        files = [f for f in files if f.suffix.lower() in extensions]

    return len(files)


def today_string() -> str:
    """Return current date as string."""
    return datetime.now().strftime("%Y-%m-%d")