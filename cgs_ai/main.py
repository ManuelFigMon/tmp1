"""
=====================================================================
  Program Name  : main.py
  Author        : Manuel Figallo
  Purpose       : Single command-line entry point for the cgs_ai toolkit.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Description:
    Dispatches to any cgs_ai function by name, so a scheduler or operator
    has one predictable command instead of nine scripts.

  Usage:
      python main.py --list
      python main.py scanFileSystem --input-folder-root "\\srv\logs" \
                                    --extract-keyword "ERROR"
      python main.py filescanPipeline
=====================================================================
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

__version__ = "1.0beta"

COMMANDS = {
    "scanFileSystem": "src.py.scanFileSystem",
    "hello": None,
    "filescanPipeline": "src.pipelines.filescan_pipeline",
}


def main(argv=None) -> int:
    """Dispatch a cgs_ai command. Parameters: argv. Returns: exit code."""
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("--list", "-l", "--help", "-h"):
        print(f"cgs_ai {__version__} -- available commands:")
        for name in COMMANDS:
            print(f"  {name}")
        print("\nRun 'python main.py <command> --help' for a command's options.")
        return 0

    command, rest = argv[0], argv[1:]
    if command == "hello":
        import cgs_ai
        print(cgs_ai.basic_hello())
        return 0
    if command == "scanFileSystem":
        from src.py.scanFileSystem import main as scanMain
        return scanMain(rest)
    if command == "filescanPipeline":
        from src.pipelines.filescan_pipeline import runFilescanPipeline
        runFilescanPipeline()
        return 0

    print(f"unknown command: {command}. Use --list to see the options.",
          file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main())
