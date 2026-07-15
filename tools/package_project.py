"""Creează o arhivă GitHub fără date, secrete sau fișiere temporare."""

from __future__ import annotations

import argparse
import hashlib
import re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[1]
EXCLUDED_DIRS = {
    ".git",
    "__pycache__",
    ".pytest_cache",
    ".ipynb_checkpoints",
    ".idea",
    ".vscode",
}
EXCLUDED_SUFFIXES = {".pyc", ".pyo", ".xls", ".xlsx", ".xlsm", ".log"}
SECRET_PATTERNS = [
    re.compile(
        r"""(?i)(?:serpapi[_-]?api[_-]?key|api[_-]?key)\s*=\s*["']([A-Za-z0-9_-]{20,})["']"""
    ),
]


def include(path: Path) -> bool:
    relative = path.relative_to(ROOT)
    if any(part in EXCLUDED_DIRS for part in relative.parts):
        return False
    if path.name == ".env":
        return False
    if path.suffix.casefold() in EXCLUDED_SUFFIXES:
        return False
    if path.name.endswith((".status.json", ".summary.json")):
        return False
    return path.is_file()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    output = args.output.expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    files = sorted(path for path in ROOT.rglob("*") if include(path))
    suspicious = []
    for path in files:
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        if any(pattern.search(text) for pattern in SECRET_PATTERNS):
            suspicious.append(str(path.relative_to(ROOT)))
    if suspicious:
        raise RuntimeError(f"Posibile secrete detectate: {suspicious}")

    with ZipFile(output, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for path in files:
            archive.write(path, Path(ROOT.name) / path.relative_to(ROOT))

    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    checksum = output.with_suffix(output.suffix + ".sha256")
    checksum.write_text(f"{digest}  {output.name}\n", encoding="ascii")
    print(f"Arhivă: {output}")
    print(f"Fișiere incluse: {len(files)}")
    print(f"Dimensiune: {output.stat().st_size / 1024:.1f} KB")
    print(f"SHA-256: {digest}")
    print(f"Fișier checksum: {checksum}")


if __name__ == "__main__":
    main()
