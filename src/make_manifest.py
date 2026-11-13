#!/usr/bin/env python
"""Create SHA-256 checksums for the complete reproducibility package."""

from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> None:
    patterns = [
        "README.md",
        "requirements.txt",
        "CITATION.cff",
        "LICENSE",
        "src/*.py",
        "checkpoints/*",
        "results/*",
        "figures/*",
        "paper/*.docx",
        "paper/*.pdf",
    ]
    files: dict[str, dict[str, object]] = {}
    for pattern in patterns:
        for path in sorted(ROOT.glob(pattern)):
            if not path.is_file() or path.name == "manifest.json":
                continue
            relative = path.relative_to(ROOT).as_posix()
            files[relative] = {
                "bytes": path.stat().st_size,
                "sha256": sha256(path),
            }
    payload = {
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "algorithm": "SHA-256",
        "files": files,
    }
    output = ROOT / "results" / "manifest.json"
    with output.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    print(f"Wrote {output} with {len(files)} file checksums.")


if __name__ == "__main__":
    main()
