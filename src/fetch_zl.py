from __future__ import annotations

import argparse
import hashlib
import sys
import urllib.request
from pathlib import Path


SOURCE_URL = "https://www.voynich.nu/data/ZL3b-n.txt"
EXPECTED_SHA256 = "bf5b6d4ac1e3a51b1847a9c388318d609020441ccd56984c901c32b09beccafc"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Download and verify the pinned ZL3b-n IVTFF transcription."
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("sources/ZL3b-n.txt"),
        help="Destination path (default: sources/ZL3b-n.txt)",
    )
    args = parser.parse_args()

    request = urllib.request.Request(
        SOURCE_URL,
        headers={"User-Agent": "Voynich-Dual-Track-AI-Edition/0.1.0"},
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        payload = response.read()

    actual = hashlib.sha256(payload).hexdigest()
    if actual != EXPECTED_SHA256:
        print(
            f"Refusing to save changed source: expected {EXPECTED_SHA256}, got {actual}",
            file=sys.stderr,
        )
        return 1

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_bytes(payload)
    print(f"Saved {args.output} ({len(payload)} bytes; SHA-256 {actual})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

