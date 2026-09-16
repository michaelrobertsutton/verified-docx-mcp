"""Generate the three PNG icons `addin/manifest.xml` references
(icon-16.png, icon-32.png, icon-80.png), stdlib-only (`zlib` + `struct`,
no Pillow) — issue #106 WP-1 deliverable 1: "generate simple PNG icons
... so no binary asset is hand-made."

Each icon is a flat navy square with a lighter diagonal mark, encoded as
an 8-bit RGB PNG: one IHDR chunk, one IDAT chunk (zlib-compressed
scanlines, filter type 0/None on every row), one IEND chunk — the
minimum a PNG decoder needs. Run directly to (re)write the three files
under `addin/`:

    python3 scripts/generate_addin_icons.py

Idempotent: re-running overwrites the same three deterministic files.
"""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

ADDIN_DIR = Path(__file__).resolve().parent.parent / "addin"

# Skyward-neutral navy + a lighter accent for the mark. Plain RGB, no
# branding claim -- this is a hello-world spike icon, not a shipped logo.
_BG = (0x0B, 0x2A, 0x4A)
_FG = (0xE8, 0xC1, 0x4C)


def _chunk(tag: bytes, data: bytes) -> bytes:
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)


def _pixel(x: int, y: int, size: int) -> tuple[int, int, int]:
    # A simple diagonal band (top-left to bottom-right) in the accent
    # color over the navy background -- enough to be visibly non-blank
    # at 16px, no drawing library involved.
    band = size / 4.0
    return _FG if abs(x - y) <= band else _BG


def make_png(size: int) -> bytes:
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)  # 8-bit depth, color type 2 = RGB
    raw = bytearray()
    for y in range(size):
        raw.append(0)  # filter type 0 (None) for every scanline
        for x in range(size):
            r, g, b = _pixel(x, y, size)
            raw.extend((r, g, b))
    idat = zlib.compress(bytes(raw), level=9)
    return b"\x89PNG\r\n\x1a\n" + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", idat) + _chunk(b"IEND", b"")


def main() -> None:
    ADDIN_DIR.mkdir(parents=True, exist_ok=True)
    for size in (16, 32, 80):
        out = ADDIN_DIR / f"icon-{size}.png"
        out.write_bytes(make_png(size))
        print(f"wrote {out} ({out.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
