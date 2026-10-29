"""Generate the large acceptance document used by the token-budget measurements
(#110 / #112): 150 headed sections, 300 paragraphs, ~20K words.

    uv run python scripts/make_large_fixture.py /private/tmp/large-acceptance.docx

The file is built from the Word-authored tests/fixtures/sections.docx by the
server's own replace_body_markdown, so it is a valid .docx that Word opens, but it
is a generated measurement artifact and is not committed. Deterministic: the same
words, in the same order, every run.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "src"))

SECTIONS = 150
WORDS_PER_PARAGRAPH = 130
VOCAB = ["program", "office", "delivery", "quarterly", "milestone", "requirement", "stakeholder", "baseline", "deliverable", "integration", "verification", "schedule", "performance", "quality", "assurance", "risk", "mitigation", "workforce", "transition", "support", "analysis", "reporting", "governance"]


def body_markdown() -> str:
    blocks: list[str] = []
    for n in range(1, SECTIONS + 1):
        words = [VOCAB[(n * 7 + i * 3) % len(VOCAB)] for i in range(WORDS_PER_PARAGRAPH)]
        sentence = " ".join(words)
        blocks.append(f"# Section {n:03d}\n\n{sentence.capitalize()}.")
    return "\n\n".join(blocks)


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit("usage: make_large_fixture.py <output.docx>")
    out = Path(sys.argv[1]).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(REPO / "tests" / "fixtures" / "sections.docx", out)
    os.environ["VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS"] = str(out.parent)
    from verified_docx_mcp import mutations

    result = mutations.execute_replace_body_markdown(str(out), body_markdown(), force=True)
    assert result["applied"], result
    print(out)


if __name__ == "__main__":
    main()
