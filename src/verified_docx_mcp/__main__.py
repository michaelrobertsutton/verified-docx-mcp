"""Run the server via ``python -m verified_docx_mcp``.

Delegates to the same entry point as the ``verified-docx-mcp`` console
script, so ``python -m verified_docx_mcp`` and
``python -m verified_docx_mcp doctor`` behave identically to the installed
command.
"""

from .server import main

if __name__ == "__main__":
    main()
