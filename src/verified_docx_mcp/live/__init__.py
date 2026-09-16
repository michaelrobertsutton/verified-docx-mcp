"""Live-mode package (issue #106 WP-1 spike): the local HTTPS bridge that
serves the ``addin/`` task-pane files to Word and answers ``/ping``.

Import-clean on purpose: importing this package (or ``live.bridge``) never
starts a server, opens a socket, or touches the filesystem beyond module
load. ``server.py`` (the FastMCP process) must be able to import
``verified_docx_mcp.live`` without side effects; WP-2 wires a *lazy* start
triggered by an actual live tool call, not by import. See
``bridge.py``'s module docstring for the CLI this WP-1 spike ships
(``--serve-only`` / ``--make-cert``) and ``docs/live-mode.md`` for the
lead's sideload runbook.
"""

from __future__ import annotations
