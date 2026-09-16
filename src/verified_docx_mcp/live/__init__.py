"""Live-mode package (issue #106: https://github.com/michaelrobertsutton/
JennyStack/issues/106): the local HTTPS + WSS bridge between
``verified_docx_mcp.server`` and the Word task pane (``addin/``).

  - ``bridge.py``    -- the HTTPS static server (WP-1: pane files, /ping,
                         /report) plus the WSS ``/ops`` channel and
                         ``start_in_background``/``stop`` (WP-2).
  - ``protocol.py``  -- the wire message schemas the two sides exchange
                         (WP-2).
  - ``session.py``   -- ``LiveSession``/``SessionRegistry``: one session
                         per connected pane, request/reply correlation,
                         heartbeat eviction (WP-2).

Import-clean on purpose: importing this package (or any module in it)
never starts a server, opens a socket, or touches the filesystem beyond
module load. ``server.py`` imports ``live.bridge`` at module load time
(for its ``live_status`` tool) with no side effect from the import
itself; the actual bridge only starts when ``live_status`` (or, from
WP-3/4 onward, a ``write_mode="live"`` tool call) calls
``bridge.start_in_background()``, which is idempotent. See
``bridge.py``'s module docstring for the CLI (``--serve-only`` /
``--make-cert``) and ``docs/live-mode.md`` for the lead's sideload
runbook and this package's architecture.
"""

from __future__ import annotations
