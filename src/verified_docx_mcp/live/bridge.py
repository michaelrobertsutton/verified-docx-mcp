"""Local HTTPS bridge for the WP-1 spike
(https://github.com/michaelrobertsutton/JennyStack/issues/106).

Two independent CLI actions, no server framework, stdlib only:

  python -m verified_docx_mcp.live.bridge --make-cert [DIR]
      Generate a self-signed certificate for `localhost` (SAN
      `DNS:localhost, IP:127.0.0.1`) via the `openssl` CLI -- no Python
      crypto dependency. Writes `localhost.pem` / `localhost-key.pem`
      into DIR (default `~/.cache/verified-docx-mcp/live-cert/`) and
      prints the exact `security add-trusted-cert` command the lead runs
      once, against their LOGIN keychain (`login.keychain-db`, not
      System), so no `sudo` is needed.

  python -m verified_docx_mcp.live.bridge --serve-only [--port 53135] [--cert DIR]
      Serve `addin/` (this repo's task-pane files) statically over HTTPS
      on 127.0.0.1, plus `GET /ping` -> `{"ok": true, "server":
      "verified-docx-mcp", "time": <epoch seconds>}`. Blocks until
      Ctrl+C. This is throwaway WP-1 scaffolding per the plan
      (docs/plans/issue-106-word-addin-bridge.md): WP-2 replaces it with
      the real asyncio bridge (`live/session.py`, `live/protocol.py`,
      a WSS ops channel) started lazily by the MCP process itself. Until
      then, `--serve-only` is what the lead runs by hand per
      docs/live-mode.md.

Importing this module never starts a server (see the package's own
`__init__.py` docstring) -- everything here is a function the CLI's
`main()` calls; `verified_docx_mcp.server` does not import this module
at all yet (WP-2 wires that, lazily, on the first live tool call).

Binds to 127.0.0.1 only. Office requires HTTPS for add-in resources even
on localhost, hence the cert step; there is no HTTP fallback here.
"""

from __future__ import annotations

import argparse
import functools
import http.server
import json
import ssl
import subprocess
import sys
import time
from pathlib import Path

DEFAULT_PORT = 53135
DEFAULT_HOST = "127.0.0.1"

# bridge.py lives at src/verified_docx_mcp/live/bridge.py; the addin/
# directory this WP-1 spike serves lives at the repo root, four levels
# up. addin/ is deliberately NOT in pyproject.toml's sdist include list
# (dev/spike tooling, not a package to ship in the wheel), so this path
# only resolves inside a checkout -- see main()'s explicit error when it
# does not, rather than a bare FileNotFoundError.
REPO_ADDIN_DIR = Path(__file__).resolve().parents[3] / "addin"

DEFAULT_CERT_DIR = Path.home() / ".cache" / "verified-docx-mcp" / "live-cert"
CERT_FILENAME = "localhost.pem"
KEY_FILENAME = "localhost-key.pem"


def trust_command(cert_path: Path) -> str:
    """The exact command the lead runs once to trust the cert -- LOGIN
    keychain, not System, so no sudo (deliverable 3's requirement)."""
    return f"security add-trusted-cert -d -r trustRoot -k ~/Library/Keychains/login.keychain-db {cert_path}"


def make_cert(cert_dir: Path, *, openssl_bin: str = "openssl") -> tuple[Path, Path]:
    """Generate a self-signed cert+key for `localhost` (SAN
    `DNS:localhost, IP:127.0.0.1`) via the `openssl` CLI. Returns
    `(cert_path, key_path)`. Always overwrites any existing cert in
    `cert_dir`; `main()` is the layer that decides whether to skip
    regeneration when one already exists."""
    cert_dir.mkdir(parents=True, exist_ok=True)
    cert_path = cert_dir / CERT_FILENAME
    key_path = cert_dir / KEY_FILENAME
    cmd = [
        openssl_bin,
        "req",
        "-x509",
        "-newkey",
        "rsa:2048",
        "-sha256",
        "-days",
        "825",
        "-nodes",
        "-keyout",
        str(key_path),
        "-out",
        str(cert_path),
        "-subj",
        "/CN=localhost",
        "-addext",
        "subjectAltName=DNS:localhost,IP:127.0.0.1",
    ]
    subprocess.run(cmd, check=True, capture_output=True, text=True)
    return cert_path, key_path


class _AddinRequestHandler(http.server.SimpleHTTPRequestHandler):
    """Serves `addin_dir` statically, plus a synthetic `/ping` JSON
    route the static handler cannot produce. `log_message` is silenced
    by default -- the lead runs this in a foreground terminal per
    docs/live-mode.md and does not need a request log for the spike."""

    def __init__(self, *args, addin_dir: Path, **kwargs):
        self._addin_dir = addin_dir
        super().__init__(*args, directory=str(addin_dir), **kwargs)

    def do_GET(self) -> None:  # stdlib override name, not our naming convention to control
        if self.path == "/ping" or self.path.startswith("/ping?"):
            self._send_ping()
            return
        super().do_GET()

    def _send_ping(self) -> None:
        payload = json.dumps({"ok": True, "server": "verified-docx-mcp", "time": time.time()}).encode(
            "utf-8"
        )
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, format: str, *args: object) -> None:  # stdlib override signature
        pass


def make_server(
    addin_dir: Path,
    *,
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    certfile: Path,
    keyfile: Path,
) -> http.server.HTTPServer:
    """Build (but do not start) an HTTPS `HTTPServer` bound to `host:port`
    serving `addin_dir`. `port=0` binds an ephemeral port -- the unit
    tests use that to avoid colliding with a real WP-1 run on
    DEFAULT_PORT; read the actual port back from `httpd.server_address`."""
    handler = functools.partial(_AddinRequestHandler, addin_dir=addin_dir)
    httpd = http.server.HTTPServer((host, port), handler)
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(certfile=str(certfile), keyfile=str(keyfile))
    httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
    return httpd


def _build_argparser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m verified_docx_mcp.live.bridge",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--serve-only",
        action="store_true",
        help="Serve addin/ and /ping over local HTTPS and block until Ctrl+C.",
    )
    parser.add_argument(
        "--make-cert",
        metavar="DIR",
        nargs="?",
        const=str(DEFAULT_CERT_DIR),
        default=None,
        help=(
            "Generate a self-signed localhost cert in DIR "
            f"(default: {DEFAULT_CERT_DIR}) and print the trust command."
        ),
    )
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help=f"default {DEFAULT_PORT}")
    parser.add_argument(
        "--cert",
        metavar="DIR",
        default=None,
        help=f"Cert directory to serve with (default: {DEFAULT_CERT_DIR}).",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_argparser()
    args = parser.parse_args(argv)

    if args.make_cert:
        cert_dir = Path(args.make_cert).expanduser()
        cert_path, key_path = make_cert(cert_dir)
        print(f"wrote {cert_path}")
        print(f"wrote {key_path}")
        print()
        print("Trust it once (login keychain, no sudo required):")
        print(f"  {trust_command(cert_path)}")
        return 0

    if args.serve_only:
        cert_dir = Path(args.cert).expanduser() if args.cert else DEFAULT_CERT_DIR
        cert_path = cert_dir / CERT_FILENAME
        key_path = cert_dir / KEY_FILENAME
        if not cert_path.exists() or not key_path.exists():
            print(f"No cert found in {cert_dir}.", file=sys.stderr)
            print("Generate one first:", file=sys.stderr)
            print(f"  python -m verified_docx_mcp.live.bridge --make-cert {cert_dir}", file=sys.stderr)
            return 1
        if not REPO_ADDIN_DIR.exists():
            print(f"addin/ not found at {REPO_ADDIN_DIR}.", file=sys.stderr)
            print("Run this from a verified-docx-mcp checkout, not an installed wheel.", file=sys.stderr)
            return 1
        httpd = make_server(REPO_ADDIN_DIR, port=args.port, certfile=cert_path, keyfile=key_path)
        bound_port = httpd.server_address[1]
        print(f"Serving {REPO_ADDIN_DIR} on https://{DEFAULT_HOST}:{bound_port}/ (Ctrl+C to stop)")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            pass
        finally:
            httpd.server_close()
        return 0

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
