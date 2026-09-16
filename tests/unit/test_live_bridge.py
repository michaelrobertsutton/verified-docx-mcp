"""Unit tests for the WP-1 spike
(https://github.com/michaelrobertsutton/JennyStack/issues/106):
`addin/manifest.xml`, the generated icons, and
`src/verified_docx_mcp/live/bridge.py`.

stdlib `unittest`, no third-party dependencies. Run:

  uv run pytest tests/unit
  # or directly:
  uv run python -m unittest tests.unit.test_live_bridge

Covers:
  - manifest.xml is well-formed XML and carries the required elements/
    attributes (Id, Version, ProviderName, DisplayName, Hosts=Document,
    WordApi MinVersion 1.4, localhost SourceLocation/AppDomain,
    ReadWriteDocument permission, a ShowTaskpane VersionOverrides
    control).
  - icon-16/32/80.png are valid PNGs: signature + parseable IHDR
    matching the claimed size.
  - `bridge.make_cert` produces a cert whose SAN includes `localhost`
    (parsed from `openssl x509 -text`). Skipped cleanly when `openssl`
    is not on PATH.
  - `bridge.make_server` starts an HTTPS server on an ephemeral port
    with that cert and serves `/ping` and `/taskpane.html` (a plain
    `ssl.CERT_NONE` client, mirroring an agent's own verification --
    the lead's real trust step is out of scope for an offline test).
"""

from __future__ import annotations

import http.client
import shutil
import ssl
import struct
import subprocess
import sys
import tempfile
import threading
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp.live import bridge

ADDIN_DIR = REPO / "addin"

_NS = {
    "o": "http://schemas.microsoft.com/office/appforoffice/1.1",
    "ov": "http://schemas.microsoft.com/office/taskpaneappversionoverrides",
    "bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0",
}


class ManifestTests(unittest.TestCase):
    def setUp(self):
        self.root = ET.parse(ADDIN_DIR / "manifest.xml").getroot()

    def test_root_is_office_app(self):
        self.assertEqual(self.root.tag, f"{{{_NS['o']}}}OfficeApp")
        self.assertEqual(self.root.attrib.get("{http://www.w3.org/2001/XMLSchema-instance}type"), "TaskPaneApp")

    def test_id_is_a_valid_guid(self):
        import uuid

        id_text = self.root.find("o:Id", _NS).text
        # Round-trips through uuid.UUID without raising -- the well-formed
        # check; ValueError would fail the test if it were not a GUID.
        uuid.UUID(id_text)

    def test_version_and_provider_and_display_name(self):
        self.assertEqual(self.root.find("o:Version", _NS).text, "0.1.0")
        self.assertEqual(self.root.find("o:ProviderName", _NS).text, "verified-docx-mcp")
        display = self.root.find("o:DisplayName", _NS)
        self.assertEqual(display.attrib.get("DefaultValue"), "verified-docx-mcp live")

    def test_hosts_is_document_only(self):
        hosts = self.root.findall("o:Hosts/o:Host", _NS)
        self.assertEqual([h.attrib.get("Name") for h in hosts], ["Document"])

    def test_requirements_set_wordapi_1_4(self):
        sets = self.root.findall("o:Requirements/o:Sets/o:Set", _NS)
        word_sets = [s for s in sets if s.attrib.get("Name") == "WordApi"]
        self.assertEqual(len(word_sets), 1)
        self.assertEqual(word_sets[0].attrib.get("MinVersion"), "1.4")

    def test_source_location_is_localhost_53135(self):
        source = self.root.find("o:DefaultSettings/o:SourceLocation", _NS)
        self.assertEqual(source.attrib.get("DefaultValue"), "https://localhost:53135/taskpane.html")

    def test_app_domain_includes_localhost(self):
        domains = [d.text for d in self.root.findall("o:AppDomains/o:AppDomain", _NS)]
        self.assertIn("https://localhost:53135", domains)

    def test_permissions_is_read_write_document(self):
        self.assertEqual(self.root.find("o:Permissions", _NS).text, "ReadWriteDocument")

    def test_icon_urls_point_at_localhost(self):
        icon_url = self.root.find("o:IconUrl", _NS).attrib.get("DefaultValue")
        hires_url = self.root.find("o:HighResolutionIconUrl", _NS).attrib.get("DefaultValue")
        self.assertEqual(icon_url, "https://localhost:53135/icon-32.png")
        self.assertEqual(hires_url, "https://localhost:53135/icon-80.png")

    def test_version_overrides_has_show_taskpane_button(self):
        actions = self.root.findall(".//ov:Action", _NS)
        types = [a.get("{http://www.w3.org/2001/XMLSchema-instance}type") for a in actions]
        self.assertIn("ShowTaskpane", types)
        button = self.root.find(".//ov:Control[@id='VerifiedDocxMcp.ShowTaskpaneButton']", _NS)
        self.assertIsNotNone(button)

    def test_resource_urls_reference_taskpane_and_icons(self):
        urls = {u.attrib.get("id"): u.attrib.get("DefaultValue") for u in self.root.findall(".//bt:Url", _NS)}
        self.assertEqual(urls.get("Taskpane.Url"), "https://localhost:53135/taskpane.html")
        images = {i.attrib.get("id"): i.attrib.get("DefaultValue") for i in self.root.findall(".//bt:Image", _NS)}
        self.assertEqual(images.get("Icon.16x16"), "https://localhost:53135/icon-16.png")
        self.assertEqual(images.get("Icon.32x32"), "https://localhost:53135/icon-32.png")
        self.assertEqual(images.get("Icon.80x80"), "https://localhost:53135/icon-80.png")


class IconPngTests(unittest.TestCase):
    def _assert_valid_png(self, path: Path, expected_size: int) -> None:
        data = path.read_bytes()
        self.assertEqual(data[:8], b"\x89PNG\r\n\x1a\n", f"{path} missing PNG signature")
        self.assertEqual(data[12:16], b"IHDR", f"{path} first chunk is not IHDR")
        width, height, bit_depth, color_type = struct.unpack(">IIBB", data[16:26])
        self.assertEqual(width, expected_size)
        self.assertEqual(height, expected_size)
        self.assertEqual(bit_depth, 8)
        self.assertEqual(color_type, 2)  # RGB, no palette

    def test_icon_16(self):
        self._assert_valid_png(ADDIN_DIR / "icon-16.png", 16)

    def test_icon_32(self):
        self._assert_valid_png(ADDIN_DIR / "icon-32.png", 32)

    def test_icon_80(self):
        self._assert_valid_png(ADDIN_DIR / "icon-80.png", 80)


_OPENSSL = shutil.which("openssl")


@unittest.skipUnless(_OPENSSL, "openssl CLI not found on PATH")
class MakeCertTests(unittest.TestCase):
    def test_make_cert_san_includes_localhost(self):
        with tempfile.TemporaryDirectory() as d:
            cert_dir = Path(d)
            cert_path, key_path = bridge.make_cert(cert_dir, openssl_bin=_OPENSSL)
            self.assertTrue(cert_path.exists())
            self.assertTrue(key_path.exists())
            out = subprocess.run(
                [_OPENSSL, "x509", "-in", str(cert_path), "-noout", "-text"],
                check=True,
                capture_output=True,
                text=True,
            ).stdout
            self.assertIn("DNS:localhost", out)
            self.assertIn("IP Address:127.0.0.1", out)

    def test_trust_command_uses_login_keychain_no_sudo(self):
        cmd = bridge.trust_command(Path("/tmp/x/localhost.pem"))
        self.assertIn("login.keychain-db", cmd)
        self.assertNotIn("sudo", cmd)
        self.assertIn("-d -r trustRoot", cmd)


@unittest.skipUnless(_OPENSSL, "openssl CLI not found on PATH")
class ServeOnlyTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        cert_dir = Path(self._tmp.name)
        self.cert_path, self.key_path = bridge.make_cert(cert_dir, openssl_bin=_OPENSSL)
        self.httpd = bridge.make_server(ADDIN_DIR, port=0, certfile=self.cert_path, keyfile=self.key_path)
        self.port = self.httpd.server_address[1]
        self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        self.thread.start()

    def tearDown(self):
        self.httpd.shutdown()
        self.httpd.server_close()
        self.thread.join(timeout=5)
        self._tmp.cleanup()

    def _get(self, path: str) -> http.client.HTTPResponse:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        conn = http.client.HTTPSConnection("127.0.0.1", self.port, context=ctx, timeout=10)
        conn.request("GET", path)
        return conn.getresponse()

    def test_ping_returns_ok_json(self):
        import json

        resp = self._get("/ping")
        self.assertEqual(resp.status, 200)
        self.assertEqual(resp.getheader("Content-Type"), "application/json")
        body = json.loads(resp.read())
        self.assertEqual(body["ok"], True)
        self.assertEqual(body["server"], "verified-docx-mcp")
        self.assertIn("time", body)

    def test_taskpane_html_served(self):
        resp = self._get("/taskpane.html")
        self.assertEqual(resp.status, 200)
        body = resp.read()
        self.assertIn(b"<!DOCTYPE html>", body)
        self.assertIn(b"office.js", body)

    def test_manifest_served_statically(self):
        resp = self._get("/manifest.xml")
        self.assertEqual(resp.status, 200)
        self.assertIn(b"<OfficeApp", resp.read())


if __name__ == "__main__":
    unittest.main()
