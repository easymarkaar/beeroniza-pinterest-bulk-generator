"""Offline tests: no network or credentials needed."""
import io
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from generate_pins import upload_local_image, GeneratorError, warn_if_background_looks_missing


class FakeResponse:
    def __init__(self, *, status_code=200, content=b"", payload=None):
        self.status_code = status_code
        self.content = content
        self._payload = payload or {}
        self.text = ""
        self.headers = {}

    def json(self):
        return self._payload

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"http {self.status_code}")

    def iter_content(self, chunk_size=1024):
        for i in range(0, len(self.content), chunk_size):
            yield self.content[i:i+chunk_size]

    def __enter__(self):
        return self

    def __exit__(self, *_):
        return False


class FakeSession:
    def __init__(self, content):
        self.content = content
        self.post_count = 0
        self.get_urls = []

    def post(self, url, headers=None, files=None, timeout=None):
        self.post_count += 1
        self.upload_content = files["file"][1].read()
        return FakeResponse(status_code=201, payload={"id": "file123", "url": "http://test:8502/api/files/file123"})

    def get(self, url, headers=None, timeout=None, stream=None):
        self.get_urls.append(url)
        return FakeResponse(content=self.content)


class BackgroundUploadTests(unittest.TestCase):
    def make_png(self):
        out = io.BytesIO()
        Image.new("RGB", (100, 140), (180, 130, 70)).save(out, "PNG")
        return out.getvalue()

    def test_local_upload_and_round_trip(self):
        data = self.make_png()
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "bg.png").write_bytes(data)
            session = FakeSession(data)
            cache = {}
            url = upload_local_image(session, "http://test:8502", "secret", "bg.png", Path(tmp), cache)
            self.assertEqual(url, "http://test:8502/api/files/file123")
            self.assertEqual(session.upload_content, data)
            self.assertEqual(session.post_count, 1)
            url2 = upload_local_image(session, "http://test:8502", "secret", "bg.png", Path(tmp), cache)
            self.assertEqual(url2, url)
            self.assertEqual(session.post_count, 1)

    def test_remote_is_imported_not_passed_through(self):
        data = self.make_png()
        session = FakeSession(data)
        url = upload_local_image(session, "http://test:8502", "secret", "https://images.test/bg.png", Path("."), {})
        self.assertEqual(url, "http://test:8502/api/files/file123")
        self.assertIn("https://images.test/bg.png", session.get_urls)
        self.assertIn("http://test:8502/api/files/file123", session.get_urls)
        self.assertEqual(session.post_count, 1)

    def test_invalid_image_is_rejected(self):
        session = FakeSession(b"<html>Not an image</html>")
        with self.assertRaises(GeneratorError):
            upload_local_image(session, "http://test:8502", "secret", "https://images.test/bad.jpg", Path("."), {})
        self.assertEqual(session.post_count, 0)


if __name__ == "__main__":
    unittest.main()
