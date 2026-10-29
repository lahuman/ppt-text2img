import re
import unittest
from html.parser import HTMLParser
from pathlib import Path

import app


ROOT = Path(__file__).resolve().parent.parent


class PageReferences(HTMLParser):
    def __init__(self):
        super().__init__()
        self.ids = set()
        self.references = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if "id" in attrs:
            self.ids.add(attrs["id"])
        for attribute in ("href", "src"):
            if attrs.get(attribute):
                self.references.append(attrs[attribute])


class ReleaseTests(unittest.TestCase):
    def test_app_installer_and_site_versions_match(self):
        installer = (ROOT / "inno-setup.iss").read_text(encoding="utf-8")
        site = (ROOT / "docs/site.js").read_text(encoding="utf-8")
        installer_version = re.search(r'#define MyAppVersion "([0-9.]+)"', installer)
        site_version = re.search(r'version: "([0-9.]+)"', site)
        self.assertIsNotNone(installer_version)
        self.assertIsNotNone(site_version)
        self.assertEqual(installer_version.group(1), app.APP_VERSION)
        self.assertEqual(site_version.group(1), app.APP_VERSION)

    def test_landing_page_local_assets_and_anchors_exist(self):
        page = PageReferences()
        page.feed((ROOT / "docs/index.html").read_text(encoding="utf-8"))
        for reference in page.references:
            with self.subTest(reference=reference):
                if reference.startswith("#"):
                    self.assertIn(reference[1:], page.ids)
                elif "://" not in reference:
                    self.assertTrue((ROOT / "docs" / reference).is_file())


if __name__ == "__main__":
    unittest.main()
