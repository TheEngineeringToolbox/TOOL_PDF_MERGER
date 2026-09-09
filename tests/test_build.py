"""Version policy tests without invoking the packager."""

import importlib.util
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location(
    "build_exe", Path(__file__).resolve().parents[1] / "build" / "build_exe.py"
)
build_exe = importlib.util.module_from_spec(spec)
spec.loader.exec_module(build_exe)


class VersionTests(unittest.TestCase):
    def test_bump_resets_lower_components(self):
        for kind, expected in [
            ("patch", "1.2.4"),
            ("minor", "1.3.0"),
            ("major", "2.0.0"),
        ]:
            with self.subTest(kind=kind):
                self.assertEqual(build_exe.bumped_version("1.2.3", kind), expected)

    def test_reject_invalid_windows_versions(self):
        for version in ["1.2", "1.2.3.4", "1.2.65536", "1.2.3-dev"]:
            with self.subTest(version=version), self.assertRaises(ValueError):
                build_exe.read_version(f'APP_VERSION = "{version}"')
        with self.assertRaises(ValueError):
            build_exe.bumped_version("1.2.65535", "patch")

    def test_read_central_version(self):
        self.assertEqual(build_exe.read_version('APP_VERSION = "0.0.3"\n'), "0.0.3")
