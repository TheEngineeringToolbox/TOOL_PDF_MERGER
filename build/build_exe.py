"""Build a checked, versioned Windows executable using the active Python."""

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VERSION_PATTERN = re.compile(r'^APP_VERSION = "(\d+\.\d+\.\d+)"$', re.MULTILINE)


def read_version(text):
    match = VERSION_PATTERN.search(text)
    if not match or any(int(part) > 65535 for part in match[1].split(".")):
        raise ValueError("APP_VERSION moet major.minor.patch zijn (elk 0 t/m 65535).")
    return match[1]


def bumped_version(version, kind):
    parts = [int(part) for part in version.split(".")]
    index = {"major": 0, "minor": 1, "patch": 2}[kind]
    parts[index] += 1
    parts[index + 1 :] = [0] * (2 - index)
    if max(parts) > 65535:
        raise ValueError("Versie overschrijdt de Windows-limiet van 65535.")
    return ".".join(map(str, parts))


def run(*args):
    subprocess.run(args, cwd=ROOT, check=True)


def git(*args):
    return subprocess.check_output(["git", *args], cwd=ROOT, text=True).strip()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bump", choices=["patch", "minor", "major"])
    args = parser.parse_args()
    app = ROOT / "app.py"
    source = app.read_text(encoding="utf-8")
    version = read_version(source)
    if args.bump:
        new_version = bumped_version(version, args.bump)
        app.write_text(
            VERSION_PATTERN.sub(f'APP_VERSION = "{new_version}"', source),
            encoding="utf-8",
        )
        print(
            f"Versie: {version} -> {new_version}. Commit deze wijziging voor een release."
        )
        return

    for command in [
        ("-m", "ruff", "format", "--check", "."),
        ("-m", "ruff", "check", "."),
        ("-m", "unittest", "discover", "-s", "tests", "-v"),
        (str(ROOT / "build" / "check_licenses.py"),),
    ]:
        run(sys.executable, *command)

    commit = git("rev-parse", "HEAD")
    branch = git("branch", "--show-current")
    dirty = bool(git("status", "--porcelain"))
    timestamp = datetime.now(timezone.utc)
    build_id = f"{version}-{commit[:8]}{'-dirty' if dirty else ''}-{timestamp:%Y%m%dT%H%M%S%fZ}"
    output = ROOT / "dist" / build_id
    work = ROOT / "temp" / build_id
    work.mkdir(parents=True, exist_ok=False)
    version_file = work / "version_info.txt"
    run(
        sys.executable,
        str(ROOT / "build" / "build_version.py"),
        str(app),
        str(version_file),
    )
    name = f"Rapportage Merger TOOL-{version}"
    run(
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name",
        name,
        "--distpath",
        str(output),
        "--workpath",
        str(work / "work"),
        "--specpath",
        str(work),
        "--version-file",
        str(version_file),
        "--hidden-import",
        "win32timezone",
        "--icon",
        str(ROOT / "Assets" / "Rapportage_PDF_Generator.ico"),
        str(app),
    )
    executable = output / f"{name}.exe"
    if not executable.is_file():
        raise RuntimeError(f"EXE ontbreekt: {executable}")
    metadata = {
        "version": version,
        "commit": commit,
        "branch": branch,
        "dirty": dirty,
        "built_at_utc": timestamp.isoformat(),
        "python": sys.version,
        "sha256": hashlib.sha256(executable.read_bytes()).hexdigest(),
    }
    (output / "build-info.json").write_text(
        json.dumps(metadata, indent=2), encoding="utf-8"
    )
    packages = subprocess.check_output(
        [sys.executable, "-m", "pip", "freeze"], text=True
    )
    (output / "dependencies.txt").write_text(packages, encoding="utf-8")
    for filename in ["LICENSE.txt", "THIRD_PARTY_NOTICES.md", "THIRD_PARTY_LICENSES.txt"]:
        source_file = ROOT / filename
        if not source_file.is_file():
            raise RuntimeError(f"Licentiebestand ontbreekt: {source_file}")
        shutil.copy2(source_file, output / filename)
    print(f"Build geslaagd: {executable}")


if __name__ == "__main__":
    main()
