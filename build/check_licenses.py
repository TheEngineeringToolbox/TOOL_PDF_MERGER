"""Validate licenses of runtime packages and generate a redistribution report."""

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPORT_PATH = ROOT / "THIRD_PARTY_LICENSES.txt"
RUNTIME_PACKAGES = ["pywin32", "pypdf", "Pillow", "lxml", "tqdm", "colorama"]

# These are license families reviewed for the current runtime dependencies.
# Additions must be reviewed before extending this allow-list.
ALLOWED_LICENSE_MARKERS = (
    "BSD",
    "MIT",
    "MIT-CMU",
    "MPL-2.0",
    "Mozilla Public License 2.0",
    "PSF",
    "Python Software Foundation",
)
FORBIDDEN_MARKERS = (
    "UNKNOWN",
    "AGPL",
    "Affero",
    "Non-Commercial",
    "Noncommercial",
    "Commons Clause",
)


def load_license_data():
    command = [
        sys.executable,
        "-m",
        "piplicenses",
        "--format=json",
        "--with-authors",
        "--with-urls",
        "--packages",
        *RUNTIME_PACKAGES,
    ]
    result = subprocess.run(
        command,
        cwd=ROOT,
        check=True,
        capture_output=True,
        text=True,
    )
    return json.loads(result.stdout)


def validate(rows):
    found = {row["Name"].lower(): row for row in rows}
    required = {name.lower() for name in RUNTIME_PACKAGES if name.lower() != "colorama"}
    missing = sorted(required - found.keys())
    problems = []
    if missing:
        problems.append("Ontbrekende runtime-pakketten: " + ", ".join(missing))

    for row in rows:
        license_text = row.get("License", "UNKNOWN") or "UNKNOWN"
        if any(marker.lower() in license_text.lower() for marker in FORBIDDEN_MARKERS):
            problems.append(f"{row['Name']}: geblokkeerde licentie: {license_text}")
            continue
        if not any(marker.lower() in license_text.lower() for marker in ALLOWED_LICENSE_MARKERS):
            problems.append(f"{row['Name']}: niet beoordeelde licentie: {license_text}")

    if problems:
        raise SystemExit("Licentiecontrole mislukt:\n- " + "\n- ".join(problems))


def write_report(rows):
    lines = [
        "TOOL PDF Merger - Third-party license report",
        "Generated from the active build environment.",
        "",
    ]
    for row in sorted(rows, key=lambda item: item["Name"].lower()):
        lines.extend(
            [
                f"Package: {row['Name']}",
                f"Version: {row['Version']}",
                f"License: {row.get('License', 'UNKNOWN')}",
                f"Author: {row.get('Author', '')}",
                f"URL: {row.get('URL', '')}",
                "",
            ]
        )
    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")


def main():
    rows = load_license_data()
    validate(rows)
    write_report(rows)
    print(f"Licentiecontrole geslaagd: {REPORT_PATH}")


if __name__ == "__main__":
    main()
