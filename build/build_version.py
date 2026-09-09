import re
import sys
from datetime import datetime
from pathlib import Path

APP_FILE = Path(sys.argv[1] if len(sys.argv) > 1 else "app.py")
OUTPUT_FILE = Path(sys.argv[2] if len(sys.argv) > 2 else "version_info.txt")

COPYRIGHT = f"© {datetime.now().year} TOOL Engineers B.V."

text = APP_FILE.read_text(encoding="utf-8")
match = re.search(r"^\s*APP_VERSION\s*=\s*[\"\']([^\"\']+)[\"\']", text, re.MULTILINE)

if not match:
    raise RuntimeError(f"APP_VERSION niet gevonden in {APP_FILE}")

version = match.group(1).strip()
parts = version.split(".")

if len(parts) != 3 or not all(part.isdigit() and int(part) <= 65535 for part in parts):
    raise RuntimeError(f"Ongeldige APP_VERSION: {version!r}")

parts = (parts + ["0", "0", "0", "0"])[:4]
v = ", ".join(parts)

version_info = f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({v}),
    prodvers=({v}),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [
          StringStruct('CompanyName', 'TOOL Engineers BV'),
          StringStruct('FileDescription', 'Rapportage PDF Generator'),
          StringStruct('FileVersion', '{version}'),
          StringStruct('ProductName', 'Rapportage PDF Generator'),
          StringStruct('ProductVersion', '{version}'),
          StringStruct('LegalCopyright', '{COPYRIGHT}'),
        ]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""

OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
OUTPUT_FILE.write_text(version_info, encoding="utf-8")
print(f"Windows versie-informatie aangemaakt: {version}")
