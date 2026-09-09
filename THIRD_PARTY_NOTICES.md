# Third-party notices

TOOL PDF Merger uses third-party software. These components are **not** licensed under the TOOL Free Use and Redistribution License 1.0; their own license terms continue to apply.

The runtime dependencies declared in `requirements.txt` are currently:

| Component | Declared license | Notes |
| --- | --- | --- |
| pywin32 | Python Software Foundation License / mixed notices | pywin32 contains differently licensed files; the license files and per-file notices in the upstream distribution are authoritative. |
| pypdf | BSD-3-Clause | Preserve the applicable copyright and license notice when redistributing. |
| Pillow | MIT-CMU | Preserve the applicable copyright and license notice when redistributing. |
| lxml | BSD-3-Clause | Preserve the applicable copyright and license notice when redistributing. |
| tqdm | MPL-2.0 AND MIT | Preserve the applicable notices. MPL-2.0 requirements continue to apply to MPL-covered files. |

Development and packaging tools, including Ruff, PyInstaller and pip-licenses, are used to build or validate the project and may themselves have separate license terms. A packaged release must include the license report generated from the actual build environment.

## Release requirement

Before distributing a build:

1. Run the license compliance check.
2. Review the generated third-party license report for `UNKNOWN`, copyleft, source-availability, non-commercial, or otherwise unexpected terms.
3. Include `LICENSE.txt`, this notice file, and the generated third-party license report next to the distributed executable.
4. Preserve any license files or notices that a dependency requires to accompany redistribution.

The automated check is a compliance aid, not a substitute for reviewing the authoritative license text of each dependency.
