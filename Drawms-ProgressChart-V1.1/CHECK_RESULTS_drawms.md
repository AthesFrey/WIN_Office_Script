# drawms verification

The standalone workbook is built with `uv sync --locked` and `uv run python scripts/build_drawms.py`. `uv run pytest -q` reports **28 passed** (the original 24 regressions plus four drawms checks). Static checks cover the single worksheet, literal text input cells, three button bindings, the three embedded VBA modules, DrawingML geometry, package relationships and release checksums. The three VBA modules pass the ANTLR VBA 7.1 parser, and the workbook passes the local ECMA-376 worksheet, DrawingML and style schemas. The release ZIP passes CRC/SHA-256 checks, and rebuilding from its bundled `source_project/` produces matching workbook parts.

The available environment has no desktop Excel. VBA parser and OOXML schema checks do not establish Excel compilation or interactive COM behavior. The optional Windows scenarios should be run from an Excel trusted location before release.

`tests/excel_drawms.ps1` parses successfully under PowerShell 7.4.6 but was not run because this Linux environment has no desktop Excel.
