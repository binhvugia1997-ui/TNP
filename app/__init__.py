"""Report Extractor - batch PPTX quality report -> Excel verification template.

Single canonical source of the application version / build number (never duplicate these strings).
``BUILD_NUMBER`` is the authoritative, numeric update ordering (PROMPT-005); it is displayed zero-padded
("Build 004").  A Git revision – when available – is internal diagnostic metadata only, never the build number.
"""

__version__ = "1.3.2"
BUILD_NUMBER = 15


def format_build(n) -> str:
    """Numeric build shown to users: 4 -> '004', 12 -> '012', 1234 -> '1234'."""
    return f"{int(n):03d}"


BUILD_ID = format_build(BUILD_NUMBER)                         # "004" – numeric only (no git hash, no prompt tag)
BUILD_LABEL = f"Build {BUILD_ID}"                             # "Build 004"
APP_NAME = "Report Extractor"
APP_TITLE = f"{APP_NAME} v{__version__}"                      # window title / about
VERSION_LABEL = f"{__version__} — {BUILD_LABEL}"              # "1.0.4 — Build 004"
VERSION_LINE = f"version={__version__} build={BUILD_ID}"      # diagnostics / log start-up line
