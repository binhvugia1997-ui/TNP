"""Report Extractor - batch PPTX quality report -> Excel verification template.

Single canonical source of the application version / build identifier (never duplicate these strings).
"""

__version__ = "1.0.1"
BUILD_ID = "PROMPT-001"
APP_NAME = "Report Extractor"
APP_TITLE = f"{APP_NAME} v{__version__}"                 # window title / about
VERSION_LINE = f"version={__version__} prompt={BUILD_ID}"  # diagnostics / log start-up line
