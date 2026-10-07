"""Small cooperative-cancellation primitives shared by the Python pipeline.

The token itself stays backend-owned. Callers pass a predicate into long stages;
checks raise this distinct exception so normal user cancellation never becomes a
processing failure or traceback.
"""
from __future__ import annotations

from typing import Callable, Optional


class CancellationRequested(Exception):
    """Raised only at an explicitly safe boundary after a user cancel request."""


def check_cancelled(should_cancel: Optional[Callable[[], bool]]) -> None:
    if should_cancel is not None and should_cancel():
        raise CancellationRequested("User requested cooperative cancellation")
