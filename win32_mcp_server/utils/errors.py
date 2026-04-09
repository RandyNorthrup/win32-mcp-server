"""
Error types and structured error responses.
"""

from typing import Any


class ToolError(Exception):
    """Raised when a tool encounters a known, actionable error.

    Attributes:
        suggestion: Optional guidance for the caller on how to fix the issue.
    """

    def __init__(self, message: str, suggestion: str | None = None):
        super().__init__(message)
        self.suggestion = suggestion

    def to_dict(self) -> dict[str, Any]:
        d = {"error": True, "message": str(self)}
        if self.suggestion:
            d["suggestion"] = self.suggestion
        return d
