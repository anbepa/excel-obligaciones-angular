"""Error handling utilities for the Excel MCP server."""

import json


class ToolError(Exception):
    """Exception that automatically converts to JSON error response.

    Usage:
        raise ToolError("Something went wrong")
        # Returns: {"status": "error", "error": "Something went wrong"}

        raise ToolError("Invalid range", code="INVALID_RANGE")
        # Returns: {"status": "error", "error": "Invalid range", "code": "INVALID_RANGE"}

        ToolError("Formula error", cells=[...]).to_json()
        # Returns: {"status": "error", "error": "Formula error", "cells": [...]}
    """

    def __init__(self, message: str, code: str | None = None, **kwargs):
        self.message = message
        self.code = code
        self.extra = kwargs
        super().__init__(message)

    def to_json(self) -> str:
        """Convert to JSON error response."""
        result = {"status": "error", "error": self.message}
        if self.code:
            result["code"] = self.code
        result.update(self.extra)
        return json.dumps(result)


class ToolSuccess:
    """Helper for creating JSON success responses.

    Usage:
        return ToolSuccess().to_json()  # {"status": "success"}
        return ToolSuccess(ids=["a", "b"]).to_json()  # {"status": "success", "ids": ["a", "b"]}
        return ToolSuccess(message="Created 5 rows").to_json()
    """

    def __init__(self, data: dict | None = None, **kwargs):
        self.data = {"status": "success"}
        if data:
            self.data.update(data)
        self.data.update(kwargs)

    def to_json(self) -> str:
        """Convert to JSON success response."""
        return json.dumps(self.data)
