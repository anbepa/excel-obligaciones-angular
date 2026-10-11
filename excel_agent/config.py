"""Configuration and logging utilities for Excel Agent evaluation."""

import dataclasses
import json
from dataclasses import asdict
from datetime import datetime
from typing import Any

from pydantic import BaseModel

SYSTEM_PROMPT = """You are Sylvian, an expert Excel spreadsheet assistant that reads and edits Excel files directly using provided tools.
RULES:
    1. ONLY update/format the cells required for the task; avoid re-formatting unless specified. 
        DO NOT move cells / add unnecessary cells unless required! We expect the cells to be in the same position as the input.
    2. Do not adjust cells for the purposes of testing, especially cells that are used in formulas for other cells.
    3. Do not exit early or ask user of input, just perform the task.
    4. Always check the output of tool calls to ensure the cells are updated correctly.
    5. After performing the task, ALWAYS EXAMINE YOUR OUTPUT AND ENSURE IT IS CORRECT BEFORE RETURNING IT.

<solution_persistence>
- Treat yourself as an autonomous senior pair-programmer: once the user gives a direction, proactively gather context, plan, implement, test, and refine without waiting for additional prompts at each step.
- Persist until the task is fully handled end-to-end within the current turn whenever feasible: do not stop at analysis or partial fixes; carry changes through implementation, verification, and a clear explanation of outcomes unless the user explicitly pauses or redirects you.
- Be extremely biased for action. If a user provides a directive that is somewhat ambiguous on intent, assume you should go ahead and make the change. If the user asks a question like "should we do x?" and your answer is "yes", you should also go ahead and perform the action. It's very bad to leave the user hanging and require them to follow up with a request to "please do it."
</solution_persistence>
"""

PROMPT_TEMPLATE = """Task: {instruction}
Input: {output_file}
The workbook is already loaded. Use the MCP tools to perform the task.
"""


class ExperimentConfig(BaseModel):
    """Configuration for the Excel agent evaluation experiment."""

    model: str
    system_prompt: str = SYSTEM_PROMPT
    prompt_template: str = PROMPT_TEMPLATE
    timeout: float = 120.0  # Overall timeout for entire agent run
    timeout_request: float = 60.0  # Timeout per model API request
    timeout_tool: float = 20.0  # Timeout per tool execution

    def format_prompt(self, instruction: str, output_file: str) -> str:
        """Format the prompt template with the given inputs."""
        return self.prompt_template.format(
            instruction=instruction, output_file=output_file
        )


class TraceCollector:
    """Collects and stores traces during evaluation runs."""

    def __init__(self):
        self._traces: dict[str, Any] = {}

    def reset(self) -> None:
        """Clear all stored traces."""
        self._traces = {}

    def add_trace(
        self,
        session_id: str,
        prompt: str,
        output: str,
        messages: list[dict[str, Any]],
        usage: dict[str, Any],
        timestamp: str | None,
    ) -> None:
        """Add a successful trace."""
        self._traces[session_id] = {
            "session_id": session_id,
            "prompt": prompt,
            "output": output,
            "messages": messages,
            "usage": usage,
            "timestamp": timestamp,
        }

    def add_error_trace(self, session_id: str, error: str) -> None:
        """Add an error trace."""
        self._traces[session_id] = {
            "session_id": session_id,
            "error": error,
            "timestamp": datetime.now().isoformat(),
        }

    @property
    def traces(self) -> dict[str, Any]:
        """Get all collected traces."""
        return self._traces


def _serialize_dataclass(obj: Any) -> Any:
    """Recursively serialize dataclasses and other objects to JSON-compatible format."""
    if dataclasses.is_dataclass(obj) and not isinstance(obj, type):
        return {k: _serialize_dataclass(v) for k, v in asdict(obj).items()}
    elif isinstance(obj, (list, tuple)):
        return [_serialize_dataclass(item) for item in obj]
    elif isinstance(obj, dict):
        return {k: _serialize_dataclass(v) for k, v in obj.items()}
    elif hasattr(obj, "isoformat"):  # datetime objects
        return obj.isoformat()
    elif hasattr(obj, "__dict__") and not callable(obj):
        return {
            k: _serialize_dataclass(v)
            for k, v in obj.__dict__.items()
            if not k.startswith("_")
        }
    else:
        try:
            json.dumps(obj)  # Test if JSON serializable
            return obj
        except (TypeError, ValueError):
            return str(obj)
